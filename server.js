#!/usr/bin/env node
/**
 * 議事録AI（株式会社エフシー用）
 * 取締役会議事録形式
 */

import express from "express";
import multer from "multer";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";
import { execFile } from "child_process";
import { writeFileSync, readFileSync, unlinkSync, mkdirSync } from "fs";
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, LevelFormat, TabStopType, TabStopPosition,
} from "docx";

const __dirname = dirname(fileURLToPath(import.meta.url));
const TMP_DIR = resolve(__dirname, "tmp");
mkdirSync(TMP_DIR, { recursive: true });

const MODEL = "gemini-3-flash-preview";
const PORT = process.env.PORT || process.env.GIJIROKU_PORT || 3456;
const MAX_RETRIES = 5;
const RETRY_DELAY = 5000;

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(resolve(__dirname, "public")));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 200 * 1024 * 1024 },
});

const AUDIO_MIME = {
  mp3: "audio/mp3", wav: "audio/wav", m4a: "audio/mp4",
  mp4: "audio/mp4", ogg: "audio/ogg", webm: "audio/webm",
  flac: "audio/flac", aac: "audio/aac",
};

// ─── 音声圧縮（ffmpeg） ───
const COMPRESS_THRESHOLD = 5 * 1024 * 1024;

function compressAudio(buffer, originalName) {
  return new Promise((resolve, reject) => {
    const ts = Date.now();
    const inputPath = `${TMP_DIR}/input-${ts}`;
    const outputPath = `${TMP_DIR}/output-${ts}.mp3`;
    writeFileSync(inputPath, buffer);
    execFile("ffmpeg", [
      "-i", inputPath, "-ac", "1", "-ab", "32k", "-ar", "16000", "-y", outputPath,
    ], { timeout: 120_000 }, (err) => {
      try { unlinkSync(inputPath); } catch {}
      if (err) { try { unlinkSync(outputPath); } catch {}; reject(new Error(`ffmpeg圧縮エラー: ${err.message}`)); return; }
      const compressed = readFileSync(outputPath);
      try { unlinkSync(outputPath); } catch {}
      resolve(compressed);
    });
  });
}

function getApiKey(req) {
  const key = req.headers["x-api-key"] || req.body?.apiKey;
  if (!key) throw new Error("Gemini APIキーが設定されていません。画面上部でAPIキーを入力してください。");
  return key;
}

// ─── Gemini File API ───
async function uploadToGeminiFileAPI(apiKey, buffer, mimeType, displayName) {
  const startUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`;
  const startRes = await fetch(startUrl, {
    method: "POST",
    headers: {
      "X-Goog-Upload-Protocol": "resumable", "X-Goog-Upload-Command": "start",
      "X-Goog-Upload-Header-Content-Length": buffer.length.toString(),
      "X-Goog-Upload-Header-Content-Type": mimeType, "Content-Type": "application/json",
    },
    body: JSON.stringify({ file: { displayName } }),
  });
  if (!startRes.ok) throw new Error(`File API 開始エラー (${startRes.status}): ${await startRes.text()}`);
  const uploadUrl = startRes.headers.get("X-Goog-Upload-URL");
  if (!uploadUrl) throw new Error("File API: Upload URLが取得できませんでした");

  const uploadRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: { "Content-Length": buffer.length.toString(), "X-Goog-Upload-Offset": "0", "X-Goog-Upload-Command": "upload, finalize" },
    body: buffer,
  });
  if (!uploadRes.ok) throw new Error(`File API アップロードエラー (${uploadRes.status}): ${await uploadRes.text()}`);

  const uploadData = await uploadRes.json();
  const fileUri = uploadData.file?.uri;
  const fileName = uploadData.file?.name;
  if (!fileUri) throw new Error("File API: file URIが取得できませんでした");

  const checkUrl = `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${apiKey}`;
  for (let i = 0; i < 60; i++) {
    const checkRes = await fetch(checkUrl);
    const checkData = await checkRes.json();
    if (checkData.state === "ACTIVE") return fileUri;
    if (checkData.state === "FAILED") throw new Error("File API: ファイル処理に失敗しました");
    await new Promise((r) => setTimeout(r, 2000));
  }
  throw new Error("File API: ファイル処理がタイムアウトしました");
}

// ─── Gemini API リトライ ───
async function callGeminiWithRetry(apiKey, model, body, timeoutMs = 60_000) {
  let lastError;
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const res = await fetch(url, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body), signal: AbortSignal.timeout(timeoutMs),
      });
      if (res.ok) { const data = await res.json(); return data.candidates?.[0]?.content?.parts?.[0]?.text || ""; }
      const errText = await res.text();
      if ((res.status === 503 || res.status === 429) && attempt < MAX_RETRIES) {
        const wait = RETRY_DELAY * attempt;
        console.log(`  → ${model} ${res.status} (${attempt}/${MAX_RETRIES}), ${wait / 1000}秒後にリトライ...`);
        await new Promise((r) => setTimeout(r, wait)); continue;
      }
      throw new Error(`Gemini API エラー (${res.status}, ${model}): ${errText}`);
    } catch (err) {
      lastError = err;
      if (attempt < MAX_RETRIES) {
        const wait = RETRY_DELAY * attempt;
        console.log(`  → ${model} エラー: ${err.message} (${attempt}/${MAX_RETRIES}), ${wait / 1000}秒後にリトライ...`);
        await new Promise((r) => setTimeout(r, wait)); continue;
      }
      throw err;
    }
  }
  throw lastError;
}

async function callGeminiSmart(apiKey, body, timeoutMs = 60_000) {
  try {
    return await callGeminiWithRetry(apiKey, MODEL, body, timeoutMs);
  } catch (err) {
    if (err.message.includes("503") || err.message.includes("UNAVAILABLE"))
      throw new Error("AIサーバーが現在混み合っています。しばらく時間を置いてから再度お試しください。");
    if (err.message.includes("fetch failed") || err.message.includes("other side closed"))
      throw new Error("AIサーバーとの接続が切れました。しばらく時間を置いてから再度お試しください。");
    if (err.message.includes("timeout") || err.message.includes("aborted"))
      throw new Error("AIサーバーからの応答がありませんでした。しばらく時間を置いてから再度お試しください。");
    throw err;
  }
}

async function callGeminiText(apiKey, prompt) {
  return callGeminiSmart(apiKey, {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 65536 },
  });
}

async function callGeminiWithFileUri(apiKey, fileUri, mimeType, prompt) {
  return callGeminiSmart(apiKey, {
    contents: [{ parts: [
      { file_data: { mime_type: mimeType, file_uri: fileUri } },
      { text: prompt },
    ]}],
    generationConfig: { temperature: 0.1, maxOutputTokens: 65536 },
  }, 120_000);
}

// ─── APIキー検証 ───
app.post("/api/verify-key", async (req, res) => {
  try {
    const apiKey = req.body.apiKey;
    if (!apiKey) return res.status(400).json({ valid: false, error: "APIキーが空です" });
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
    const checkRes = await fetch(url);
    res.json(checkRes.ok ? { valid: true } : { valid: false, error: "無効なAPIキーです" });
  } catch (err) { res.json({ valid: false, error: err.message }); }
});

// ─── ステップ1: 文字起こし ───
app.post("/api/transcribe", upload.single("audio"), async (req, res) => {
  req.setTimeout(600_000);
  res.setTimeout(600_000);
  try {
    const apiKey = req.headers["x-api-key"];
    if (!apiKey) return res.status(400).json({ error: "APIキーが設定されていません" });
    if (!req.file) return res.status(400).json({ error: "音声ファイルが指定されていません" });

    const file = req.file;
    let audioBuffer = file.buffer;
    let mimeType = file.mimetype;
    if (mimeType === "audio/mpeg") mimeType = "audio/mp3";
    const ext = file.originalname?.split(".").pop()?.toLowerCase();
    if (ext && AUDIO_MIME[ext]) mimeType = AUDIO_MIME[ext];

    const sizeMB = (file.size / (1024 * 1024)).toFixed(1);
    console.log(`文字起こし開始: ${file.originalname} (${mimeType}, ${sizeMB}MB)`);

    if (file.size > COMPRESS_THRESHOLD) {
      console.log(`  → ${sizeMB}MB > 5MB — ffmpegで圧縮中...`);
      audioBuffer = await compressAudio(file.buffer, file.originalname);
      mimeType = "audio/mp3";
      console.log(`  → 圧縮完了: ${sizeMB}MB → ${(audioBuffer.length / (1024 * 1024)).toFixed(1)}MB`);
    }

    const fileUri = await uploadToGeminiFileAPI(apiKey, audioBuffer, mimeType, file.originalname);
    console.log(`  → アップロード完了, 文字起こし中...`);

    let transcript = "";
    for (let attempt = 1; attempt <= 3; attempt++) {
      transcript = await callGeminiWithFileUri(apiKey, fileUri, mimeType, TRANSCRIBE_PROMPT);
      if (transcript.trim().length > 10) break;
      console.log(`  → 空レスポンス (${attempt}/3), リトライ中...`);
      await new Promise((r) => setTimeout(r, 3000));
    }
    if (!transcript.trim()) throw new Error("文字起こし結果が空でした。再度お試しください。");

    console.log(`文字起こし完了: ${transcript.length}文字`);
    res.json({ transcript });
  } catch (err) {
    const detail = err.cause ? `${err.message} (${err.cause.message || err.cause})` : err.message;
    console.error("文字起こしエラー:", detail);
    res.status(500).json({ error: detail });
  }
});

// ─── ステップ2: 議事録生成（取締役会形式） ───
app.post("/api/generate", async (req, res) => {
  try {
    const apiKey = getApiKey(req);
    const { transcript, memo, meetingTitle, participants, date } = req.body;
    if (!transcript && !memo) return res.status(400).json({ error: "文字起こしテキストまたはメモを入力してください" });

    const prompt = buildMinutesPrompt({ transcript, memo, meetingTitle, participants, date });
    const result = await callGeminiText(apiKey, prompt);
    res.json({ minutes: result });
  } catch (err) {
    console.error("生成エラー:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Word出力（取締役会議事録形式・A3） ───
app.post("/api/export/docx", async (req, res) => {
  const { markdown, meetingTitle } = req.body;
  if (!markdown) return res.status(400).json({ error: "Markdownが空です" });

  try {
    const paragraphs = markdownToDocxParagraphs(markdown);
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Century", size: 22 }, // 11pt Century
            paragraph: { spacing: { line: 300, after: 0 } },
          },
        },
      },
      sections: [{
        properties: {
          page: {
            size: { width: 16838, height: 23811 }, // A3
            margin: {
              top: 1701, right: 1701, bottom: 1701, left: 1701, // 30mm全方向
            },
          },
        },
        children: paragraphs,
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = encodeURIComponent(meetingTitle || "取締役会議事録") + ".docx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("Word生成エラー:", err.message);
    res.status(500).json({ error: `Word生成エラー: ${err.message}` });
  }
});

// ─── Markdown → docx（取締役会形式） ───
function markdownToDocxParagraphs(md) {
  const lines = md.split("\n").reduce((acc, line) => {
    if (line.trim() === "" && acc.length > 0 && acc[acc.length - 1].trim() === "") return acc;
    acc.push(line);
    return acc;
  }, []);
  const paragraphs = [];

  for (const line of lines) {
    // # タイトル → 中央揃え・太字・大きめ
    if (line.startsWith("# ")) {
      paragraphs.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: line.slice(2), bold: true, size: 32, font: "MS Mincho" })],
      }));
    // ## 議案見出し → 太字
    } else if (line.startsWith("## ")) {
      paragraphs.push(new Paragraph({
        spacing: { before: 200, after: 60 },
        children: [new TextRun({ text: line.slice(3), bold: true, size: 24, font: "MS Mincho" })],
      }));
    // ### 小見出し
    } else if (line.startsWith("### ")) {
      paragraphs.push(new Paragraph({
        spacing: { before: 100, after: 0 },
        children: [new TextRun({ text: line.slice(4), bold: true, size: 22, font: "MS Mincho" })],
      }));
    // --- → スキップ
    } else if (line.startsWith("---")) {
      // スキップ
    // 空行 → スキップ
    } else if (line.trim() === "") {
      // スキップ
    // 署名行（インデント付き：「　　　　」で始まる行）
    } else if (line.startsWith("　　")) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, font: "MS Mincho", size: 22 })],
      }));
    // 通常段落
    } else {
      paragraphs.push(new Paragraph({
        indent: { firstLine: 420 }, // 1文字分字下げ
        children: parseInline(line),
      }));
    }
  }
  return paragraphs;
}

function parseInline(text) {
  const runs = [];
  const regex = /\*\*(.+?)\*\*/g;
  let lastIndex = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) runs.push(new TextRun({ text: text.slice(lastIndex, match.index), font: "MS Mincho", size: 22 }));
    runs.push(new TextRun({ text: match[1], bold: true, font: "MS Mincho", size: 22 }));
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) runs.push(new TextRun({ text: text.slice(lastIndex), font: "MS Mincho", size: 22 }));
  if (runs.length === 0) runs.push(new TextRun({ text, font: "MS Mincho", size: 22 }));
  return runs;
}

// ─── プロンプト ───
const TRANSCRIBE_PROMPT = `会議は日本語音声です。この音声ファイルの文字起こしをしてください。

## ルール
- 複数人が話しているので、話す人ごとに改行してください
- 話者が判別できる場合は「話者A:」「話者B:」のように区別してください
- 「えー」や「あー」など話す人特有の癖は文字にしないでください
- 改行、句読点を付けてもっと詳細に書いてください
- 発言者ごとに改行してください
- 聞き取れない箇所は「[聞き取り不可]」と記載してください`;

function buildMinutesPrompt({ transcript, memo, meetingTitle, participants, date }) {
  return `あなたは取締役会の議事録を作成する専門アシスタントです。以下の文字起こしテキストから、正式な取締役会議事録を作成してください。

## 会議情報
- 会議名: ${meetingTitle || "取締役会"}
- 日時: ${date || new Date().toLocaleDateString("ja-JP")}
- 参加者: ${participants || "（未設定）"}

## 文字起こしテキスト
${transcript || "（なし）"}

## 手書きメモ・補足
${memo || "（なし）"}

## 出力フォーマット
以下の取締役会議事録の形式でMarkdownを作成してください：

# 取　締　役　会　議　事　録

{日時（令和○年○月○日（○曜日）午前/午後○時より）}、{場所}において、取締役{○名}出席のもとに取締役会を開催し、{議長名}が議長となり次の議案につき慎重に協議した結果、全会一致をもって、下記のとおり可決確定したので、{終了時刻}散会した。

## 議案及び議決内容

## 第1号議案　{議案名}について

{議案の内容を段落で記述。発言者と発言内容を具体的に書く。}

## 第2号議案　{議案名}について

{同様に記述}

（議案の数は会議の内容に応じて増減）

---

以上の決議を明確にするため、この議事録を作り、出席取締役の全員がこれに記名押印する。

{日付（令和○年○月○日）}
　　　　{会社名}
　　　　　　　　　　出席取締役　　{名前}
　　　　　　　　　　出席取締役　　{名前}
　　　　　　　　　　出席取締役　　{名前}
（出席者全員を列挙）

---

### ルール
- 文字起こしテキストから日時・場所・出席者・議長を判断してください
- 議案番号（第1号議案、第2号議案...）は会議の議題に応じて自動で振ってください
- 各議案の内容は、誰がどのような発言をしたかを具体的に記載してください
- 冒頭の定型文（「○○において、取締役○名出席のもとに〜」）は必ず含めてください
- 末尾の署名欄は、出席者全員の名前を列挙してください
- 不明な点は「※要確認」と注記してください
- 格式のある正式なビジネス文書として記述してください
- **「承知いたしました」「以下に〜」などの前置き・挨拶文は一切不要です。いきなり「# 取　締　役　会　議　事　録」から始めてください**`;
}

// ─── サーバー起動 ───
const server = app.listen(PORT, () => {
  console.log(`\n🎙️  議事録AI（エフシー用）サーバー起動`);
  console.log(`   http://localhost:${PORT}\n`);
});

server.timeout = 600_000;
server.keepAliveTimeout = 600_000;
