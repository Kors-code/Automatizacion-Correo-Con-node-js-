const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

const CREDENTIALS_PATH = path.join(__dirname, "credentials.json");
const TOKEN_PATH = path.join(__dirname, "token.json");

const SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];
const allowedSender = "reportes@skyfreeshop.org";

// Reglas de clasificación
const FILTER_RULES = [
  {
    name: "reporte-75-dfp",
    matchText: "reporte ventas 75 dfp",
    folder: "dfp",
    type: "DFP",
  },
  {
    name: "reporte-75-ldc",
    matchText: "reporte ventas 75 ldc",
    folder: "ldc",
    type: "LDC",
  },
];

function loadToken() {
  if (!fs.existsSync(TOKEN_PATH)) {
    throw new Error("No existe token.json. Debes autorizar primero la app.");
  }

  return JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"));
}

function loadCredentials() {
  const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, "utf8"));
  const data = credentials.installed || credentials.web;

  if (!data) {
    throw new Error("credentials.json no tiene la estructura esperada.");
  }

  return data;
}

function getHeader(headers, name) {
  const found = headers?.find(
    (h) => h.name?.toLowerCase() === name.toLowerCase()
  );
  return found ? found.value : "";
}

function sanitizeFilename(name) {
  return String(name || "")
    .replace(/[<>:"/\\|?*]+/g, "_")
    .trim();
}

function collectAttachments(payload, result = []) {
  if (!payload) return result;

  if (payload.filename && payload.body && payload.body.attachmentId) {
    result.push({
      filename: payload.filename,
      attachmentId: payload.body.attachmentId,
      mimeType: payload.mimeType,
    });
  }

  if (Array.isArray(payload.parts)) {
    for (const part of payload.parts) {
      collectAttachments(part, result);
    }
  }

  return result;
}

function extractTextFromPayload(payload) {
  if (!payload) return "";

  let text = "";

  if (payload.body && payload.body.data) {
    const base64 = payload.body.data.replace(/-/g, "+").replace(/_/g, "/");
    try {
      text += Buffer.from(base64, "base64").toString("utf8") + " ";
    } catch {
      // ignorar
    }
  }

  if (Array.isArray(payload.parts)) {
    for (const part of payload.parts) {
      const mimeType = (part.mimeType || "").toLowerCase();

      if (mimeType === "text/plain" || mimeType === "text/html") {
        if (part.body && part.body.data) {
          const base64 = part.body.data.replace(/-/g, "+").replace(/_/g, "/");
          try {
            text += Buffer.from(base64, "base64").toString("utf8") + " ";
          } catch {
            // ignorar
          }
        }
      }

      text += extractTextFromPayload(part) + " ";
    }
  }

  return text.trim();
}

function getMatchedRule(subject, bodyText) {
  const haystack = `${subject} ${bodyText}`.toLowerCase();

  for (const rule of FILTER_RULES) {
    if (haystack.includes(rule.matchText.toLowerCase())) {
      return rule;
    }
  }

  return null;
}

async function authorize() {
  const { client_secret, client_id, redirect_uris } = loadCredentials();

  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );

  const token = loadToken();
  oAuth2Client.setCredentials(token);

  return oAuth2Client;
}

async function getRecentZipMessage(auth, processedIds = []) {
  const gmail = google.gmail({ version: "v1", auth });

  const listRes = await gmail.users.messages.list({
    userId: "me",
    maxResults: 20,
    q: `from:${allowedSender} has:attachment`,
  });

  const messages = listRes.data.messages || [];
  if (messages.length === 0) {
    return null;
  }

  for (const msg of messages) {
    if (processedIds.includes(msg.id)) {
      continue;
    }

    const fullMsg = await gmail.users.messages.get({
      userId: "me",
      id: msg.id,
      format: "full",
    });

    const payload = fullMsg.data.payload;
    const headers = payload?.headers || [];

    const subject = getHeader(headers, "Subject");
    const from = getHeader(headers, "From");
    const date = getHeader(headers, "Date");
    const snippet = fullMsg.data.snippet || "";

    if (!from.toLowerCase().includes(allowedSender.toLowerCase())) {
      continue;
    }

    const bodyText = extractTextFromPayload(payload) || "";
    const combinedText = `${subject} ${snippet} ${bodyText}`;

    const matchedRule = getMatchedRule(subject, combinedText);
    if (!matchedRule) {
      continue;
    }

    const attachments = collectAttachments(payload);

    const zipAttachment = attachments.find((a) =>
      a.filename?.toLowerCase().endsWith(".zip")
    );

    if (!zipAttachment) {
      continue;
    }

    return {
      gmail,
      messageId: msg.id,
      threadId: msg.threadId,
      subject,
      from,
      date,
      bodyText: combinedText,
      rule: matchedRule,
      attachment: zipAttachment,
    };
  }

  return null;
}

async function downloadAttachment(gmail, messageId, attachmentId, outputPath) {
  const attRes = await gmail.users.messages.attachments.get({
    userId: "me",
    messageId,
    id: attachmentId,
  });

  const base64Url = attRes.data.data;
  const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
  const buffer = Buffer.from(base64, "base64");

  fs.writeFileSync(outputPath, buffer);
  return outputPath;
}

async function downloadLatestZip(processedIds = []) {
  const auth = await authorize();
  const found = await getRecentZipMessage(auth, processedIds);

  if (!found) {
    console.log("No se encontró ningún ZIP que coincida con las reglas.");
    return null;
  }

  const storageDir = path.join(
    __dirname,
    "storage",
    "incoming-zips",
    found.rule.folder
  );
  fs.mkdirSync(storageDir, { recursive: true });

  const safeSubject = sanitizeFilename(found.subject || "sin-asunto");
  const safeAttachment = sanitizeFilename(
    found.attachment.filename || "adjunto.zip"
  );

  const zipName = `${Date.now()}-${found.rule.name}-${safeSubject}-${safeAttachment}`.replace(
    /\s+/g,
    "_"
  );

  const zipPath = path.join(storageDir, zipName);

  await downloadAttachment(
    found.gmail,
    found.messageId,
    found.attachment.attachmentId,
    zipPath
  );

  return {
    ...found,
    zipPath,
  };
}

module.exports = {
  downloadLatestZip,
};