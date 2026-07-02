require("dotenv").config();
const fs = require("fs");
const axios = require("axios");

const DEFAULT_ONEDRIVE_USER_ID = "8590d404-d339-4552-a498-ad0047d0243e";
const DEFAULT_UPLOAD_RETRIES = 5;
const DEFAULT_RETRY_DELAY_MS = 2500;

function encodeDrivePath(path) {
  return String(path || "")
    .split("/")
    .filter(Boolean)
    .map((part) => encodeURIComponent(part))
    .join("/");
}

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("client_id", process.env.CLIENT_ID);
  params.append("client_secret", process.env.CLIENT_SECRET);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("grant_type", "client_credentials");

  const res = await axios.post(url, params);
  return res.data.access_token;
}

async function ensureOneDriveFolder(token, userId, folderPath) {
  const parts = String(folderPath || "")
    .replace(/\\/g, "/")
    .split("/")
    .filter(Boolean);

  let parentPath = "";

  for (const folderName of parts) {
    const url = parentPath
      ? `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${encodeDrivePath(parentPath)}:/children`
      : `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/children`;

    try {
      await axios.post(
        url,
        {
          name: folderName,
          folder: {},
          "@microsoft.graph.conflictBehavior": "fail",
        },
        {
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
        }
      );
    } catch (error) {
      if (error.response?.status !== 409) {
        throw error;
      }
    }

    parentPath = parentPath ? `${parentPath}/${folderName}` : folderName;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRetryableOneDriveUploadError(error) {
  const status = error.response?.status;
  const code = error.response?.data?.error?.code;

  return (
    (status === 409 && code === "resourceModified") ||
    !status ||
    ["ECONNRESET", "ETIMEDOUT", "ECONNABORTED", "EPIPE"].includes(error.code) ||
    [423, 429, 500, 502, 503, 504].includes(status)
  );
}

async function putOneDriveContentWithRetry(url, fileBuffer, headers) {
  const maxAttempts = Number(
    process.env.ONEDRIVE_UPLOAD_RETRIES || DEFAULT_UPLOAD_RETRIES
  );
  const retryDelayMs = Number(
    process.env.ONEDRIVE_RETRY_DELAY_MS || DEFAULT_RETRY_DELAY_MS
  );

  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await axios.put(url, fileBuffer, { headers });
    } catch (error) {
      lastError = error;

      if (!isRetryableOneDriveUploadError(error) || attempt >= maxAttempts) {
        throw error;
      }

      console.log(
        `OneDrive retry ${attempt}/${maxAttempts - 1}: ${
          error.response?.status || error.code || error.message
        } ${error.response?.data?.error?.code || ""}`.trim()
      );

      if (retryDelayMs > 0) {
        await sleep(retryDelayMs * attempt);
      }
    }
  }

  throw lastError;
}

async function uploadToOneDrive(localPath, fileName, folderPath) {
  const token = await getAccessToken();
  const fileBuffer = fs.readFileSync(localPath);
  const userId = process.env.ONEDRIVE_USER_ID || DEFAULT_ONEDRIVE_USER_ID;

  const cleanFolderPath = String(folderPath || "")
    .replace(/\\/g, "/")
    .replace(/^\/+|\/+$/g, "");

  await ensureOneDriveFolder(token, userId, cleanFolderPath);

  const encodedPath = encodeDrivePath(`${cleanFolderPath}/${fileName}`);

  const url =
    `https://graph.microsoft.com/v1.0/users/${userId}/drive/root:/${encodedPath}:/content` +
    "?@microsoft.graph.conflictBehavior=replace";

  const res = await putOneDriveContentWithRetry(url, fileBuffer, {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/octet-stream",
  });

  return res.data;
}

module.exports = {
  uploadToOneDrive,
  isRetryableOneDriveUploadError,
  putOneDriveContentWithRetry,
};
