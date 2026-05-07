require("dotenv").config();
const fs = require("fs");
const axios = require("axios");

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

async function uploadToOneDrive(localPath, fileName, folderPath) {
  const token = await getAccessToken();
  const fileBuffer = fs.readFileSync(localPath);

  const url = `https://graph.microsoft.com/v1.0/users/8590d404-d339-4552-a498-ad0047d0243e/drive/root:/${folderPath}/${fileName}:/content`;

  const res = await axios.put(url, fileBuffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/octet-stream",
    },
  });

  return res.data;
}

module.exports = { uploadToOneDrive };