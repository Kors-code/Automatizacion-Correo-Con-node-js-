require("dotenv").config();
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

async function getSite() {
  const token = await getAccessToken();

  const url = `https://graph.microsoft.com/v1.0/sites/${process.env.DOMAIN}:/sites/${process.env.SITE_NAME}`;

  const res = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  console.log("🔥 SITE_ID:");
  console.log(res.data.id);
}

getSite().catch(err => {
  console.error("❌ ERROR:");
  console.error(err.response?.data || err.message);
});