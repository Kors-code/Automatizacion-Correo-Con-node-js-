require("dotenv").config();

const fs = require("fs");
const path = require("path");
const axios = require("axios");
const FormData = require("form-data");

async function uploadExcelToLaravel(localPath, storeId) {
  const form = new FormData();

  form.append(
    "file",
    fs.createReadStream(localPath),
    path.basename(localPath)
  );

  form.append("store_id", String(storeId));

  const response = await axios.post(
    process.env.LARAVEL_IMPORT_URL,
    form,
    {
      headers: {
        ...form.getHeaders(),
        "X-Automation-Token": process.env.IMPORT_AUTOMATION_TOKEN,
      },

      maxBodyLength: Infinity,
      maxContentLength: Infinity,

      timeout: 1000 * 60 * 10,
    }
  );

  return response.data;
}

module.exports = {
  uploadExcelToLaravel,
};