require("dotenv").config();

const fs = require("fs");
const path = require("path");
const axios = require("axios");
const FormData = require("form-data");

const DEFAULT_LARAVEL_RETRIES = 4;
const DEFAULT_LARAVEL_RETRY_DELAY_MS = 5000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRetryableLaravelError(error) {
  const status = error.response?.status;
  const code = error.code;

  return (
    !status ||
    ["ECONNRESET", "ETIMEDOUT", "ECONNABORTED", "EPIPE"].includes(code) ||
    [408, 409, 423, 429, 500, 502, 503, 504].includes(status)
  );
}

async function withLaravelRetry(label, operation) {
  const maxAttempts = Number(
    process.env.LARAVEL_UPLOAD_RETRIES || DEFAULT_LARAVEL_RETRIES
  );
  const retryDelayMs = Number(
    process.env.LARAVEL_RETRY_DELAY_MS || DEFAULT_LARAVEL_RETRY_DELAY_MS
  );

  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await operation(attempt);
    } catch (error) {
      lastError = error;

      if (!isRetryableLaravelError(error) || attempt >= maxAttempts) {
        throw error;
      }

      const reason = error.response?.status || error.code || error.message;
      console.log(
        `${label} retry ${attempt}/${maxAttempts - 1}: ${reason}`
      );

      if (retryDelayMs > 0) {
        await sleep(retryDelayMs * attempt);
      }
    }
  }

  throw lastError;
}

function buildSalesChunkUrl() {
  const baseUrl = process.env.LARAVEL_SALES_IMPORT_CHUNK_URL;

  if (baseUrl) {
    return baseUrl;
  }

  return process.env.LARAVEL_SALES_IMPORT_URL.replace(/\/+$/, "") + "/chunk";
}

function buildCatalogImportUrl() {
  const startUrl = process.env.LARAVEL_CATALOG_IMPORT_START_URL;

  if (startUrl) {
    return startUrl;
  }

  const baseUrl =
    process.env.LARAVEL_CATALOG_IMPORT_URL ||
    "http://127.0.0.1:8000/api/automation/import-catalog";

  const normalized = baseUrl.replace(/\/+$/, "");

  return normalized.endsWith("/start") ? normalized : normalized + "/start";
}

function buildCatalogChunkUrl() {
  const chunkUrl = process.env.LARAVEL_CATALOG_IMPORT_CHUNK_URL;

  if (chunkUrl) {
    return chunkUrl;
  }

  const baseUrl =
    process.env.LARAVEL_CATALOG_IMPORT_URL ||
    "http://127.0.0.1:8000/api/automation/import-catalog";

  const normalized = baseUrl.replace(/\/+$/, "");

  if (normalized.endsWith("/chunk")) {
    return normalized;
  }

  return normalized.endsWith("/start")
    ? normalized.replace(/\/start$/, "/chunk")
    : normalized + "/chunk";
}

function buildAutomationHeaders(extraHeaders = {}) {
  return {
    ...extraHeaders,
    "X-Automation-Token": process.env.IMPORT_AUTOMATION_TOKEN,
  };
}

async function uploadInventoryExcelToLaravel(localPath, storeId) {
  const response = await withLaravelRetry("Laravel inventario", async () => {
    const form = new FormData();

    form.append("file", fs.createReadStream(localPath), {
      filename: path.basename(localPath),
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    form.append("store_id", String(storeId));

    return axios.post(process.env.LARAVEL_IMPORT_URL, form, {
      headers: {
        ...form.getHeaders(),
        ...buildAutomationHeaders(),
      },
      maxBodyLength: Infinity,
      maxContentLength: Infinity,
      timeout: 1000 * 60 * 10,
    });
  });

  return response.data;
}

async function uploadSalesExcelToLaravel(localPath, storeId) {
  const response = await withLaravelRetry("Laravel ventas inicio", async () => {
    const form = new FormData();

    form.append("file", fs.createReadStream(localPath), {
      filename: path.basename(localPath),
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    form.append("store_id", String(Number(storeId || 0)));
    form.append("replace_existing", "1");

    const headers = form.getHeaders();

    return axios.post(process.env.LARAVEL_SALES_IMPORT_URL, form, {
      headers: {
        ...buildAutomationHeaders(headers),
      },

      maxBodyLength: Infinity,
      maxContentLength: Infinity,
      timeout: 1000 * 60 * 10,
    });
  });

  return processSalesChunks(response.data);
}

async function uploadCatalogExcelToLaravel(localPath) {
  const response = await withLaravelRetry("Laravel catalogo inicio", async () => {
    const form = new FormData();

    form.append("file", fs.createReadStream(localPath), {
      filename: path.basename(localPath),
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    return axios.post(buildCatalogImportUrl(), form, {
      headers: {
        ...form.getHeaders(),
        ...buildAutomationHeaders(),
      },
      maxBodyLength: Infinity,
      maxContentLength: Infinity,
      timeout: 1000 * 60 * 10,
    });
  });

  return processCatalogChunks(response.data);
}

async function processSalesChunks(initialState) {
  if (!initialState || initialState.done || !initialState.batch_id) {
    return initialState;
  }

  const chunkUrl = buildSalesChunkUrl();
  let state = initialState;
  let chunkNumber = 0;

  while (!state.done) {
    chunkNumber += 1;

    console.log(
      `Procesando chunk ventas #${chunkNumber}: fila ${state.next_row} de ${state.total_rows}`
    );

    const response = await withLaravelRetry(
      `Laravel ventas chunk #${chunkNumber}`,
      async () =>
        axios.post(
          chunkUrl,
          {
            path: state.path,
            batch_id: state.batch_id,
            store_id: state.store_id || 0,
            next_row: state.next_row,
            total_rows: state.total_rows,
            chunk_size: state.chunk_size || 100,
          },
          {
            headers: buildAutomationHeaders({
              Accept: "application/json",
            }),
            maxBodyLength: Infinity,
            maxContentLength: Infinity,
            timeout: 1000 * 60 * 3,
          }
        )
    );

    state = {
      ...state,
      ...response.data,
    };

    console.log(
      `Chunk ventas OK: procesadas ${response.data.processed_rows || 0}, siguiente ${state.next_row}, done=${state.done}`
    );
  }

  return state;
}

async function processCatalogChunks(initialState) {
  if (!initialState || initialState.done || !initialState.path) {
    return initialState;
  }

  const chunkUrl = buildCatalogChunkUrl();
  const chunkSize = Number(
    process.env.LARAVEL_CATALOG_IMPORT_CHUNK_SIZE ||
      initialState.chunk_size ||
      100
  );
  let state = initialState;
  let chunkNumber = 0;
  let totalProcessedRows = 0;
  const summaries = [];

  while (!state.done) {
    chunkNumber += 1;

    console.log(
      `Procesando chunk catalogo #${chunkNumber}: fila ${state.next_row} de ${state.total_rows}`
    );

    const response = await withLaravelRetry(
      `Laravel catalogo chunk #${chunkNumber}`,
      async () =>
        axios.post(
          chunkUrl,
          {
            path: state.path,
            next_row: state.next_row,
            total_rows: state.total_rows,
            chunk_size: chunkSize,
          },
          {
            headers: buildAutomationHeaders({
              Accept: "application/json",
            }),
            maxBodyLength: Infinity,
            maxContentLength: Infinity,
            timeout: 1000 * 60 * 3,
          }
        )
    );

    state = {
      ...state,
      ...response.data,
    };
    totalProcessedRows += Number(response.data.processed_rows || 0);
    summaries.push(response.data.summary || {});

    console.log(
      `Chunk catalogo OK: procesadas ${response.data.processed_rows || 0}, siguiente ${state.next_row}, done=${state.done}`
    );
  }

  return {
    ...state,
    total_processed_rows: totalProcessedRows,
    chunk_count: chunkNumber,
    summaries,
  };
}
module.exports = {
  uploadInventoryExcelToLaravel,
  uploadSalesExcelToLaravel,
  uploadCatalogExcelToLaravel,
  isRetryableLaravelError,
  withLaravelRetry,
};
