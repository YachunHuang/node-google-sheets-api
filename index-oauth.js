// --- 建立Express伺服器 ---
const express = require("express");
// 建立 Express 伺服器
const app = express();
const port = 3001;

app.listen(port, () => {
  console.log(`Express server is running on port ${port}`);
});

// --- 進行驗證並串接Google Sheets API ---
const fs = require("fs").promises;
const path = require("path");
const process = require("process");
const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");
// Google Sheets API 設定
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const TOKEN_PATH = path.join(process.cwd(), "token.json");
const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

// 已存在的試算表 ID
const SPREADSHEET_ID = "13x26zuu_uK1ELuT0mMSoLY5cM5N5gbDu44Z10yGqKVw";

/**
 * 檢查 TOKEN_PATH 是否存放 refresh token
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * 讀取 CREDENTIALS_PATH 中的 Credentials 進行驗證，產生 refresh token 並存入 TOKEN_PATH
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * 驗證：如果有 refresh token，則用 refresh token 驗證取得 access token；如果沒有則產生 refresh token
 */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

// --- 操作 Google Sheets API ---
/** 新增試算表 */
async function createSheet(auth, title) {
  try {
    const service = google.sheets({ version: "v4", auth });
    const resource = {
      properties: { title },
    };

    const spreadsheet = await service.spreadsheets.create({
      resource,
      fields: "spreadsheetId",
    });

    console.log(`新增的試算表 ID: ${spreadsheet.data.spreadsheetId}`);
    return spreadsheet.data.spreadsheetId;
  } catch (error) {
    console.error(error);
  }
}

/** 新增資料（不覆蓋原有資料） */
async function appendToSheet(auth, spreadsheetId) {
  const sheets = google.sheets({ version: "v4", auth });

  const values = [
    ["A1", "B1", "C1"],
    ["A2", "B2", "C2"],
  ];

  const resource = { values };

  try {
    const result = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Sheet1!A1", // 只指定 A 欄，Google Sheets 會自動往下追加
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS", // 插入新行
      resource,
    });
    console.log(`${result.data.updates.updatedCells} cells appended.`);
  } catch (error) {
    console.error(error);
  }
}

// --- 使用 refresh_token 來取得 access_token ---
// 設定 OAuth2 認證
const oauth2Client = new google.auth.OAuth2(
  "962386445344-o27is2g5che7sgis8qh7bjsa7gmrimpk.apps.googleusercontent.com",
  "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"
);

// 設定 refresh_token
oauth2Client.setCredentials({
  refresh_token:
    "1//0eQwW-n_oiyfJCgYIARAAGA4SNwF-L9IrpcZN-Xpz9A7DMQSzfk-yeIzXiOlZULlX3CoagWS8yUfwv-i_ubZQLfP8WyhKmcuQtM8",
});

// 使用 refresh_token 來取得新的 access_token
async function getNewAccessToken() {
  try {
    const response = await oauth2Client.refreshAccessToken();
    const newAccessToken = response.credentials.access_token;
    console.log("New Access Token:", newAccessToken);
  } catch (error) {
    console.error("發生錯誤:", error);
  }
}

// --- 透過 post url append row ---
const axios = require("axios");

/** 透過 post url append row */
async function appendToSheetByUrl() {
  const range = "Sheet1!A1:C";
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
  
  // 透過 refresh token 取得有時效性的 access token
  const accessToken =
    "ya29.a0AZYkNZijEmzpJ9vbusJWrn8vDzSBZ9bZyHUtj_UDGI0PpH_l_Cj3Iy0byV-6Q7gi3JiLkanuVnFY3vh7YIawDx7uPZMu-wbdHaDliIUgEBetamDjnO1-8DC1fxisu-8MVyV90OrytUKaaXoJ0sCE9BphnufaeuSetelDaldBaCgYKAVQSARESFQHGX2MilxZQNERBuODjmbYh5M4DTw0175";

  // 要新增的資料
  const data = {
    values: [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"],
    ],
  };

  axios.post(url, data, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    })
    .then((response) => {
      console.log("資料成功寫入:", response.data);
    })
    .catch((error) => {
      console.error("寫入失敗:", error.response.data);
    });
}

/** 執行 */
async function run() {
  try {
    const auth = await authorize();

    let sheetId = SPREADSHEET_ID;
    if (!sheetId) {
      // 如果沒有指定試算表 ID，則新增試算表
      sheetId = await createSheet(auth, 'mysheets');
    }

    // await appendToSheet(auth, sheetId);
    await getNewAccessToken();
    // await appendToSheetByUrl();
  } catch (err) {
    console.error(err);
  }
}

run();
