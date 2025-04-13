// --- 建立Express伺服器 ---
const express = require("express");
// 建立 Express 伺服器
const app = express();
const port = 3001;

app.listen(port, () => {
  console.log(`Express server is running on port ${port}`);
});

const { google } = require("googleapis");
const axios = require("axios");
const credentials = require("./service-account-key.json");

const SPREADSHEET_ID = "13x26zuu_uK1ELuT0mMSoLY5cM5N5gbDu44Z10yGqKVw";

async function getAccessToken() {
  const jwtClient = new google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    ["https://www.googleapis.com/auth/spreadsheets"]
  );

  await jwtClient.authorize();
  const newAccessToken = jwtClient.credentials.access_token;
  console.log("New Access Token:", newAccessToken);
  return newAccessToken;
}

/** 新增資料（不覆蓋原有資料） */
async function appendToSheet() {
  const accessToken = await getAccessToken();
  const range = "有誰報名!A1:C";
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;

  const data = {
    values: [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"],
    ],
  };

  try {
    const res = await axios.post(url, data, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    console.log("資料成功寫入:", res.data);
  } catch (err) {
    console.error("寫入失敗:", err.response?.data || err);
  }
}


/** 執行 */
async function run() {
  await appendToSheet();
  console.log("資料寫入完成");
}

run();