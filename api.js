// ===================================================
// api.js - Google Sheets連携の共通モジュール
// 全HTMLページから読み込んで使用する
// ===================================================

// ⚠️ ここにGoogle Apps ScriptのデプロイURLを貼り付けてください
const GAS_URL = 'https://script.google.com/macros/s/AKfycbyPixJZWnLQMbrlfm2uR-hIX1A8_lw6rvpl7nsBbvO2rPQD8P8BaSKFXUn4p8ERUvom/exec';

// ===================================================
// 共通送信関数
// ===================================================
async function postToSheets(type, data) {
  try {
    const res = await fetch(GAS_URL, {
      method: 'POST',
      body: JSON.stringify({ type, ...data })
    });
    const json = await res.json();
    if (json.status !== 'ok') throw new Error(json.message);
    return json.id;
  } catch(e) {
    console.error('Sheets送信エラー:', e);
    throw e;
  }
}

async function getFromSheets(type, params = {}) {
  try {
    const q = new URLSearchParams({ type, ...params }).toString();
    const res = await fetch(`${GAS_URL}?${q}`);
    const json = await res.json();
    if (json.status !== 'ok') throw new Error(json.message);
    return json.data;
  } catch(e) {
    console.error('Sheets取得エラー:', e);
    throw e;
  }
}

// ===================================================
// 各データの保存関数
// ===================================================

// 案件登録
async function saveAnken(formData) {
  return await postToSheets('anken', formData);
}

// 作業員登録（新規入場者）
async function saveWorker(formData) {
  return await postToSheets('worker', formData);
}

// 入退場記録
async function saveNyutaijo(formData) {
  // 「type」フィールドがGASのtype判定と競合するため「kind」に変換して送る
  const { type: kind, ...rest } = formData;
  return await postToSheets('nyutaijo', { ...rest, kind });
}

// 安全日誌
async function saveNisshi(formData) {
  return await postToSheets('nisshi', formData);
}

// ===================================================
// 各データの取得関数
// ===================================================

// 案件一覧取得
async function fetchAnkenList() {
  return await getFromSheets('anken');
}

// 作業員一覧取得（案件IDで絞り込み）
async function fetchWorkers(ankenId) {
  return await getFromSheets('workers', { ankenId });
}

// 入退場記録取得（案件ID・日付で絞り込み）
async function fetchNyutaijo(ankenId, date) {
  return await getFromSheets('nyutaijo', { ankenId, date });
}

// 安全日誌一覧取得
async function fetchNisshi(ankenId) {
  return await getFromSheets('nisshi', { ankenId });
}

// ===================================================
// ユーティリティ
// ===================================================

// 送信中インジケーター表示
function showLoading(btn, text = '送信中...') {
  btn.disabled = true;
  btn._origText = btn.textContent;
  btn.textContent = text;
}

// 送信中インジケーター解除
function hideLoading(btn) {
  btn.disabled = false;
  btn.textContent = btn._origText || '送信';
}

// 今日の日付文字列（YYYY-MM-DD）
function todayStr() {
  const n = new Date();
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')}`;
}

// 現在時刻文字列（HH:MM）
function nowTimeStr() {
  const n = new Date();
  return `${String(n.getHours()).padStart(2,'0')}:${String(n.getMinutes()).padStart(2,'0')}`;
}
