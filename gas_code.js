// ===================================================
// 安全書類管理システム - Google Apps Script
// スプレッドシートの「拡張機能」→「Apps Script」に貼り付けてください
// ===================================================

const SHEET_NAMES = {
  anken: '案件マスタ',
  worker: '作業員マスタ',
  nyutaijo: '入退場記録',
  nisshi: '安全日誌'
};

// ===================================================
// シート初期化（初回のみ実行）
// ===================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 案件マスタ
  let sh = ss.getSheetByName(SHEET_NAMES.anken) || ss.insertSheet(SHEET_NAMES.anken);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['案件ID','建物名','工事名称','現場住所','着工日','竣工日','発注会社','担当部署','担当者名','病院名','病院TEL','作業内容','登録日時']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,13).setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 作業員マスタ
  sh = ss.getSheetByName(SHEET_NAMES.worker) || ss.insertSheet(SHEET_NAMES.worker);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['作業員ID','案件ID','氏名','ふりがな','生年月日','年齢','血液型','現住所','所属会社','次数','職種','役割','雇用形態','雇入年月日','経験年数','血圧上','血圧下','血圧判定','健康診断日','健康診断種類','家族氏名','家族続柄','保険_健康','保険_年金','保険_雇用','保険_特別労災','中退共','建退共','資格','入場日','登録日時']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,31).setBackground('#057a55').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 入退場記録
  sh = ss.getSheetByName(SHEET_NAMES.nyutaijo) || ss.insertSheet(SHEET_NAMES.nyutaijo);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['記録ID','案件ID','作業員ID','氏名','所属会社','職種','種別','記録日','記録時刻','備考','登録日時']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,11).setBackground('#c81e1e').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 安全日誌
  sh = ss.getSheetByName(SHEET_NAMES.nisshi) || ss.insertSheet(SHEET_NAMES.nisshi);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['日誌ID','案件ID','作業日','天気','気温','作業責任者','入場者数','作業内容','KY安全対策','ヒヤリハット','チェックリスト_結果','担当者名','担当者確認日時','承認者名','承認日時','承認ステータス','登録日時']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,17).setBackground('#5145cd').setFontColor('#ffffff').setFontWeight('bold');
  }

  SpreadsheetApp.getUi().alert('✅ シートの初期化が完了しました！');
}

// ===================================================
// メニューに「初期化」ボタンを追加
// ===================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 安全書類システム')
    .addItem('シートを初期化する', 'initSheets')
    .addToUi();
}

// ===================================================
// WebアプリとしてのエンドポイントURL発行用
// ===================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type;
    let result;

    if (type === 'anken')      result = saveAnken(data);
    else if (type === 'worker')    result = saveWorker(data);
    else if (type === 'nyutaijo')  result = saveNyutaijo(data);
    else if (type === 'nisshi')    result = saveNisshi(data);
    else throw new Error('不明なtype: ' + type);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', id: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const type = e.parameter.type;
  try {
    let result;
    if (type === 'anken')       result = getAnken();
    else if (type === 'workers')    result = getWorkers(e.parameter.ankenId);
    else if (type === 'nyutaijo')   result = getNyutaijo(e.parameter.ankenId, e.parameter.date);
    else if (type === 'nisshi')     result = getNisshi(e.parameter.ankenId);
    else throw new Error('不明なtype: ' + type);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===================================================
// 案件マスタ 保存・取得
// ===================================================
function saveAnken(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.anken);
  const id = 'ANK-' + Date.now();
  const now = new Date().toLocaleString('ja-JP');
  sh.appendRow([
    id,
    data.tatemono || '',
    data.kojimei || '',
    data.jusho || '',
    data.chakko || '',
    data.shunko || '',
    data.hatchusha || '大和ライフネクスト株式会社',
    data.busho || '',
    data.tanto || '',
    data.byoin || '',
    data.byoinTel || '',
    data.sagyo || '',
    now
  ]);
  return id;
}

function getAnken() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.anken);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ===================================================
// 作業員マスタ 保存・取得
// ===================================================
function saveWorker(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.worker);
  const id = 'WRK-' + Date.now();
  const now = new Date().toLocaleString('ja-JP');
  sh.appendRow([
    id,
    data.ankenId || '',
    data.name || '',
    data.kana || '',
    data.birth || '',
    data.age || '',
    data.ketsueki || '',
    data.jusho || '',
    data.kaisha || '',
    data.jisu || '',
    data.shokushu || '',
    (data.yakuwari || []).join('・'),
    data.koyo || '',
    data.koyobi || '',
    data.keiken || '',
    data.bpH || '',
    data.bpL || '',
    data.bpResult || '',
    data.kenkoDate || '',
    data.kenkoType || '',
    data.kazokuName || '',
    data.kazokuZokugara || '',
    data.insKenko || '',
    data.insNenkin || '',
    data.insKoyo || '',
    data.insTokurosa || '',
    data.chutaiko || '',
    data.kentaiko || '',
    (data.shikaku || []).join('・'),
    data.nyujoDate || '',
    now
  ]);
  return id;
}

function getWorkers(ankenId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.worker);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => !ankenId || row[1] === ankenId)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

// ===================================================
// 入退場記録 保存・取得
// ===================================================
function saveNyutaijo(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.nyutaijo);
  const id = 'NYT-' + Date.now();
  const now = new Date().toLocaleString('ja-JP');
  const today = new Date().toLocaleDateString('ja-JP');
  sh.appendRow([
    id,
    data.ankenId || '',
    data.workerId || '',
    data.name || '',
    data.kaisha || '',
    data.shokushu || '',
    data.kind || '',
    data.date || today,
    data.time || '',
    data.biko || '',
    now
  ]);
  return id;
}

function getNyutaijo(ankenId, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.nyutaijo);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => (!ankenId || row[1] === ankenId) && (!date || row[7] === date))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

// ===================================================
// 安全日誌 保存・取得
// ===================================================
function saveNisshi(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.nisshi);
  const id = 'NSS-' + Date.now();
  const now = new Date().toLocaleString('ja-JP');
  sh.appendRow([
    id,
    data.ankenId || '',
    data.date || '',
    data.tenki || '',
    data.temp || '',
    data.sekinin || '',
    data.nyujoCount || '',
    data.sagyo || '',
    (data.kyItems || []).join('\n'),
    data.hiyari || '',
    JSON.stringify(data.checklist || {}),
    data.tantouName || '',
    data.tantouTime || '',
    data.approverName || '',
    data.approverTime || '',
    data.approvalStatus || '未承認',
    now
  ]);
  return id;
}

function getNisshi(ankenId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.nisshi);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => !ankenId || row[1] === ankenId)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}
