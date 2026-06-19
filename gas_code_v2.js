// ============================================================
// GX 浜松CK チェックシートアプリ - Google Apps Script
// スプレッドシートで項目管理版
// ============================================================

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var p = e.parameter || {};
  var body = e.postData ? JSON.parse(e.postData.contents) : {};
  var action = p.action || body.action || "";
  var result;
  try {
    switch (action) {
      case "getStaff": result = getStaff(); break;
      case "getItems": result = getItems(); break;
      case "submitCheck": result = submitCheck(body); break;
      case "getRecords": result = getRecords(p); break;
      case "deleteRecord": result = deleteRecord(body); break;
      default: result = { success: false, error: "不明なアクション: " + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- 社員一覧 ---
function getStaff() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("社員");
  if (!ws) return { success: true, staff: [] };
  var last = ws.getLastRow();
  if (last <= 1) return { success: true, staff: [] };
  var data = ws.getRange(2, 1, last - 1, 1).getValues();
  var staff = data.map(function(r){return r[0]}).filter(function(v){return v!==""});
  return { success: true, staff: staff };
}

// --- 項目取得 ---
function getItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};

  // 作業前
  result.before = readSheet_before(ss);
  // 残留塩素
  result.chlorine = readSheet_chlorine(ss);
  // 最終確認
  result.final = readSheet_final(ss);
  // 清掃
  result.clean = readSheet_clean(ss);

  return { success: true, items: result };
}

function readSheet_before(ss) {
  var ws = ss.getSheetByName("項目_作業前");
  if (!ws || ws.getLastRow() <= 1) return null;
  var data = ws.getRange(2, 1, ws.getLastRow()-1, 2).getValues();
  var items = [];
  data.forEach(function(r, i) {
    if (r[0]) items.push({ id:"b"+String(i+1).padStart(2,"0"), text:r[0], type:"check", note:r[1]||"" });
  });
  return { title:"🌅 作業前確認チェックシート", badge:"① 作業前確認", hasTantou:false, isChlorine:false, sections:[{label:null, items:items}] };
}

function readSheet_chlorine(ss) {
  var ws = ss.getSheetByName("項目_残留塩素");
  if (!ws || ws.getLastRow() <= 1) return null;
  var data = ws.getRange(2, 1, ws.getLastRow()-1, 3).getValues();
  var items = [];
  data.forEach(function(r, i) {
    if (!r[0]) return;
    var type = "check";
    if (r[1] === "選択") type = "select";
    else if (r[1] === "○×") type = "ox";
    items.push({ id:"cl"+String(i+1).padStart(2,"0"), text:r[0], type:type, note:r[2]||"" });
  });
  return { title:"💧 残留塩素確認（水質検査記録）", badge:"② 残留塩素確認", hasTantou:false, isChlorine:true, sections:[{label:null, items:items}] };
}

function readSheet_final(ss) {
  var ws = ss.getSheetByName("項目_最終確認");
  if (!ws || ws.getLastRow() <= 1) return null;
  var data = ws.getRange(2, 1, ws.getLastRow()-1, 4).getValues();
  var items = [];
  data.forEach(function(r, i) {
    if (!r[0]) return;
    var type = (r[1] === "数値") ? "number" : "check";
    items.push({ id:"f"+String(i+1).padStart(2,"0"), text:r[0], type:type, unit:r[2]||"", note:r[3]||"" });
  });
  return { title:"🌙 最終確認チェックシート", badge:"③ 最終確認", hasTantou:true, isChlorine:false, sections:[{label:null, items:items}] };
}

function readSheet_clean(ss) {
  var ws = ss.getSheetByName("項目_清掃");
  if (!ws || ws.getLastRow() <= 1) return null;
  var data = ws.getRange(2, 1, ws.getLastRow()-1, 3).getValues();
  var sections = [];
  var currentArea = null;
  var currentItems = [];
  var areaIndex = 0;
  var areaKeys = ["toilet","hallway","kitchen","processing","break","area6","area7","area8"];
  var itemCount = 0;

  data.forEach(function(r) {
    if (!r[1] && !r[0]) return;
    var area = r[0] || currentArea;
    if (area !== currentArea) {
      if (currentArea && currentItems.length > 0) {
        sections.push({ label:currentArea, area:areaKeys[areaIndex]||("area"+areaIndex), items:currentItems });
        areaIndex++;
      }
      currentArea = area;
      currentItems = [];
    }
    if (r[1]) {
      itemCount++;
      currentItems.push({ id:"c"+String(itemCount).padStart(2,"0"), text:r[1], type:"check", note:r[2]||"" });
    }
  });
  if (currentArea && currentItems.length > 0) {
    sections.push({ label:currentArea, area:areaKeys[areaIndex]||("area"+areaIndex), items:currentItems });
  }

  return { title:"🧹 清掃確認チェックシート", badge:"④ 清掃確認", hasTantou:true, isChlorine:false, sections:sections };
}

// --- チェック送信 ---
function submitCheck(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetType = body.sheetType;
  var sheetName = "記録_" + sheetType;

  var ws = ss.getSheetByName(sheetName);
  if (!ws) {
    ws = ss.insertSheet(sheetName);
    var headers = ["送信日時", "担当者", "シート種別"];
    var items = body.items || [];
    items.forEach(function(item) {
      if (item.type === "check") {
        if (sheetType !== "作業前") headers.push(item.label + "_担当");
        headers.push(item.label + "_完了");
      } else if (item.type === "number") {
        headers.push(item.label);
      } else if (item.type === "text") {
        headers.push(item.label);
      } else if (item.type === "select") {
        headers.push(item.label);
      } else if (item.type === "ox") {
        headers.push(item.label);
      }
    });
    headers.push("備考");
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    ws.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    ws.setFrozenRows(1);
  }

  var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  var row = [now, body.staffName || "", sheetType];

  var items = body.items || [];
  items.forEach(function(item) {
    if (item.type === "check") {
      if (sheetType !== "作業前") row.push(item.tantou || "");
      row.push(item.done ? "✓" : "");
    } else if (item.type === "number") {
      row.push(item.value || "");
    } else if (item.type === "text") {
      row.push(item.value || "");
    } else if (item.type === "select") {
      row.push(item.value ? item.value + " mg/L" : "");
    } else if (item.type === "ox") {
      row.push(item.ox || "");
    }
  });
  row.push(body.remarks || "");

  ws.appendRow(row);
  return { success: true, message: sheetType + " を記録しました" };
}

// --- 記録取得 ---
function getRecords(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetType = p.sheetType || "作業前";
  var ws = ss.getSheetByName("記録_" + sheetType);
  if (!ws || ws.getLastRow() <= 1) return { success: true, headers: [], records: [] };

  var lastRow = ws.getLastRow();
  var lastCol = ws.getLastColumn();
  var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  var allData = ws.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var dateFilter = p.date || "";
  var filtered = allData;
  if (dateFilter) {
    filtered = allData.filter(function(row) {
      var d = row[0] ? Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "yyyy-MM-dd") : "";
      return d === dateFilter;
    });
  }

  var records = filtered.slice(-50).reverse().map(function(row) {
    return { rowIndex: allData.indexOf(row) + 2, data: row };
  });

  return { success: true, headers: headers, records: records };
}

// --- 記録削除 ---
function deleteRecord(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("記録_" + (body.sheetType || "作業前"));
  if (!ws) return { success: false, error: "シートが見つかりません" };
  var rowIndex = body.rowIndex;
  if (rowIndex && rowIndex > 1) {
    ws.deleteRow(rowIndex);
    return { success: true, message: "削除しました" };
  }
  return { success: false, error: "行番号が不正です" };
}

// --- 初期セットアップ ---
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 社員シート
  if (!ss.getSheetByName("社員")) {
    var ws = ss.insertSheet("社員");
    ws.getRange("A1").setValue("社員名");
    ws.getRange(1,1,1,1).setFontWeight("bold");
    ws.getRange("A2").setValue("サンプル太郎");
  }

  // 項目_作業前
  if (!ss.getSheetByName("項目_作業前")) {
    var w1 = ss.insertSheet("項目_作業前");
    var h1 = [["項目名","備考"]];
    w1.getRange(1,1,1,2).setValues(h1).setFontWeight("bold");
    var d1 = [
      ["冷凍庫（雷神）通電チェック",""],["冷蔵庫（風神）通電チェック",""],["温度っちのランプは青色か",""],
      ["水たまり、雑草の確認",""],["外壁の破損、隙間はないか",""],["鳥の巣は作られていないか",""],
      ["廃棄物置き場は清潔か",""],["郵便BOX 確認",""],["健康チェックシートに未記載の人はいないか",""],
      ["マットに除菌液をかける（2枚）",""],["水質検査を行う（残留塩素チェック）",""],
      ["テント（しげお）内に虫の発生はないか",""],["薬品BOXの安全確認",""],
      ["休憩室のエアコンは消えているか",""],["休憩室の鍵は閉まっているか",""]
    ];
    w1.getRange(2,1,d1.length,2).setValues(d1);
  }

  // 項目_残留塩素
  if (!ss.getSheetByName("項目_残留塩素")) {
    var w2 = ss.insertSheet("項目_残留塩素");
    w2.getRange(1,1,1,3).setValues([["項目名","種別（選択/○×）","備考"]]).setFontWeight("bold");
    var d2 = [
      ["残留塩素濃度","選択","基準値: 0.1〜1.0 mg/L"],
      ["色","○×",""],["臭い","○×",""],["味","○×",""],["濁り","○×",""]
    ];
    w2.getRange(2,1,d2.length,3).setValues(d2);
  }

  // 項目_最終確認
  if (!ss.getSheetByName("項目_最終確認")) {
    var w3 = ss.insertSheet("項目_最終確認");
    w3.getRange(1,1,1,4).setValues([["項目名","種別（チェック/数値）","単位","備考"]]).setFontWeight("bold");
    var d3 = [
      ["ガス元栓OFF（スープケトル）","チェック","",""],["ガス＆電源OFF（撹拌機2号）","チェック","",""],
      ["ガス＆電源OFF（撹拌機3号）","チェック","",""],["換気口（シャラシャラ）閉鎖","チェック","",""],
      ["シーラー 電源OFF","チェック","",""],["全IHコンロ、電源OFF","チェック","",""],
      ["全換気扇、エアコン電源OFF","チェック","",""],["ティファール電源OFF","チェック","",""],
      ["ラジオOFF","チェック","",""],["まな板殺菌庫 ON、タイマー、スタート","チェック","",""],
      ["蛇口（ちゃんと閉めた？）チェック","チェック","",""],
      ["各シンク 洗剤、アルコール、ペーパー補充OK？","チェック","",""],
      ["冷凍庫（雷神）通電チェック","チェック","",""],["冷蔵庫（風神）通電チェック","チェック","",""],
      ["温度っちのランプは青色か","チェック","",""],["テントの入り口を締める","チェック","",""],
      ["コンテナ内薬品管理記録表の確認","チェック","",""],["コンテナストレージの施錠確認","チェック","",""],
      ["厨房内の洗剤は所定の位置においてあるか","チェック","",""],
      ["水質検査記録表の検証を行う（日付入力）","チェック","",""],
      ["家庭ごみ 明日収集日？","チェック","","該当曜日を確認"],["郵便BOX 確認","チェック","",""],
      ["湯沸かしポット（コード抜いた？お湯捨てた？）","チェック","",""],
      ["入口 シャッターを下ろす","チェック","",""],
      ["電気 使用量確認 100V","数値","kWh",""],["電気 使用量確認 200V","数値","kWh",""],
      ["水道 使用量確認","数値","㎥",""],["ガス 使用量確認","数値","㎥",""]
    ];
    w3.getRange(2,1,d3.length,4).setValues(d3);
  }

  // 項目_清掃
  if (!ss.getSheetByName("項目_清掃")) {
    var w4 = ss.insertSheet("項目_清掃");
    w4.getRange(1,1,1,3).setValues([["エリア","項目名","備考"]]).setFontWeight("bold");
    var d4 = [
      ["トイレ","便器（汚れ、詰まりはないか）",""],["トイレ","床（掃き掃除）",""],
      ["トイレ","トイレットペーパー（替えの補充）",""],
      ["廊下","床（水拭き）","週1回"],["廊下","マット（掃除機）6枚",""],
      ["廊下","棚（結露、水たまり、虫の発生防止）",""],["廊下","靴箱（掃き掃除、ホコリ取り）",""],
      ["厨房","シンク内（生ゴミが残っていないか）",""],["厨房","床（ゴミが落ちていないか）",""],
      ["厨房","ゴシゴシ（床水洗い）","週1回"],["厨房","掃除道具確認（トンボ3本、ブラシ3本）",""],
      ["厨房","シーラー（ダスター汚れ、電源OFF）",""],["厨房","ゴミ箱（ゴミ袋入れ替え）",""],
      ["厨房","ゴミ箱（水洗い）","週1回・木曜予定"],
      ["加工室","シンク内（生ゴミが残っていないか）",""],["加工室","床（ゴミが落ちていないか）",""],
      ["加工室","ゴシゴシ（床水洗い）","週1回"],["加工室","ゴミ箱（ゴミ袋入れ替え）",""],
      ["加工室","ゴミ箱（水洗い）","週1回・木曜予定"],["加工室","スライサー（破損確認）",""],
      ["休憩室","掃き掃除 or 掃除機",""]
    ];
    w4.getRange(2,1,d4.length,3).setValues(d4);
  }
}
