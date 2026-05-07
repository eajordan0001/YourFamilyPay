// ═══════════════════════════════════════════════════════════
//  辦公室代買記帳系統 — Google Apps Script 後端
//  版本：2.1  |  成員：15人版
// ═══════════════════════════════════════════════════════════

// ── 設定區（部署後請修改） ──────────────────────────────────
const CONFIG = {
  TOKEN_ADMIN : "admin2024",      // A同事（記帳）用的 token
  TOKEN_VIEW  : "view2024",       // 其他人（查詢）用的 token
  LOW_BALANCE : 50,               // 低餘額警示下限（元）
  SHEET_ID    : "",               // 留空 = 使用當前試算表；或填入指定 ID
};

// ── 成員設定（共 15 人）────────────────────────────────────
const MEMBERS = [
  { id:"YF001", name:"成員01", emoji:"👦" },
  { id:"YF002", name:"成員02", emoji:"👧" },
  { id:"YF003", name:"成員03", emoji:"🧑" },
  { id:"YF004", name:"成員04", emoji:"👨" },
  { id:"YF005", name:"成員05", emoji:"👩" },
  { id:"YF006", name:"成員06", emoji:"🧔" },
  { id:"YF007", name:"成員07", emoji:"👱‍♀️" },
  { id:"YF008", name:"成員08", emoji:"💪" },
  { id:"YF009", name:"成員09", emoji:"🤓" },
  { id:"YF010", name:"成員10", emoji:"🤖" },
  { id:"YF011", name:"成員11", emoji:"🧑‍💼" },
  { id:"YF012", name:"成員12", emoji:"👩‍💼" },
  { id:"YF013", name:"成員13", emoji:"🧑‍🎓" },
  { id:"YF014", name:"成員14", emoji:"👨‍🎓" },
  { id:"YF015", name:"成員15", emoji:"👩‍🎓" },
];

// ── 分頁名稱常數 ─────────────────────────────────────────────
const SHEET_MEMBERS = "成員總覽";
const SHEET_LOG     = "Log";

// ══════════════════════════════════════════════════════════════
//  進入點
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  const p      = e.parameter;
  const token  = p.token  || "";
  const action = p.action || "";

  const isAdmin = token === CONFIG.TOKEN_ADMIN;
  const isView  = token === CONFIG.TOKEN_VIEW || isAdmin;

  if (!isView) return json({ status:"error", message:"權限不足或 token 錯誤" });

  try {
    switch (action) {
      case "members"     : return json(getMembers());
      case "mydata"      : return json(getMyData(p.id));
      case "overview"    : if (!isAdmin) return json({ status:"error", message:"需要管理員權限" });
                           return json(getOverview());
      case "records"     : if (!isAdmin) return json({ status:"error", message:"需要管理員權限" });
                           return json(getRecords(p.startDate, p.endDate, p.memberId));
      case "addTx"       : if (!isAdmin) return json({ status:"error", message:"需要管理員權限" });
                           return json(addTransaction(p));
      case "topup"       : if (!isAdmin) return json({ status:"error", message:"需要管理員權限" });
                           return json(topUp(p.id, Number(p.amount), p.note || ""));
      case "initSheets"  : if (!isAdmin) return json({ status:"error", message:"需要管理員權限" });
                           return json(initSheets());
      case "addNewMember": if (!isAdmin) return json({ status:"error", message:"需要管理員權限" });
                           return json(addNewMember());
      default            : return json({ status:"error", message:"未知的 action: " + action });
    }
  } catch(err) {
    return json({ status:"error", message:err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════
//  取得試算表
// ══════════════════════════════════════════════════════════════
function getSpreadsheet() {
  return CONFIG.SHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

// ══════════════════════════════════════════════════════════════
//  初始化試算表（全新部署時使用）
// ══════════════════════════════════════════════════════════════
function initSheets() {
  const ss = getSpreadsheet();

  // ── 建立「成員總覽」分頁 ──
  let s1 = ss.getSheetByName(SHEET_MEMBERS);
  if (!s1) {
    s1 = ss.insertSheet(SHEET_MEMBERS, 0);
  } else {
    s1.clearContents();
  }
  const memberHeaders = ["編號","姓名","Emoji","目前餘額","累積儲值","最後更新時間"];
  s1.getRange(1,1,1,memberHeaders.length).setValues([memberHeaders]);
  styleHeader(s1, 1, memberHeaders.length);

  MEMBERS.forEach((m, i) => {
    s1.getRange(i+2, 1, 1, 6).setValues([[m.id, m.name, m.emoji, 0, 0, now()]]);
  });
  centerAll(s1);
  s1.setFrozenRows(1);

  // ── 建立「Log」分頁 ──
  let s2 = ss.getSheetByName(SHEET_LOG);
  if (!s2) {
    s2 = ss.insertSheet(SHEET_LOG, 1);
  } else {
    s2.clearContents();
  }
  const logHeaders = ["交易序號(TID)","交易時間","成員編號","姓名","交易類型","交易項目","備註","原餘額","變動金額","新餘額","回溯交易日期"];
  s2.getRange(1,1,1,logHeaders.length).setValues([logHeaders]);
  styleHeader(s2, 1, logHeaders.length);
  centerAll(s2);
  s2.setFrozenRows(1);

  // ── 建立個人分頁（YF001–YF015）──
  const personalHeaders = ["交易序號(TID)","交易時間","交易類型","交易項目","備註","原餘額","變動金額","新餘額","回溯交易日期"];
  MEMBERS.forEach((m, i) => {
    let ps = ss.getSheetByName(m.id);
    if (!ps) {
      ps = ss.insertSheet(m.id, i + 2);
    } else {
      ps.clearContents();
    }
    ps.getRange(1,1,1,personalHeaders.length).setValues([personalHeaders]);
    styleHeader(ps, 1, personalHeaders.length);
    centerAll(ps);
    ps.setFrozenRows(1);
  });

  return { status:"success", message:"試算表初始化完成！共 " + MEMBERS.length + " 位成員" };
}

// ══════════════════════════════════════════════════════════════
//  安全新增成員（不影響現有資料）
// ══════════════════════════════════════════════════════════════
function addNewMember() {
  const ss          = getSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!memberSheet) return { status:"error", message:"請先執行 initSheets 初始化" };

  const existingIds = memberSheet.getDataRange().getValues()
    .slice(1).map(r => r[0]);

  const personalHeaders = ["交易序號(TID)","交易時間","交易類型","交易項目","備註","原餘額","變動金額","新餘額","回溯交易日期"];
  const added = [];

  MEMBERS.forEach((m, i) => {
    if (!existingIds.includes(m.id)) {
      memberSheet.appendRow([m.id, m.name, m.emoji, 0, 0, now()]);
      const lastRow = memberSheet.getLastRow();
      memberSheet.getRange(lastRow, 1, 1, 6)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      added.push(m.id);
    }
    if (!ss.getSheetByName(m.id)) {
      const ps = ss.insertSheet(m.id, i + 2);
      ps.getRange(1,1,1,personalHeaders.length).setValues([personalHeaders]);
      styleHeader(ps, 1, personalHeaders.length);
      centerAll(ps);
      ps.setFrozenRows(1);
    }
  });

  return {
    status  : "success",
    message : added.length > 0
      ? "已新增成員分頁：" + added.join(", ")
      : "所有成員分頁皆已存在，無需新增",
    added,
  };
}

// ══════════════════════════════════════════════════════════════
//  取得所有成員（餘額）
// ══════════════════════════════════════════════════════════════
function getMembers() {
  const s = getSheet(SHEET_MEMBERS);
  if (!s) return { status:"error", message:"請先執行 initSheets 初始化" };
  const rows = s.getDataRange().getValues().slice(1);
  const result = rows.map(r => ({
    id         : r[0],
    name       : r[1],
    emoji      : r[2],
    balance    : Number(r[3]),
    topupTotal : Number(r[4]),
    isLow      : Number(r[3]) < CONFIG.LOW_BALANCE,
    updated    : r[5],
  }));
  return { status:"success", members: result, lowLimit: CONFIG.LOW_BALANCE };
}

// ══════════════════════════════════════════════════════════════
//  取得個人資料（查詢用）
// ══════════════════════════════════════════════════════════════
function getMyData(memberId) {
  if (!memberId) return { status:"error", message:"缺少 id 參數" };
  const s = getSheet(memberId);
  if (!s) return { status:"error", message:"找不到此成員分頁：" + memberId };

  const memberInfo = getMemberInfo(memberId);
  const rows       = s.getDataRange().getValues().slice(1);
  const records    = rows.reverse().map(r => ({
    tid    : r[0],
    time   : r[1],
    type   : r[2],
    item   : r[3],
    note   : r[4],
    before : Number(r[5]),
    change : Number(r[6]),
    after  : Number(r[7]),
    txDate : r[8] || "",
  }));
  return { status:"success", member: memberInfo, records };
}

// ══════════════════════════════════════════════════════════════
//  總覽（管理員用）
// ══════════════════════════════════════════════════════════════
function getOverview() {
  const s    = getSheet(SHEET_MEMBERS);
  const rows = s.getDataRange().getValues().slice(1);
  const members = rows.map(r => ({
    id         : r[0],
    name       : r[1],
    emoji      : r[2],
    balance    : Number(r[3]),
    topupTotal : Number(r[4]),
    isLow      : Number(r[3]) < CONFIG.LOW_BALANCE,
  }));
  return { status:"success", members, lowLimit: CONFIG.LOW_BALANCE };
}

// ══════════════════════════════════════════════════════════════
//  取得交易紀錄（管理員用，支援篩選）
// ══════════════════════════════════════════════════════════════
function getRecords(startDate, endDate, memberId) {
  const s    = getSheet(SHEET_LOG);
  const rows = s.getDataRange().getValues().slice(1);
  let records = rows.map(r => ({
    tid    : r[0],
    time   : r[1],
    id     : r[2],
    name   : r[3],
    type   : r[4],
    item   : r[5],
    note   : r[6],
    before : Number(r[7]),
    change : Number(r[8]),
    after  : Number(r[9]),
    txDate : r[10] || "",
  })).filter(r => r.tid !== "");

  if (memberId)   records = records.filter(r => r.id === memberId);
  if (startDate) {
    const sd = new Date(startDate);
    records = records.filter(r => new Date(r.time) >= sd);
  }
  if (endDate) {
    const ed = new Date(endDate);
    ed.setHours(23,59,59);
    records = records.filter(r => new Date(r.time) <= ed);
  }

  records.reverse();
  return { status:"success", records };
}

// ══════════════════════════════════════════════════════════════
//  新增交易（扣款，支援多品項＋回溯日期）
// ══════════════════════════════════════════════════════════════
function addTransaction(p) {
  const memberId = p.id;
  const itemsRaw = p.items;
  const txDate   = p.txDate || "";

  if (!memberId || !itemsRaw) return { status:"error", message:"缺少必要參數" };

  let items;
  try { items = JSON.parse(itemsRaw); }
  catch(e) { return { status:"error", message:"items 格式錯誤" }; }

  const memberRow = findMemberRow(memberId);
  if (!memberRow) return { status:"error", message:"找不到成員：" + memberId };

  const memberSheet   = getSheet(SHEET_MEMBERS);
  const logSheet      = getSheet(SHEET_LOG);
  const personalSheet = getSheet(memberId);
  const results       = [];

  items.forEach(entry => {
    const amount         = -Math.abs(Number(entry.amount));
    const item           = entry.item || "未指定";
    const note           = entry.note || "";
    const currentBalance = Number(memberSheet.getRange(memberRow, 4).getValue());
    const newBalance     = currentBalance + amount;
    const tid            = generateTID();
    const timestamp      = now();

    memberSheet.getRange(memberRow, 4).setValue(newBalance);
    memberSheet.getRange(memberRow, 6).setValue(timestamp);

    const logRow = [tid, timestamp, memberId, getMemberName(memberId), "扣款", item, note,
                    currentBalance, amount, newBalance, txDate];
    logSheet.appendRow(logRow);
    centerLastRow(logSheet, logRow.length);

    const personalRow = [tid, timestamp, "扣款", item, note,
                         currentBalance, amount, newBalance, txDate];
    personalSheet.appendRow(personalRow);
    centerLastRow(personalSheet, personalRow.length);

    results.push({ tid, item, note, amount, before:currentBalance, after:newBalance,
                   isLow: newBalance < CONFIG.LOW_BALANCE });
  });

  return {
    status       : "success",
    memberId,
    memberName   : getMemberName(memberId),
    txDate,
    results,
    finalBalance : results[results.length-1].after,
    isLow        : results[results.length-1].after < CONFIG.LOW_BALANCE,
  };
}

// ══════════════════════════════════════════════════════════════
//  儲值
// ══════════════════════════════════════════════════════════════
function topUp(memberId, amount, note) {
  if (!memberId || !amount || amount <= 0)
    return { status:"error", message:"缺少必要參數或金額無效" };

  const memberRow = findMemberRow(memberId);
  if (!memberRow) return { status:"error", message:"找不到成員：" + memberId };

  const memberSheet   = getSheet(SHEET_MEMBERS);
  const logSheet      = getSheet(SHEET_LOG);
  const personalSheet = getSheet(memberId);

  const currentBalance = Number(memberSheet.getRange(memberRow, 4).getValue());
  const newBalance     = currentBalance + amount;
  const currentTopup   = Number(memberSheet.getRange(memberRow, 5).getValue());
  const tid            = generateTID();
  const timestamp      = now();

  memberSheet.getRange(memberRow, 4).setValue(newBalance);
  memberSheet.getRange(memberRow, 5).setValue(currentTopup + amount);
  memberSheet.getRange(memberRow, 6).setValue(timestamp);

  const logRow = [tid, timestamp, memberId, getMemberName(memberId), "儲值", "儲值", note,
                  currentBalance, amount, newBalance, ""];
  logSheet.appendRow(logRow);
  centerLastRow(logSheet, logRow.length);

  const personalRow = [tid, timestamp, "儲值", "儲值", note,
                       currentBalance, amount, newBalance, ""];
  personalSheet.appendRow(personalRow);
  centerLastRow(personalSheet, personalRow.length);

  return {
    status:"success", tid, memberId,
    memberName: getMemberName(memberId),
    amount, before:currentBalance, after:newBalance,
  };
}

// ══════════════════════════════════════════════════════════════
//  工具函式
// ══════════════════════════════════════════════════════════════
function getMemberInfo(id) {
  const s    = getSheet(SHEET_MEMBERS);
  const rows = s.getDataRange().getValues().slice(1);
  const r    = rows.find(r => r[0] === id);
  if (!r) return null;
  return { id:r[0], name:r[1], emoji:r[2],
           balance:Number(r[3]), topupTotal:Number(r[4]),
           isLow: Number(r[3]) < CONFIG.LOW_BALANCE };
}

function getMemberName(id) {
  const m = MEMBERS.find(m => m.id === id);
  return m ? m.name : id;
}

function findMemberRow(id) {
  const s    = getSheet(SHEET_MEMBERS);
  const rows = s.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (rows[i][0] === id) return i+1;
  }
  return null;
}

function generateTID() {
  const d   = new Date();
  const pad = n => String(n).padStart(2,'0');
  return `TXN${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}${Math.floor(Math.random()*1000).toString().padStart(3,'0')}`;
}

function now() {
  const d   = new Date();
  const pad = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function styleHeader(sheet, row, cols) {
  const range = sheet.getRange(row, 1, 1, cols);
  range.setBackground("#1a1a2e");
  range.setFontColor("#e9c46a");
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
}

function centerAll(sheet) {
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(),1), Math.max(sheet.getMaxColumns(),1))
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
}

function centerLastRow(sheet, cols) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, cols)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
}
