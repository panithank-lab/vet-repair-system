// ============================================================
// VET REPAIR SYSTEM — Google Apps Script Backend
// คณะสัตวแพทยศาสตร์ มหาวิทยาลัยสงขลานครินทร์
// วิธีใช้: วางโค้ดทั้งหมดนี้ใน Apps Script แล้ว Deploy as Web App
// Execute as: Me | Who has access: Anyone
// ============================================================

const TICKETS_SHEET  = 'Tickets';
const HISTORY_SHEET  = 'History';

const TICKET_HEADERS = [
  'ID', 'CreatedAt', 'Name', 'Phone', 'Dept', 'Email',
  'Building', 'Room', 'Category', 'Priority', 'Desc',
  'EstimatedCost', 'ActualCost', 'Status',
  'Assignees', 'ApprovedAt', 'ApprovedBy', 'ApprovedBudget', 'CompletedAt'
];

const HISTORY_HEADERS = ['TicketID', 'Status', 'Date', 'Note', 'By'];

// ─── Entry Points ──────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const params = e.parameter || {};
    const body   = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    const action = params.action || body.action;

    let result;
    switch (action) {
      case 'init':       result = initSheets(); initConfigSheets();     break;
      case 'getAll':     result = getAllTickets();                      break;
      case 'add':        result = addTicket(body.ticket);               break;
      case 'update':     result = updateTicket(body.id, body.fields);  break;
      case 'addHistory': result = addHistory(body.ticketId, body.entry); break;
      case 'getConfig':  result = getConfig();                          break;
      case 'saveConfig': result = saveConfig(body.config);              break;
      default:           result = { error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ─── initSheets ────────────────────────────────────────────
// สร้าง Sheet Tickets และ History พร้อม Header และ Format
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tickets Sheet
  let ts = ss.getSheetByName(TICKETS_SHEET);
  if (!ts) {
    ts = ss.insertSheet(TICKETS_SHEET);
    ts.appendRow(TICKET_HEADERS);
    ts.getRange(1, 1, 1, TICKET_HEADERS.length)
      .setBackground('#1F3864')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    ts.setFrozenRows(1);
    ts.setColumnWidth(1,  160);   // ID
    ts.setColumnWidth(11, 300);   // Desc
    ts.setColumnWidth(15, 200);   // Assignees
  }

  // History Sheet
  let hs = ss.getSheetByName(HISTORY_SHEET);
  if (!hs) {
    hs = ss.insertSheet(HISTORY_SHEET);
    hs.appendRow(HISTORY_HEADERS);
    hs.getRange(1, 1, 1, HISTORY_HEADERS.length)
      .setBackground('#2d4a6b')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11);
    hs.setFrozenRows(1);
    hs.setColumnWidth(1, 160);   // TicketID
    hs.setColumnWidth(4, 300);   // Note
  }

  return { message: 'Sheets initialized successfully' };
}

// ═══════════ ระบบตั้งค่า (อาคาร / สถานะ / ประเภทงาน) ═══════════
const DEFAULT_BUILDINGS = [
  ['b1','อาคารจุฬาภรณการุณยรักษ์','จุฬาภร','🏢',120000,'FALSE'],
  ['b2','อาคารสหเวชศาสตร์','สหเวช','🏛',80000,'FALSE'],
  ['b3','โรงพยาบาลปศุสัตว์','โรงพยาบาล','🏥',200000,'FALSE'],
  ['b4','อาคารวิจัยสัตวน้ำและสัตว์ปีก','วิจัย','🔬',100000,'FALSE'],
];
const DEFAULT_STATUSES = [
  ['pending_approval','รอวิศวกรอนุมัติ','b-amber','FALSE','FALSE','FALSE'],
  ['assigned','มอบหมายแล้ว','b-purple','FALSE','FALSE','FALSE'],
  ['survey','กำลังสำรวจหน้างาน','b-amber','FALSE','FALSE','FALSE'],
  ['purchase','อยู่ระหว่างจัดซื้อ','b-purple','FALSE','FALSE','FALSE'],
  ['working','กำลังดำเนินการซ่อม','b-blue','FALSE','FALSE','FALSE'],
  ['insurance','ส่งประกันเคลม','b-teal','TRUE','FALSE','FALSE'],
  ['contractor','ส่งผู้รับเหมา','b-teal','TRUE','FALSE','FALSE'],
  ['done','เสร็จเรียบร้อย','b-green','FALSE','TRUE','FALSE'],
  ['cannot','ซ่อมไม่ได้','b-red','FALSE','TRUE','FALSE'],
  ['cancelled','ยกเลิกใบแจ้ง','b-gray','FALSE','TRUE','FALSE'],
];
const DEFAULT_CATEGORIES = [
  ['ระบบไฟฟ้า',5,'FALSE'],
  ['ระบบประปา/สุขาภิบาล',5,'FALSE'],
  ['ระบบปรับอากาศ',7,'FALSE'],
  ['งานโครงสร้างอาคาร',30,'FALSE'],
  ['ระบบสื่อสาร/เครือข่าย IT',7,'FALSE'],
  ['งานครุภัณฑ์/เฟอร์นิเจอร์',14,'FALSE'],
  ['อื่นๆ',14,'FALSE'],
];
const CFG = {
  buildings:  { sheet:'Buildings',  headers:['ID','Name','Short','Icon','Budget','Hidden'],        seed:DEFAULT_BUILDINGS },
  statuses:   { sheet:'Statuses',   headers:['ID','Label','Class','Paused','Closed','Hidden'],       seed:DEFAULT_STATUSES },
  categories: { sheet:'Categories', headers:['Name','SLA','Hidden'],                                 seed:DEFAULT_CATEGORIES },
};

function initConfigSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CFG).forEach(function(k){
    const c = CFG[k];
    let sh = ss.getSheetByName(c.sheet);
    if (!sh) {
      sh = ss.insertSheet(c.sheet);
      sh.appendRow(c.headers);
      sh.getRange(1,1,1,c.headers.length).setBackground('#0f6e56').setFontColor('#FFFFFF').setFontWeight('bold');
      sh.setFrozenRows(1);
      c.seed.forEach(function(row){ sh.appendRow(row); });
    }
  });
  return { message: 'Config sheets ready' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // สร้าง+seed ถ้ายังไม่มี
  const missing = Object.keys(CFG).some(function(k){ return !ss.getSheetByName(CFG[k].sheet); });
  if (missing) initConfigSheets();

  function read(k){
    const c = CFG[k];
    const sh = ss.getSheetByName(c.sheet);
    const d = sh.getDataRange().getValues();
    const out = [];
    for (let i=1;i<d.length;i++){ if(d[i][0]==='' || d[i][0]===null) continue; out.push(d[i]); }
    return out;
  }
  const B = read('buildings').map(function(r){ return {id:String(r[0]),name:r[1],short:r[2],icon:r[3],budget:Number(r[4])||0,hidden:String(r[5]).toUpperCase()==='TRUE'}; });
  const S = read('statuses').map(function(r){ return {id:String(r[0]),label:r[1],cls:r[2],paused:String(r[3]).toUpperCase()==='TRUE',closed:String(r[4]).toUpperCase()==='TRUE',hidden:String(r[5]).toUpperCase()==='TRUE'}; });
  const C = read('categories').map(function(r){ return {name:String(r[0]),sla:Number(r[1])||14,hidden:String(r[2]).toUpperCase()==='TRUE'}; });
  return { buildings:B, statuses:S, categories:C };
}

function saveConfig(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initConfigSheets();
  function writeAll(k, rows){
    const c = CFG[k];
    const sh = ss.getSheetByName(c.sheet);
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2,1,last-1,c.headers.length).clearContent();
    if (rows.length) sh.getRange(2,1,rows.length,c.headers.length).setValues(rows);
  }
  if (cfg.buildings) writeAll('buildings', cfg.buildings.map(function(b){ return [b.id,b.name,b.short,b.icon,b.budget||0, b.hidden?'TRUE':'FALSE']; }));
  if (cfg.statuses)  writeAll('statuses',  cfg.statuses.map(function(s){ return [s.id,s.label,s.cls,s.paused?'TRUE':'FALSE',s.closed?'TRUE':'FALSE',s.hidden?'TRUE':'FALSE']; }));
  if (cfg.categories)writeAll('categories',cfg.categories.map(function(c){ return [c.name,c.sla||14,c.hidden?'TRUE':'FALSE']; }));
  return { message: 'Config saved' };
}

// ─── getAllTickets ──────────────────────────────────────────
// ดึงข้อมูลทั้งหมดจาก Tickets + History ส่งกลับเป็น JSON
function getAllTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = ss.getSheetByName(TICKETS_SHEET);
  const hs = ss.getSheetByName(HISTORY_SHEET);

  if (!ts) return { tickets: [], history: {} };

  const tData = ts.getDataRange().getValues();
  const hData = hs ? hs.getDataRange().getValues() : [HISTORY_HEADERS];

  // แปลง rows เป็น ticket objects
  const tickets = [];
  for (let i = 1; i < tData.length; i++) {
    const r = tData[i];
    if (!r[0]) continue; // ข้าม row ว่าง
    tickets.push({
      id:             String(r[0]),
      createdAt:      r[1]  ? new Date(r[1]).toISOString()  : '',
      name:           r[2]  || '',
      phone:          r[3]  || '',
      dept:           r[4]  || '',
      email:          r[5]  || '',
      building:       r[6]  || '',
      room:           r[7]  || '',
      category:       r[8]  || '',
      priority:       r[9]  || 'low',
      desc:           r[10] || '',
      estimatedCost:  Number(r[11]) || 0,
      actualCost:     Number(r[12]) || 0,
      status:         r[13] || 'pending_approval',
      assignees:      r[14] ? String(r[14]).split(',').map(s => s.trim()).filter(Boolean) : [],
      approvedAt:     r[15] ? new Date(r[15]).toISOString() : null,
      approvedBy:     r[16] || null,
      approvedBudget: Number(r[17]) || 0,
      completedAt:    r[18] ? new Date(r[18]).toISOString() : null,
    });
  }

  // จัดกลุ่ม history ตาม ticketId
  const history = {};
  for (let i = 1; i < hData.length; i++) {
    const r = hData[i];
    if (!r[0]) continue;
    const tid = String(r[0]);
    if (!history[tid]) history[tid] = [];
    history[tid].push({
      status: r[1] || '',
      date:   r[2] ? new Date(r[2]).toISOString() : '',
      note:   r[3] || '',
      by:     r[4] || '',
    });
  }

  return { tickets, history };
}

// ─── addTicket ─────────────────────────────────────────────
// เพิ่มใบแจ้งซ่อมใหม่ลงใน Tickets Sheet
function addTicket(ticket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ts = ss.getSheetByName(TICKETS_SHEET);
  if (!ts) {
    initSheets();
    ts = ss.getSheetByName(TICKETS_SHEET);
  }

  const row = [
    ticket.id          || '',
    ticket.createdAt   || new Date().toISOString(),
    ticket.name        || '',
    ticket.phone       || '',
    ticket.dept        || '',
    ticket.email       || '',
    ticket.building    || '',
    ticket.room        || '',
    ticket.category    || '',
    ticket.priority    || 'low',
    ticket.desc        || '',
    Number(ticket.estimatedCost) || 0,
    0,                             // actualCost เริ่มต้น 0
    ticket.status      || 'pending_approval',
    (ticket.assignees  || []).join(', '),
    '',                            // approvedAt
    '',                            // approvedBy
    '',                            // approvedBudget
    '',                            // completedAt
  ];

  ts.appendRow(row);

  // ระบายสีตาม priority
  const lastRow = ts.getLastRow();
  const bgColor = ticket.priority === 'high' ? '#FDE8E8'
                : ticket.priority === 'med'  ? '#FFF8E6'
                :                              '#FFFFFF';
  ts.getRange(lastRow, 1, 1, TICKET_HEADERS.length).setBackground(bgColor);

  return { inserted: lastRow };
}

// ─── updateTicket ──────────────────────────────────────────
// อัปเดต field ที่ระบุของใบแจ้งซ่อม และเปลี่ยนสีตามสถานะ
function updateTicket(id, fields) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = ss.getSheetByName(TICKETS_SHEET);
  if (!ts) return { error: 'Tickets sheet not found' };

  const data   = ts.getDataRange().getValues();
  const colMap = {
    name:           3,
    phone:          4,
    dept:           5,
    email:          6,
    building:       7,
    room:           8,
    category:       9,
    priority:       10,
    desc:           11,
    estimatedCost:  12,
    actualCost:     13,
    status:         14,
    assignees:      15,
    approvedAt:     16,
    approvedBy:     17,
    approvedBudget: 18,
    completedAt:    19,
  };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(id)) continue;

    const rowNum = i + 1;

    // อัปเดตทีละ field
    Object.entries(fields).forEach(([key, val]) => {
      const col = colMap[key];
      if (!col) return;
      const v = Array.isArray(val) ? val.join(', ') : (val === null ? '' : String(val));
      ts.getRange(rowNum, col).setValue(v);
    });

    // เปลี่ยนสีพื้นหลังตามสถานะ
    if (fields.status) {
      const statusColors = {
        pending_approval: '#FFF8E6',
        assigned:         '#E6F1FB',
        survey:           '#FFF8E6',
        purchase:         '#EEEDFE',
        working:          '#E6F1FB',
        done:             '#EAF3DE',
        cannot:           '#FCEBEB',
      };
      const bg = statusColors[fields.status] || '#FFFFFF';
      ts.getRange(rowNum, 1, 1, TICKET_HEADERS.length).setBackground(bg);
    }

    return { updated: rowNum };
  }

  return { error: 'Ticket not found: ' + id };
}

// ─── addHistory ────────────────────────────────────────────
// บันทึกประวัติการเปลี่ยนสถานะลงใน History Sheet
function addHistory(ticketId, entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hs = ss.getSheetByName(HISTORY_SHEET);
  if (!hs) {
    initSheets();
    hs = ss.getSheetByName(HISTORY_SHEET);
  }

  hs.appendRow([
    ticketId        || '',
    entry.status    || '',
    entry.date      || new Date().toISOString(),
    entry.note      || '',
    entry.by        || '',
  ]);

  return { ok: true };
}

// ============================================================
// นำเข้าข้อมูลแจ้งซ่อมย้อนหลัง (จาก Google Form เดิม) — รันครั้งเดียว
// วิธีใช้: หลังวาง Code.gs แล้ว เลือกฟังก์ชัน importInitialData ที่แถบด้านบน
//         แล้วกด Run (▶) หนึ่งครั้ง — ระบบจะเพิ่ม 20 รายการลง Sheet
// หากรันซ้ำจะข้ามรายการที่มีอยู่แล้ว (กันข้อมูลซ้ำ)
// ============================================================
function importInitialData() {
  initSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ts = ss.getSheetByName(TICKETS_SHEET);

  // เก็บ ID ที่มีอยู่แล้ว เพื่อกันข้อมูลซ้ำ
  var existing = {};
  var data = ts.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0]) existing[String(data[i][0])] = true; }

  var SEED = [
  {
    "id": "REP-H001",
    "createdAt": "2026-04-12T18:43:39",
    "name": "ธนิษฐา เอชะนิยม",
    "phone": "0622541444",
    "dept": "บุคลากร/อาจารย์",
    "email": "thanidtha.vet.psu@gmail.com",
    "building": "b2",
    "room": "ชั้น 10 · ห้องน้ำหญิง ห้องที่ 2 และ 3",
    "category": "ระบบประปา/สุขาภิบาล",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "เต้าใต้ใช้ไม่ได้ ห้องที่ 2 ไม่มีท่อประปาเข้าชักโครก ห้องที่ 3 ไม่มีสายชำระ",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H002",
    "createdAt": "2026-04-21T10:10:44",
    "name": "สิริลักษณ์ แก้วมณี",
    "phone": "0873345930",
    "dept": "บุคลากร/อาจารย์",
    "email": "siriluk.ka@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 4 · ห้องปฏิบัติการ 4",
    "category": "ระบบปรับอากาศ",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "แอร์ไม่ทำงาน 1 ตัว ขึ้นไฟกระพริบ อีก 1 ตัวใช้งานได้แต่ไม่สามารถทำให้ห้องเย็นได้ และเกรงว่าตัวที่ใช้ได้อาจเสียตามเพราะทำงานหนักเกินไป [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H003",
    "createdAt": "2026-04-28T09:41:07",
    "name": "บารมี ชาญบุญญานนท์",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "baramee.c@psu.ac.th",
    "building": "b3",
    "room": "ชั้น 1 · คอกเลี้ยงวัว (รหัสทรัพย์สิน 9613 / อาคารเลี้ยงปศุสัตว์)",
    "category": "งานโครงสร้างอาคาร",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "1. ดอกเหล็กเป็นสนิม ผุกร่อน จำนวน 4 ดอก  2. ประตูเปิด-ปิดคอกชำรุด จำนวน 1 ประตู  3. ขอเสริมโครงเหล็กกั้นคอกวัว เนื่องจากปัจจุบันช่องว่างมากจนวัวหลุดออกมาได้",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H004",
    "createdAt": "2026-04-30T14:50:33",
    "name": "สภาวพร ประจันฉะเสน",
    "phone": "0860281609",
    "dept": "บุคลากร/อาจารย์",
    "email": "sakaoporn.p@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 4 · ห้องปฏิบัติการ 1",
    "category": "ระบบประปา/สุขาภิบาล",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "น้ำรั่ว",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H005",
    "createdAt": "2026-05-06T13:30:34",
    "name": "นภาพร พงศ์อินทร์",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "napaporn.b@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 3 · ห้องการศึกษา (รหัสทรัพย์สิน 9623)",
    "category": "ระบบปรับอากาศ",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "เครื่องปรับอากาศมีเสียงดัง (เหมือนเสียงลูกหนู) ตลอดเวลา รบกวนช่วยตรวจสอบก่อนเกิดอาการเสียหายมากกว่านี้ [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H006",
    "createdAt": "2026-05-07T16:57:50",
    "name": "ราชกุล วีระพราน",
    "phone": "0839944706",
    "dept": "บุคลากร/อาจารย์",
    "email": "ratchakul.w@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 3 · ห้องพักนักวิทยาศาสตร์",
    "category": "ระบบปรับอากาศ",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "แอร์ตัด น้ำรั่ว",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H007",
    "createdAt": "2026-05-12T14:35:52",
    "name": "บารมี ชาญบุญญานนท์",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "baramee.c@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 1 · ห้องเก็บซากสัตว์เล็ก (รหัสทรัพย์สิน 9613)",
    "category": "ระบบประปา/สุขาภิบาล",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "ก๊อกน้ำที่อ่างล้างมือรั่ว ไม่สามารถปิดสนิทได้ (ไม่มีเลขครุภัณฑ์)",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H008",
    "createdAt": "2026-05-19T12:58:15",
    "name": "ราชกุล วีระพราน",
    "phone": "0839944706",
    "dept": "บุคลากร/อาจารย์",
    "email": "ratchakul.w@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 4 · แลป 1",
    "category": "ระบบปรับอากาศ",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "น้ำรั่วออกจากแอร์ (เลขครุภัณฑ์ 4/2)",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H009",
    "createdAt": "2026-05-19T15:17:02",
    "name": "เสาวคนธ์ อินทร์ด้วง",
    "phone": "0869633577",
    "dept": "บุคลากร/อาจารย์",
    "email": "saowakon.i@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 3 · ห้องพักนักวิทยาศาสตร์",
    "category": "ระบบไฟฟ้า",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "หลอดไฟดับ",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H010",
    "createdAt": "2026-05-19T16:10:45",
    "name": "บารมี ชาญบุญญานนท์",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "baramee.c@psu.ac.th",
    "building": "b1",
    "room": "ชั้นใต้ดิน (-1) · ห้องสัตว์น้ำ (รหัสทรัพย์สิน 9613)",
    "category": "ระบบไฟฟ้า",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "หลอดไฟเสีย จำนวน 2 ดวง [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H011",
    "createdAt": "2026-05-25T10:56:05",
    "name": "บารมี ชาญบุญญานนท์",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "baramee.c@psu.ac.th",
    "building": "b4",
    "room": "ถังเก็บน้ำ ด้านหลังตึก (รหัสทรัพย์สิน 9613 / อาคารสัตว์น้ำและสัตว์ปีก)",
    "category": "ระบบประปา/สุขาภิบาล",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "ฝาถังเก็บน้ำชำรุด หลุดออกจากตำแหน่ง [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H012",
    "createdAt": "2026-06-01T10:20:51",
    "name": "เมโอ เอชะนิยม",
    "phone": "0918944370",
    "dept": "บุคลากร/อาจารย์",
    "email": "meo.pompurin@gmail.com",
    "building": "b2",
    "room": "ชั้น 10 · ห้อง 10C1",
    "category": "ระบบไฟฟ้า",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "หลอดไฟกระพริบ ติด-ดับ สลับไป-มา",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H013",
    "createdAt": "2026-06-02T19:40:04",
    "name": "เมโอ เอชะนิยม",
    "phone": "0918944370",
    "dept": "บุคลากร/อาจารย์",
    "email": "meo.pompurin@gmail.com",
    "building": "b2",
    "room": "ชั้น 10 · ห้องพักอาจารย์ 9 (10C1)",
    "category": "ระบบไฟฟ้า",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "ไฟเสียทั้ง 2 ถาด (วันที่ 1 แจ้งไปหนึ่งถาดแล้ว) รวมถึงหน้ากากไฟ (แจ้งเครื่องปรับอากาศ และไฟ) [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H014",
    "createdAt": "2026-06-04T08:39:28",
    "name": "นายสุทธิพงษ์ ปัญญาภีรี",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "suttipong.p@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 1 · ห้องบริการการศึกษา (รหัสทรัพย์สิน 9608 / อาคารบริรักษ์การศึกษา)",
    "category": "ระบบปรับอากาศ",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "แอร์ไม่ทำงาน [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H015",
    "createdAt": "2026-06-15T11:29:44",
    "name": "นายพีระพล สรยิง",
    "phone": "0816090693",
    "dept": "บุคลากร/อาจารย์",
    "email": "peerapon.s@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 1 · ห้องปฏิบัติการผสมสูตรอาหารสัตว์",
    "category": "งานครุภัณฑ์/เฟอร์นิเจอร์",
    "priority": "low",
    "estimatedCost": 0,
    "desc": "บานพับกั้นสายเตา (สีเขียว) ชำรุดเสียหาย (ไม่มีรหัสทรัพย์สิน)",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H016",
    "createdAt": "2026-06-15T11:33:05",
    "name": "นายพีระพล สรยิง",
    "phone": "0816090693",
    "dept": "บุคลากร/อาจารย์",
    "email": "peerapon.s@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 1 · ห้องผสมสูตรอาหารสัตว์เล็ก",
    "category": "งานครุภัณฑ์/เฟอร์นิเจอร์",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "พัดลมดูดผนังเสีย มอเตอร์ไม่ทำงาน จำนวน 5 ตัว",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H017",
    "createdAt": "2026-06-16T10:28:23",
    "name": "เสาวคนธ์ อินทร์ด้วง",
    "phone": "0869633577",
    "dept": "บุคลากร/อาจารย์",
    "email": "saowakon.i@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 4 · ห้องแลป 1",
    "category": "ระบบปรับอากาศ",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "แอร์ในแลปไม่ทำความเย็น 1 ตัว",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H018",
    "createdAt": "2026-06-19T13:40:45",
    "name": "พีรพัฒน์ แก้วมณี",
    "phone": "0856732121",
    "dept": "บุคลากร/อาจารย์",
    "email": "peerapat.ka@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 2 · ห้องน้ำชั้นสอง (ตรงตู้)",
    "category": "ระบบประปา/สุขาภิบาล",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "ท่อระบายน้ำตัน น้ำดันย้อนขึ้นไม่ระบาย [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H019",
    "createdAt": "2026-06-23T10:35:40",
    "name": "ราชกุล วีระพราน",
    "phone": "0839944706",
    "dept": "บุคลากร/อาจารย์",
    "email": "ratchakul.w@psu.ac.th",
    "building": "b1",
    "room": "แลป 1",
    "category": "ระบบปรับอากาศ",
    "priority": "med",
    "estimatedCost": 0,
    "desc": "เครื่องดับ รีโมทไม่ติด",
    "status": "pending_approval",
    "assignees": []
  },
  {
    "id": "REP-H020",
    "createdAt": "2026-06-26T08:04:31",
    "name": "นัสรินทร์ เจ๊ะปิ",
    "phone": "",
    "dept": "บุคลากร/อาจารย์",
    "email": "nasrin.c@psu.ac.th",
    "building": "b1",
    "room": "ชั้น 3 · ห้องคณบดี (รหัสทรัพย์สิน 9627)",
    "category": "งานโครงสร้างอาคาร",
    "priority": "high",
    "estimatedCost": 0,
    "desc": "น้ำฟอร์มาลีนแช่ซากสัตว์จากชั้น 4 ย้อยลงมาในฝ้าเพดานห้องคณบดี (ต้องเร่งตรวจสอบด่วน เนื่องจากฟอร์มาลีนเป็นสารอันตราย) [มีรูปหลักฐานแนบใน Google Drive]",
    "status": "pending_approval",
    "assignees": []
  }
];

  var added = 0;
  SEED.forEach(function(t) {
    if (existing[t.id]) return; // ข้ามถ้ามีแล้ว
    addTicket(t);
    addHistory(t.id, {
      status: 'pending_approval',
      date: t.createdAt,
      note: 'นำเข้าข้อมูลย้อนหลังจาก Google Form (ก่อนเริ่มใช้ระบบ)',
      by: 'นำเข้าข้อมูล'
    });
    added++;
  });

  Logger.log('นำเข้าสำเร็จ ' + added + ' รายการ (ข้ามซ้ำ ' + (SEED.length - added) + ' รายการ)');
  return { added: added, skipped: SEED.length - added };
}
