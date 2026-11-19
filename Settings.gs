/**
 * ===== ส่วนตั้งค่า สาขา / เหตุผลการย้าย =====
 * วางไฟล์นี้เป็นไฟล์ใหม่ (เช่น Settings.gs) ในโปรเจกต์ Apps Script เดิม
 *
 * ขั้นตอนติดตั้ง
 * 1) รันฟังก์ชัน setupAdminPin() 1 ครั้ง เพื่อตั้งรหัสผู้ดูแล (เก็บใน Script properties ชื่อ ADMIN_PIN)
 *    ถ้าต้องการเปลี่ยนรหัสภายหลัง แก้ค่าใน setupAdminPin() แล้วรันใหม่
 *    หรือแก้ที่ Project Settings > Script properties > ADMIN_PIN โดยตรง
 * 2) รันฟังก์ชัน seedSettings() 1 ครั้ง เพื่อสร้างชีต "Branches" และ "Reasons" พร้อมข้อมูลเดิม
 * 3) ใน doPost(e) เดิม ให้เพิ่มบรรทัดนี้เป็นบรรทัดแรก:
 *        var s = handleSettings(e); if (s) return s;
 *    และถ้ามี doGet(e) ให้เพิ่มบรรทัดเดียวกัน (ถ้าไม่มี doGet ใช้ doGet ด้านล่างได้เลย)
 * 4) Deploy > Manage deployments > Edit > New version (URL เดิมใช้ต่อได้)
 */

var BRANCH_SHEET = 'Branches';
var REASON_SHEET = 'Reasons';

/** รันครั้งเดียว: ตั้งรหัสผู้ดูแลสำหรับหน้า settings.html */
function setupAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', 'tg98765432');
  Logger.log('ตั้งรหัสผู้ดูแลเรียบร้อย');
}

// ถ้าโปรเจกต์เดิมมี doGet อยู่แล้ว ให้ลบฟังก์ชันนี้ แล้วใส่บรรทัด handleSettings ใน doGet เดิมแทน
function doGet(e) {
  var s = handleSettings(e);
  if (s) return s;
  return jsonOut_({ status: 'error', message: 'unknown opt' });
}

/** คืนค่า TextOutput ถ้าเป็นคำสั่งของหน้าตั้งค่า, คืน null ถ้าไม่ใช่ (ให้ doPost เดิมทำงานต่อ) */
function handleSettings(e) {
  var p = (e && e.parameter) || {};
  var opt = p.opt || '';
  try {
    if (opt === 'getsettings') {
      return jsonOut_({ status: 'success', branches: readBranches_(), reasons: readReasons_() });
    }
    if (opt === 'checkpin') {
      requirePin_(p.pin);
      return jsonOut_({ status: 'success' });
    }
    var actions = {
      branch_add: branchAdd_, branch_update: branchUpdate_, branch_delete: branchDelete_,
      reason_add: reasonAdd_, reason_update: reasonUpdate_, reason_delete: reasonDelete_
    };
    if (!actions[opt]) return null;

    requirePin_(p.pin);
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      actions[opt](p);
    } finally {
      lock.releaseLock();
    }
    return jsonOut_({ status: 'success', branches: readBranches_(), reasons: readReasons_() });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err.message || err) });
  }
}

// ---------- สาขา ----------
function branchAdd_(p) {
  var code = clean_(p.code).toUpperCase(), name = clean_(p.name);
  if (!code || !name) throw new Error('กรุณากรอกรหัสและชื่อสาขา');
  if (findRow_(BRANCH_SHEET, code) > 0) throw new Error('รหัสสาขา ' + code + ' มีอยู่แล้ว');
  appendText_(sheet_(BRANCH_SHEET, ['code', 'name']), [code, name]);
}

function branchUpdate_(p) {
  var oldCode = clean_(p.oldCode).toUpperCase();
  var code = clean_(p.code).toUpperCase(), name = clean_(p.name);
  if (!code || !name) throw new Error('กรุณากรอกรหัสและชื่อสาขา');
  var row = findRow_(BRANCH_SHEET, oldCode);
  if (row < 1) throw new Error('ไม่พบสาขา ' + oldCode);
  if (code !== oldCode && findRow_(BRANCH_SHEET, code) > 0) throw new Error('รหัสสาขา ' + code + ' มีอยู่แล้ว');
  sheet_(BRANCH_SHEET).getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[code, name]]);
}

function branchDelete_(p) {
  var code = clean_(p.code).toUpperCase();
  var row = findRow_(BRANCH_SHEET, code);
  if (row < 1) throw new Error('ไม่พบสาขา ' + code);
  sheet_(BRANCH_SHEET).deleteRow(row);
}

// ---------- เหตุผลการย้าย ----------
function reasonAdd_(p) {
  var name = clean_(p.name);
  if (!name) throw new Error('กรุณากรอกเหตุผล');
  if (findRow_(REASON_SHEET, name) > 0) throw new Error('เหตุผลนี้มีอยู่แล้ว');
  appendText_(sheet_(REASON_SHEET, ['name']), [name]);
}

function reasonUpdate_(p) {
  var oldName = clean_(p.oldName), name = clean_(p.name);
  if (!name) throw new Error('กรุณากรอกเหตุผล');
  var row = findRow_(REASON_SHEET, oldName);
  if (row < 1) throw new Error('ไม่พบเหตุผล ' + oldName);
  if (name !== oldName && findRow_(REASON_SHEET, name) > 0) throw new Error('เหตุผลนี้มีอยู่แล้ว');
  sheet_(REASON_SHEET).getRange(row, 1).setNumberFormat('@').setValue(name);
}

function reasonDelete_(p) {
  var name = clean_(p.name);
  var row = findRow_(REASON_SHEET, name);
  if (row < 1) throw new Error('ไม่พบเหตุผล ' + name);
  sheet_(REASON_SHEET).deleteRow(row);
}

// ---------- อ่านข้อมูล ----------
function readBranches_() {
  var sh = sheet_(BRANCH_SHEET, ['code', 'name']);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { code: r[0], name: r[1] }; });
}

function readReasons_() {
  var sh = sheet_(REASON_SHEET, ['name']);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues()
    .map(function (r) { return r[0]; })
    .filter(function (v) { return v; });
}

// ---------- ตัวช่วย ----------
function requirePin_(pin) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!real) throw new Error('ยังไม่ได้ตั้งค่า ADMIN_PIN ใน Script properties');
  if (String(pin || '') !== real) throw new Error('รหัสผู้ดูแลไม่ถูกต้อง');
}

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** หาแถวที่คอลัมน์ A ตรงกับ key (เทียบแบบไม่สนตัวพิมพ์) คืนเลขแถว หรือ -1 */
function findRow_(sheetName, key) {
  var sh = sheet_(sheetName);
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var vals = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  var k = String(key).trim().toUpperCase();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === k) return i + 2;
  }
  return -1;
}

function clean_(v) {
  var s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  return s.slice(0, 200);
}

/** เขียนแถวใหม่เป็นข้อความล้วน (กันไม่ให้ชีตตีความเป็นสูตร) */
function appendText_(sh, values) {
  sh.getRange(sh.getLastRow() + 1, 1, 1, values.length).setNumberFormat('@').setValues([values]);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งเดียว: สร้างชีตและใส่ข้อมูลตั้งต้น (จะไม่ทับถ้ามีข้อมูลแล้ว) */
function seedSettings() {
  var b = sheet_(BRANCH_SHEET, ['code', 'name']);
  if (b.getLastRow() < 2) {
    b.getRange(2, 1, DEFAULT_BRANCHES.length, 2).setNumberFormat('@').setValues(DEFAULT_BRANCHES);
  }
  var r = sheet_(REASON_SHEET, ['name']);
  if (r.getLastRow() < 2) {
    r.getRange(2, 1, DEFAULT_REASONS.length, 1).setNumberFormat('@').setValues(DEFAULT_REASONS.map(function (x) { return [x]; }));
  }
}

var DEFAULT_REASONS = ['ย้ายถาวร', 'หน่วยแทน', 'ย้ายชั่วคราว'];

var DEFAULT_BRANCHES = [
    ["BACR1", "tg บิ๊กซีอำนาจเจริญ"],
    ["BBDG1", "tg บิ๊กซีบ้านดุง อุดรธานี"],
    ["BBPE1", "tg บิ๊กซีบางพลี"],
    ["BBPE2", "Realme Shop บิ๊กซีบางพลี Kiosk"],
    ["BC001", "Consignment Big C : สาขาอำนาจเจริญ"],
    ["BC004", "Consignment Big C : สาขาเมกาบางนา"],
    ["BC005", "Consignment Big C : สาขาบางบอน"],
    ["BC007", "Consignment Big C : สาขาบางพลี"],
    ["BC010", "Consignment Big C : สาขาบางปะกอก"],
    ["BC011", "Consignment Big C : สาขาบางใหญ่"],
    ["BC014", "Consignment Big C : สาขาฉะเชิงเทรา ร้าน 1"],
    ["BC015", "Consignment Big C : สาขาฉะเชิงเทรา ร้าน 2"],
    ["BC016", "Consignment Big C : สาขาเชียงใหม่ ร้าน 2"],
    ["BC017", "Consignment Big C : สาขาแจ้งวัฒนะ ร้าน 2"],
    ["BC018", "Consignment Big C : สาขาชัยภูมิ"],
    ["BC020", "Consignment Big C : สาขาดาวคะนอง"],
    ["BC021", "Consignment Big C : สาขาดอนเมือง"],
    ["BC022", "Consignment Big C : สาขาเอกมัย"],
    ["BC023", "Consignment Big C : สาขาแฟชั่น ไอแลนด์"],
    ["BC027", "Consignment Big C : สาขาขอนแก่น"],
    ["BC033", "Consignment Big C : สาขาลำลูกกา"],
    ["BC037", "Consignment Big C : สาขาอิมพีเรียลลาดพร้าว"],
    ["BC038", "Consignment Big C : สาขาลาดพร้าว"],
    ["BC039", "Consignment Big C : สาขามหาชัย"],
    ["BC043", "Consignment Big C : สาขานครปฐม"],
    ["BC044", "Consignment Big C : สาขาศรีนครินทร์"],
    ["BC045", "Consignment Big C : สาขานครสวรรค์"],
    ["BC046", "Consignment Big C : สาขาหนองบัว"],
    ["BC047", "Consignment Big C : สาขาหนองจอก"],
    ["BC048", "Consignment Big C : สาขาหนองกี่"],
    ["BC056", "Consignment Big C : สาขาเพชรเกษม"],
    ["BC057", "Consignment Big C : สาขาโพนพิสัย หนองคาย"],
    ["BC061", "Consignment Big C : สาขาราชบุรี"],
    ["BC062", "Consignment Big C : สาขารัชดาฯ"],
    ["BC063", "Consignment Big C : สาขารามอินทรา"],
    ["BC064", "Consignment Big C : สาขาพระราม 2 ร้าน 2"],
    ["BC065", "Consignment Big C : สาขาร่มเกล้า"],
    ["BC067", "Consignment Big C : สาขารังสิต"],
    ["BC071", "Consignment Big C : สาขาศรีมหาโพธิ์ ปราจีน"],
    ["BC072", "Consignment Big C : สาขาสกลนคร"],
    ["BC073", "Consignment Big C : สาขาสุขาภิบาล 3 ร้าน 2"],
    ["BC074", "Consignment Big C : สาขาสุโขทัย"],
    ["BC075", "Consignment Big C : สาขาสมุย"],
    ["BC076", "Consignment Big C : สาขาสุขาภิบาล 3"],
    ["BC077", "Consignment Big C : สาขาสะพานควาย"],
    ["BC078", "Consignment Big C : สาขาสำโรง"],
    ["BC079", "Consignment Big C : สาขาสุรินทร์"],
    ["BC080", "Consignment Big C : สาขาศรีสะเกษ"],
    ["BC081", "Consignment Big C : สาขาสุขสวัสดิ์"],
    ["BC086", "Consignment Big C : สาขาลำลูกกา คลอง 4"],
    ["BC087", "Consignment Big C : สาขาวิเชียรบุรี เพชรบูรณ์"],
    ["BC090", "Consignment Big C : สาขาวงศ์สว่าง"],
    ["BC091", "Consignment Big C : สาขาบางนา"],
    ["BC094", "Consignment Big C : สาขาแจ้งวัฒนะ ร้าน 1"],
    ["BC096", "Consignment Big C : สาขาหางดง"],
    ["BC105", "Consignment Big C : สาขาพระราม 2 ร้าน 1"],
    ["BC106", "Consignment Big C : สาขารัตนาธิเบศร์ 1"],
    ["BC108", "Consignment Big C : สาขาสวนหลวง"],
    ["BC110", "Consignment Big C : สาขาสุขาภิบาล 5"],
    ["BC111", "Consignment Big C : สาขาทุ่งครุ"],
    ["BC116", "Consignment Big C : สาขาสระบุรี"],
    ["BC117", "Consignment Big C : สาขายะลา"],
    ["BC118", "Consignment Big C : สาขาอ้อมใหญ่"],
    ["BC119", "Consignment Big C : สาขาภูเก็ต"],
    ["BC120", "Consignment Big C : สาขาบ้านดุง"],
    ["BC121", "Consignment Big C : สาขานครพนม"],
    ["BC122", "Consignment Big C : สาขาบ้านบึง"],
    ["BC123", "Consignment Big C : สาขาแพร่"],
    ["BC124", "Consignment Big C : สาขาเชียงราย"],
    ["BC125", "Consignment Big C : สาขาสุราษฎร์ธานี"],
    ["BC126", "Consignment Big C : สาขาไทรน้อย"],
    ["BC127", "Consignment Big C : สาขาเพชรบุรี"],
    ["BC128", "Consignment Big C : สาขากาฬสินธุ์"],
    ["BC129", "Consignment Big C : สาขาน่าน"],
    ["BC130", "Consignment Big C : สาขาติวานนท์"],
    ["BC131", "Consignment Big C : สาขาอิสรภาพ"],
    ["BC132", "Consignment Big C : สาขาอ่อนนุช"],
    ["BC133", "Consignment Big C : สาขาราชดำริ"],
    ["BC134", "Consignment Big C : สาขาเพชรเกษม หลักสอง"],
    ["BC135", "Consignment Big C : สาขาภูเก็ต 2"],
    ["BC136", "Consignment Big C : สาขาพระราม 4"],
    ["BC137", "Consignment Big C : สาขาปัตตานี"],
    ["BC138", "Consignment Big C : สาขารังสิตคลอง 3"],
    ["BC139", "Consignment Big C : สาขาเชียงใหม่"],
    ["BC140", "Consignment Big C : สาขาพัทยากลาง"],
    ["BC141", "Consignment Big C : สาขาหัวหมาก"],
    ["BC142", "Consignment Big C : สาขาสุวินทวงศ์"],
    ["BC143", "Consignment Big C : สาขาอุดรธานี"],
    ["BC144", "Consignment Big C : สาขาฟิวเจอร์พาร์ค รังสิต"],
    ["BC145", "Consignment Big C : สาขากระนวน"],
    ["BC146", "Consignment Big C : สาขาอุบลราชธานี"],
    ["BC147", "Consignment Big C : สาขาบุรีรัมย์"],
    ["BC148", "Consignment Big C : สาขาชลบุรี"],
    ["BC149", "Consignment Big C : สาขาร้อยเอ็ด"],
    ["BC150", "Consignment Big C : สาขามุกดาหาร"],
    ["BC151", "Consignment Big C : สาขาสมุทรสงคราม"],
    ["BC152", "Consignment Big C : สาขาสุไหงโก-ลก"],
    ["BCRI1", "tg บิ๊กซี เชียงราย"],
    ["BHMK1", "tg บิ๊กซีหัวหมาก Town Center"],
    ["BHMK4", "Samsung (SPS) บิ๊กซีหัวหมาก Town Center Kiosk"],
    ["BHMK5", "Oppo Shop บิ๊กซีหัวหมาก Town Center Kiosk"],
    ["BHMK6", "Vivo Shop บิ๊กซีหัวหมาก Town Center"],
    ["BHMK7", "Honor Shop บิ๊กซีหัวหมาก Town Center"],
    ["BKPT1", "tg บิ๊กซีกำแพงเพชร"],
    ["BNSW3", "tg บิ๊กซีนครสวรรค์ Kiosk"],
    ["BNSW5", "tg วี-สแควร์พลาซ่า บิ๊กซีนครสวรรค์ Kiosk"],
    ["BPTY1", "tg บิ๊กซีพัทยากลาง"],
    ["BPTY2", "Samsung (SPS) บิ๊กซีพัทยากลาง"],
    ["BRAD2", "tg บิ๊กซีราชบุรี"],
    ["BSIK2", "Samsung (SIS) บิ๊กซีสีคิ้ว"],
    ["BSUK3", "Vivo Shop บิ๊กซีสุขสวัสดิ์"],
    ["BSWG1", "Vivo Shop บิ๊กซีสุวินทวงศ์"],
    ["BTAK1", "tg บิ๊กซีท่าตะโก"],
    ["BTAK2", "Samsung (SPS) บิ๊กซีท่าตะโก"],
    ["BWCB1", "tg บิ๊กซีวิเชียรบุรี"],
    ["CAYA1", "Xiaomi Shop เซ็นทรัลอยุธยา"],
    ["CAYA2", "tg เซ็นทรัลอยุธยา"],
    ["CAYA3", "Samsung (SES) เซ็นทรัลอยุธยา"],
    ["CBNA2", "tg เซ็นทรัล บางนา (ร้าน2)"],
    ["CBYI1", "Oppo Shop เซ็นทรัลเวสเกต"],
    ["CBYI2", "tg เซ็นทรัลเวสเกต"],
    ["CBYI3", "tg Smart Life เซ็นทรัลเวสเกต"],
    ["CBYI4", "Realme Shop เซ็นทรัลเวสเกต Kiosk"],
    ["CCAP1", "tg เซ็นทรัลแอร์พอร์ทเชียงใหม่"],
    ["CCAP2", "Vivo Shop เซ็นทรัลแอร์พอร์ตเชียงใหม่ Kiosk"],
    ["CCRI1", "tg เซ็นทรัลพลาซ่าเชียงราย"],
    ["CFTV1", "tg เซ็นทรัลเฟสติวัลเชียงใหม่"],
    ["CFTV2", "tg เซ็นทรัลเฟสติวัลเชียงใหม่ Kiosk"],
    ["CHDY1", "tg เซ็นทรัลหาดใหญ่"],
    ["CKBI1", "tg เซ็นทรัลกระบี่"],
    ["CKKN1", "tg เซ็นทรัลขอนแก่น"],
    ["CKKN2", "Samsung (SES) เซ็นทรัลขอนแก่น"],
    ["CLPA2", "tg เซ็นทรัลลำปาง Kiosk"],
    ["CLPW3", "Samsung (SES) เซ็นทรัลลาดพร้าว ชั้น 3"],
    ["CLPW7", "Huawei Shop เซ็นทรัลลาดพร้าว"],
    ["CMCH1", "tg เซ็นทรัลมหาชัย"],
    ["CNKO1", "tg เซ็นทรัลนครศรีธรรมราช"],
    ["CNKO2", "tg เซ็นทรัลนครศรีธรรมราช Kiosk"],
    ["CNPT1", "tg เซ็นทรัลนครปฐม"],
    ["CNSM1", "tg เซ็นทรัลนครราชสีมา"],
    ["CNSM2", "Oppo Shop เซ็นทรัลนครราชสีมา"],
    ["CNSM3", "Samsung (SES) เซ็นทรัลนครราชสีมา"],
    ["CNSM4", "tg เซ็นทรัลนครราชสีมา Kiosk"],
    ["CNSW1", "tg เซ็นทรัลนครสวรรค์"],
    ["CNSW2", "Realme Shop เซ็นทรัลนครสวรรค์"],
    ["CPKO2", "tg เซ็นทรัลปิ่นเกล้า ชั้น 3"],
    ["CPRK1", "tg เซ็นทรัลพาร์ค Kiosk"],
    ["CPSL2", "tg เซ็นทรัลพิษณุโลก Kiosk"],
    ["CRIN1", "tg เซ็นทรัลรามอินทรา"],
    ["CRIN3", "tg เซ็นทรัลเฟสติวัลเลียบทางด่วนรามอินทรา"],
    ["CRIN4", "Samsung (SES) เซ็นทรัลเฟสติวัลเลียบทางด่วนรามอินทรา"],
    ["CRM21", "tg เซ็นทรัลพระราม 2 ร้าน 1"],
    ["CRM22", "Samsung (SES) เซ็นทรัลพระราม 2 ชั้น 2"],
    ["CRM23", "Samsung (SES) เซ็นทรัลพระราม 2 ชั้น 3"],
    ["CRM24", "tg เซ็นทรัลพระราม 2 ชั้น 3"],
    ["CRM27", "Vivo Shop เซ็นทรัลพระราม 2 ชั้น 2"],
    ["CRM32", "tg เซ็นทรัลพระราม 3"],
    ["CRM91", "tg เซ็นทรัล พระราม 9"],
    ["CRM92", "Vivo Shop เซ็นทรัลพระราม 9 Kiosk"],
    ["CRM93", "Samsung (SES) เซ็นทรัลพระราม 9  ชั้น G"],
    ["CRM94", "tg Service เซ็นทรัลพระราม 9 ชั้น 4"],
    ["CRYG1", "tg เซ็นทรัลระยอง"],
    ["CSIR1", "tg เซ็นทรัลศรีราชา"],
    ["CSIR2", "Xiaomi Shop เซ็นทรัลศรีราชา"],
    ["CSUM1", "tg เซ็นทรัลเฟสติวัลสมุย"],
    ["CSUM2", "Samsung (SES) เซ็นทรัลเฟสติวัลสมุย"],
    ["CSUR1", "tg เซ็นทรัลสุราษฎร์ธานี"],
    ["CUBN1", "tg เซ็นทรัลอุบลราชธานี"],
    ["CUDO2", "Huawei Shop เซ็นทรัลอุดรธานี Kiosk"],
    ["CWAT1", "tg เซ็นทรัลแจ้งวัฒนะ ชั้น 4"],
    ["CWEV1", "tg เซ็นทรัลเวสวิลล์"],
    ["CWEV2", "Realme Shop เซ็นทรัลเวสวิลล์"],
    ["CWOR3", "Samsung (SES) เซ็นทรัลเวิลด์ ชั้น 4"],
    ["CWOR5", "tg เซ็นทรัลเวิลด์ ( ชั้น5 )"],
    ["CWOR6", "tg เซ็นทรัลเวิลด์ Kiosk"],
    ["CWSW1", "tg เซ็นทรัลวงศ์สว่าง Town Center"],
    ["FRDA1", "tg บิ๊กซีรัชดาฯ"],
    ["FRM42", "tg บิ๊กซีพระราม 4"],
    ["KMBK1", "tg มาบุญครอง (MBK) ชั้น 4"],
    ["KMBK5", "Samsung (SES) มาบุญครอง (MBK) ชั้น 4"],
    ["KMBK6", "Honor Shop มาบุญครอง (MBK) ชั้น 4"],
    ["KMBK7", "tg Service มาบุญครอง (MBK) ชั้น 5"],
    ["LATG1", "tg โลตัสอ่างทอง"],
    ["LBAB1", "tg โลตัสบ้านบึง Kiosk"],
    ["LBAB2", "Vivo Shop โลตัสบ้านบึง"],
    ["LBAB3", "AIS Buddy โลตัสบ้านบึง"],
    ["LBAB4", "Oppo Shop โลตัสบ้านบึง"],
    ["LBAB5", "Samsung (SPS) โลตัสบ้านบึง"],
    ["LBAB6", "Huawei Shop โลตัสบ้านบึง Kiosk"],
    ["LBBO1", "tg โลตัสบางบ่อ Kiosk"],
    ["LBKE3", "Samsung (SPS) โลตัสบางแค"],
    ["LBKR3", "tg โลตัสบางกรวย-ไทรน้อย Kiosk"],
    ["LBLN1", "tg โลตัสบางเลน"],
    ["LBOW2", "Samsung (SIS) โลตัสบ่อวิน"],
    ["LBPK2", "tg โลตัสบางปะกอก ชั้น 3"],
    ["LBSP1", "tg โลตัสบึงสามพัน"],
    ["LBUH1", "tg โลตัสบัวใหญ่"],
    ["LBWN1", "tg โลตัสบ่อวิน"],
    ["LBYI1", "tg โลตัสบางใหญ่"],
    ["LBYI2", "Samsung (SES) โลตัสบางใหญ่"],
    ["LCAM1", "tg สาขาโลตัสชะอำ Kiosk"],
    ["LCAN2", "tg โลตัสจันทบุรี Kiosk"],
    ["LCHP1", "tg โลตัสชุมพร Kiosk"],
    ["LCHP2", "Samsung (SPS) โลตัสชุมพร"],
    ["LCHY1", "tg โลตัสเชียงยืน"],
    ["LCKM1", "tg โลตัสเชียงคำ พะเยา"],
    ["LCMT1", "Telewiz โลตัสจอมทอง เชียงใหม่"],
    ["LCMT2", "tg โลตัสจอมทอง เชียงใหม่"],
    ["LCNT1", "tg โลตัสชัยนาท Kiosk"],
    ["LCPE1", "tg โลตัสชุมแพ"],
    ["LCRI1", "Telewiz โลตัสแม่จัน"],
    ["LCYP2", "Realme Shop โลตัสชัยภูมิ Kiosk"],
    ["LCYP3", "tg โลตัสชัยภูมิ"],
    ["LDAN1", "tg โลตัสด่านขุนทด"],
    ["LHAT1", "tg โลตัสหัวทะเล"],
    ["LHAT2", "tg โลตัสหัวทะเล Kiosk"],
    ["LJOH1", "tg โลตัสจอหอ"],
    ["LKBI2", "tg โลตัสกระบี่"],
    ["LKKN2", "Samsung (SPS) โลตัสขอนแก่น Kiosk"],
    ["LKTB1", "Telewiz โลตัสกระทุ่มแบน สมุทรสาคร"],
    ["LKUM1", "Samsung (SPS) โลตัสคำเที่ยง"],
    ["LKUM3", "tg โลตัสคำเที่ยง"],
    ["LKUM4", "Vivo Shop โลตัสคำเที่ยง"],
    ["LKUM5", "Realme Shop โลตัสคำเที่ยง"],
    ["LKUP1", "tg โลตัสกุมภวาปี อุดรธานี"],
    ["LKUP2", "tg (H2O) โลตัสกุมภวาปี"],
    ["LLAM1", "tg โลตัสลำปลายมาศ"],
    ["LLKL2", "tg โลตัสเลียบคลอง 2"],
    ["LLSI1", "tg โลตัสหลักสี่"],
    ["LMEN2", "tg โลตัสมีนบุรี Kiosk"],
    ["LNAR1", "tg โลตัสนางรอง"],
    ["LNAW3", "tg โลตัส นวนคร ชั้น2"],
    ["LNCS1", "tg โลตัสนครชัยศรี"],
    ["LNKO1", "tg โลตัสนครศรีธรรมราช"],
    ["LNKR1", "Oppo Shop โลตัสศรีนครินทร์ Kiosk"],
    ["LNOJ1", "Samsung (SPS) โลตัสหนองจอก Kiosk"],
    ["LNRP2", "AIS SHOP โลตัส นอร์ทราชพฤกษ์"],
    ["LPAC1", "tg โลตัสประโคนชัย"],
    ["LPBR2", "tg โลตัสปราณบุรี"],
    ["LPBR3", "Samsung (SPS) โลตัสปราณบุรี"],
    ["LPCT1", "tg โลตัสท่าหลวง พิจิตร"],
    ["LPKT2", "tg โลตัสภูเก็ต"],
    ["LPNK1", "tg โลตัสพนมสารคาม"],
    ["LPNK2", "tg (SIS70) โลตัสพนมสารคาม"],
    ["LPSL1", "tg โลตัสพิษณุโลก"],
    ["LPSL2", "Realme Shop โลตัสท่าทอง พิษณุโลก"],
    ["LPTC1", "tg โลตัสปักธงชัย"],
    ["LPTM1", "tg โลตัสปทุมธานี Kiosk"],
    ["LPTN1", "tg โลตัสพัฒนาการ"],
    ["LPTO1", "tg (SIS70) โลตัสปักธงชัย"],
    ["LPTO2", "Vivo Shop โลตัสปักธงชัย"],
    ["LPTY1", "Samsung (SPS) โลตัสพัทยาใต้"],
    ["LPYA1", "tg โลตัสพะเยา Kiosk"],
    ["LRM11", "tg โลตัสพระราม 1"],
    ["LRM32", "tg โลตัสพระราม 3 Kiosk"],
    ["LSIK1", "tg โลตัสสีคิ้ว"],
    ["LSSK1", "tg โลตัสสมุทรสงคราม Kiosk"],
    ["LSUG1", "tg โลตัสสังขะ"],
    ["LSUR3", "tg โลตัสสุราษฎร์ธานี"],
    ["LTAK1", "tg โลตัสตาก"],
    ["LTAK2", "tg โลตัสตาคลี"],
    ["LTPN2", "tg โลตัสธาตุพนม"],
    ["LTPN3", "Oppo Shop โลตัสธาตุพนม"],
    ["LTPN4", "Samsung (SPS) โลตัสธาตุพนม"],
    ["LTWN2", "tg โลตัส ติวานนท์"],
    ["LUBN2", "tg โลตัสอุบลราชธานี"],
    ["LUBN3", "H2O โลตัสอุบลราชธานี"],
    ["LUBN4", "Samsung (SPS) โลตัสอุบลราชธานี"],
    ["LUDI1", "tg โลตัสอุตรดิตถ์"],
    ["LYLA1", "tg โลตัสยะลา Kiosk"],
    ["MBKE2", "tg เดอะมอลล์บางแค ชั้น 2"],
    ["MBKE5", "Samsung (SES) เดอะมอลล์บางแค ชั้น 2"],
    ["MBKE6", "Realme Shop เดอะมอลล์บางแค"],
    ["MBKP5", "tg เดอะมอลล์บางกะปิ"],
    ["MBKP7", "Vivo Shop เดอะมอลล์บางกะปิ"],
    ["MBKP8", "tg เดอะมอลล์บางกะปิ Kiosk"],
    ["MNGW1", "tg เดอะมอลล์งามวงศ์วาน"],
    ["MNGW3", "Samsung (SES) เดอะมอลล์งามวงศ์วาน"],
    ["MNGW4", "Honor Shop เดอะมอลล์งามวงศ์วาน Kiosk"],
    ["MNSM1", "tg เดอะมอลล์นครราชสีมา"],
    ["MNSM3", "Samsung (SES) เดอะมอลล์นครราชสีมา"],
    ["MTPA3", "tg เดอะมอลล์ ท่าพระ ชั้น 3"],
    ["OAYA2", "Oppo Shop อยุธยาซิติ้พาร์ค สาขา 1"],
    ["OAYA3", "Oppo Shop อยุธยาซิติ้พาร์ค สาขา 2"],
    ["OBNA1", "tg เมกะบางนา"],
    ["OBRM2", "Samsung (SES) ทวีกิจบุรีรัมย์"],
    ["OBRM3", "Oppo Shop ทวีกิจบุรีรัมย์ สาขา 1"],
    ["OBRM4", "tg ทวีกิจบุรีรัมย์"],
    ["OBRM5", "Oppo Shop ทวีกิจบุรีรัมย์ สาขา 2"],
    ["OBSN1", "Samsung (SPS) แหลมทองบางแสน"],
    ["OCUD1", "tg เซ็นทรัลอุดรธานี"],
    ["OCUD4", "tg ตึกคอมอุดรธานี"],
    ["OCUD5", "Samsung (SPS) ตึกคอมอุดรธานี"],
    ["OESR2", "tg อิมพิเรียลเวิลด์สำโรง"],
    ["OFIL1", "tg แฟชั่นไอส์แลนด์ ชั้น 3"],
    ["OFNK2", "tg Asawann Shopping Complex หนองคาย"],
    ["OFNK3", "Huawei Shop Asawann Shopping Complex หนองคาย"],
    ["OFPK1", "tg ฟิวเจอร์พาร์ค รังสิต ชั้น 3"],
    ["OFPK3", "Samsung (SES) ฟิวเจอร์พาร์ค รังสิต"],
    ["OFPK4", "tg Service ฟิวเจอร์พาร์ครังสิต"],
    ["OFPK5", "Honor Shop ฟิวเจอร์พาร์ครังสิต Kiosk"],
    ["OFPK6", "Oppo Shop ฟิวเจอร์พาร์ครังสิต Kiosk"],
    ["OFPK7", "Vivo Shop ฟิวเจอร์พาร์ครังสิต Kiosk"],
    ["OFPK8", "Soon Box ฟิวเจอร์พาร์ครังสิต Kiosk"],
    ["OITS3", "tg ไอทีสแควร์"],
    ["OITS4", "Samsung (SPS) ไอทีสแควร์หลักสี่"],
    ["OPTY1", "tg ตึกคอมพัทยา"],
    ["OPTY3", "tg Terminal 21 พัทยา"],
    ["OSBK3", "tg ซีคอนบางแค"],
    ["OSSQ6", "tg ซีคอนสแควร์ ศรีนครินทร์ ชั้น 2"],
    ["OSSQ7", "tg ซีคอนสแควร์ ศรีนครินทร์ ชั้น G"],
    ["OSSQ8", "tg Smart Club ซีคอนสแควร์ ศรีนครินทร์"],
    ["OSTP1", "tg เสริมไทยคอมเพล็กซ์"],
    ["OSTP2", "Realme Shop เสริมไทยคอมเพล็กซ์ Kiosk"],
    ["OTMI2", "Samsung (SES) Terminal 21 โคราช"],
    ["OUCS1", "tg:สาขา service ยูไนเต็ด สีลม (ชั้น 3)"],
    ["OVHH1", "tg มาร์เก็ตวิลเลจหัวหิน Kiosk"],
    ["OVSB1", "tg มาร์เก็ตวิลเลจสุวรรณภูมิ"],
    ["OVSB2", "tg มาร์เก็ตวิลเลจสุวรรณภูมิ Kiosk"],
    ["OVSB3", "Honor Shop มาร์เก็ตวิลเลจสุวรรณภูมิ Kiosk"],
    ["PM001", "Power Mall สาขาเดอะมอลล์รามคำแหง (M3)"],
    ["PM002", "Power Mall สาขาเดอะมอลล์โคราช (M10)"],
    ["PM003", "Power Mall สาขาเดอะมอลล์งามวงศ์วาน (M6)"],
    ["PM004", "Power Mall สาขาเดอะมอลล์บางแค (M7)"],
    ["PM005", "Power Mall สาขาเดอะมอลลบางกะปิ (M8)"],
    ["PM006", "Power Mall สาขาพารากอน (M11)"],
    ["PM007", "Power Mall สาขาเอ็มโพเรี่ยม (M9)"],
    ["PM008", "Power Mall สาขาเดอะมอลล์ท่าพระ (M5)"],
    ["PRYG1", "Realme Kiosk แพชชั่น ช้อปปิ้งเดสติเนชั่น ระยอง"],
    ["PRYG2", "tg แพชชั่น ช้อปปิ้งเดสติเนชั่น ระยอง Kiosk"],
    ["RBOW1", "Realme Shop โรบินสันบ่อวิน Kiosk"],
    ["RBOW2", "tg โรบินสันบ่อวิน Kiosk"],
    ["RBRM1", "Samsung (SES) โรบินสันบุรีรัมย์"],
    ["RCCO1", "tg โรบินสันฉะเชิงเทรา Kiosk"],
    ["RCCO2", "Honor Shop โรบินสันฉะเชิงเทรา Kiosk"],
    ["RCHL1", "tg โรบินสันฉลอง"],
    ["RCYP1", "Samsung (SES) โรบินสันชัยภูมิ"],
    ["RCYP2", "tg โรบินสันชัยภูมิ Kiosk"],
    ["RCYP3", "Realme โรบินสันชัยภูมิ"],
    ["RKRI1", "tg โรบินสันกาญจนบุรี Kiosk"],
    ["RLKB1", "tg โรบินสันลาดกระบัง Kiosk"],
    ["RMDH1", "tg โรบินสันมุกดาหาร Kiosk"],
    ["RPBI1", "tg โรบินสันเพชรบุรี Kiosk"],
    ["RRET1", "Oppo Shop โรบินสันร้อยเอ็ด"],
    ["RRYG1", "tg โรบินสันบ้านฉางระยอง"],
    ["RSKN1", "Samsung (SPS) โรบินสันสกลนคร"],
    ["RSKN2", "tg โรบินสันสกลนคร"],
    ["RSPB1", "tg โรบินสันสุพรรณบุรี Kiosk"],
    ["RSPB2", "Honor Shop โรบินสันสุพรรณบุรี"],
    ["RSPK1", "tg โรบินสันสมุทรปราการ Kiosk"],
    ["RSRI1", "tg โรบินสันสระบุรี"],
    ["RSRI2", "Honor Shop โรบินสันสระบุรี Kiosk"],
    ["RSRN1", "Samsung (SES) โรบินสันสุรินทร์"],
    ["RSRN2", "Realme Shop โรบินสันสุรินทร์ Kiosk"],
    ["RSRN3", "Oppo Shop โรบินสันสุรินทร์ Kiosk"],
    ["RTRG1", "tg โรบินสันตรัง Kiosk"],
    ["SCMA1", "tg Service ฟิวเจอร์เชียงใหม่"],
    ["SHAY1", "tg Service เมกาหาดใหญ่"],
    ["SKKN1", "tg Service SY Thai ขอนแก่น"],
    ["SKOR1", "tg Service STS โคราช"],
    ["SPKT1", "tg Service SMT ภูเก็ต"],
    ["SPSL1", "tg Service ฟิวเจอร์พิษณุโลก"],
    ["SSUR1", "tg Service SMTสุราษฎร์ธานี"],
    ["SUBN2", "tg Service AIT อุบลราชธานี"],
    ["SUDO1", "tg Service ASM อุดรธานี"],
    ["CKKC1", "Samsung (SPS) เซ็นทรัลขอนแก่น แคมปัส"],
    ["LKKN4", "Samsung (SPS) โลตัสขอนแก่น Kiosk"],
    ["BNPT1", "tg บิ๊กซีนครปฐม"],
    ["BPKT1", "tg บิ๊กซีภูเก็ต"],
    ["BHDY1", "tg บิ๊กซีหาดใหญ่ Kiosk"],
    ["LBAB7", "Telewiz โลตัสบ้านบึง"],
    ["MBKE7", "Tecno Shop เดอะมอลล์บางแค Kiosk"],
    ["CNVL1", "tg เซ็นทรัลนอร์ทวิลล์"],
    ["RSRN4", "tg โรบินสันสุรินทร์ Kiosk"],
    ["LPCT2", "Telewiz โลตัสพิจิตร ท่าหลวง"],
    ["LKBI3", "tg โลตัสกระบี่ (ร้าน2)"],
    ["KMBK8", "Huawei Shop มาบุญครอง (MBK) ชั้น 4"],
    ["LUDN1", "tg โลตัสอุดรธานี"],
    ["LUDN2", "Huawei Shop โลตัสอุดรธานี Kiosk"],
    ["LCEI1", "LCEI1 tg โลตัสเชียงราย"],
];
