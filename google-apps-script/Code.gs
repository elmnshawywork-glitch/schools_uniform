/*******************************************************************
 * تسجيل الزي المدرسي — مدارس ابدأ الوطنية للعلوم التقنية (بدر ودمياط)
 * Backend: Google Apps Script Web App  ←→  فورم HTML على GitHub Pages
 *
 * التبويبات في الشيت:
 *  - «الصف الأول / الثاني / الثالث» : تسجيلات كل صف
 *  - «الإجمالي» : تجميع الكميات لكل قطعة × مقاس + القيمة التقديرية
 *  - «الطلاب»  : قائمة الأسماء المستوردة + حالة التسجيل (منع التكرار)
 *  - «الإعدادات» : إعدادات النموذج (مدارس/قطع/أسعار/مستخدمون) كـ JSON
 *
 * النشر: Deploy ▸ New deployment ▸ Web app
 *        Execute as: Me   |   Who has access: Anyone
 *        انسخ رابط /exec وضعه في GAS_URL داخل index.html
 * مهم: بعد أي تعديل هنا اعمل Deploy ▸ Manage deployments ▸ Edit ▸
 *      Version: New version ▸ Deploy.
 *******************************************************************/

var GRADE_SHEETS = { "1":"الصف الأول", "2":"الصف الثاني", "3":"الصف الثالث" };
var SUMMARY_SHEET = "الإجمالي";
var ROSTER_SHEET  = "الطلاب";
var CONFIG_SHEET  = "الإعدادات";
var RECEIPTS_SHEET = "الإيصالات";
var SECURITY_EMAIL = "elmnshawy.work@gmail.com"; // بريد الآمن (يمكن تجاوزه من الإعدادات)
var SIZES = ["S","M","L","XL","2XL","3XL","4XL",
             "36","37","38","39","40","41","42","43","44","45","46","موحد"];

/* الإعدادات الافتراضية (تطابق DEFAULT_CONFIG في الفورم) */
var DEFAULT_CONFIG = {
  academicYear: "2026 / 2027",
  schools: [
    { key:"badr",     name:"مدرسة ابدأ الوطنية للعلوم التقنية", branch:"بدر" },
    { key:"damietta", name:"مدرسة ابدأ الوطنية للعلوم التقنية", branch:"دمياط" }
  ],
  pieces: [
    { key:"afarol",    name:"أفارول قطعتين عواكس (تطريز صدر + كم)",             price:1045, size:"letter" },
    { key:"balto",     name:"بالطو أبيض (تطريز جيب + صدر)",                    price:495,  size:"letter" },
    { key:"sport",     name:"طقم رياضي ثلاث قطع + تي شيرت تطريز",               price:1100, size:"letter" },
    { key:"cap",       name:"كاب تطريز",                                       price:165,  size:"one" },
    { key:"pants",     name:"بنطلون جبردين بيج باجي بجيوب",                     price:550,  size:"letter" },
    { key:"polo_half", name:"بولو شيرت ½ كم أخضر (تطريز صدر)",                  price:440,  size:"letter" },
    { key:"polo_full", name:"بولو شيرت كم أخضر (تطريز صدر)",                    price:495,  size:"letter" },
    { key:"hoodie",    name:"بولو شيرت هودي ميلتون سبن أخضر (تطريز صدر + كم)",   price:660,  size:"letter" },
    { key:"safety_vest",   name:"فيست أمان (2 لون) بجيوب",                     price:0, size:"letter", grade1Only:true },
    { key:"safety_shoes",  name:"حذاء سيفتي",                                  price:0, size:"shoe",   grade1Only:true },
    { key:"safety_helmet", name:"خوذة سيفتي V-Gard (أصفر - طالب)",             price:0, size:"one",    grade1Only:true },
    { key:"safety_gloves", name:"قفازات مقاومة للقطع (رمادي)",                 price:0, size:"letter", grade1Only:true }
  ],
  users: [ { username:"admin", password:"2027" } ],
  inventory: {},
  supervisorHelmets: {},
  cashRecipients: [],
  securityEmail: "elmnshawy.work@gmail.com"
};

/* ===================== نقاط الدخول ===================== */
function doGet(e){
  var p = (e && e.parameter) || {};
  var cb = p.callback, action = p.action, out;
  try{
    if(action === "roster"){ out = { ok:true, students: getRoster_(p.school, p.grade, p.year) }; }
    else if(action === "check"){ out = { ok:true, registered: isRegistered_(p.school, p.grade, p.code, p.name, p.year) }; }
    else if(action === "config"){ out = { ok:true, config: getConfig_() }; }
    else if(action === "receipts"){ out = { ok:true, receipts: getReceipts_() }; }
    else { out = { ok:true, service:"NTSS uniform endpoint" }; }
  }catch(err){ out = { ok:false, error:String(err) }; }
  return reply_(out, cb);
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);

    if(data.type === "saveConfig"){ saveConfig_(data.config); return reply_({ ok:true }, null); }
    if(data.type === "import"){ var n = importStudents_(data.students || []); return reply_({ ok:true, imported:n }, null); }
    if(data.type === "deliver"){ markDelivery_(data); rebuildDeliveries_(SpreadsheetApp.getActiveSpreadsheet()); return reply_({ ok:true }, null); }
    if(data.type === "resetYear"){ resetYearData_(data.year || ""); return reply_({ ok:true }, null); }
    if(data.type === "receipt"){ var rs = recordReceipt_(data); return reply_({ ok:true, serial: rs }, null); }

    // تسجيل طالب — منع التكرار (لنفس العام الدراسي)
    var year = data.year || getConfig_().academicYear;
    if(isRegistered_(data.school, data.grade, data.code, data.name, year)){ return reply_({ ok:false, duplicate:true }, null); }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = GRADE_SHEETS[String(data.grade)] || ("صف " + data.grade);
    var sh = ensureGradeSheet_(ss, sheetName);
    appendSubmission_(sh, data);
    markRegistered_(data.school, data.grade, data.code, data.name, year);
    rebuildSummary_(ss);
    return reply_({ ok:true }, null);

  }catch(err){
    return reply_({ ok:false, error:String(err) }, null);
  }finally{
    try{ lock.releaseLock(); }catch(e2){}
  }
}

function reply_(obj, cb){
  var t = JSON.stringify(obj);
  if(cb){ return ContentService.createTextOutput(cb + "(" + t + ")").setMimeType(ContentService.MimeType.JAVASCRIPT); }
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== الإعدادات ===================== */
var _cfgCache = null;
function getConfig_(){
  if(_cfgCache) return _cfgCache;
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG_SHEET);
    if(sh){
      var v = sh.getRange(1,1).getValue();
      if(v){ _cfgCache = JSON.parse(v); return _cfgCache; }
    }
  }catch(e){}
  _cfgCache = DEFAULT_CONFIG;
  return _cfgCache;
}
function saveConfig_(cfg){
  if(!cfg) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  sh.getRange(1,1).setValue(JSON.stringify(cfg));
  sh.getRange(3,1).setValue("لا تعدّل هذا التبويب يدويًا — يُدار من تبويب الإعدادات في النموذج.");
  _cfgCache = cfg;
}
function pieceNames_(){ return getConfig_().pieces.map(function(p){ return p.name; }); }
function priceMap_(){ var m={}; getConfig_().pieces.forEach(function(p){ m[p.name]=Number(p.price)||0; }); return m; }
function schoolDisplay_(key){
  var c=getConfig_();
  for(var i=0;i<c.schools.length;i++){ if(c.schools[i].key===key) return c.schools[i].branch || c.schools[i].name; }
  return key;
}
function normSchool_(s){
  s = String(s||"").trim();
  var c=getConfig_();
  for(var i=0;i<c.schools.length;i++){ var sc=c.schools[i]; if(s===sc.key || s===sc.name || s===sc.branch) return sc.key; }
  return s;
}
function normGrade_(g){
  g = String(g||"").trim();
  if(g.indexOf("أول")>=0 || g.indexOf("اول")>=0 || g==="1") return "1";
  if(g.indexOf("ثاني")>=0 || g.indexOf("ثان")>=0 || g==="2") return "2";
  if(g.indexOf("ثالث")>=0 || g==="3") return "3";
  return g;
}
function identity_(school, grade, code, name, year){
  school = normSchool_(school); grade = normGrade_(grade);
  year = String(year||"").trim();
  var id = (code && String(code).trim()) ? ("c:"+String(code).trim()) : ("n:"+String(name||"").trim());
  return school + "|" + grade + "|" + year + "|" + id;
}

/* ===================== شيتات الصفوف ===================== */
function ensureGradeSheet_(ss, name){
  var sh = ss.getSheetByName(name);
  var pieces = pieceNames_();
  if(!sh){ sh = ss.insertSheet(name); }
  if(sh.getLastRow() === 0){
    var header = ["م","العام الدراسي","التاريخ","المدرسة","اسم الطالب","كود الطالب","الفصل","ولي الأمر","الموبايل"]
                 .concat(pieces).concat(["عدد القطع","الإجمالي (ج.م)"]);
    sh.getRange(1,1,1,header.length).setValues([header])
      .setFontWeight("bold").setBackground("#143653").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}
function appendSubmission_(sh, data){
  var pieces = pieceNames_();
  var head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var seq = Math.max(0, sh.getLastRow()-1) + 1;
  var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Africa/Cairo", "yyyy-MM-dd HH:mm");
  var bySize = {};
  (data.items||[]).forEach(function(it){ bySize[it.name] = it.size; });
  var year = data.year || getConfig_().academicYear;
  var row = [];
  for(var c=0;c<head.length;c++){
    var col = head[c];
    if(col==="م") row.push(seq);
    else if(col==="العام الدراسي") row.push(year);
    else if(col==="التاريخ") row.push(when);
    else if(col==="المدرسة") row.push(schoolDisplay_(normSchool_(data.school)));
    else if(col==="اسم الطالب") row.push(data.name||"");
    else if(col==="كود الطالب") row.push(data.code||"");
    else if(col==="الفصل") row.push(data.section||"");
    else if(col==="ولي الأمر") row.push(data.guardian||"");
    else if(col==="الموبايل") row.push("'"+(data.phone||""));
    else if(col==="عدد القطع") row.push((data.items||[]).length);
    else if(col==="الإجمالي (ج.م)") row.push(normGrade_(data.grade)==="1" ? 0 : (data.total||0));
    else row.push(bySize[col] || "");  // أعمدة القطع
  }
  sh.appendRow(row);
}

/* ===================== شيت الإجمالي ===================== */
function rebuildSummary_(ss){
  var sh = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  sh.clear(); sh.setRightToLeft(true);
  var PIECES = pieceNames_();
  var PRICE = priceMap_();
  var counts = {}, paid = {};
  PIECES.forEach(function(p){ counts[p]={}; paid[p]={}; });

  Object.keys(GRADE_SHEETS).forEach(function(g){
    var gs = ss.getSheetByName(GRADE_SHEETS[g]);
    if(!gs || gs.getLastRow() < 2) return;
    var values = gs.getDataRange().getValues();
    var head = values[0], pieceCol = {};
    var yCol = head.indexOf("العام الدراسي");
    var curYear = getConfig_().academicYear;
    PIECES.forEach(function(p){ pieceCol[p] = head.indexOf(p); });
    for(var r=1; r<values.length; r++){
      if(yCol >= 0){ var ry = String(values[r][yCol]||"").trim(); if(ry && ry !== curYear) continue; }
      var isFree = (g === "1");
      PIECES.forEach(function(p){
        var c = pieceCol[p]; if(c < 0) return;
        var size = values[r][c];
        if(size === "" || size == null) return;
        counts[p][size] = (counts[p][size]||0) + 1;
        if(!isFree) paid[p][size] = (paid[p][size]||0) + 1;
      });
    }
  });

  var header = ["القطعة"].concat(SIZES).concat(["إجمالي الكمية","المدفوع (صف2+3)","سعر القطعة","القيمة التقديرية (ج.م)"]);
  var out = [header], grandQty=0, grandPaid=0, grandVal=0;
  PIECES.forEach(function(p){
    var row=[p], totalQty=0, paidQty=0;
    SIZES.forEach(function(s){ var v=counts[p][s]||0; row.push(v||""); totalQty+=v; });
    SIZES.forEach(function(s){ paidQty += (paid[p][s]||0); });
    var price = PRICE[p]||0, value = paidQty*price;
    row.push(totalQty, paidQty, price, value);
    out.push(row); grandQty+=totalQty; grandPaid+=paidQty; grandVal+=value;
  });
  var footer=["الإجمالي"]; for(var i=0;i<SIZES.length;i++) footer.push("");
  footer.push(grandQty, grandPaid, "", grandVal); out.push(footer);

  sh.getRange(1,1,out.length,header.length).setValues(out);
  sh.getRange(1,1,1,header.length).setFontWeight("bold").setBackground("#12938B").setFontColor("#ffffff");
  sh.getRange(out.length,1,1,header.length).setFontWeight("bold").setBackground("#E7F4F2");
  sh.setFrozenRows(1);
}

/* ===================== شيت الطلاب (الروستر) ===================== */
function ensureRosterSheet_(ss){
  var sh = ss.getSheetByName(ROSTER_SHEET);
  if(!sh){ sh = ss.insertSheet(ROSTER_SHEET); }
  if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,6).setValues([["المدرسة","الصف","كود الطالب","اسم الطالب","الحالة","العام الدراسي"]])
      .setFontWeight("bold").setBackground("#2E6E8E").setFontColor("#ffffff");
    sh.setFrozenRows(1); sh.setRightToLeft(true);
  }
  return sh;
}
function importStudents_(students){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureRosterSheet_(ss);
  var data = sh.getDataRange().getValues(), existing = {};
  for(var r=1; r<data.length; r++){ existing[ identity_(data[r][0], data[r][1], data[r][2], data[r][3], data[r][5]) ] = true; }
  var toAppend=[], added=0;
  var defaultYear = getConfig_().academicYear;
  students.forEach(function(st){
    var school=normSchool_(st.school), grade=normGrade_(st.grade), year=String(st.year||defaultYear).trim();
    var code=(st.code||"").toString().trim(), name=(st.name||"").toString().trim();
    if(!name) return;
    var id=identity_(school, grade, code, name, year);
    if(existing[id]) return;
    toAppend.push([ schoolDisplay_(school), grade, code, name, "", year ]);
    existing[id]=true; added++;
  });
  if(toAppend.length){ sh.getRange(sh.getLastRow()+1,1,toAppend.length,6).setValues(toAppend); }
  return added;
}
function getRoster_(school, grade, year){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ROSTER_SHEET);
  if(!sh || sh.getLastRow() < 2) return [];
  school = normSchool_(school); grade = normGrade_(grade); year = String(year||"").trim();
  var data = sh.getDataRange().getValues(), out = [];
  for(var r=1; r<data.length; r++){
    var rs = normSchool_(data[r][0]), rg = normGrade_(data[r][1]), ry = String(data[r][5]||"").trim();
    if((school && rs !== school) || (grade && rg !== grade)) continue;
    if(year && ry && ry !== year) continue;
    out.push({ code:String(data[r][2]||""), name:String(data[r][3]||""), registered: String(data[r][4]||"").trim() !== "" });
  }
  return out;
}
function isRegistered_(school, grade, code, name, year){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ROSTER_SHEET);
  var target = identity_(school, grade, code, name, year);
  if(sh && sh.getLastRow() >= 2){
    var data = sh.getDataRange().getValues();
    for(var r=1; r<data.length; r++){
      if(identity_(data[r][0], data[r][1], data[r][2], data[r][3], data[r][5]) === target){
        return String(data[r][4]||"").trim() !== "";
      }
    }
  }
  return false;
}
function markRegistered_(school, grade, code, name, year){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureRosterSheet_(ss);
  var target = identity_(school, grade, code, name, year);
  var data = sh.getDataRange().getValues();
  for(var r=1; r<data.length; r++){
    if(identity_(data[r][0], data[r][1], data[r][2], data[r][3], data[r][5]) === target){
      sh.getRange(r+1, 5).setValue("مُسجّل"); return;
    }
  }
  sh.appendRow([ schoolDisplay_(normSchool_(school)), normGrade_(grade), code||"", name||"", "مُسجّل", String(year||"") ]);
}

/* ===================== مسح بيانات عام دراسي ===================== */
function resetYearData_(year){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  year = String(year||"").trim();

  // 1) حذف صفوف التسجيلات المطابقة من شيتات الصفوف
  Object.keys(GRADE_SHEETS).forEach(function(g){
    var sh = ss.getSheetByName(GRADE_SHEETS[g]);
    if(!sh || sh.getLastRow() < 2) return;
    var values = sh.getDataRange().getValues(), head = values[0];
    var yCol = head.indexOf("العام الدراسي");
    if(yCol < 0) return;
    for(var r=values.length-1; r>=1; r--){
      var ry = String(values[r][yCol]||"").trim();
      if(ry === year){ sh.deleteRow(r+1); }
    }
  });

  // 2) حذف صفوف الطلاب (الروستر) المطابقة لهذا العام
  var rsh = ss.getSheetByName(ROSTER_SHEET);
  if(rsh && rsh.getLastRow() >= 2){
    var rdata = rsh.getDataRange().getValues();
    for(var rr=rdata.length-1; rr>=1; rr--){
      var ry2 = String(rdata[rr][5]||"").trim();
      if(ry2 === year){ rsh.deleteRow(rr+1); }
    }
  }

  // 3) إعادة بناء شيتات الإجمالي والتسليمات
  rebuildSummary_(ss);
  rebuildDeliveries_(ss);
}

/* ===================== التسليمات ===================== */
var DELIV_SHEET = "التسليمات";
function stripCheck_(v){ return String(v==null?"":v).replace(/\s*✓\s*$/,"").trim(); }

// data: { grade, name, code, delivered:{ "اسم القطعة":true, ... } }
function markDelivery_(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = GRADE_SHEETS[String(normGrade_(data.grade))]; if(!name) return;
  var sh = ss.getSheetByName(name); if(!sh || sh.getLastRow() < 2) return;
  var values = sh.getDataRange().getValues(), head = values[0];
  var nameCol = head.indexOf("اسم الطالب"), codeCol = head.indexOf("كود الطالب"), yCol = head.indexOf("العام الدراسي");
  var PIECES = pieceNames_(), pieceCol = {};
  PIECES.forEach(function(p){ pieceCol[p] = head.indexOf(p); });
  var delivered = data.delivered || {};
  var wantCode = String(data.code||"").trim(), wantName = String(data.name||"").trim(), wantYear = String(data.year||"").trim();
  for(var r=1; r<values.length; r++){
    if(wantYear && yCol>=0){ var ry = String(values[r][yCol]||"").trim(); if(ry && ry !== wantYear) continue; }
    var rowName = String(values[r][nameCol]||"").trim();
    var rowCode = codeCol>=0 ? String(values[r][codeCol]||"").trim() : "";
    var match = wantCode ? (rowCode === wantCode) : (rowName === wantName);
    if(!match) continue;
    PIECES.forEach(function(p){
      var c = pieceCol[p]; if(c < 0) return;
      var base = stripCheck_(values[r][c]);
      if(base === "") return;                       // لم تُطلب هذه القطعة
      var isDel = delivered[p] === true || delivered[p] === "true";
      sh.getRange(r+1, c+1).setValue(isDel ? (base + " ✓") : base);
    });
    break;
  }
}

function rebuildDeliveries_(ss){
  var sh = ss.getSheetByName(DELIV_SHEET) || ss.insertSheet(DELIV_SHEET);
  sh.clear(); sh.setRightToLeft(true);
  var PIECES = pieceNames_(), ordered = {}, delivered = {};
  PIECES.forEach(function(p){ ordered[p]=0; delivered[p]=0; });
  Object.keys(GRADE_SHEETS).forEach(function(g){
    var gs = ss.getSheetByName(GRADE_SHEETS[g]);
    if(!gs || gs.getLastRow() < 2) return;
    var values = gs.getDataRange().getValues(), head = values[0], pc = {};
    var yCol = head.indexOf("العام الدراسي");
    var curYear = getConfig_().academicYear;
    PIECES.forEach(function(p){ pc[p] = head.indexOf(p); });
    for(var r=1; r<values.length; r++){
      if(yCol >= 0){ var ry = String(values[r][yCol]||"").trim(); if(ry && ry !== curYear) continue; }
      PIECES.forEach(function(p){
        var c = pc[p]; if(c < 0) return;
        var cell = String(values[r][c]||"");
        if(stripCheck_(cell) === "") return;
        ordered[p]++;
        if(cell.indexOf("✓") >= 0) delivered[p]++;
      });
    }
  });
  var out = [["القطعة","مطلوب","مُسلَّم","متبقٍ","النسبة %"]];
  var gO=0,gD=0;
  PIECES.forEach(function(p){
    var o=ordered[p], d=delivered[p], rem=o-d, pct=o?Math.round(d/o*100):0;
    out.push([p,o,d,rem,pct]); gO+=o; gD+=d;
  });
  out.push(["الإجمالي", gO, gD, gO-gD, gO?Math.round(gD/gO*100):0]);
  sh.getRange(1,1,out.length,5).setValues(out);
  sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#0C6E68").setFontColor("#ffffff");
  sh.getRange(out.length,1,1,5).setFontWeight("bold").setBackground("#E7F4F2");
  sh.setFrozenRows(1);
}

/* ===================== إيصالات القيمة المادية ===================== */
function ensureReceiptsSheet_(ss){
  var sh = ss.getSheetByName(RECEIPTS_SHEET);
  if(!sh){ sh = ss.insertSheet(RECEIPTS_SHEET); }
  if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,9).setValues([["رقم الإيصال","العام الدراسي","التاريخ","المدرسة","الصف","اسم الطالب","كود الطالب","القيمة (ج.م)","أصدره"]])
      .setFontWeight("bold").setBackground("#0C6E68").setFontColor("#ffffff");
    sh.setFrozenRows(1); sh.setRightToLeft(true);
  }
  return sh;
}
function padSerial_(n){ n = String(n); while(n.length < 4) n = "0" + n; return n; }
/* يسجّل الإيصال في الشيت ويُرسل إشعارًا لبريد الآمن. يحترم الرقم القادم من الواجهة إن وُجد، وإلا يحسبه. */
function recordReceipt_(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureReceiptsSheet_(ss);
  var serial = String(data.serial || "").trim();
  if(!serial){
    var last = Math.max(0, sh.getLastRow() - 1); // عدد الإيصالات الحالية
    serial = padSerial_(last + 1);
  }
  var gradeName = GRADE_SHEETS[String(normGrade_(data.grade))] || String(data.grade||"");
  var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Africa/Cairo", "yyyy-MM-dd HH:mm");
  var schoolName = data.schoolName || schoolDisplay_(normSchool_(data.school));
  var year = data.year || getConfig_().academicYear;
  sh.appendRow([ serial, year, when, schoolName, gradeName, data.name||"", data.code||"", Number(data.amount)||0, data.issuedBy||"" ]);

  // إشعار بريدي فوري للآمن
  try{
    var to = (data.securityEmail || (getConfig_().securityEmail) || SECURITY_EMAIL || "").trim();
    if(to){
      var subject = "إيصال قيمة مادية رقم " + serial + " - " + (data.name||"");
      var body =
        "تم إصدار إيصال قيمة مادية:\n\n" +
        "رقم الإيصال: " + serial + "\n" +
        "الاسم: " + (data.name||"") + "\n" +
        "المدرسة: " + schoolName + "\n" +
        "الصف: " + gradeName + "\n" +
        "كود الطالب: " + (data.code||"-") + "\n" +
        "القيمة: " + (Number(data.amount)||0) + " ج.م\n" +
        "العام الدراسي: " + year + "\n" +
        "التاريخ: " + when + "\n" +
        "أصدره: " + (data.issuedBy||"-");
      MailApp.sendEmail(to, subject, body);
    }
  }catch(mailErr){ /* تجاهل خطأ الإرسال حتى لا يفشل التسجيل */ }

  return serial;
}
function getReceipts_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RECEIPTS_SHEET);
  if(!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues(), out = [];
  for(var r=1; r<data.length; r++){
    if(!data[r][0]) continue;
    out.push({
      serial: String(data[r][0]),
      year: String(data[r][1]||""),
      schoolName: String(data[r][3]||""),
      grade: normGrade_(data[r][4]),
      name: String(data[r][5]||""),
      code: String(data[r][6]||""),
      amount: Number(data[r][7])||0,
      issuedBy: String(data[r][8]||"")
    });
  }
  return out;
}
