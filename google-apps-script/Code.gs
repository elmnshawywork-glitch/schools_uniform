/*******************************************************************
 * تسجيل الزي المدرسي — مدارس ابدأ الوطنية للعلوم التقنية (بدر ودمياط)
 * Backend: Google Apps Script Web App  ←→  فورم HTML على GitHub Pages
 *
 * الوظائف:
 *  - كل صف دراسي في تبويب (شيت) منفصل: «الصف الأول/الثاني/الثالث»
 *  - تبويب «الإجمالي»: تجميع الكميات المطلوبة لكل قطعة × مقاس
 *  - تبويب «الطلاب»: قائمة الأسماء المستوردة + حالة التسجيل (لمنع التكرار)
 *  - قراءة القائمة والتحقق من التكرار عبر JSONP (GET)
 *  - كتابة التسجيلات والاستيراد عبر POST
 *
 * النشر: Deploy ▸ New deployment ▸ Web app
 *        Execute as: Me   |   Who has access: Anyone
 *        انسخ رابط /exec وضعه في GAS_URL داخل index.html
 * مهم: بعد أي تعديل هنا، اعمل Deploy ▸ Manage deployments ▸ Edit ▸
 *      Version: New version ▸ Deploy حتى تُطبَّق التغييرات.
 *******************************************************************/

var GRADE_SHEETS = { "1":"الصف الأول", "2":"الصف الثاني", "3":"الصف الثالث" };
var SUMMARY_SHEET = "الإجمالي";
var ROSTER_SHEET  = "الطلاب";
var SCHOOL_AR = { badr:"بدر", damietta:"دمياط" };
var SIZES = ["S","M","L","XL","2XL","3XL","4XL","موحد"];

// أسماء القطع (يجب أن تطابق CATALOG في الفورم)
var PIECES = [
  "أفارول قطعتين كحلي عواكس (تطريز صدر + كم)",
  "بالطو أبيض (تطريز جيب + صدر)",
  "طقم رياضي ثلاث قطع + تي شيرت تطريز",
  "كاب تطريز",
  "بنطلون جبردين كحلي",
  "بولو شيرت ½ كم أخضر (تطريز صدر)",
  "بولو شيرت كم (تطريز صدر)",
  "بولو شيرت هودي ميلتون سبن كحلي (تطريز صدر + كم)"
];

/* ===================== نقاط الدخول ===================== */
function doGet(e){
  var p = (e && e.parameter) || {};
  var cb = p.callback, action = p.action, out;
  try{
    if(action === "roster"){ out = { ok:true, students: getRoster_(p.school, p.grade) }; }
    else if(action === "check"){ out = { ok:true, registered: isRegistered_(p.school, p.grade, p.code, p.name) }; }
    else { out = { ok:true, service:"NTSS uniform endpoint" }; }
  }catch(err){ out = { ok:false, error:String(err) }; }
  return reply_(out, cb);
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);

    if(data.type === "import"){
      var n = importStudents_(data.students || []);
      return reply_({ ok:true, imported:n }, null);
    }

    // تسجيل طالب — منع التكرار
    if(isRegistered_(data.school, data.grade, data.code, data.name)){
      return reply_({ ok:false, duplicate:true }, null);
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = GRADE_SHEETS[String(data.grade)] || ("صف " + data.grade);
    var sh = ensureGradeSheet_(ss, sheetName);
    appendSubmission_(sh, data);
    markRegistered_(data.school, data.grade, data.code, data.name);
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

/* ===================== أدوات تطبيع ===================== */
function normGrade_(g){
  g = String(g||"").trim();
  if(g.indexOf("أول")>=0 || g.indexOf("اول")>=0 || g==="1") return "1";
  if(g.indexOf("ثاني")>=0 || g.indexOf("ثان")>=0 || g==="2") return "2";
  if(g.indexOf("ثالث")>=0 || g==="3") return "3";
  return g;
}
function normSchool_(s){
  s = String(s||"").trim();
  if(s==="badr" || s.indexOf("بدر")>=0) return "badr";
  if(s==="damietta" || s.indexOf("دمياط")>=0) return "damietta";
  return s;
}
function identity_(school, grade, code, name){
  school = normSchool_(school); grade = normGrade_(grade);
  var id = (code && String(code).trim()) ? ("c:"+String(code).trim()) : ("n:"+String(name||"").trim());
  return school + "|" + grade + "|" + id;
}

/* ===================== شيتات الصفوف ===================== */
function ensureGradeSheet_(ss, name){
  var sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); }
  if(sh.getLastRow() === 0){
    var header = ["م","التاريخ","المدرسة","اسم الطالب","كود الطالب","الفصل","ولي الأمر","الموبايل"]
                 .concat(PIECES).concat(["عدد القطع","الإجمالي (ج.م)"]);
    sh.getRange(1,1,1,header.length).setValues([header])
      .setFontWeight("bold").setBackground("#143653").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function appendSubmission_(sh, data){
  var seq = Math.max(0, sh.getLastRow()-1) + 1;
  var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Africa/Cairo", "yyyy-MM-dd HH:mm");
  var bySize = {};
  (data.items||[]).forEach(function(it){ bySize[it.name] = it.size; });
  var row = [
    seq, when, SCHOOL_AR[normSchool_(data.school)] || data.school,
    data.name || "", data.code || "", data.section || "", data.guardian || "", "'" + (data.phone||"")
  ];
  PIECES.forEach(function(pn){ row.push(bySize[pn] || ""); });
  row.push((data.items||[]).length);
  row.push(normGrade_(data.grade)==="1" ? 0 : (data.total||0));
  sh.appendRow(row);
}

/* ===================== شيت الإجمالي ===================== */
function rebuildSummary_(ss){
  var sh = ss.getSheetByName(SUMMARY_SHEET);
  if(!sh){ sh = ss.insertSheet(SUMMARY_SHEET); }
  sh.clear();
  sh.setRightToLeft(true);

  // counts[piece][size] , paid[piece][size]
  var counts = {}, paid = {}, prices = {};
  PIECES.forEach(function(p){ counts[p] = {}; paid[p] = {}; });

  Object.keys(GRADE_SHEETS).forEach(function(g){
    var gs = ss.getSheetByName(GRADE_SHEETS[g]);
    if(!gs || gs.getLastRow() < 2) return;
    var values = gs.getDataRange().getValues();
    var head = values[0];
    var pieceCol = {};
    PIECES.forEach(function(p){ pieceCol[p] = head.indexOf(p); });
    var totalCol = head.indexOf("الإجمالي (ج.م)");
    for(var r=1; r<values.length; r++){
      var isFree = (g === "1");
      PIECES.forEach(function(p){
        var c = pieceCol[p];
        if(c < 0) return;
        var size = values[r][c];
        if(size === "" || size == null) return;
        counts[p][size] = (counts[p][size]||0) + 1;
        if(!isFree) paid[p][size] = (paid[p][size]||0) + 1;
      });
    }
  });

  var header = ["القطعة"].concat(SIZES).concat(["إجمالي الكمية","المدفوع (صف2+3)","سعر القطعة","القيمة التقديرية (ج.م)"]);
  var out = [header];
  var PRICE = {
    "أفارول قطعتين كحلي عواكس (تطريز صدر + كم)":1045,
    "بالطو أبيض (تطريز جيب + صدر)":495,
    "طقم رياضي ثلاث قطع + تي شيرت تطريز":1100,
    "كاب تطريز":165,
    "بنطلون جبردين كحلي":550,
    "بولو شيرت ½ كم أخضر (تطريز صدر)":440,
    "بولو شيرت كم (تطريز صدر)":495,
    "بولو شيرت هودي ميلتون سبن كحلي (تطريز صدر + كم)":660
  };
  var grandQty = 0, grandPaidQty = 0, grandValue = 0;
  PIECES.forEach(function(p){
    var row = [p], totalQty = 0, paidQty = 0;
    SIZES.forEach(function(s){ var v = counts[p][s]||0; row.push(v||""); totalQty += v; });
    SIZES.forEach(function(s){ paidQty += (paid[p][s]||0); });
    var price = PRICE[p] || 0;
    var value = paidQty * price;
    row.push(totalQty, paidQty, price, value);
    out.push(row);
    grandQty += totalQty; grandPaidQty += paidQty; grandValue += value;
  });
  var footer = ["الإجمالي"];
  for(var i=0;i<SIZES.length;i++) footer.push("");
  footer.push(grandQty, grandPaidQty, "", grandValue);
  out.push(footer);

  sh.getRange(1,1,out.length,header.length).setValues(out);
  sh.getRange(1,1,1,header.length).setFontWeight("bold").setBackground("#12938B").setFontColor("#ffffff");
  sh.getRange(out.length,1,1,header.length).setFontWeight("bold").setBackground("#E7F4F2");
  sh.setFrozenRows(1);
  sh.autoResizeColumn(1);
}

/* ===================== شيت الطلاب (الروستر) ===================== */
function ensureRosterSheet_(ss){
  var sh = ss.getSheetByName(ROSTER_SHEET);
  if(!sh){ sh = ss.insertSheet(ROSTER_SHEET); }
  if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,5).setValues([["المدرسة","الصف","كود الطالب","اسم الطالب","الحالة"]])
      .setFontWeight("bold").setBackground("#2E6E8E").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function importStudents_(students){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureRosterSheet_(ss);
  var data = sh.getDataRange().getValues();
  var existing = {};
  for(var r=1; r<data.length; r++){
    existing[ identity_(data[r][0], data[r][1], data[r][2], data[r][3]) ] = true;
  }
  var toAppend = [], added = 0;
  students.forEach(function(st){
    var school = normSchool_(st.school), grade = normGrade_(st.grade);
    var code = (st.code||"").toString().trim(), name = (st.name||"").toString().trim();
    if(!name) return;
    var id = identity_(school, grade, code, name);
    if(existing[id]) return;             // تجاهل المكرر
    toAppend.push([ SCHOOL_AR[school]||school, grade, code, name, "" ]);
    existing[id] = true; added++;
  });
  if(toAppend.length){
    sh.getRange(sh.getLastRow()+1, 1, toAppend.length, 5).setValues(toAppend);
  }
  return added;
}

function getRoster_(school, grade){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ROSTER_SHEET);
  if(!sh || sh.getLastRow() < 2) return [];
  school = normSchool_(school); grade = normGrade_(grade);
  var data = sh.getDataRange().getValues();
  var out = [];
  for(var r=1; r<data.length; r++){
    var rs = normSchool_(data[r][0]), rg = normGrade_(data[r][1]);
    if((school && rs !== school) || (grade && rg !== grade)) continue;
    out.push({ code:String(data[r][2]||""), name:String(data[r][3]||""), registered: String(data[r][4]||"").trim() !== "" });
  }
  return out;
}

function isRegistered_(school, grade, code, name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ROSTER_SHEET);
  var target = identity_(school, grade, code, name);
  if(sh && sh.getLastRow() >= 2){
    var data = sh.getDataRange().getValues();
    for(var r=1; r<data.length; r++){
      if(identity_(data[r][0], data[r][1], data[r][2], data[r][3]) === target){
        return String(data[r][4]||"").trim() !== "";
      }
    }
  }
  return false; // غير موجود بالقائمة ⇒ مسموح (طالب مكتوب يدويًا)
}

function markRegistered_(school, grade, code, name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureRosterSheet_(ss);
  var target = identity_(school, grade, code, name);
  var data = sh.getDataRange().getValues();
  for(var r=1; r<data.length; r++){
    if(identity_(data[r][0], data[r][1], data[r][2], data[r][3]) === target){
      sh.getRange(r+1, 5).setValue("مُسجّل");
      return;
    }
  }
  // طالب مكتوب يدويًا (غير موجود بالقائمة): أضِفه كمُسجّل لمنع تكراره لاحقًا
  sh.appendRow([ SCHOOL_AR[normSchool_(school)]||school, normGrade_(grade), code||"", name||"", "مُسجّل" ]);
}

/* ===================== اختبار سريع (اختياري) ===================== */
function _selftest(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  importStudents_([{school:"badr",grade:"2",code:"1001",name:"طالب تجريبي"}]);
  Logger.log("roster badr/2: " + JSON.stringify(getRoster_("badr","2")));
  ensureGradeSheet_(ss, GRADE_SHEETS["2"]);
  appendSubmission_(ss.getSheetByName(GRADE_SHEETS["2"]), {
    school:"badr", grade:"2", name:"طالب تجريبي", code:"1001", phone:"01000000000",
    items:[{name:"كاب تطريز",size:"موحد"},{name:"بنطلون جبردين كحلي",size:"L"}], total:650
  });
  markRegistered_("badr","2","1001","طالب تجريبي");
  rebuildSummary_(ss);
  Logger.log("registered? " + isRegistered_("badr","2","1001","طالب تجريبي"));
}
