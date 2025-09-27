/**
 * Google Apps Script backend (Web App) لبوابة الموظف
 *
 * ملاحظات الإعداد:
 * 1) أنشئ Google Sheet مع ورقتين:
 *    - "Employees" للأعمدة الأساسية (مثال بالأسفل).
 *    - "Attendance" لسجلات الحضور والانصراف (الترويسات: Timestamp, ID, Type, Note).
 * 2) حدّث أسماء الأعمدة/المؤشرات أدناه لتطابق شيتك.
 * 3) أنشئ مجلد Drive لحفظ ملفات السيرة الذاتية وضع ID المجلد في CV_FOLDER_ID.
 * 4) انشري السكربت كويب آب:
 *    Publish > Deploy as web app (أو New deployment) > access: Anyone with the link.
 * 5) ضعي رابط الويب آب في index.html داخل الثابت API_URL.
 */

const SHEET_EMP = "Employees";
const SHEET_ATT = "Attendance";

// غيّري هذه المعرفات لتطابق الأعمدة في ورقة Employees لديك
const COLS = {
  id_number: "National ID",      // رقم الهوية
  password:  "Password",         // كلمة المرور
  name:      "Full Name",        // الاسم
  email:     "Email",            // البريد
  phone:     "Phone",            // الهاتف
  department:"Department",       // جهة التكليف/القسم
  leave:     "Leave Balance",    // رصيد الإجازة
  degree:    "Latest Degree",    // آخر مؤهل
  address:   "National Address", // العنوان الوطني
  cvFileId:  "CV File Id"        // يخزّن فيه ملف السيرة (Drive File Id)
};

// مجلد حفظ السير الذاتية
const CV_FOLDER_ID = "PUT_DRIVE_FOLDER_ID_HERE";

// تفعيل CORS
function _cors(e){
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers":"Content-Type",
    "Access-Control-Allow-Methods":"POST, OPTIONS"
  };
}

function doOptions(e){
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders(_cors(e));
}

function doPost(e){
  try{
    const headers = _cors(e);
    const data = JSON.parse(e.postData.contents || "{}");
    const action = data.action;

    let out = {};
    if(action === "login"){
      out = handleLogin(data);
    }else if(action === "updateCV"){
      out = handleUpdateCV(data);
    }else if(action === "attendance"){
      out = handleAttendance(data);
    }else{
      out = { ok:false, message:"Unknown action" };
    }

    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders(headers);
  }catch(err){
    const headers = _cors(e);
    return ContentService.createTextOutput(JSON.stringify({ ok:false, message:String(err) }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders(headers);
  }
}

function handleLogin(data){
  const id = String(data.id_number || "").trim();
  const pw = String(data.password || "").trim();
  if(!id || !pw){ return { ok:false, message:"required" }; }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMP);
  const rng = sh.getDataRange();
  const values = rng.getValues();

  // map headers
  const header = values[0];
  const idx = {};
  header.forEach((h, i) => { idx[h] = i; });

  const idCol = idx[COLS.id_number];
  const pwCol = idx[COLS.password];

  for(let r = 1; r < values.length; r++){
    const row = values[r];
    if(String(row[idCol]) === id && String(row[pwCol]) === pw){
      const profile = {
        id_number: id,
        name:      row[idx[COLS.name]] || "",
        email:     row[idx[COLS.email]] || "",
        phone:     row[idx[COLS.phone]] || "",
        department:row[idx[COLS.department]] || "",
        leave_balance: row[idx[COLS.leave]] || "",
        latest_degree: row[idx[COLS.degree]] || "",
        national_address: row[idx[COLS.address]] || ""
      };
      const cvFileId = row[idx[COLS.cvFileId]] || "";
      if(cvFileId){
        try{
          const f = DriveApp.getFileById(String(cvFileId));
          profile.cv_url = "https://drive.google.com/file/d/" + f.getId() + "/view";
        }catch(e){ /* ignore */ }
      }

      // pull latest 20 attendance rows for this id
      profile.attendance = fetchAttendanceFor(id, 20);

      return { ok:true, profile };
    }
  }
  return { ok:false, message:"الهوية أو كلمة المرور غير صحيحة." };
}

function handleUpdateCV(data){
  const id = String(data.id_number || "").trim();
  const fileName = String(data.file_name || "cv.pdf");
  const fileB64 = String(data.file_b64 || "");
  if(!id || !fileB64){ return { ok:false, message:"بيانات غير مكتملة." }; }

  const blob = Utilities.newBlob(Utilities.base64Decode(fileB64), MimeType.PDF, fileName);
  const folder = DriveApp.getFolderById(CV_FOLDER_ID);
  const file = folder.createFile(blob);
  const fileId = file.getId();

  // update sheet cell
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMP);
  const rng = sh.getDataRange();
  const values = rng.getValues();
  const header = values[0];
  const idx = {};
  header.forEach((h, i) => { idx[h] = i; });

  const idCol = idx[COLS.id_number];
  const cvCol = idx[COLS.cvFileId];
  for(let r = 1; r < values.length; r++){
    if(String(values[r][idCol]) === id){
      sh.getRange(r+1, cvCol+1).setValue(fileId);
      break;
    }
  }

  return {
    ok:true,
    cv_url: "https://drive.google.com/file/d/" + fileId + "/view"
  };
}

function handleAttendance(data){
  const id = String(data.id_number || "").trim();
  const kind = String(data.kind || "").trim(); // "in" | "out"
  if(!id || !kind){ return { ok:false, message:"بيانات غير مكتملة." }; }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATT);
  const now = new Date();
  sh.appendRow([ now, id, kind === "in" ? "IN" : "OUT", "" ]);

  return { ok:true };
}

function fetchAttendanceFor(id, limit){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATT);
  const rng = sh.getDataRange();
  const vals = rng.getValues();
  const out = [];
  // headers: Timestamp, ID, Type, Note
  for(let i = vals.length - 1; i >= 1 && out.length < limit; i--){
    const row = vals[i];
    if(String(row[1]) === id){
      out.push({
        date: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        time: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "HH:mm"),
        type: row[2] === "IN" ? "حضور" : "انصراف",
        note: row[3] || ""
      });
    }
  }
  return out;
}
