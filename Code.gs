/* =========================================================
   WASHROOM REPORTS — Google Apps Script backend
   Setup: Google Sheet খুলো -> Extensions -> Apps Script -> এই কোড পেস্ট করো
   -> Deploy -> New deployment -> Web app -> Execute as: Me, Access: Anyone
   -> /exec URL টা index.html এর CONFIG.SCRIPT_URL এ বসাও।
   ========================================================= */

const SHEET_NAME = "Reports";
const ADMIN_PIN  = "CHANGE_ME_1234";   // ⚠️ নিজের admin PIN দাও
const ADMIN_EMAIL = "";  // নতুন রিপোর্ট এলে এই ইমেইলে নোটিফিকেশন যাবে (খালি রাখলে বন্ধ)
const HEADERS = ["ID","Timestamp","Name","Position","Department","Location",
                 "Category","Severity","Description","Photo","Status","AdminNote","UpdatedAt"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0F3D3E").setFontColor("#fff");
  }
  return sh;
}

function now_() {
  return Utilities.formatDate(new Date(), "Asia/Dhaka", "yyyy-MM-dd HH:mm:ss");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// পাবলিক ডেটা পড়া (শিট প্রাইভেট থাকলেও চলবে)
function doGet() {
  const rows = getSheet_().getDataRange().getValues();
  const head = rows.shift().map(h => String(h).toLowerCase());
  const data = rows.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = (r[i] instanceof Date) ? Utilities.formatDate(r[i], "Asia/Dhaka", "yyyy-MM-dd HH:mm:ss") : r[i]);
    return o;
  });
  return json_({ status: "ok", data: data.reverse() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = e.parameter;
    const sh = getSheet_();

    if (p.action === "submit") {
      const required = ["name", "position", "location", "category", "description"];
      for (const k of required) if (!p[k] || !String(p[k]).trim()) return json_({ status: "error", message: "Missing: " + k });
      const id = "WR-" + Utilities.formatDate(new Date(), "Asia/Dhaka", "yyMMdd") + "-" + String(sh.getLastRow()).padStart(4, "0");
      sh.appendRow([id, now_(), p.name, p.position, p.department || "", p.location,
                    p.category, p.severity || "Medium", p.description, p.photo || "", "Pending", "", now_()]);
      if (ADMIN_EMAIL) {
        try { MailApp.sendEmail(ADMIN_EMAIL, "[" + (p.severity || "Medium") + "] New washroom report " + id,
          p.name + " (" + p.position + ", " + (p.department || "-") + ")\n" + p.location + " — " + p.category + "\n\n" + p.description + "\n\nPhoto: " + (p.photo || "-")); } catch (x) {}
      }
      return json_({ status: "ok", id: id });
    }

    if (p.action === "updateStatus") {
      if (p.pin !== ADMIN_PIN) return json_({ status: "error", message: "Wrong PIN" });
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat();
      const idx = ids.indexOf(p.id);
      if (idx < 0) return json_({ status: "error", message: "Report not found" });
      const row = idx + 2;
      sh.getRange(row, 11).setValue(p.status);
      if (p.note !== undefined) sh.getRange(row, 12).setValue(p.note);
      sh.getRange(row, 13).setValue(now_());
      return json_({ status: "ok" });
    }

    return json_({ status: "error", message: "Unknown action" });
  } catch (err) {
    return json_({ status: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}
