/**
 * =========================================================================
 * سكريبت جوجل لتخزين بيانات مسجلي "فُرَصَك" (Forasak) في Google Sheets تلقائيًا
 * رابط الشيت المستهدف:
 * https://docs.google.com/spreadsheets/d/1IOYYl6qf15DkYO0Jtky9nweYBY3xXdTM0_yiNUq_caU/edit?usp=sharing
 * 
 * ترتيب الأعمدة في الشيت:
 * - العمود الأول (A): الاسم (Name)
 * - العمود الثاني (B): البريد الإلكتروني (Email)
 * - العمود الثالث (C): تاريخ ووقت التسجيل (Registration Timestamp)
 * =========================================================================
 * 
 * ⚠️ سبب عدم ظهور البيانات لديك عند تعبئة النموذج:
 * حتى تذهب البيانات إلى شيت جوجل، يجب نشر هذا السكريبت والحصول على رابط Web App URL
 * ووضعه في ملف index.html مكان:
 * const FORM_ENDPOINT = "ضع_الرابط_هنا";
 * (إذا كان FORM_ENDPOINT فارغاً، المتصفح يحفظ البيانات تجريبياً على جهازك فقط ولا يرسلها إلى جوجل).
 * 
 * طريقة التركيب السريعة (في دقيقة واحدة):
 * 1. افتح شيت جوجل الخاص بك في المتصفح.
 * 2. من القائمة العلوية اضغط على: Extensions (ملحقات) ثم اختر Apps Script.
 * 3. احذف أي كود موجود في المحرر وضع هذا الكود بالكامل مكانه.
 * 4. اضغط على أيقونة الحفظ 💾 (أو Ctrl + S).
 * 5. اضغط على الزر الأزرق أعلى اليمين: Deploy (نشر) -> New deployment (نشر جديد).
 * 6. بجانب "Select type" (اختر النوع)، اضغط على الترس ⚙️ واختر Web app (تطبيق ويب).
 * 7. اضبط الإعدادات كالتالي:
 *    - Description: Forasak Form Webhook
 *    - Execute as: Me (حسابي)
 *    - Who has access: Anyone (أي شخص) <--- مهم جداً ليتمكن زوار الموقع من التسجيل
 * 8. اضغط على زر Deploy (نشر).
 * 9. وافق على الأذونات (Authorize access -> اختر حسابك -> Advanced -> Go to ... (unsafe) -> Allow).
 * 10. انسخ الرابط المسمى Web app URL (ينتهي بـ /exec) وضعه في ملف index.html في المتغير:
 *     const FORM_ENDPOINT = "الرابط_المنسوخ_هنا";
 * =========================================================================
 */

function doPost(e) {
  // قفل لحماية الشيت من الكتابة المتزامنة في نفس اللحظة
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    // فتح الشيت الحالي أو عبر المعرّف
    var ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {}
    
    if (!ss) {
      ss = SpreadsheetApp.openById("1IOYYl6qf15DkYO0Jtky9nweYBY3xXdTM0_yiNUq_caU");
    }
    var sheet = ss.getActiveSheet();

    var name = "";
    var email = "";
    var timestamp = new Date();

    // استخراج البيانات سواء أُرسلت كـ JSON أو كـ Form URL Encoded
    if (e && e.postData && e.postData.contents) {
      try {
        var parsed = JSON.parse(e.postData.contents);
        name = parsed.name || "";
        email = parsed.email || "";
      } catch (err) {
        name = (e.parameter && e.parameter.name) ? e.parameter.name : "";
        email = (e.parameter && e.parameter.email) ? e.parameter.email : "";
      }
    } else if (e && e.parameter) {
      name = e.parameter.name || "";
      email = e.parameter.email || "";
    }

    // تنسيق التاريخ والوقت بتوقيت مصر
    var formattedDate = Utilities.formatDate(timestamp, "Africa/Cairo", "yyyy-MM-dd hh:mm:ss a");

    // إذا كان الشيت فارغاً تماماً، نضيف صف العناوين:
    // العمود الأول: الاسم | العمود الثاني: البريد الإلكتروني | العمود الثالث: تاريخ ووقت التسجيل
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["الاسم", "البريد الإلكتروني", "تاريخ ووقت التسجيل"]);
      var headerRange = sheet.getRange(1, 1, 1, 3);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1B263B"); // كحلي فُرَصَك
      headerRange.setFontColor("#FFFFFF");
      headerRange.setHorizontalAlignment("center");
      sheet.setRowHeight(1, 35);
      sheet.setColumnWidth(1, 220); // العمود 1: الاسم
      sheet.setColumnWidth(2, 260); // العمود 2: البريد الإلكتروني
      sheet.setColumnWidth(3, 200); // العمود 3: التاريخ
    }

    // إضافة البيانات: الاسم في العمود الأول، والبريد الإلكتروني في العمود الثاني
    sheet.appendRow([name, email, formattedDate]);

    // محاذاة النص بالوسط
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 3).setHorizontalAlignment("center");

    // إرجاع استجابة نجاح بصيغة JSON
    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", name: name, email: email }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// دالة doGet للتحقق من عمل السكريبت عند فتحه في المتصفح مباشرة
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      status: "ok", 
      service: "Forasak Google Sheet Webhook", 
      message: "الخدمة تعمل بنجاح وجاهزة لاستقبال البيانات من موقع فُرَصَك." 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// دالة تجريبية لاختبار الكود يدوياً من داخل محرر جوجل
function testAddRow() {
  doPost({
    postData: {
      contents: JSON.stringify({
        name: "تجربة فُرَصَك",
        email: "test@forasak.com"
      })
    }
  });
}
