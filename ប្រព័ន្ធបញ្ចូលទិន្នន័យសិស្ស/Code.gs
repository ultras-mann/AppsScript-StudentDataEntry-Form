/**
 * បង្កើតម៉ឺនុយបញ្ជា (Menu) ថ្មីក្នុង Google Sheet នៅពេលបើកឯកសារ
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📝 បញ្ចូលទិន្នន័យសិស្ស')
      .addItem('បើកទម្រង់បញ្ចូល (Open Form)', 'showStudentForm')
      .addToUi();
}

/**
 * បង្ហាញទំព័រ HTML Form ជា Modal Dialog (Popup) នៅកណ្តាលអេក្រង់ Google Sheet
 */
function showStudentForm() {
  var html = HtmlService.createHtmlOutputFromFile('Student_list')
      .setWidth(500)   
      .setHeight(600); 
 
  SpreadsheetApp.getUi().showModalDialog(html, 'ប្រព័ន្ធបញ្ចូលទិន្នន័យសិស្ស');
}

/**
 * បង្ហាញទំព័រ HTML Form នៅពេលបើកតាម Web App URL
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Student_list')
      .setTitle('ប្រព័ន្ធបញ្ចូលទិន្នន័យសិស្ស')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * មុខងារទាញយកបញ្ជីឈ្មោះសិស្សទាំងអស់សម្រាប់ដាក់ក្នុង Dropdown
 * Mapping ផ្អែកលើរូបភាព: Col A=ID, Col B=នាមខ្លួន, Col C=គោត្តនាម
 */
function getAllStudents() {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 10) return []; 

  var range = sheet.getRange(10, 1, lastRow - 9, 9);
  var values = range.getValues();
  var students = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[0]) { // បើមាន ID / លេខរៀង
      students.push({
        id: row[0],         // លេខរៀង (ID) នៅ Column A
        firstName: row[1],  // នាមខ្លួន Column B
        lastName: row[2]    // គោត្តនាម Column C
      });
    }
  }
  return students;
}

/**
 * មុខងារទាញយកព័ត៌មានលម្អិតរបស់សិស្សម្នាក់តាមរយៈ ID (លេខរៀង)
 * Mapping ផ្អែកលើរូបភាព: 
 * A: ID, B: នាមខ្លួន, C: គោត្តនាម, D: ភូមិ, E: ឃុំ/សង្កាត់, F: ស្រុក/ខណ្ឌ, G: រាជធានី/ខេត្ត, H: លេខទូរស័ព្ទ, I: ប្រព័ន្ធទូរស័ព្ទ
 */
function getStudentById(studentId) {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 10) return null;

  var range = sheet.getRange(10, 1, lastRow - 9, 9);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[0] == studentId) {
      return {
        id: row[0],
        firstName: row[1] || "",
        lastName: row[2] || "",
        village: row[3] || "",       // Column D
        commune: row[4] || "",       // Column E
        district: row[5] || "",      // Column F
        province: row[6] || "កំពង់ចាម", // Column G
        phone: row[7] || "",         // Column H
        network: row[8] || ""        // Column I
      };
    }
  }
  return null;
}

/**
 * មុខងារសម្រាប់ទទួលទិន្នន័យពី Form ហើយបញ្ចូលទៅកាន់ Google Sheet (Save)
 */
function saveStudent(data) {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("រកមិនឃើញ Sheet ដែលមានឈ្មោះ '" + sheetName + "' ទេ។");
  }
  
  var lastRow = sheet.getLastRow();
  var autoId = 1;
  if (lastRow >= 10) {
    var existingIds = sheet.getRange(10, 1, lastRow - 9, 1).getValues();
    var maxId = 0;
    for (var i = 0; i < existingIds.length; i++) {
      var val = Number(existingIds[i][0]);
      if (val > maxId) maxId = val;
    }
    autoId = maxId + 1;
  }

  // រៀបចំលំដាប់ទិន្នន័យឱ្យត្រូវគ្នានឹង Column A ដល់ I ក្នុង Sheet របស់អ្នក
  var rowData = [
    autoId,                 // A: ID
    data.firstName || "",   // B: នាមខ្លួន
    data.lastName || "",    // C: គោត្តនាម
    data.village || "",     // D: ភូមិ
    data.commune || "",     // E: ឃុំ / សង្កាត់
    data.district || "",    // F: ស្រុក / ក្រុង / ខណ្ឌ
    data.province || "",    // G: រាជធានី / ខេត្ត
    data.phone || "",       // H: លេខទូរស័ព្ទ
    data.network || ""      // I: ប្រព័ន្ធទូរស័ព្ទ
  ];

  sheet.appendRow(rowData);
  return "Success";
}

/**
 * មុខងារសម្រាប់កែប្រែទិន្នន័យសិស្ស (Update)
 */
function updateStudent(data) {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("រកមិនឃើញ Sheet ដែលមានឈ្មោះ '" + sheetName + "' ទេ។");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 10) throw new Error("គ្មានទិន្នន័យសម្រាប់កែប្រែទេ!");

  var range = sheet.getRange(10, 1, lastRow - 9, 9);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == data.id) {
      var rowIndex = 10 + i; 
      
      sheet.getRange(rowIndex, 2).setValue(data.firstName || ""); // Column B
      sheet.getRange(rowIndex, 3).setValue(data.lastName || "");  // Column C
      sheet.getRange(rowIndex, 4).setValue(data.village || "");   // Column D
      sheet.getRange(rowIndex, 5).setValue(data.commune || "");   // Column E
      sheet.getRange(rowIndex, 6).setValue(data.district || "");  // Column F
      sheet.getRange(rowIndex, 7).setValue(data.province || "");  // Column G
      sheet.getRange(rowIndex, 8).setValue(data.phone || "");     // Column H
      sheet.getRange(rowIndex, 9).setValue(data.network || "");   // Column I
      
      return "Updated successfully";
    }
  }
  throw new Error("រកមិនឃើញ ID សិស្សនេះទេ!");
}

/**
 * មុខងារសម្រាប់លុបទិន្នន័យសិស្ស (Delete)
 */
function deleteStudent(studentId) {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("រកមិនឃើញ Sheet ដែលមានឈ្មោះ '" + sheetName + "' ទេ។");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 10) throw new Error("គ្មានទិន្នន័យសម្រាប់លុបទេ!");

  var range = sheet.getRange(10, 1, lastRow - 9, 9);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == studentId) {
      var rowIndex = 10 + i;
      sheet.deleteRow(rowIndex);
      return "Deleted successfully";
    }
  }
  throw new Error("រកមិនឃើញ ID សិស្សនេះដើម្បីលុបទេ!");
}
