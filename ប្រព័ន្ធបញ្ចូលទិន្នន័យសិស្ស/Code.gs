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
 */
function getAllStudents() {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 10) return []; // ប្រសិនបើគ្មានទិន្នន័យពីជួរទី ១០ ចុះក្រោម

  var range = sheet.getRange(10, 1, lastRow - 9, 9);
  var values = range.getValues();
  var students = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[0]) { // បើមាន ID/ លេខរៀង
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
        province: row[3] || "កំពង់ចាម",
        district: row[4] || "",       
        commune: row[5] || "",        
        village: row[6] || "",        
        phone: row[7] || "",          
        network: row[8] || ""         
      };
    }
  }
  return null;
}

/**
 * មុខងារសម្រាប់ទទួលទិន្នន័យពី Form ហើយបញ្ចូលទៅកាន់ Google Sheet
 */
function saveStudent(data) {
  var sheetName = "តារាងពិនិត្យឈ្មោះសិស្ស"; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("រកមិនឃើញ Sheet ដែលមានឈ្មោះ '" + sheetName + "' ទេ។ សូមពិនិត្យមើលឈ្មោះឡើងវិញ!");
  }
  
  var lastRow = sheet.getLastRow();
  // កែតម្រូវការគណនា ID ស្វ័យប្រវត្តិឱ្យបានត្រឹមត្រូវ
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

  var rowData = [
    autoId,
    data.firstName || "",
    data.lastName || "",
    data.province || "",
    data.district || "",
    data.commune || "",
    data.village || "",
    data.phone || "",
    data.network || ""
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
      
      sheet.getRange(rowIndex, 2).setValue(data.firstName || "");
      sheet.getRange(rowIndex, 3).setValue(data.lastName || "");
      sheet.getRange(rowIndex, 4).setValue(data.province || "");
      sheet.getRange(rowIndex, 5).setValue(data.district || "");
      sheet.getRange(rowIndex, 6).setValue(data.commune || "");
      sheet.getRange(rowIndex, 7).setValue(data.village || "");
      sheet.getRange(rowIndex, 8).setValue(data.phone || "");
      sheet.getRange(rowIndex, 9).setValue(data.network || "");
      
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
