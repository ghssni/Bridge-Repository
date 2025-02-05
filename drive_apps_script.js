
// Script ini digunakan untuk otomasi pembuatan folder pada url google drive 
// sesuai nama tim yang ada pada table spreadsheet

// Copy code ini dan jalankan pada app script di spreadsheet:
//     - Extension -> apps script


function createFoldersFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var folderNames = sheet.getRange("B2:B48").getValues(); // Sesuaikan rentang data letak nama tim

  var parentFolder = DriveApp.getFolderById("");

  folderNames.forEach(function (name) {
    if (name[0]) {
      parentFolder.createFolder(name[0]);
    }
  });

  Logger.log("Folders created successfully in the specific folder!");
}
