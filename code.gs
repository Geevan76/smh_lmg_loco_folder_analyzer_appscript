// This function will run when the spreadsheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Creates a custom menu in the Google Sheets UI
  ui.createMenu('H10 Report Dashboard')
      .addItem('🗂️ Update Folders & Files List', 'listFoldersAndFiles')
      .addToUi();
}

function listFoldersAndFiles() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  sheet.clearContents();
  sheet.clearFormats();

  // Set up headers and depth input cell
  var headers = ["Folder Name", "File Name", "Level", "File Path"];
  sheet.appendRow(["Max Depth:"]);
  sheet.getRange("A1").setNote("Specify the maximum depth of folder traversal here").setValue(6); // Default max depth
  sheet.appendRow(headers);

  // Read the max depth from the cell A1
  var maxDepth = sheet.getRange("A1").getValue();
  
  var folderId = '1of3N1OflLGIzun111w_cP-q_J2BprRU9';
  var rootFolder = DriveApp.getFolderById(folderId);

  var data = [];
  var subFolders = rootFolder.getFolders();
  
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    traverseFolder(subFolder, data, "", 1, maxDepth); // Start at level 1, read max depth from sheet
  }

  // Write all the data to the sheet
  if (data.length > 0) {
    var range = sheet.getRange(3, 1, data.length, headers.length);
    range.setValues(data);
  }

  // Convert file paths to clickable links
  for (var i = 0; i < data.length; i++) {
    if (data[i][3] !== "") {
      var cell = sheet.getRange(i + 3, 4);
      cell.setFormula('=HYPERLINK("' + data[i][3] + '", "' + data[i][3] + '")');
    }
  }

  // Apply alternating background colors and styling for each folder's data
  applyStyling(sheet, data);

  // Auto-resize columns for content
  for (var j = 1; j <= headers.length; j++) {
    sheet.autoResizeColumn(j);
  }

  // Auto-resize header columns for content
  for (var k = 1; k <= headers.length; k++) {
    sheet.autoResizeColumn(k);
  }

  // Set background color and border for the header
  sheet.getRange(2, 1, 1, headers.length).setBackground('#3467eb').setFontColor('#FFFFFF').setFontWeight('bold').setBorder(true, true, true, true, true, true);
}

// Recursive function to traverse folders up to a certain depth
function traverseFolder(folder, data, path, depth, maxDepth) {
  if (depth > maxDepth) {
    return;
  }

  var folderName = folder.getName();
  var folderPath = path ? path + "/" + folderName : folderName;
  
  // Add a blank row before each folder's data
  if (data.length > 0) {
    data.push(["", "", "", ""]);
  }
  
  // Get all files in the current folder
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var filePath = file.getUrl(); // Get the URL of the file
    data.push([folderName, fileName, depth, filePath]);
  }

  // Get all subfolders in the current folder and recurse into them
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    traverseFolder(subFolder, data, folderPath, depth + 1, maxDepth);
  }
}

function applyStyling(sheet, data) {
  var colors = ['#b4faf4', '#e6e8e8']; // Alternating colors: aqua and light gray
  var currentColorIndex = 0;
  
  for (var i = 0; i < data.length; i++) {
    var rowRange = sheet.getRange(i + 3, 1, 1, 3); // Apply styling to the first three columns
    
    if (data[i][0] === "" && data[i][1] === "" && data[i][2] === "" && data[i][3] === "") { // This is the blank row before a new folder
      currentColorIndex = (currentColorIndex + 1) % colors.length; // Alternate color
    } else {
      rowRange.setBackground(colors[currentColorIndex]);
      if (data[i][1] === "") { // Folder name row
        rowRange.setFontWeight('bold').setBorder(true, true, true, true, true, true); // Bold and border for folder name row
      }
    }
  }
}
