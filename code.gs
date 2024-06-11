// This function will run when the spreadsheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Creates a custom menu in the Google Sheets UI
  // The custom menu is named "H10 Report Dashboard" and contains an item to trigger the listFoldersAndFiles function
  ui.createMenu('🟢 H10 Report Dashboard')
      .addItem('🟢 Update Folders & Files List', 'listFoldersAndFiles')
      .addToUi();
}

function listFoldersAndFiles() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  // Clears any existing content and formatting in the active sheet
  sheet.clearContents();
  sheet.clearFormats();

  // Set up headers and depth input cell
  // The first row is used to specify the maximum depth for folder traversal
  var headers = ["Folder Name", "File Name", "Level", "File Path"];
  sheet.appendRow(["Max Depth:"]);
  // Cell A1 is used to input the maximum depth, with a default value of 6
  sheet.getRange("A1").setNote("Specify the maximum depth of folder traversal here").setValue(6); // Default max depth
  // The second row is for the headers of the data table
  sheet.appendRow(headers);

  // Read the max depth from the cell A1
  var maxDepth = sheet.getRange("A1").getValue();
  
  // ID of the root folder to be listed
  var folderId = '1of3N1OflLGIzun111w_cP-q_J2BprRU9';
  // Retrieve the root folder using its ID
  var rootFolder = DriveApp.getFolderById(folderId);

  // Array to hold the folder and file data
  var data = [];
  // Get all subfolders of the root folder
  var subFolders = rootFolder.getFolders();
  
  // Iterate through each subfolder
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    // Recursively traverse each subfolder and collect data
    traverseFolder(subFolder, data, "", 1, maxDepth); // Start at level 1, read max depth from sheet
  }

  // Write all the collected data to the sheet
  if (data.length > 0) {
    var range = sheet.getRange(3, 1, data.length, headers.length);
    range.setValues(data);
  }

  // Convert file paths to clickable links
  for (var i = 0; i < data.length; i++) {
    if (data[i][3] !== "") {
      var cell = sheet.getRange(i + 3, 4);
      // Set the file path as a hyperlink in the cell
      cell.setFormula('=HYPERLINK("' + data[i][3] + '", "' + data[i][3] + '")');
    }
  }

  // Apply alternating background colors and styling for each folder's data
  applyStyling(sheet, data);

  // Auto-resize columns to fit the content
  for (var j = 1; j <= headers.length; j++) {
    sheet.autoResizeColumn(j);
  }

  // Auto-resize header columns to fit the content
  for (var k = 1; k <= headers.length; k++) {
    sheet.autoResizeColumn(k);
  }

  // Set background color and border for the header row
  sheet.getRange(2, 1, 1, headers.length).setBackground('#3467eb').setFontColor('#FFFFFF').setFontWeight('bold').setBorder(true, true, true, true, true, true);
}

// Recursive function to traverse folders up to a certain depth
function traverseFolder(folder, data, path, depth, maxDepth) {
  // Stop if the current depth exceeds the maximum depth
  if (depth > maxDepth) {
    return;
  }

  var folderName = folder.getName();
  // Create the full path of the current folder
  var folderPath = path ? path + "/" + folderName : folderName;
  
  // Add a blank row before each folder's data for separation
  if (data.length > 0) {
    data.push(["", "", "", ""]);
  }
  
  // Get all files in the current folder
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    // Get the URL of the file
    var filePath = file.getUrl();
    // Add the folder name, file name, depth, and file path to the data array
    data.push([folderName, fileName, depth, filePath]);
  }

  // Get all subfolders in the current folder and recurse into them
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    // Recursively traverse each subfolder
    traverseFolder(subFolder, data, folderPath, depth + 1, maxDepth);
  }
}

function applyStyling(sheet, data) {
  var colors = ['#b4faf4', '#e6e8e8']; // Alternating colors: aqua and light gray
  var currentColorIndex = 0;
  
  for (var i = 0; i < data.length; i++) {
    // Apply styling to the first three columns (Folder Name, File Name, Level)
    var rowRange = sheet.getRange(i + 3, 1, 1, 3);
    
    if (data[i][0] === "" && data[i][1] === "" && data[i][2] === "" && data[i][3] === "") { // This is the blank row before a new folder
      // Alternate the background color
      currentColorIndex = (currentColorIndex + 1) % colors.length;
    } else {
      // Set the background color for the current row
      rowRange.setBackground(colors[currentColorIndex]);
      if (data[i][1] === "") { // Folder name row
        // Set the folder name row to bold and add borders
        rowRange.setFontWeight('bold').setBorder(true, true, true, true, true, true);
      }
    }
  }
}
