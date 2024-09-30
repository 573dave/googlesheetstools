function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Drive Tree')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var rootFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  var dir = rootFolder.getName(); // Get current directory name

  var template = HtmlService.createTemplateFromFile('Sidebar');
  template.dir = dir;

  const html = template.evaluate()
    .setTitle('File Tree Generator')
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}

function generateTree() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Check if the sheet is empty
  if (!isSheetEmpty(sheet)) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Warning',
      'The sheet is not empty. Do you want to overwrite the content?',
      ui.ButtonSet.YES_NO
    );

    // If user clicks 'No', stop the process
    if (response == ui.Button.NO) {
      return "Tree generation cancelled.";
    }
  }

  // Clear the sheet and proceed if user clicked 'Yes' or sheet is empty
  sheet.clear();
  styleSheet(sheet);

  try {
    var rootFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
    var rootFolders = getFoldersInRoot(rootFolder);

    // Write root folder
    sheet.getRange(1, 1).setValue(formatItem({ name: rootFolder.getName(), id: rootFolder.getId(), type: 'folder', level: 0 }));

    var currentRow = 2;
    for (var i = 0; i < rootFolders.length; i++) {
      var folder = rootFolders[i];
      currentRow = processFolder(folder, sheet, currentRow, 1);

      // Resize columns for tree-like structure
      resizeColumns(sheet);

      // Update progress
      var progress = Math.round((i + 1) / rootFolders.length * 100);
      setProgress(progress, folder.getName());
    }

    return "Success";
  } catch (error) {
    throw new Error("Failed to generate tree: " + error.message);
  }
}

// Helper function to check if the sheet is empty
function isSheetEmpty(sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== "") {
        return false;
      }
    }
  }
  return true;
}


function processFolder(folder, sheet, startRow, level) {
  var currentRow = startRow;

  // Write folder
  sheet.getRange(currentRow, level + 1).setValue(formatItem({ name: folder.getName(), id: folder.getId(), type: 'folder', level: level }));
  currentRow++;

  // Process files
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    sheet.getRange(currentRow, level + 2).setValue(formatItem({ name: file.getName(), id: file.getId(), type: file.getMimeType(), level: level + 1 }));
    currentRow++;
  }

  // Process subfolders
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    currentRow = processFolder(subFolder, sheet, currentRow, level + 1);
  }

  return currentRow;
}

function formatItem(item) {
  var icon = getIconForType(item.type);
  var url = 'https://drive.google.com/open?id=' + item.id;
  return '=HYPERLINK("' + url + '", "' + icon + ' ' +  item.name + '")';
}

function getIconForType(type) {
  switch (type) {
    case 'folder':
      return '📂'; // Folder icon
    case 'application/vnd.google-apps.document':
      return '📝'; // Document icon
    case 'application/vnd.google-apps.spreadsheet':
      return '📊'; // Spreadsheet icon
    case 'application/vnd.google-apps.presentation':
      return '📑'; // Presentation icon
    case 'application/pdf':
      return '📕'; // PDF icon
    case 'image/jpeg':
    case 'image/png':
    case 'image/gif':
      return '🖼️'; // Image icon
    case 'video/mp4':
    case 'video/quicktime':
      return '🎥'; // Video icon
    case 'audio/mpeg':
    case 'audio/wav':
      return '🎧'; // Audio icon
    case 'application/zip':
    case 'application/x-zip-compressed':
      return '📦'; // Compressed file icon
    case 'text/plain':
      return '📃'; // Text file icon
    default:
      return '📄'; // Default file icon
  }
}

function resizeColumns(sheet) {
  var lastColumn = sheet.getLastColumn();
  for (var i = 1; i <= lastColumn; i++) {
    sheet.setColumnWidth(i, 23); // Narrower columns for tree-like visualization
  }
  sheet.setColumnWidth(i, 200);
}

function styleSheet(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground('#202124'); // Dark background
  sheet.getRange('A1:Z').setFontColor('#E8EAED'); // Light font color
  sheet.getRange('A1:Z').setFontFamily('Ubuntu'); // Modern font
  sheet.getRange('A1:Z').setFontSize(10);
  sheet.getRange('A1:Z').setHorizontalAlignment('left'); // Tree structure alignment
}

function getProgress() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return JSON.parse(scriptProperties.getProperty('treeProgress') || '{"percent": 0, "folderName": ""}');
}

function setProgress(percent, folderName) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('treeProgress', JSON.stringify({percent: percent, folderName: folderName}));
}

function getFoldersInRoot(rootFolder) {
  var folders = [];
  var folderIterator = rootFolder.getFolders();
  while (folderIterator.hasNext()) {
    folders.push(folderIterator.next());
  }
  return folders;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
