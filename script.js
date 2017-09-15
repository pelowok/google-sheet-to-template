
function onOpen() {
  var menuEntries = [ {name: "Create Autofilled Template", functionName: "AutofillDocFromTemplate"}];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Create", menuEntries);
}

function AutofillDocFromTemplate(){
  //var templateId = "1d9lAkpYnyt6sIP2uewt_KVqwpoklylOxxXe0ERmZNDo"; // this is the template doc file id. You can find this in the URL of the google document template. For example, if your URL Looks like this: https://docs.google.com/document/d/1SDTSW2JCItWMGkA8cDZGwZdAQa13sSpiYhiH-Kla6VA/edit, THEN the ID would be 1SDTSW2JCItWMKkA8cDZGwZdAQa13sSpiYhiH-Kla6VA
  var templateId = "1d9lAkpYnyt6sIP2uewt_KVqwpoklylOxxXe0ERmZNDo";
  var mainFolderId = "0B1S8sXGxQgoXMlhycUpKN1JmR00"; // folder id of main folder
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  //create a subfolder named after the sheet tab you are using
  var folder = DriveApp.getFolderById(mainFolderId).createFolder(sheet.getSheetName());
  
  var startRow = 3; // what row of the active sheet should the range start on?
  var stopRow = sheet.getLastRow() - startRow; // this is set dynamically, so the loop is right-sized
  
  var range = sheet.getRange(startRow, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues(); // define the range as a block of rows and columns
  
  
  
  for (var i = 1; i <= stopRow; i++) { // loop that applies this code block to each ROW in the range
  // test conditions
  //for (var i = 1; i <= 3; i++) { // loop that applies this code block to each ROW in the range
        
    var childName = range[i][1]; // get the child name from the data range[row][column]
    
    var newDoc = DocumentApp.create("Storybook - " + childName); //create a blank doc with the name of the child
                           
    var newDocID = DriveApp.getFileById(newDoc.getId());
    
    // move file to right folder
    folder.addFile(newDocID);
    
    var templateCopyId = DriveApp.getFileById(templateId).makeCopy().getId();
    var doc = DocumentApp.openById(templateCopyId);
    
    var body = doc.getActiveSection();
    var activeRow = range[i];
    
    var formattedDate = Utilities.formatDate(activeRow[2], 'America/New_York', 'MMMM dd, yyyy');
    Logger.log(formattedDate);

    // use the value in the current row that is in column [number] of the RANGE (range may be shifted from first column, look at definition of range above)
    body.replaceText("%NAME%", activeRow[1]); 
    body.replaceText("%DATE%", Utilities.formatDate(activeRow[2], 'America/New_York', 'MMMM dd, yyyy') );
    body.replaceText("%FAMILY%", activeRow[3]); 
    body.replaceText("%PLACE%", activeRow[4]);
    body.replaceText("%LANGUAGE%", activeRow[6]);
    body.replaceText("%FARAWAY%", activeRow[5]);
    body.replaceText("%LOVE%", activeRow[7]);
    body.replaceText("%FOODS%", activeRow[8]);
    body.replaceText("%SPECIAL%", activeRow[9]);
    
    body.replaceText("%CLASS%", activeRow[33]);

    //appendToDoc(doc, newDoc);
    var body = doc.getActiveSection();
    var newBody = newDoc.getActiveSection();
    appendToDoc(body, newBody);
    
    doc.saveAndClose()
    newDoc.saveAndClose()
 
    DriveApp.getFileById(templateCopyId).setTrashed(true);
    
  }
  ss.toast("It worked. it's alive. IT'S ALIVE!!");
}

function appendToDoc(src, dst) {
  for (var i = 0; i < src.getNumChildren(); i++) {
    appendElementToDoc(dst, src.getChild(i));
  }
}

function appendElementToDoc(doc, object) {
  var type = object.getType();
  var element = object.copy();
  Logger.log("Element type is "+type);
  if (type == "PARAGRAPH") {
    doc.appendParagraph(element);
  } else if (type == "TABLE") {
    doc.appendTable(element);
  } 
}
