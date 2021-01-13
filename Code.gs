function getScriptFolder()
{
  var scriptId = ScriptApp.getScriptId();
  console.info('scriptId = ' + scriptId);
  
  var file = DriveApp.getFileById(scriptId);
  var folders = file.getParents();
  if (folders.hasNext())
  {
    var folder = folders.next();
    var name = folder.getId();
    console.info('script folder name = ' + name);  
    return name  
  }  
}

function makeCopies(templateFilename, rangeArr, filenameFormat) {
  // Get current folder
  // TODO: explore how to get folder dynamically, make it a param
  var thisFolder = DriveApp.getFolderById(getScriptFolder());
  Logger.log(thisFolder.getName());

  // Use pre-created template file
  // TODO: explore feeding this as a parameter
  var files = thisFolder.getFilesByName(templateFilename);
  console.info("files: " + files)
  var template = files.next();

  // Get the names of the companies
  // In this case, there is only 1 Google Sheet file in the folder
  // TODO: otherwise explore making the filename a param as well
  var spreadsheetFiles = thisFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var companyFile = spreadsheetFiles.next();
  var companySheet = SpreadsheetApp.open(companyFile);//(MimeType.GOOGLE_SHEETS).getSheetByName("Sheet1");
  Logger.log(companySheet);

  // TODO: make the Range a param as well
  var companyNames = companySheet.getSheetByName("Sheet1").getSheetValues(rangeArr[0],rangeArr[1],rangeArr[2], rangeArr[3]);
  //.getSheetValues(2,1,9,1);//.getRange(2,1,11,1);
  //*var companyName = companyNames.getCell(1,1);
  Logger.log(companyNames);

  // iterate through company names and create customised copies of the slide deck proposal
  // TODO: make the filename format of the copies a param to the fn too
  for (companyName of companyNames) {
    var index = filenameFormat[0];
    var copyFilename = '';
    var i = 0;
    for (chunk of filenameFormat[1]) {
      if (i == index) {
        copyFilename += companyName;
      }
      copyFilename += filenameFormat[1][i];
      i++;
    }
    var copy = template.makeCopy(copyFilename);
    var copyId = copy.getId();
    // TODO: explore using a pop-up visual selector to indicate where the object is in the presentation/document.
    //        Then work to be able to edit this dynamically.
    SlidesApp.openById(copyId).getSlides()[0].getShapes()[0].getText().appendText(companyName);
  }
}

function run() {
  var templateFilename = "For Potential Sponsors || ASG 1920 | AUG 20' Recruitment Sponsorships Proposal";
  var rangeArr = [2,1,9,1];
  var copyFiletype = MimeType.GOOGLE_SLIDES;
  filenameFormat = [1, ["For "," || ASG 1920 | AUG 20' Recruitment Sponsorships Proposal"]];
  // null represents the position to insert the dynamic values
  //"For " + companyName + " || ASG 1920 | AUG 20' Recruitment Sponsorships Proposal"
  makeCopies(templateFilename,rangeArr,filenameFormat)
}

function clearCopies(templateFilename, copyFiletype) {
  // TODO: explore how to get folder dynamically, make it a param
  var thisFolder = DriveApp.getFolderById(getScriptFolder());
  Logger.log(thisFolder.getName());

  // TODO: make the file type a param, or at least base it on the template filetype
  var files = thisFolder.getFilesByType(copyFiletype);
  while (files.hasNext()) {
    var f = files.next();
    // TODO: trash all but the template file (don't hardcode template filename)
    if (f.getName() != templateFilename) {
       f.setTrashed(true);
    }
  }    
}

function clear() {
  var templateFilename = "For Potential Sponsors || ASG 1920 | AUG 20' Recruitment Sponsorships Proposal";
  var rangeArr = [2,1,9,1];
  var copyFiletype = MimeType.GOOGLE_SLIDES;
  filenameFormat = [1, ["For "," || ASG 1920 | AUG 20' Recruitment Sponsorships Proposal"]]
  // first element is the position to insert the dynamic values
  // second element is a list with the other chunks
  clearCopies(templateFilename, copyFiletype)
}