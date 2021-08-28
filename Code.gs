/** 
 * This function transfers the information from the Google Forms connected Sheets document to a Backend Notes Template 
 * generating a new copy in the process. It works by scanning the Google Form for the highlighting (purple) associated with 
 * the main performers and fetching their associated data. In order to enter the performers in line-up order, 
 * drag the rows around in the Google Forms Sheet so that the main performers are order. It does not matter
 * whether there are non-main performers above, below, or in-between the main performer rows, as long as the highlighting is 
 * correct. If any changes are made to the Google Form that affect the organization of the spreadsheet, then this code may need
 * to be modified. To select which spreadsheet to run the function on, just replace the link below with the link to the 
 * correct spreadsheet. Then just click the "Run" button above.
*/

function transferSheetToBackendTemplate() {
  /**                               REPLACE THIS GOOGLE DRIVE LINK with the desired form submission spreadsheet - 
   * Make sure it has quotation marks around it in the parentheses. */
  var spreadsheetId = getIdFromUrl("");

  // This link is for the folder the populated template will be sent to. You can change this link if you would like to select a different folder.
  var templateFolderId = getIdFromUrl("");

  // This link is for the template the function uses for the Backend Notes. As long as the template contains {{Emcee Act Notes}}, {{Backline Act Notes}}, and {{Sound Act Notes}} then the function should have no problems filling the template. These ARE case sensitive.
  var templateId = getIdFromUrl("");


  /** Unless you are familiar with Google Apps Script, I would not alter anything below this point. */


  // Grabs template and source files
  var templateFile = DriveApp.getFileById(templateId);
  var populatedTemplateFolder = DriveApp.getFolderById(templateFolderId);
  var submissionSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var actSheet = submissionSpreadsheet.getSheets()[0];

  // Creates copy of template in the template folder and name it with current date
  var date = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy");
  var copy = templateFile.makeCopy("Backend Notes " + date, populatedTemplateFolder);

  // Loops through the form submission sheet and grabs main performer data based on highlighting
  var performerList = [];
  var firstColumn = actSheet.getRange(8, getColByName("Main"), actSheet.getLastRow());
  for (var i = 1; i <= firstColumn.getNumRows(); i++) {
    if (firstColumn.getCell(i, 1).getBackground() == "#c27ba0") {
      // Adjusted by 7 to compensate for firstColumn and actSheet having separate ranges (where main begins)
      var performer = {
        stageName : actSheet.getRange(i + 7, getColByName("Stage Name ")).getValue(),
        emceeNotes: actSheet.getRange(i + 7, getColByName("Emcee Notes")).getValue(),
        crewNotes : actSheet.getRange(i + 7, getColByName("Crew notes - for physical backstage performer support")).getValue(),
        stageLocation : actSheet.getRange(i + 7, getColByName("Stage Preferences")).getValue(),
        song : actSheet.getRange(i + 7, getColByName("MP3 backing track - *REQUIRED*")).getValue(),
        techRequirements: actSheet.getRange(i + 7, getColByName("Tech Requirements - stuff our nerds need to know ")).getValue()
      }
      performerList.push(performer);
    }
  }
  // Returns the index of a column in the submission sheet by name (+1 bc of 0-indexing on the indexOf method)
  function getColByName (colName) {
    var data = actSheet.getRange(7,1,1,actSheet.getLastColumn()).getValues();
    return data[0].indexOf(colName) + 1;
  }

  // Grabs the copy's document body
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();
  
  // Loops through the paragraphs in the body and finds the ones to replace
  var paraArray = body.getParagraphs();
  var parasToReplace = [];
  paraArray.forEach(checkReplacement);
  function checkReplacement(currPara) {
    if (currPara.findText("{{.*}}") != null) {
      parasToReplace.push(currPara);
    }
  }

  // Goes through each paragraph to replace and fills in the act info based on the section
  parasToReplace.forEach(replaceWithActInfo);
  function replaceWithActInfo(toReplace) {
    var currIdx = body.getChildIndex(toReplace);

    // For each performer, add their emcee notes
    if (toReplace.findText("Emcee Act Notes") != null) {
      toReplace.clear();
      for (var i = performerList.length - 1; i >= 0; i--) {
        var currListItem = body.insertListItem(currIdx, performerList[i].stageName);
        currListItem.editAsText().setBold(true);
        var currPara = body.insertParagraph(currIdx + 1, performerList[i].emceeNotes + "\n");
        currPara.editAsText().setBold(false);
        currPara.setIndentFirstLine(36);
      }
    }

    // For each performer, add their backline notes
    else if (toReplace.findText("Backline Act Notes") != null) {
      toReplace.clear();
      for (var i = performerList.length - 1; i >= 0; i--) {
        var currListItem = body.insertListItem(currIdx, performerList[i].stageName);
        currListItem.editAsText().setBold(true);
        var currPara = body.insertParagraph(currIdx + 1, "Stage Location: " + performerList[i].stageLocation);
        currPara.editAsText().setBold(false);
        currPara = body.insertParagraph(currIdx + 2, "Crew Notes: " + performerList[i].crewNotes + "\n");
      }
    }

    // For each performer, add their sound cue notes
    else if (toReplace.findText("Sound Act Notes") != null) {
      toReplace.clear();
      for (var i = performerList.length - 1; i >= 0; i--) {
        var currListItem = body.insertListItem(currIdx, performerList[i].stageName);
        currListItem.editAsText().setBold(true);
        currListItem.editAsText().setItalic(false);
        var currPara = body.insertParagraph(currIdx + 1, "Song: " + DriveApp.getFileById(getIdFromUrl(performerList[i].song)).getName());
        currPara.editAsText().setBold(false);
        var currPara = body.insertParagraph(currIdx + 2, "MP3: ");
        currPara.appendText(performerList[i].song).setLinkUrl(performerList[i].song);
        currPara = body.insertParagraph(currIdx + 3, "Sound Cue: " + performerList[i].techRequirements + "\n");
      }
    }
  }

  // Returns the file id from the Google Drive url (thx Stack Overflow)
  function getIdFromUrl(url) { 
    return url.match(/[-\w]{25,}/); 
  }

  doc.saveAndClose();
}
