// Google sheets script to assign contribution number and send an email to the
// person requesting a number for a new publication

// get the value of the cell with the next contribution number (max of current #s plus one)
function onFormSubmit(sheet, row) {
  // https://webapps.stackexchange.com/questions/15449/can-i-add-an-autoincrement-field-to-a-google-spreadsheet-based-on-a-google-form
  // Get the next ID value.  NOTE: This cell should be set to: =MAX(A:A)+1
  var id = sheet.getRange("N1").getValue();
  // Check of ID column is empty
  if (sheet.getRange(row, 1).getValue() == "") {
    // Set new ID value
    sheet.getRange(row, 1).setValue(id);
  }
}

/*Send an email with data from the current spreadsheet.
 https://developers.google.com/apps-script/articles/sending_emails
 Email is bcc'ed to marinegeo resource accounts
 */
function sendEmails(sheet, row) {
  var dataRange = sheet.getRange(row, 1, 1, 12);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues()[0];
  Logger.log(data);
  var id = data[0];
  var ts = data[1];
  var emailAddress = data[2];
  var name = data[3];
  var title = data[4];
  var journal = data[5];
  var citation = data[7];
  var doi = data[8];
  var url = data[10];

  var subject = 'MarineGEO publication number: ' + id;

  var message = "Hi " + name + ", \n\n" +
     "You have successfully requested a MarineGEO publication number.\n\n" +
     "Title: " + title + "\n" +
     "Journal: " + journal + "\n"+
     "MarineGEO publication number: " + id + "\n" +
      "Citation: " + citation +"\n" +
      "DOI: " + doi + "\n\n" +

     "Please include the following language in the acknowledgement section: \n\n" +
     "This is contribution " + id + " from the Smithsonian’s MarineGEO Network." + "\n\n"+
     "Contact marinegeo@si.edu with any questions and please let us know when the paper is published."+ "\n\n" +
     "Thank you, \n\n"+
     "MarineGEO" + "\n\n\n"+
      "Update/Edit Response: " + url;
  Logger.log(message);
  MailApp.sendEmail(emailAddress, subject, message, {bcc:"marinegeo-data@si.edu, marinegeo@si.edu"});

}

/*
Adds a note to the spreadsheet if the email confirmation was sent
*/
function confirmation(sheet, row){
  // Check of emailConfimation column is empty
  if (sheet.getRange(row, 10).getValue() == "") {
    sheet.getRange(row, 10).setValue("yes");
}
}

// Add url for user to edit their reponse
function addFormResponseUrl(sheet, row) {
  // https://ctrlq.org/code/20540-edit-form-response-spreadsheet-url
  // Get the Google Form linked to the response
  var googleFormUrl = sheet.getFormUrl();
  var googleForm = FormApp.openByUrl(googleFormUrl);

  // Get the form response based on the timestamp
  var timestamp = new Date(sheet.getRange(row, 2).getValue());
  var formResponse = googleForm.getResponses(timestamp).pop();

  // Get the Form response URL and add it to the Google Spreadsheet
  var responseUrl = formResponse.getEditResponseUrl();
  var responseColumn = 11; // Column where the response URL is recorded.
  sheet.getRange(row, responseColumn).setValue(responseUrl);
}


// triggers script when the form is submitted
function main(e){
  // selet the active sheet
  //var sheet = SpreadsheetApp.getActiveSheet();
  var sheet = e.range.getSheet();

  // automatically select the last row
  //var row = sheet.getLastRow();
  var row = e.range.getRow();

  onFormSubmit(sheet, row);
  addFormResponseUrl(sheet, row);
  sendEmails(sheet, row);
  confirmation(sheet, row);
}

// Given a publication number return the row in the spreadsheet
//https://stackoverflow.com/questions/32565859/find-cell-matching-value-and-return-rownumber
function rowOfPubNumb(pubnumber){
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for(var i = 0; i<data.length;i++){
    if(data[i][0] == pubnumber){
      Logger.log((i+1))
      return i+1;
    }
  }
}

// manually resend email
function manualEmail(pubNumber){
  var sheet = SpreadsheetApp.getActiveSheet();
  // get row number from publiction ID
  row = rowOfPubNumb(pubNumber)
  sendEmails(sheet, row);
}

// uncomment code below to resend an email
// manualEmail(27)
