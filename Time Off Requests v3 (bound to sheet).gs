

function createPDF() {
  const pdfFolder = DriveApp.getFolderById('1aW9-DKkA-KLJP9aUrFkFm5fgaOI-bQoK');
  const tempFolder = DriveApp.getFolderById('1JyjOkmekOIL1PzuoV7kLQMQ_gFVoaB2p');
  const templateDoc = DriveApp.getFileById('1CcRJAPJiquGmBONotSoKiyXCfaIdkDzT6fNc5IC49Ck');

  const spreadsheet = SpreadsheetApp.openById('1_BumT1wq0XUwvtAYagqofdQuxJ2eT_LoEKwZOy_aYkw');
  const sheet = spreadsheet.getSheetByName('Form Responses 1');
  const lastRow = sheet.getLastRow();

  // From Spreadsheet Data

  const timestamp = sheet.getRange(lastRow, 1).getValue().toString();
  const emailAddress = sheet.getRange(lastRow, 2).getValue().toString();
  const name = sheet.getRange(lastRow, 3).getValue().toString();
  const selectYourOrganization = sheet.getRange(lastRow, 4).getValue().toString();
  const managerEmail =  sheet.getRange(lastRow, 5).getValue().toString();
  const numberOfDaysRequested = sheet.getRange(lastRow, 6).getValue().toString();
  const startingOn = sheet.getRange(lastRow, 7).getValue().toString();
  const endingOn = sheet.getRange(lastRow, 8).getValue().toString();
  const dateOfReturn = sheet.getRange(lastRow, 9).getValue().toString();
  const reason = sheet.getRange(lastRow, 10).getValue().toString();
  const signatureWriteYourFullName = sheet.getRange(lastRow, 12).getValue().toString();
  const year = sheet.getRange(lastRow, 13).getValue().toString();
  const eaEmail = "marga.ea.jjproductions@gmail.com";
  const hrEmail = "hrdept@jjproductions.online";

  // end of From Spreadsheet Data

  // Process code

  const newTempfile = templateDoc.makeCopy(tempFolder);
  const openDoc = DocumentApp.openById(newTempfile.getId());
  const body = openDoc.getBody();

  // End of process code

  // Date format code

  const spreadsheetTimeZone = spreadsheet.getSpreadsheetTimeZone();

  const startingOnDate = Utilities.formatDate(new Date(Date.parse(startingOn)), spreadsheetTimeZone, "MM-dd-yyyy");
  const endingOnDate = Utilities.formatDate(new Date(Date.parse(endingOn)), spreadsheetTimeZone, "MM-dd-yyyy");
  const dateOfReturnDate = Utilities.formatDate(new Date(Date.parse(dateOfReturn)), spreadsheetTimeZone, "MM-dd-yyyy");

  // End of date format code


  body.replaceText("{Timestamp}", timestamp);
  body.replaceText("{Email Address}", emailAddress);
  body.replaceText("{Name}", name);
  body.replaceText("{Select your organization}", selectYourOrganization);
  body.replaceText("{Number of Days Requested}", numberOfDaysRequested);
  body.replaceText("{Starting On}", startingOnDate);
  body.replaceText("{Ending On}", endingOnDate);
  body.replaceText("{Date of Return}", dateOfReturnDate);
  body.replaceText("{Reason}", reason);
  body.replaceText("{Signature: Write your full name }", signatureWriteYourFullName);
  body.replaceText("{Year}", year);



  openDoc.saveAndClose();

  const blobPDF = newTempfile.getAs(MimeType.PDF);
  const fileName = reason + ', ' + name + '.pdf'
  const pdfFile = pdfFolder.createFile(blobPDF).setName(fileName);


  //code to send email 
  
 const emailSubject = 'Time Off Request - ' + reason + ', ' + name;
  const htmlBody = "Dear " + name + ",\n\n" +
    "Attached below is a copy of the Time Off Request you submitted.\n\n" +
    "This is still subject to your director/manager's approval.";

  const to = emailAddress + ','+ managerEmail + ',' + eaEmail + ',' + hrEmail;
  const cc = eaEmail;
  const attachments = [pdfFile];
  const ename = 'J&J Productions Automated Emailer';

GmailApp.sendEmail(to,emailSubject,htmlBody, {
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: 'J&J Productions Automated Emailer'
});


}
