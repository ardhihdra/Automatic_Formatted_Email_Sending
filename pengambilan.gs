
function sendEmailPengambilan() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[1];
  var startRow = 1; // First row of data to process
  var numRows = sheet.getLastRow(); // Number of rows to process
  
  var dataRange = sheet.getRange(startRow, 1, numRows, 11);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  // initialize nomor order
 
  for (var i = 1; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[3];
    var emailSent = row[10];
    var konfirmasi = row[9];
    if (emailSent != EMAIL_SENT && konfirmasi == 'OKE') { // Prevents sending duplicates
      var subject = 'Pengambilan Barang';
      var blob = HtmlService.createHtmlOutputFromFile("emailPengambilan").getBlob();
      var a = blob.getDataAsString();
      //generate random code
      //var brand = generateRandom(); 
    
      a=a.replace('{{Nama}}',row[2]);
      a=a.replace('{{Nama}}',row[2]);
      a=a.replace('{{Jumlah Pesanan}}',row[5]);
      a=a.replace('{{Kode}}',row[4]);
      a=a.replace('{{Kode}}',row[4]);
      //Logger.log(b);
      MailApp.sendEmail({to: emailAddress,subject: subject,htmlBody: a});
      
      sheet.getRange(startRow +i, 11).setValue(EMAIL_SENT);
      //sheet.getRange(startRow +i, 5).setValue(brand);
      // Make sure the cell is updated right away in case the script is interrupted {to: 'ardhi.rofi@gmail.com',subject: subject,htmlBody: blob.getDataAsString()}
      SpreadsheetApp.flush();
    }
  } 
}
