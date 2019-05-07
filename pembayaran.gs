// Mengirim Email otomatis ketika terdapat submit baru di suatu form
// ardhihdra ardhi.rofi@gmail.com

var EMAIL_SENT = 'EMAIL_SENT';
var OKE = 'OKE';

function sendEmailPembayaran(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1; // First row of data to process
  var numRows = sheet.getLastRow(); // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 13);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  // initialize nomor order
 
  for (var i = 2; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[2]; // third column
    
    
      // to avoid email sent 
    //sheet.getRange(startRow +i, 13).setValue(EMAIL_SENT);
    var emailSent = row[12]; // twelfth column
    if (emailSent != EMAIL_SENT) { // Prevents sending duplicates
      var subject = 'Pemesanan Barang';
      var blob = HtmlService.createHtmlOutputFromFile("emailPembayaran").getBlob();
      var a = blob.getDataAsString();
      //generate random code
      var brand = generateRandom(); 
      //generate order number
      var nomororder = data[i-1][9] + 1;
      var inc =  (data[i-1][10] - 40000*data[i-1][5]) + 1 ;
      if(inc>999) inc=1;
      var total =  40000*(data[i][5]) + inc;
      
      a=a.replace('{{Nama}}',row[1]);
      a=a.replace('{{Nama}}',row[1]);
      a=a.replace('{{Nomor Telepon}}',row[3]);
      a=a.replace('{{Jumlah Pesanan}}',row[5]);
      a=a.replace('{{Kode}}',nomororder);
      a=a.replace('{{Total}}',total);
      //Logger.log(b);
      
      //getRange index is direct index
      sheet.getRange(startRow+i, 10).setValue(nomororder);
      sheet.getRange(startRow+i, 11).setValue(total);
      sheet.getRange(startRow +i, 12).setValue(brand);
      
      MailApp.sendEmail({to: emailAddress,subject: subject,htmlBody: a});
      
      sheet.getRange(startRow +i, 13).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted {to: 'ardhi.rofi@gmail.com',subject: subject,htmlBody: blob.getDataAsString()}
      SpreadsheetApp.flush();
    }
  }
}

// generate random passowrd, original code from Alan Wells, edited
function generateRandom() {
  // some random function to generate code, hidden
  return text;
}


