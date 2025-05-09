function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

// Validasi login berdasarkan ID dan password
function validateLogin(id, password) {
  const sheet = SpreadsheetApp.openById('1sn_oxGo8zf9xvc0OzQOly20IUkMeO7cGWmE1yiZ4fOk')
                              .getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id && data[i][3] == password) { // password di kolom D (index 3)
      return {
        valid: true,
        id: data[i][0],
        name: data[i][1],
        class: data[i][2]
      };
    }
  }
  return { valid: false };
}

// Ambil header bulan + data pembayaran berdasarkan ID
function getStudentMonthlyData(id) {
  const sheet = SpreadsheetApp.openById('1sn_oxGo8zf9xvc0OzQOly20IUkMeO7cGWmE1yiZ4fOk')
                              .getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].slice(4, 16); // kolom C - O (bulan)

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const pembayaran = data[i].slice(4, 16); // data pembayaran siswa
      return { bulan: headers, pembayaran: pembayaran };
    }
  }
  return null;
}
