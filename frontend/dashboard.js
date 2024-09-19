function processSpreadsheet() {
  var fileInput = document.getElementById("fileInput");
  var file = fileInput.files[0];

  if (file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      //converting the file to developer readable json
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: "array" });
      var sheetName = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[sheetName];
      var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Process the spreadsheet data
      var emailMap = mapEmails(jsonData);
      const stringifiedBody = JSON.stringify(emailMap);

      //displaying the data
      const content = document.getElementById("content");
      content.innerText = stringifiedBody;
    };
    reader.readAsArrayBuffer(file);
  } else {
    alert("Please upload a spreadsheet file.");
  }
}

function mapEmails(jsonData) {
  var mappedEmails = {};
  for (let obj of jsonData) {
    if (!obj.length) break;
    else if (obj[4] != "NEW") mappedEmails[obj[4]] = obj[5];
  }
  return mappedEmails;
}
