function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function fetchData(rollNumber) {
  var ss = SpreadsheetApp.openById("// YOUR ID  ////");
  var sheet = ss.getSheetByName("Data");
  var lastRow = sheet.getLastRow();
  var htmlData = "";
  var paperName = ""; 
  var semesterName = ""; 

  for (var i = 2; i <= lastRow; i++) {
    var currentRollNumber = sheet.getRange(i, 6).getValue(); 
    if (currentRollNumber == rollNumber) {
      paperName = sheet.getRange(i, 7).getValue(); 
      semesterName = sheet.getRange(i, 5).getValue();
      break; 
    }
  }

  htmlData += "<div class='container mt-5'>";
 
  htmlData += "<div class='card shadow'>";
  htmlData += "<div class='card-header bg-primary text-white'>";
 htmlData += "<div style='display: flex; justify-content: center;'>";
htmlData += "<h4 class='mb-0'>ASTOE COLLEGE : COLLEGE OF LAHORE</h4>";
htmlData += "</div>";

  htmlData += "</div>";
  htmlData += "<div style='display: flex; justify-content: center;'>";
  htmlData += "<img src='https://cdn.bio.link/uploads/profile_pictures/2023-04-16/LtkOkQHLljUSvYcfJpKC8NRFH1q3MfXZ.png' alt='Sample Logo' style='width: 150px; height: auto; margin-bottom: 20px;'>";
  htmlData += "</div>";

  htmlData += "<div class='card-body'>";
  htmlData += "<div id='table-data' class='table-responsive mt-3'>";
htmlData += "<div style='display: flex; justify-content: space-between;'>";
htmlData += "<p><strong>UPC :</strong> " + rollNumber + "</p>";
htmlData += "<p><strong>Date:</strong> " + new Date().toLocaleDateString() + "</p>";
htmlData += "</div>";
htmlData += "<div style='display: flex; justify-content: space-between;'>";
htmlData += "<p><strong>PaperName:</strong> " + paperName + "</p>";
htmlData += "<p><strong>Semester:</strong> " + semesterName + "</p>";
htmlData += "</div>";
  htmlData += "<table class='table table-bordered table-striped excel-table'>";
  htmlData += "<thead>";
  htmlData += "<tr>";
  htmlData += "<th>SrNo</th>";
  htmlData += "<th>Student Name</th>";
  htmlData += "<th>Exam Roll Number</th>";
  htmlData += "<th>Program Name</th>";
  htmlData += "<th>Max Marks</th>";
  htmlData += "<th>Obtained Marks</th>";
  htmlData += "<th>Signature</th>";
  htmlData += "</tr>";
  htmlData += "</thead>";
  htmlData += "<tbody>";

  // Loop through the rows to find matching roll number and generate table rows
  for (var i = 2; i <= lastRow; i++) {
    var currentRollNumber = sheet.getRange(i, 6).getValue(); // Assuming Roll Number is in column F
    if (currentRollNumber == rollNumber) {
      var serialNumber = i - 1; // Serial number starting from 1
      var studentName = sheet.getRange(i, 1).getValue(); // Student Name from column A
      var examRollNumber = sheet.getRange(i, 2).getValue(); // Exam Roll Number from column B
      var progName = sheet.getRange(i, 4).getValue();
      var maxmarks = sheet.getRange(i, 8).getValue();
      var omarks = sheet.getRange(i, 9).getValue();
      var sig = sheet.getRange(i, 10).getValue(); // Program Name from column D
      // Append table row to HTML data
      htmlData += "<tr>";
      htmlData += "<td>" + serialNumber + "</td>";
      htmlData += "<td>" + studentName + "</td>";
      htmlData += "<td>" + examRollNumber + "</td>";
      htmlData += "<td>" + progName + "</td>";
      htmlData += "<td>" + maxmarks + "</td>";
      htmlData += "<td>" + omarks + "</td>";
      htmlData += "<td>" + sig + "</td>";
      htmlData += "</tr>";
    }
  }

  htmlData += "</tbody>";
  htmlData += "</table>";
  htmlData += "</div>"; // Close table-responsive
  htmlData += "</div>"; // Close card-body
  htmlData += "</div>"; // Close card
  htmlData += "</div>"; // Close container

  // If no matching roll number found
  if (htmlData == "") {
    htmlData = "<div class='container mt-5 text-center'>" +
      "<div class='alert alert-danger' role='alert'>" +
      "<strong>Error:</strong> Roll Number not found!" +
      "</div>" +
      "</div>";
  }

  return htmlData;
}

