var Excel = require("exceljs");

var workbook = new Excel.Workbook();

var keywords = [
  "question",
  "minute",
  "task",
  "workshop",
  "structural",
  "i-405",
  "vernon",
  "ug3",
  "ug4",
  "la brea",
  "hindry",
  "cos",
  "passage",
  "wall",
  "slab"
];
var valuesarray = [];
workbook.xlsx.readFile("Prebid_HNTB_Review.xlsx").then(function() {
  var worksheet = workbook.getWorksheet(1);
  var column = worksheet.getColumn(5);
  console.log(column.length);
  column.eachCell({ includeEmpty: false }, function(cell, rowNumber) {
    var celltext = cell._value.value;
    if (celltext !== null) {
      var cellwords = celltext.split(" ");
      for (i = 0; i <= cellwords.length; i++) {
        if (cellwords[i] !== undefined) {
          var cellwordLc = cellwords[i].toLowerCase();
          for (j = 0; j <= keywords.length; j++) {
            if (keywords[j] === cellwordLc) {
              worksheet.getCell("G" + rowNumber).value = capitalizeFirstLetter(
                cellwordLc
              );
              console.log(capitalizeFirstLetter(cellwordLc));
              console.log(rowNumber);
            }
          }
        }
      }
    }
  });
  workbook.xlsx.writeFile("Prebid_HNTB_Review.xlsx").then(function() {
    console.log("xls file is written.");
  });
});

function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}
