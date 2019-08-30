var Excel = require("exceljs");

var workbook = new Excel.Workbook();

// Keywords to Search for
var keywords = [
  "question",
  "questions",
  "minute",
  "minutes",
  "task",
  "tasks",
  "workshop",
  "structural",
  "i-405",
  "vernon",
  "ug3",
  "ug4",
  "la brea",
  "labrea",
  "hindry",
  "cos",
  "passage",
  "wall",
  "slab",
  "memo",
  "memos",
  "rebar",
  "cidh",
  "west",
  "soe",
  "ug1",
  "expo",
  "greenline"
];

// Reading and overwriting the xlsx file
workbook.xlsx.readFile("Prebid_HNTB_Review.xlsx").then(function() {
  //The File name you want to access

  // Worksheet by number
  var worksheet = workbook.getWorksheet(1);

  // Column by number you want to go through
  var column = worksheet.getColumn(5);
  console.log(column.length);

  // Going through the Cells
  column.eachCell({ includeEmpty: false }, function(cell, rowNumber) {
    var celltext = cell._value.value;
    if (celltext !== null) {
      var cellwords = celltext
        .split(" ")
        .join(",")
        .split(".")
        .join(",")
        .split(")")
        .join(",")
        .split("_")
        .join(",")
        .split("-")
        .join(",")
        .split(",");
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
  // Writing the new information to the file again
  workbook.xlsx.writeFile("Prebid_HNTB_Review.xlsx").then(function() {
    console.log("xls file is written.");
  });
});

function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}
