// Assumptions
// 1. The sheet with the students is called Sheet1
// 2. The certificate name placeholder is 'Name PlaceHolder'

var studentsId = "16FCTdyJ2ONyXPZvm6fn_o8KP6zjDlSyThXrvpMrrnVU";
var sheetName = "Sheet1";
var namePlaceholder = "Name PlaceHolder";
var courses = [
  {
    name: "Web App Development with PHP & MySQL",
    slides: "1ZQCC-o1f_29U6YLpaE2wpsCBkE-HpD8Ka83hiAdGNh0",
  },
  {
    name: "Web App Development with PHP & MySQL",
    slides: "1ZQCC-o1f_29U6YLpaE2wpsCBkE-HpD8Ka83hiAdGNh0",
  },
];

function main() {
  const getByName = (colName, sheetName, id) => {
    var sheet = SpreadsheetApp.openById(id).getSheetByName(sheetName);
    var data = sheet.getRange("A1:1").getValues();
    var col = data[0].indexOf(colName);
    if (col != -1) {
      var unfiltered = sheet.getRange(2, col + 1, sheet.getMaxRows()).getValues();
      var arrFiltered = unfiltered.filter(function (x) {
        return !x.every((element) => element === (undefined || null || ""));
      });
      return arrFiltered.filter(function (x) {
        return x.toString();
      });
    }
  };

  const duplicateSlides = (certSlide, students) => {
    certSlide.forEach(function (slide) {
      for (let i = 1; i < students.length; i++) {
        slide.duplicate();
      }
    });
  };

  const fillData = (certs, students) => {
    for (var i = 0; i < students.length; i++) {
      var student = students[i].toString();
      var shapes = certs[i].getShapes();
      shapes.forEach(function (shape) {
        shape.getText().replaceAllText(namePlaceholder, student);
      });
    }
  };

  for (const course of courses) {
    let certs = SlidesApp.openById(course.slides).getSlides();
    let students = getByName(course.name, sheetName, studentsId);
    duplicateSlides(certs, students);
    fillData(certs, students);
  }
}
