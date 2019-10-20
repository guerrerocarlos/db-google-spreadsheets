function intToExcelCol(i) {
  i = i + 1;
  var result = "";
  while (i > 0) {
    var modulo = (i - 1) % 26;
    result = String.fromCharCode(65 + modulo) + result;
    i = parseInt((i - modulo) / 26);
  }
  return result;
}

function excelColToInt(colName) {
  var digits = colName.toUpperCase().split(""),
    number = 0;

  for (var i = 0; i < digits.length; i++) {
    number +=
      (digits[i].charCodeAt(0) - 64) * Math.pow(26, digits.length - i - 1);
  }

  return number;
}

module.exports = intToExcelCol;

// if (require.main === module) {
//   var i = 0;
//   while (i < 200) {
//     console.log(i, intToExcelCol(i));
//     i++;
//   }
// }
