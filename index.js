
const excelToJson = require('convert-excel-to-json');

sourceFile= __dirname + "/IT_old.xls";

const result = excelToJson({
    sourceFile: sourceFile,
});

console.log(JSON.stringify(result, undefined, 2));