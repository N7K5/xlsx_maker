
const fs= require("fs");

const excelToJson = require('convert-excel-to-json');

const convert_to_json= require("./convert_to_json");

// const json2xls = require('json2xls');

const filename= "lt.xls";

sourceFile= __dirname + "/"+ filename;


var req_funs= {
    "Name":(obj) => obj.C,
    "Institution": (obj) => "Govt. College Of Engg And Leather Tech",
    "Roll no": (obj) => obj.AI,
    "Gender": (obj) => obj.D,
    "DOB": (obj) => obj.E,
    "10th Marks": (obj) => obj.H,
    "10th Year of Passing": (obj) => obj.I,
    "12th Marks": (obj) => obj.K,
    "12th Year of Passing": (obj) => obj.L,
    "Grad. avg Marks": (obj) => ((obj.W+ obj.X+ obj.Y+ obj.Z+ obj.AA)/5),
    "Grad year of passing": (obj) => 2019,
    "Stream": (obj) => "IT",
    "Degree": (obj) => "B. Tech",
    "BackLogs": (obj) => "N/A",
    "Placement Status": (obj) =>"N/A",
    "Certification": (obj) => "N/A",
    "E-Mail": (obj) => obj.F,
    "Contact": (obj) => obj.G,
}



const result = convert_to_json.convert(sourceFile, req_funs, 2);

console.log(JSON.stringify(result, undefined, 2));

// fs.writeFile("./res.json", JSON.stringify(result, undefined, 2), (err)=> {
//     if(err) {
//         console.log(err);
//     }
// });


// var xls = json2xls(result);

// fs.writeFileSync('result.xlsx', xls, 'binary');
try {
    require('json2xlsx').write("Created_" + Math.floor(Math.random()*1000) + "_" + filename, "xlxs", result);
} catch(e) {
    console.log("\n\n\tCould not create file... Permition problem\n\n");
}
