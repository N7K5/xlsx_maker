
const fs= require("fs");

const excelToJson = require('convert-excel-to-json');

const convert_to_json= require("./convert_to_json");

const filename= ["cse.xls", "cse_lat.xls", "it.xls", "it_lat.xls", "lt.xls"];

// sourceFile= __dirname + "/"+ filename;


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

var result= {
    data: [],
};

for(var i=0; i<filename.length; i++) {
    var sourceFile= __dirname +"/" +filename[i];

    if(i==0 || i==1) {
        req_funs.Stream= (obj) => "Computer Science And Engineering";
    }
    else if(i==2 || i==2) {
        req_funs.Stream= (obj) => "Information Technology";
    }
    else if(i==4) {
        req_funs.Stream= (obj) => "Leather Technology";
    }

    var res= convert_to_json.convert(sourceFile, req_funs, 2);

    for(var j=0; j<res.data.length; j++) {
        if(res.data[j].B && res.data[j].B.length<2) {
            console.log(res.data[j]);
            continue;
        }
        result.data.push(res.data[j]);
    }

}


try {
    require('json2xlsx').write("all_combine.xls", "xls", result);
} catch(e) {
    console.log("\n\n\tCould not create file... Permition problem\n\n");
}


// const result = convert_to_json.convert(sourceFile, req_funs, 2);

// console.log(JSON.stringify(result, undefined, 2));

// fs.writeFile("./res.json", JSON.stringify(result, undefined, 2), (err)=> {
//     if(err) {
//         console.log(err);
//     }
// });


// var xls = json2xls(result);

// fs.writeFileSync('result.xlsx', xls, 'binary');
// try {
//     require('json2xlsx').write("Created_" + Math.floor(Math.random()*1000) + "_" + filename, "xlxs", result);
// } catch(e) {
//     console.log("\n\n\tCould not create file... Permition problem\n\n");
// }
