const excelToJson = require('convert-excel-to-json');


var convert= function(file, req_funs, starting_row) {
    const ini_tot_result = excelToJson({
        sourceFile: file,
    });

    var result= {
        data: [],
    }

    var blocks= [ "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
                     "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT"];

    var ori_res= ini_tot_result[Object.keys(ini_tot_result)[0]];
    
    var property_len= Object.keys(req_funs).length;

    var obj= {A: "S. No"};
    for(var i=0; i<property_len; i++) {
        obj[blocks[i]]= Object.keys(req_funs)[i];
    }
    result.data.push(obj);
    
    
    for(var cur= 0; cur<ori_res.length; cur++) {
        var index= cur+1;
        if(starting_row) {
            if(cur<starting_row) {
                continue;
            }
            index-=starting_row;
        }
        var stud= ori_res[cur];
        // console.log("\n\n"+JSON.stringify(stud, undefined, 2)+ "\n\n");
        obj= {"A": index};

        for(var i=0; i<property_len; i++) {
            var fn= req_funs[Object.keys(req_funs)[i]];
            if(typeof(fn)!= "function") {
                obj[blocks[i]]= fn;
            }
            else {
                let tmp_res= fn(stud);
                if(tmp_res== null || tmp_res== undefined) {
                    obj[blocks[i]]= " ";
                }
                else {
                    if(typeof(tmp_res)== "number") {
                        tmp_res= round(tmp_res, 2);
                    }
                    obj[blocks[i]]= tmp_res;
                }
            }
        }
        result.data.push(obj);
    }

    return result;

}

function round(value, decimals) {
    return Number(Math.round(value+'e'+decimals)+'e-'+decimals);
}
  
  

module.exports= {
    convert,
}