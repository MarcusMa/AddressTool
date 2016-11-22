console.log("...Start...");

var xlsx = require("node-xlsx");
var fs = require('fs');
var request = require('sync-request');
var querystring = require('querystring');

var queryData = {
    qt:"s",
    c:289,
    wd:"",
    rn:10,
    ie:"utf-8",
    oue:1,
    res:"api"
}

var excelNameSet = [];

fs.readdir('./input', function(err, files) {
    if (err) {
        throw err;
    }
    // 1. get all file name
    for(var k= 0;k<files.length;k++){
        var tmpFileName = files[k].split(".")[0];
        if(""!=tmpFileName)
            excelNameSet.push(tmpFileName);
    }
    console.log(excelNameSet);
    // 2. get all item in the file
    for(var i = 0;i<excelNameSet.length;i++){
        // Read excel
        var excelName = excelNameSet[0];
        var excelPath = "./input/" + excelName +".xlsx";
        console.log("Read the xlsx file: " + excelPath)
        var sheets = xlsx.parse("./" + excelPath );
        console.log(typeof sheets[0]);
        var dataSet = sheets[0].data;
        var rowNum = dataSet.length;
        console.log(dataSet[0].length);
        console.log("All row:" + rowNum);
        dataSet[0].push("lng");
        dataSet[0].push("lat");
        for(var j= 1 ;j<rowNum;j++){
            console.log("Do:"+ j + "/" + rowNum);
            var tempData = dataSet[j];
            var templng = 0; // 经度
            var templat = 0; // 纬度

            queryData.wd = tempData[0].split(" ")[0];
            var searchString = querystring.stringify(queryData);
            var reqPath="http://api.map.baidu.com/?"+ searchString;
            console.log("searchString:" + reqPath);
            try{
                var res = request('GET', reqPath,{
                    timeout:5000
                });
            }catch (e){
                j--;
                continue;
            }
            var data = JSON.parse(res.getBody('utf8'));
            if(data.content!=undefined&&data.content.length >=1){
                templng = getRealLocationData(data.content[0]["navi_x"]); //经度
                templat = getRealLocationData(data.content[0]["navi_y"]);
                console.log(queryData.wd +" : "+ templng +" , "+ templat);
                // 加入数据
                dataSet[j].push(templng);
                dataSet[j].push(templat);
            }
        }
        console.log("Write to the new xlsx file.....")
        var newExcelName = excelName + "_new";
        var newExcelPath = "./output/"+newExcelName + ".xlsx";
        var buffer = xlsx.build([
            {
                name:'sheet1',
                data:dataSet
            }
        ]);
        console.log("New fileName: "+newExcelPath);
        fs.writeFileSync(newExcelPath,buffer,{'flag':'w'});
    }
    console.log("... All End...");

});

function getRealLocationData(point){
    return point/20037508.3427892*180
}
