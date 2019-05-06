var xlsx = require('node-xlsx').default;
var fs = require('fs');

var directoryPath = `${__dirname}/xlsx/import`;
var rows = [['#', 'File name', 'Sheet name', 'Total rows', 'White rows']];
var index = 1;

fs.readdir(directoryPath, function (err, files) {
    if (err) {
        console.error(err);
    }
    files.forEach(function (file) {
        var worksheet = xlsx.parse(fs.readFileSync(`${directoryPath}/${file}`));
        var worksheetSize = worksheet.length;
        var sheetData = [];
        var sheetName;
        var row, i, totalRow, whiteRow;
        for (i = 2; i < worksheetSize; i++) {
            sheetName = worksheet[i].name;
            sheetData = worksheet[i].data;
            totalRow = sheetData.length;
            whiteRow = 0;
            for (row = 0; row < totalRow; row++) {
                if (sheetData[row].length === 0) {
                    ++whiteRow;
                }
            }
            if (i === 2) {
                rows.push([index, file, sheetName, totalRow, whiteRow]);
            } else {
                rows.push([index, null, sheetName, totalRow, whiteRow]);
            }
            index++;
        }
    });

    const option = {'!cols': [{ wch: 5 }, { wch: 50 }, { wch: 20 }, { wch: 10 }, { wch: 10 }]};
    var buffer = xlsx.build([{name: "result", data: rows}], option);

    var exportFileName = `${__dirname}/xlsx/export/result_${new Date().getTime()}.xlsx`;

    fs.writeFile(exportFileName, buffer, function(err) {
        if(err) {
            return console.log(err);
        }
        console.log(`File was export to: ${exportFileName}`);
    }); 
});