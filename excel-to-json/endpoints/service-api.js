'user strinct';
var express = require('express');
var cfenv = require('cfenv');
var appEnv = cfenv.getAppEnv();
var XLSX = require("xlsx");
var fs = require("fs");
var fileUpload = require('express-fileupload');
var app = express();

app.use(fileUpload());

var server = app.listen(appEnv.port, function () {
    var host = server.address().address;
    var port = server.address().port;
    console.log("Device-api listening at http://%s:%s", host, port);
});

app.post('/uploadFile',  function (req, res, next) {
    // A simple trick, because of bufferArray problem.
    let file = req.files.file;

    let dir = './tmp';

    if (!fs.existsSync(dir)){
        fs.mkdirSync(dir);
    }

    // Use the mv() method to place the file somewhere on your server
    file.mv(dir+ '/file.xlsm', function(err) {
        if (err)
            return res.status(500).send(err);

        getExcelData(dir+ '/file.xlsm',function (fileDataToObject) {
            res.send(fileDataToObject);
        })
    });
});


function getExcelData(file, callback) {
    var workbook = XLSX.readFile(file);

    var first_sheet_name = workbook.SheetNames[0];

    /* Get worksheet */
    var worksheet = workbook.Sheets[first_sheet_name];

    callback(XLSX.utils.sheet_to_json(worksheet));
}
