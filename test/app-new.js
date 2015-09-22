//Important Security advice: 
//don't add multer as a middleware to all requests. 
//If you do this, people will be able to upload files
//in ALL YOUR 'post' handlers!!! 

var Multer = require('multer');
var Parse = require('csv-parse');
var fs = require('fs')

function parseCSVFile(sourceFilePath, columns, onNewRecord, handleError, done){
    var source = fs.createReadStream(sourceFilePath);

    var linesRead = 0;

    var parser = Parse({
        delimiter: ',', 
        columns:columns
    });

    parser.on("readable", function(){
        var record;
        while (record = parser.read()) {
            linesRead++;
            onNewRecord(record);
        }
    });

    parser.on("error", function(error){
        handleError(error)
    });

    parser.on("end", function(){
        done(linesRead);
    });

    source.pipe(parser);
}

//We will call this once Multer's middleware processed the request
//and stored file in req.files.fileFormFieldName

function parseFile(req, res, next){
    var filePath = req.files.file.path;
    console.log(filePath);
    function onNewRecord(record){
        console.log(record)
    }

    function onError(error){
        console.log(error)
    }

    function done(linesRead){
        res.send(200, linesRead)
    }

    var columns = true; 
    parseCSVFile(filePath, columns, onNewRecord, onError, done);

}

//this is the route handler with two middlewares. 
//First:  Multer middleware to download file. At some point,
//this middleware calls next() so process continues on to next middleware
//Second: use the file as you need

app.post('/upload', [Multer({dest:'./uploads'}), parseFile]);