var path = require('path'); 
var fs = require('fs');
var express = require('express');
var app = express();

// Initialize variables. 
var port = process.env.PORT || 8080;  

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname));

// Set up route to get the file names.
app.get('/getfilenames', function (req, res) {

    // Returns an array of file names in the stories directory.
    fs.readdir(path.join(__dirname, 'stories'), function(err, fileList) {

        // Send the file list back to the add-in. Sends the file list back in this
        // form: ["filename1.docx","filename2.docx","filename3.docx"]
        res.send(fileList);    
    });
});

// Set up route to get the file and send it in the response.
// Query parameter: ?filename
app.get('/getfile', function (req, res) {

    // Get the value of the filename query parameter. The expected request form is
    // http://127.0.0.1:8080/getfile?filename=The%20smallest%20story.docx
    var filename = req.query.filename;
    
    // Create path to get the file. 'stories' is the directory where the stories are stored.
    var pathToFile = path.join(__dirname, 'stories', filename);
    
    // Read the file, convert it to base64, and return the base64 file in the response to the add-in.
    fs.readFile(pathToFile, function(err, data) {
        var fileData = new Buffer(data).toString('base64');
        res.send(fileData);
    });
});

// Set the route to the index.html file.
app.get('/', function (req, res) {
	
    var homepage = path.join(__dirname, 'index.html');
    res.sendFile(homepage);
});

// Start the app.  
app.listen(port);
console.log('Listening on port ' + port + '...');
