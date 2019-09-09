const fs = require('fs');
const express = require('express');
const fileUpload = require('express-fileupload');
const Duplex = require('stream').Duplex;  
const XLSX = require('xlsx');

const app = express();
const port = 3000;

const output_file_name = 'out.csv';

function handleWorkbook(wb) {
  //  get sheet from wb
  const sheet = Object.values(wb.Sheets).shift();
  // convert sheet to JS Object
  const sheetObj = XLSX.utils.sheet_to_json(sheet);
  // add col data
  const result = sheetObj.map(obj => {
    
    obj['cat_name'] = 'cat snack';
    obj['iso_date'] = new Date().toISOString();

    return obj;
  })

  const newSheet = XLSX.utils.json_to_sheet(result);
  // convert to csv
  const csv = XLSX.stream.to_csv(newSheet);
  
  // write to disk
  csv.pipe(fs.createWriteStream(output_file_name));
}

function process_RS(stream/*:ReadStream*/, cb/*:(wb:Workbook)=>void*/)/*:void*/{
  var buffers = [];
  stream.on('data', function(data) { buffers.push(data); });
  stream.on('end', function() {
    var buffer = Buffer.concat(buffers);
    var workbook = XLSX.read(buffer, {type:"buffer"});

    /* DO SOMETHING WITH workbook IN THE CALLBACK */
    cb(workbook);
  });
}

function bufferToStream(buffer) {  
  const stream = new Duplex();
  stream.push(buffer);
  stream.push(null);
  return stream;
}
// default options

app.post('/upload', fileUpload(), (req, res) => {
  const st = bufferToStream(req.files.example.data);

  process_RS(st, handleWorkbook);

  res.send('Hello World!');
});

app.listen(port, () => console.log(`Example app listening on port ${port}!`))