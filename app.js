const express = require('express');
const upload = require('express-fileupload');
const Excel = require('exceljs');
const fetch = require('node-fetch');

const port = 5000;
const app = express();
app.use(upload())


async function postController(req , res){
  if(req.files){

    var file = req.files.product;
    var filename = file.name;

    file.mv('./uploads/'+filename , function(err){
      if(err){
        res.send(err);
      }});
    await readWrite(filename);
    res.statusMessage = "successfull added price to the product's excel sheet";
     res.download(__dirname + '/uploads/'+filename);
  }
}


async function readWrite(_filename){
  var filename = './uploads/'+_filename;

  var workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filename);
  let worksheet = workbook.getWorksheet("Sheet1");

  console.log(worksheet.getRow('2').getCell('A').value);

  var api = 'https://api.storerestapi.com/products/';
  for(let i = 2; i <= worksheet._rows.length; i++){
    var productId = worksheet.getRow(i).getCell('A').value;
    await fetch(api + productId)
      .then(response => response.json())
      .then(json => {
         worksheet.getRow(i).getCell('B').value = json.data.price;
         worksheet.getRow(i).commit();
         console.log(worksheet.getRow(i).getCell('B').value);
       })
      .catch(err => console.error(err));
  }
  await workbook.xlsx.writeFile('uploads/product_list.xlsx');
}
app.get('/' , (req , res) => {
  res.sendFile(__dirname + '/home.html');
});
app.post('/' , postController);

app.listen(port , function() {
  console.log(`Express server running at ${port}`)
});
