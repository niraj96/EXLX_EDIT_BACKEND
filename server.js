const express = require('express');
const formidabe = require('express-formidable');
const excelToJson = require('convert-excel-to-json');
const cors = require('cors');
var Excel = require('exceljs');
const fs = require('fs');
 
const app = express();
const PORT = 4000;

const corsOptions = {
    origin: "*"
};
  
app.use(cors(corsOptions));
app.use(formidabe({
    encoding: 'utf-8',
    uploadDir: 'uploadDir',
    multiples: true
}));

app.post('/json', (req, res)=>{

    let fileArr = req.files.excels;
    let jsonArr = [];

    for(let i=0; i<fileArr.length; i++){

        let file = fileArr[i];

        const result = excelToJson({
            source: fs.readFileSync(file.path)
        });

        fs.unlinkSync(file.path);
        jsonArr.push({name:file.name, data:result});

    }
  
    res.send(jsonArr);

});

app.post('/generate',(req,res)=>{

    let config = req.fields.config;
    let dataset = req.fields.dataset;

    let title = dataset['name'];
    let data =  dataset['data'];
   
    let worksheetArr = Object.keys(data);
    let workbook = new Excel.Workbook();

    worksheetArr.forEach((sheet, sheetIndex)=>{
        let worksheet = workbook.addWorksheet(sheet);
        let fieldData = data[sheet];
        fieldData.forEach((rowdata,i)=>{
          
            let index = i+1;
            let colArr = Object.keys(rowdata);
            let dataField = [];
            colArr.forEach(col=>{
                dataField.push(rowdata[col]);
            });

            worksheet.addRow(dataField);
           
        });

       


        if(sheet==config.subsheet){
            console.log(sheet);
            worksheet.getCell(config.col+ (parseInt(config.row)+1)).value = config.fieldName;
            let RowCount = worksheet.getColumn(config.col).values.length;

            for(let j=(parseInt(config.row)+2); j<RowCount; j++){
            
                let rowNumber = j;
                let cVal = worksheet.getCell(config.col+rowNumber).value;
                let newval = '';
                if(config.dataType=='string'){
                    newval = cVal.toString();
                }else if(config.dataType=='number'){
                    newval = parseInt(cVal);
                }else if(config.dataType=='date'){
                    newval = new Date(cVal);
                }

                worksheet.getCell(config.col+rowNumber).value = newval;

            }

        }
        
       

        

    })

   
    
    // res is a Stream object
    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
        "Content-Disposition",
        "attachment; filename=" + title
    );
    
    return workbook.xlsx.write(res).then(function () {
        res.status(200).end();
    });


})
app.listen(process.PORT || PORT, ()=> console.log('server started on port'+PORT))
