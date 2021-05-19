const express = require('express');
const app = express();
// port
const port = process.env.PORT || 3000;

// file sytem
const fs = require('fs');

// json to excel convertor library
const json2xls = require('json2xls');

// json cheking library
const isJson = require('is-json');

// json paser inbuilt in node
const bodyParser = require('body-parser');
const { json } = require('express');

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));

// template engine
app.set('view engine', 'ejs');

//routes
app.get('/', (req, res)=>{
    res.render('index', {title:`json to excel data file convertor`});
})

// post request
app.post('/jsontoexcel', (req, res)=>{
    
    let jsondata = req.body.json;

    let exceloutput = Date.now() + "excelfile.xlsx";

    if(isJson(jsondata)){
        
        let xls = json2xls(JSON.parse(jsondata));
        fs.writeFileSync(exceloutput, xls, 'binary');
        // download file into excel
        res.download(exceloutput, (error)=>{
            if(error){
            fs.unlinkSync(exceloutput);
            res.send('unable to download the excel file');
            }
            fs.unlinkSync(exceloutput);
        })
    }else{
        res.send("json data is not valid");
    }

})


// port listen
app.listen(port, ()=>{
    console.log(`port runing on ${port}`);
})