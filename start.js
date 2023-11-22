// Require library
var excel = require('excel4node');
const oracledb = require('oracledb');
const nodemailer = require('nodemailer');
fs = require('fs');
var CronJob = require('cron').CronJob;
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const { readFileSync } = require('fs');
require('log-timestamp');

const app = express();
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

const port = 3000;

let healthy = true;

let DB_USER = process.env.DB_USER || '/var/config/DB_USER';
if (!path.isAbsolute(DB_USER)) { DB_USER = path.resolve(__dirname, DB_USER); }
let hasDB_USER = fs.existsSync(DB_USER);
app.locals.hasDB_USER = hasDB_USER;

let DB_PASSWORD = process.env.DB_PASSWORD || '/var/secret/DB_PASSWORD';
if (!path.isAbsolute(DB_PASSWORD)) { DB_PASSWORD = path.resolve(__dirname, DB_PASSWORD); }
let hasDB_PASSWORD = fs.existsSync(DB_PASSWORD);
app.locals.hasDB_PASSWORD = hasDB_PASSWORD;

let DB_CONNECTION_STRING = process.env.DB_CONNECTION_STRING || '/var/config/DB_CONNECTION_STRING';
if (!path.isAbsolute(DB_CONNECTION_STRING)) { DB_CONNECTION_STRING = path.resolve(__dirname, DB_CONNECTION_STRING); }
let hasDB_CONNECTION_STRING = fs.existsSync(DB_CONNECTION_STRING);
app.locals.hasDB_CONNECTION_STRING = hasDB_CONNECTION_STRING;

let EMAIL_PASSWORD = process.env.EMAIL_PASSWORD || '/var/secret/EMAIL_PASSWORD';
if (!path.isAbsolute(EMAIL_PASSWORD)) { EMAIL_PASSWORD = path.resolve(__dirname, EMAIL_PASSWORD); }
let hasEMAIL_PASSWORD = fs.existsSync(EMAIL_PASSWORD);
app.locals.hasEMAIL_PASSWORD = hasEMAIL_PASSWORD;



app.listen(port, () => console.log(`Report4C starded and listening on port ${port}!`));

if(hasEMAIL_PASSWORD && hasDB_PASSWORD){

    EMAIL_PASSWORD = fs.readFileSync(EMAIL_PASSWORD,{ encoding: 'utf8', flag: 'r' });
    DB_PASSWORD = fs.readFileSync(DB_PASSWORD,{ encoding: 'utf8', flag: 'r' });

    /*
    fs.readFile(EMAIL_PASSWORD, "utf8", function (err, contents) {
        if (err) {
            console.error('secret not found');
            console.error('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
            EMAIL_PASSWORD = contents;
            console.log("- Found Email Password");
        }
    });
    
    fs.readFile(DB_PASSWORD, "utf8", function (err, contents) {
        if (err) {
            console.error('secret not found');
            console.error('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
            DB_PASSWORD = contents;
            console.log("- Found DB Password");
        }
    });*/


}else{
    console.log("Please check your secret configuration. Variable or bind not setted.");
}

if(hasDB_USER && hasDB_CONNECTION_STRING){

    DB_USER = fs.readFileSync(DB_USER,{ encoding: 'utf8', flag: 'r' });
    DB_CONNECTION_STRING = fs.readFileSync(DB_CONNECTION_STRING,{ encoding: 'utf8', flag: 'r' });
    /*
    fs.readFile(DB_USER, "utf8", function (err, contents) {
        if (err) {
            console.error('secret not found');
            console.error('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
            DB_USER = contents;
            console.log("- Found DB User");
        }
    });
    fs.readFile(DB_CONNECTION_STRING, "utf8", function (err, contents) {
        if (err) {
            console.error('secret not found');
            console.error('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
            DB_CONNECTION_STRING = contents;
            console.log("- Found DB Connection String");
        }
    });*/
}else{
    console.log("Please check your config map. Variable or bind not setted.");
}

console.log(EMAIL_PASSWORD);
console.log(DB_PASSWORD);
console.log(DB_USER);
console.log(DB_CONNECTION_STRING);



function healthStatus() {
    if (healthy) {
      return "Service is UP";
    } else {
      return "Service is DOWN";
    }
  }



// Create a new instance of a Workbook class
var workbook = new excel.Workbook();
var workbookPAZ = new excel.Workbook();

// Add Worksheets to the workbook
var worksheetPAZ = workbookPAZ.addWorksheet('FARMACI_PAZIENTE');

// Create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#000000',
    size: 10
  }});

var stylePAZ = workbookPAZ.createStyle({
  font: {
    color: '#000000',
    size: 10
  }});
  


async function generaReportTerapia(reparto,res) {
  //console.log("nosologico: "+reparto);
  let ts = Date.now();
  
  let date_ob = new Date(ts);
  let date = date_ob.getDate();
  let month = date_ob.getMonth() + 1;
  let year = date_ob.getFullYear();  
  let hour = date_ob.getHours();
  let minutes = date_ob.getMinutes();
  
  let ts1 = Date.now();
  
  let date_ob1 = new Date(ts1);
  date_ob1.setDate(date_ob1.getDate()-180);
  
  let date1 = date_ob1.getDate();
  let month1 = date_ob1.getMonth() + 1;
  let year1 = date_ob1.getFullYear();  
  let hour1 = date_ob1.getHours();
  let minutes1 = date_ob1.getMinutes();
  
  let connection;
    try {
      connection = await oracledb.getConnection({ user: DB_USER, password: DB_PASSWORD, connectionString: DB_CONNECTION_STRING });

      //Primo Foglio
      result = await connection.execute(
          `SELECT to_char(extension) extension,
                  to_char(item_code) item_code,
                  to_char(item_desc) item_desc,
                  CASE WHEN (NVL(qty_uom,'')) is NULL then ' ' ELSE TO_CHAR(NVL(qty_uom,'')) END qty_uom,
                  CASE WHEN (NVL(qty,'')) is NULL then ' ' ELSE TO_CHAR(NVL(qty,'')) END qty,
                  TO_CHAR(administered_start, 'YYYY-MM-DD HH24:MM') administered_start,
                  to_char(route_desc) route_desc 
            FROM V_SOMM_PAZ_NOS WHERE extension = '`+reparto+`'`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
    
        const rs = result.resultSet;
        let row;
        let riga = 1;
  
        worksheetPAZ.cell(riga,1).string('NOSOLOGICO').style(stylePAZ);
        worksheetPAZ.cell(riga,2).string('AIC').style(stylePAZ);
        worksheetPAZ.cell(riga,3).string('FARMACO').style(stylePAZ);
        worksheetPAZ.cell(riga,4).string('UNITA').style(stylePAZ);
        worksheetPAZ.cell(riga,5).string('QUANTITA').style(stylePAZ);
        worksheetPAZ.cell(riga,6).string('INIZIO SOMMINISTRAZIONE').style(stylePAZ);
        worksheetPAZ.cell(riga,7).string('VIA DI SOMMINISTRAZIONE').style(stylePAZ);
  
        riga++;
  
  
        while ((row = await rs.getRow())) {
          //console.log(riga);  
          //console.log(row);
          //console.log(row.ISTITUTO);
          worksheetPAZ.cell(riga,1).string(row.EXTENSION).style(stylePAZ);
          worksheetPAZ.cell(riga,2).string(row.ITEM_CODE).style(stylePAZ);
          worksheetPAZ.cell(riga,3).string(row.ITEM_DESC).style(stylePAZ);
          worksheetPAZ.cell(riga,4).string(row.QTY_UOM).style(stylePAZ);
          worksheetPAZ.cell(riga,5).string(row.QTY).style(stylePAZ);
          worksheetPAZ.cell(riga,6).string(row.ADMINISTERED_START).style(stylePAZ);
          worksheetPAZ.cell(riga,7).string(row.ROUTE_DESC).style(stylePAZ);
          riga++;
        }
  
        //workbook.write('statistiche.xlsx', res);
        workbookPAZ.write(reparto+" "+ date + "-" + month + "-" + year+" ore " + hour+"-" + minutes+".xlsx", res);
/*
        workbook.writeToBuffer().then(function(buffer) {
          //console.log(buffer);
          console.log(102);
          res.send(buffer);
         
        });
  
        */
  
    
        await rs.close();
        
    } catch (err) {
      console.error(err);
    } finally {
      if (connection) {
        try {
          await connection.close();
        } catch (err) {
          console.error(err);
        }
      }
    }
  }



/* Chiamata REST API per generazione excel terapia */
app.get('/', (req, res) => {
  res.send('STATS is up!');
});


app.engine('html', require('ejs').renderFile);
app.set('view engine', 'html');

app.get('/ward', (req, res) => {
  
  console.log("Richiesta ricevuta per report reparto:");
  console.log(req.query);

  const unitCode = req.query.unitCode;
  const idUser = req.query.idUser;

  //generaReportTerapia(req.params.wsd1+'/'+req.params.wsd2, res);
  //res.sendFile(__dirname + "/index.html");
  res.render(path.join(__dirname, '/', 'ward.html'), {unitCode: unitCode, idUser: idUser });

});

app.get('/encounter', (req, res) => {
  
  console.log("Richiesta ricevuta per report paziente:");
  console.log(req.query);
  generaReportTerapia(req.query.encounterCode, res);
});


app.get('/health', function(request, response) {
    //console.log("Invocazione health");
    if( healthy ) {
      response.status(200);
    } else {
      response.status(500);
    }
    let status = healthStatus();
    response.send(status);
  });  
