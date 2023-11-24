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

app.use('/css', express.static(path.join(__dirname, 'node_modules/bootstrap/dist/css')));
app.use('/js', express.static(path.join(__dirname, 'node_modules/bootstrap/dist/js')));
app.use('/js', express.static(path.join(__dirname, 'node_modules/jquery/dist')));
app.use(express.static(__dirname + '/public'));
app.use(cors())

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
      //http://report4c-devade.apps.ccedr.gsd.local/encounter/?encounterCode=23DEG000009%40030122%2FPSD_PSD&idUser=p4cdoctor&unitCode=PSD_PSD_071
      //http://localhost:3000/encounter/?encounterCode=23DEG000009%40030122%2FPSD_PSD&idUser=p4cdoctor&unitCode=PSD_PSD_071
      /*
      SELECT to_char(extension) extension,
                  to_char(item_code) item_code,
                  to_char(item_desc) item_desc,
                  CASE WHEN (NVL(qty_uom,'')) is NULL then ' ' ELSE TO_CHAR(NVL(qty_uom,'')) END qty_uom,
                  CASE WHEN (NVL(qty,'')) is NULL then ' ' ELSE TO_CHAR(NVL(qty,'')) END qty,
                  CASE WHEN TO_CHAR(NVL(qty_uom,'')) IN ('cpr') THEN  TO_CHAR(ROUND(qty,0)) ELSE TO_CHAR(NVL(qty,''))  END qty_arr,                  
                  TO_CHAR(administered_start, 'YYYY-MM-DD HH24:MM') administered_start,
                  to_char(route_desc) route_desc 
            FROM V_SOMM_PAZ_NOS WHERE extension = '23DEG000009@030122/PSD_PSD'
      */
      //Primo Foglio
      /*result = await connection.execute(
          `SELECT to_char(extension) extension,
                  to_char(item_code) item_code,
                  to_char(item_desc) item_desc,
                  CASE WHEN (NVL(qty_uom,'')) is NULL then ' ' ELSE TO_CHAR(NVL(qty_uom,'')) END qty_uom,
                  CASE WHEN (NVL(qty,'')) is NULL then ' ' ELSE TO_CHAR(NVL(qty,'')) END qty,
                  TO_CHAR(administered_start, 'YYYY-MM-DD HH24:MM') administered_start,
                  to_char(route_desc) route_desc 
            FROM V_SOMM_PAZ_NOS WHERE extension = '`+reparto+`'`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });*/

          result = await connection.execute(
            `SELECT
            to_char(struttura) struttura,
            to_char(codice_reparto_assistenziale) codice_reparto_assistenziale,
            to_char(reparto_assistenziale) reparto_assistenziale,
            to_char(codice_reparto_giuridico) codice_reparto_giuridico,
            to_char(reparto_giuridico) reparto_giuridico,
            to_char(id_people) id_people,
            to_char(EXTENSION) EXTENSION,
            to_char(data_inizio_prescrizione) data_inizio_prescrizione,
            to_char(data_fine_prescrizione) data_fine_prescrizione,
            to_char(codice_farmaco_prescritto) codice_farmaco_prescritto,
            to_char(descrizione_farmacto_prescritto) descrizione_farmacto_prescritto,
            to_char(forma_farmaceutica_prescritta) forma_farmaceutica_prescritta,
            to_char(codice_farmaco_somministrato) codice_farmaco_somministrato,
            to_char(descrizione_farmacto_somministrato) descrizione_farmacto_somministrato,
            CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
            CASE WHEN (NVL(quantita,'')) is NULL then ' ' ELSE TO_CHAR(NVL(quantita,'')) END quantita,
            to_char(stato) stato,
            to_char(data_inizio_somministrazione_pianificata) data_inizio_somministrazione_pianificata,
            CASE WHEN (NVL(data_inizio_somministrazione_efettuata,'')) is NULL then 'x' ELSE TO_CHAR(NVL(data_inizio_somministrazione_efettuata,'')) END data_inizio_somministrazione_efettuata,
            to_char(route_desc) route_desc
        FROM
            v_somm_paz_nos_v2 WHERE extension = '`+reparto+`'`,
            [],
            { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });

    
        const rs = result.resultSet;
        let row;
        let riga = 1;
  
        worksheetPAZ.cell(riga,1).string('STRUTTURA').style(stylePAZ);
        worksheetPAZ.cell(riga,2).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZ.cell(riga,3).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZ.cell(riga,4).string('CODICE_REPARTO_GIURIDICO').style(stylePAZ);
        worksheetPAZ.cell(riga,5).string('REPARTO_GIURIDICO').style(stylePAZ);
        worksheetPAZ.cell(riga,6).string('ID_PEOPLE').style(stylePAZ);
        worksheetPAZ.cell(riga,7).string('EXTENSION').style(stylePAZ);
        worksheetPAZ.cell(riga,8).string('DATA_INIZIO_PRESCRIZIONE').style(stylePAZ);
        worksheetPAZ.cell(riga,9).string('DATA_FINE_PRESCRIZIONE').style(stylePAZ);
        worksheetPAZ.cell(riga,10).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
        worksheetPAZ.cell(riga,11).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
        worksheetPAZ.cell(riga,12).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
        worksheetPAZ.cell(riga,13).string('CODICE_FARMACO_SOMMINISTRATO').style(stylePAZ);
        worksheetPAZ.cell(riga,14).string('DESCRIZIONE_FARMACTO_SOMMINISTRATO').style(stylePAZ);        
        worksheetPAZ.cell(riga,15).string('UNITA_DI_MISURA').style(stylePAZ);
        worksheetPAZ.cell(riga,16).string('QUANTITA').style(stylePAZ);
        worksheetPAZ.cell(riga,17).string('STATO').style(stylePAZ);
        worksheetPAZ.cell(riga,18).string('DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA').style(stylePAZ);
        worksheetPAZ.cell(riga,19).string('DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA').style(stylePAZ);
        worksheetPAZ.cell(riga,20).string('ROUTE_DESC').style(stylePAZ);
 
  
  
        riga++;
  
  
        while ((row = await rs.getRow())) {
          //console.log(riga);  
          //console.log(row);
          //console.log(row.ISTITUTO);
          worksheetPAZ.cell(riga,1).string(row.STRUTTURA).style(stylePAZ);
          worksheetPAZ.cell(riga,2).string(row.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZ.cell(riga,3).string(row.REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZ.cell(riga,4).string(row.CODICE_REPARTO_GIURIDICO).style(stylePAZ);
          worksheetPAZ.cell(riga,5).string(row.REPARTO_GIURIDICO).style(stylePAZ);
          worksheetPAZ.cell(riga,6).string(row.ID_PEOPLE).style(stylePAZ);
          worksheetPAZ.cell(riga,7).string(row.EXTENSION).style(stylePAZ);
          worksheetPAZ.cell(riga,8).string(row.DATA_INIZIO_PRESCRIZIONE).style(stylePAZ);
          worksheetPAZ.cell(riga,9).string(row.DATA_FINE_PRESCRIZIONE).style(stylePAZ);
          worksheetPAZ.cell(riga,10).string(row.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
          worksheetPAZ.cell(riga,11).string(row.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
          worksheetPAZ.cell(riga,12).string(row.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
          worksheetPAZ.cell(riga,13).string(row.CODICE_FARMACO_SOMMINISTRATO).style(stylePAZ);
          worksheetPAZ.cell(riga,14).string(row.DESCRIZIONE_FARMACTO_SOMMINISTRATO).style(stylePAZ);
          worksheetPAZ.cell(riga,15).string(row.UNITA_DI_MISURA).style(stylePAZ);
          worksheetPAZ.cell(riga,16).string(row.QUANTITA).style(stylePAZ);
          worksheetPAZ.cell(riga,17).string(row.STATO).style(stylePAZ);
          worksheetPAZ.cell(riga,18).string(row.DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA).style(stylePAZ);
          worksheetPAZ.cell(riga,19).string(row.DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA).style(stylePAZ);
          worksheetPAZ.cell(riga,20).string(row.ROUTE_DESC).style(stylePAZ);
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


async function generaReportReparto(dati,res) {
  console.log("dati passati: ");
  console.log(dati.dataIniziale);
  console.log(dati.dataFinale);
  console.log(dati.funzione);
  console.log(dati.idUser);
  console.log(dati.unitCode);


  
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
      //http://report4c-devade.apps.ccedr.gsd.local/encounter/?encounterCode=23DEG000009%40030122%2FPSD_PSD&idUser=p4cdoctor&unitCode=PSD_PSD_071
      //http://localhost:3000/encounter/?encounterCode=23DEG000009%40030122%2FPSD_PSD&idUser=p4cdoctor&unitCode=PSD_PSD_071
      

       if (dati.funzione == 'pazienti'){

        result = await connection.execute(
          `SELECT
          to_char(struttura) struttura,
          CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
          CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
          to_char(codice_reparto_giuridico) codice_reparto_giuridico,
          to_char(reparto_giuridico) reparto_giuridico,
          to_char(id_people) id_people,
          to_char(EXTENSION) EXTENSION,
          to_char(data_inizio_prescrizione) data_inizio_prescrizione,
          CASE WHEN (NVL(data_fine_prescrizione,'')) is NULL then ' ' ELSE TO_CHAR(NVL(data_fine_prescrizione,'')) END data_fine_prescrizione,
          to_char(codice_farmaco_prescritto) codice_farmaco_prescritto,
          to_char(descrizione_farmacto_prescritto) descrizione_farmacto_prescritto,
          CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,
          CASE WHEN (NVL(codice_farmaco_somministrato,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_farmaco_somministrato,'')) END codice_farmaco_somministrato,
          CASE WHEN (NVL(descrizione_farmacto_somministrato,'')) is NULL then ' ' ELSE TO_CHAR(NVL(descrizione_farmacto_somministrato,'')) END descrizione_farmacto_somministrato,
          CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
          CASE WHEN (NVL(quantita,'')) is NULL then ' ' ELSE TO_CHAR(NVL(quantita,'')) END quantita,
          to_char(stato) stato,
          to_char(data_inizio_somministrazione_pianificata) data_inizio_somministrazione_pianificata,
          CASE WHEN (NVL(data_inizio_somministrazione_efettuata,'')) is NULL then ' ' ELSE TO_CHAR(NVL(data_inizio_somministrazione_efettuata,'')) END data_inizio_somministrazione_efettuata,
          to_char(route_desc) route_desc
      FROM
          V_SOMM_PAZ_WARD WHERE data_inizio_somministrazione_pianificata between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+`','DD/MM/YYYY') + (86399/86400) and (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });

  
      const rs = result.resultSet;
      let row;
      let riga = 1;

      worksheetPAZ.cell(riga,1).string('STRUTTURA').style(stylePAZ);
      worksheetPAZ.cell(riga,2).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,3).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,4).string('CODICE_REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,5).string('REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,6).string('ID_PEOPLE').style(stylePAZ);
      worksheetPAZ.cell(riga,7).string('EXTENSION').style(stylePAZ);
      worksheetPAZ.cell(riga,8).string('DATA_INIZIO_PRESCRIZIONE').style(stylePAZ);
      worksheetPAZ.cell(riga,9).string('DATA_FINE_PRESCRIZIONE').style(stylePAZ);
      worksheetPAZ.cell(riga,10).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,11).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,12).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
      worksheetPAZ.cell(riga,13).string('CODICE_FARMACO_SOMMINISTRATO').style(stylePAZ);
      worksheetPAZ.cell(riga,14).string('DESCRIZIONE_FARMACTO_SOMMINISTRATO').style(stylePAZ);        
      worksheetPAZ.cell(riga,15).string('UNITA_DI_MISURA').style(stylePAZ);
      worksheetPAZ.cell(riga,16).string('QUANTITA').style(stylePAZ);
      worksheetPAZ.cell(riga,17).string('STATO').style(stylePAZ);
      worksheetPAZ.cell(riga,18).string('DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA').style(stylePAZ);
      worksheetPAZ.cell(riga,19).string('DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA').style(stylePAZ);
      worksheetPAZ.cell(riga,20).string('ROUTE_DESC').style(stylePAZ);



      riga++;


      while ((row = await rs.getRow())) {
        //console.log(riga);  
        //console.log(row);
        //console.log(row.ISTITUTO);
        worksheetPAZ.cell(riga,1).string(row.STRUTTURA).style(stylePAZ);
        worksheetPAZ.cell(riga,2).string(row.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZ.cell(riga,3).string(row.REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZ.cell(riga,4).string(row.CODICE_REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZ.cell(riga,5).string(row.REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZ.cell(riga,6).string(row.ID_PEOPLE).style(stylePAZ);
        worksheetPAZ.cell(riga,7).string(row.EXTENSION).style(stylePAZ);
        worksheetPAZ.cell(riga,8).string(row.DATA_INIZIO_PRESCRIZIONE).style(stylePAZ);
        worksheetPAZ.cell(riga,9).string(row.DATA_FINE_PRESCRIZIONE).style(stylePAZ);
        worksheetPAZ.cell(riga,10).string(row.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,11).string(row.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,12).string(row.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
        worksheetPAZ.cell(riga,13).string(row.CODICE_FARMACO_SOMMINISTRATO).style(stylePAZ);
        worksheetPAZ.cell(riga,14).string(row.DESCRIZIONE_FARMACTO_SOMMINISTRATO).style(stylePAZ);
        worksheetPAZ.cell(riga,15).string(row.UNITA_DI_MISURA).style(stylePAZ);
        worksheetPAZ.cell(riga,16).string(row.QUANTITA).style(stylePAZ);
        worksheetPAZ.cell(riga,17).string(row.STATO).style(stylePAZ);
        worksheetPAZ.cell(riga,18).string(row.DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA).style(stylePAZ);
        worksheetPAZ.cell(riga,19).string(row.DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA).style(stylePAZ);
        worksheetPAZ.cell(riga,20).string(row.ROUTE_DESC).style(stylePAZ);
        riga++;
      }
  
        //workbook.write('statistiche.xlsx', res);
        workbookPAZ.write(dati.funzione+" "+dati.unitCode+" "+ date + "-" + month + "-" + year+" ore " + hour+"-" + minutes+".xlsx", res);
        await rs.close();

      }

      if (dati.funzione == 'carrello' || dati.funzione == 'farmacia'){
        result = await connection.execute(
          `WITH appoggio as (
            select * from v_somm_presc_ward
            ) 
            select 
              struttura,
              CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
              CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
              codice_reparto_giuridico,
              reparto_giuridico,
              codice_farmaco_prescritto,
              descrizione_farmacto_prescritto,
              unita_di_misura,
              CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
              atc_code, 
              CASE WHEN (NVL(sum(qty_arr),'')) is NULL then ' ' ELSE TO_CHAR(NVL(sum(qty_arr),'')) END qty_arrotondata
              from appoggio 
            where appoggio.planned_start between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+`','DD/MM/YYYY') + (86399/86400) and (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')
            group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_reparto_giuridico,reparto_giuridico,codice_farmaco_prescritto,descrizione_farmacto_prescritto,unita_di_misura,forma_farmaceutica_prescritta,atc_code
            order by descrizione_farmacto_prescritto`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });

  
      const rs2 = result.resultSet;
      let row;
      let riga = 1;

      worksheetPAZ.cell(riga,1).string('DA').style(stylePAZ);
      worksheetPAZ.cell(riga,2).string('A').style(stylePAZ);
      worksheetPAZ.cell(riga,3).string('STRUTTURA').style(stylePAZ);
      worksheetPAZ.cell(riga,4).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,5).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,6).string('CODICE_REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,7).string('REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,8).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,9).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,10).string('UNITA_DI_MISURA').style(stylePAZ);
      worksheetPAZ.cell(riga,11).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
      worksheetPAZ.cell(riga,12).string('ATC_CODE').style(stylePAZ);        
      worksheetPAZ.cell(riga,13).string('QTY_ARROTONDATA').style(stylePAZ);

      riga++;


      while ((row = await rs2.getRow())) {
        //console.log(riga);  
        //console.log(row);
        //console.log(row.ISTITUTO);
        worksheetPAZ.cell(riga,1).string(dati.dataIniziale).style(stylePAZ);
        worksheetPAZ.cell(riga,2).string(dati.dataFinale).style(stylePAZ);
        worksheetPAZ.cell(riga,3).string(row.STRUTTURA).style(stylePAZ);
        worksheetPAZ.cell(riga,4).string(row.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZ.cell(riga,5).string(row.REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZ.cell(riga,6).string(row.CODICE_REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZ.cell(riga,7).string(row.REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZ.cell(riga,8).string(row.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,9).string(row.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,10).string(row.UNITA_DI_MISURA).style(stylePAZ);
        worksheetPAZ.cell(riga,11).string(row.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
        worksheetPAZ.cell(riga,12).string(row.ATC_CODE).style(stylePAZ);
        worksheetPAZ.cell(riga,13).string(row.QTY_ARROTONDATA).style(stylePAZ);

        riga++;
      }

      //workbook.write('statistiche.xlsx', res);
      workbookPAZ.write(dati.funzione+" "+dati.unitCode+" "+ date + "-" + month + "-" + year+" ore " + hour+"-" + minutes+".xlsx", res);
      await rs2.close();

    }
  
    
        
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

app.get('/wardreport', (req, res) => {
  
    console.log("Richiesta ricevuta per wardreport:");
    console.log(req.query);
  
    const unitCode = req.query.unitCode;
    const idUser = req.query.idUser;
  
    //generaReportTerapia(req.params.wsd1+'/'+req.params.wsd2, res);
    //res.sendFile(__dirname + "/index.html");
    //res.render(path.join(__dirname, '/', 'ward.html'), {unitCode: unitCode, idUser: idUser });
  

    generaReportReparto(req.query, res);
    

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
