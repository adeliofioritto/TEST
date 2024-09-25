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
app.use('/images', express.static(process.cwd() + '/images'))
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


  async function recuperaListaLetti(reparto,res) {

    var listaLetti = [];
    
    
    let connection;
      try {
        connection = await oracledb.getConnection({ user: DB_USER, password: DB_PASSWORD, connectionString: DB_CONNECTION_STRING });
        
        /* Modificata in data 27/11/2023, eliminati codici reparti ass e giu, filtro solo per COMPLETED */
        result6 = await connection.execute(
          `SELECT DISTINCT nome_stanza from v_nosologico_stanza_reparto where (CODICE_REPARTO_ASSISTENZIALE = '`+reparto+`' OR CODICE_REPARTO_GIURIDICO = '`+reparto+`')`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });    
      
          const rs6 = result6.resultSet;
          let row6;
          let riga = 1;
    
          riga++;

          while ((row6 = await rs6.getRow())) {
            //console.log(riga);  
            //console.log(row6);
            //console.log(row6.NOME_STANZA);
            listaLetti.push(row6.NOME_STANZA);
            //worksheetPAZ.cell(riga,18).string(row.ROUTE_DESC).style(stylePAZ);
            riga++;
          }
      
          await rs6.close();
          //console.log(listaLetti);
          let csvListaLetti = listaLetti.map(e => JSON.stringify(e)).join(",");
          //console.log(csvListaLetti);
          return csvListaLetti;
          
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
 
async function generaReportTerapia(reparto,res) {

  var workbook = new excel.Workbook();
  var workbookPAZ = new excel.Workbook();
  var worksheetPAZ = workbookPAZ.addWorksheet('FARMACI_PAZIENTE');
  var style = workbook.createStyle({font: {color: '#000000',size: 10}});
  var stylePAZ = workbookPAZ.createStyle({font: {color: '#000000',size: 10}});

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

          /*
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
            CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,
            to_char(codice_farmaco_somministrato) codice_farmaco_somministrato,
            to_char(descrizione_farmacto_somministrato) descrizione_farmacto_somministrato,
            CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
            CASE WHEN (NVL(quantita,'')) is NULL then ' ' ELSE TO_CHAR(NVL(quantita,'')) END quantita,
            to_char(stato) stato,
            to_char(data_inizio_somministrazione_pianificata) data_inizio_somministrazione_pianificata,
            CASE WHEN (NVL(data_inizio_somministrazione_efettuata,'')) is NULL then ' ' ELSE TO_CHAR(NVL(data_inizio_somministrazione_efettuata,'')) END data_inizio_somministrazione_efettuata,
            to_char(route_desc) route_desc
        FROM
            v_somm_paz_nos_v2 WHERE extension = '`+reparto+`'`,
            [],
            { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
*/

      /* Modificata in data 27/11/2023, eliminati codici reparti ass e giu, filtro solo per COMPLETED */
      result = await connection.execute(
        `SELECT
        CASE WHEN (NVL(tipo_fornitura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(tipo_fornitura,'')) END tipo_fornitura,
        to_char(struttura) struttura,
        to_char(reparto_assistenziale) reparto_assistenziale,
        to_char(reparto_giuridico) reparto_giuridico,
        to_char(id_people) id_people,
        to_char(EXTENSION) EXTENSION,
        to_char(data_inizio_prescrizione) data_inizio_prescrizione,
        to_char(data_fine_prescrizione) data_fine_prescrizione,
        to_char(codice_farmaco_prescritto) codice_farmaco_prescritto,
        to_char(descrizione_farmacto_prescritto) descrizione_farmacto_prescritto,
        CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,
        to_char(codice_farmaco_somministrato) codice_farmaco_somministrato,
        to_char(descrizione_farmacto_somministrato) descrizione_farmacto_somministrato,
        CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
        CASE WHEN (NVL(quantita,'')) is NULL then ' ' ELSE TO_CHAR(NVL(quantita,'')) END quantita,
        to_char(stato) stato,
        to_char(data_inizio_somministrazione_pianificata) data_inizio_somministrazione_pianificata,
        CASE WHEN (NVL(data_inizio_somministrazione_efettuata,'')) is NULL then ' ' ELSE TO_CHAR(NVL(data_inizio_somministrazione_efettuata,'')) END data_inizio_somministrazione_efettuata,
        to_char(route_desc) route_desc,
        CASE WHEN (NVL(sum_num_strength_val,'')) is NULL then ' ' ELSE TO_CHAR(NVL(sum_num_strength_val,'')) END tot_pa,
        CASE WHEN (NVL(CODE_UOM,'')) is NULL then ' ' ELSE TO_CHAR(NVL(CODE_UOM,'')) END unita_riferimento_pa,
        CASE WHEN (NVL(DESCRIZIONE_ESTESA_CONTENITORE,'')) is NULL then ' ' ELSE TO_CHAR(NVL(DESCRIZIONE_ESTESA_CONTENITORE,'')) END DESCRIZIONE_ESTESA_CONTENITORE,
        CASE WHEN (NVL(DESCRIZIONE_FORMA_FARMACEUTICA,'')) is NULL then ' ' ELSE TO_CHAR(NVL(DESCRIZIONE_FORMA_FARMACEUTICA,'')) END DESCRIZIONE_FORMA_FARMACEUTICA
      FROM
      v_somm_paz_nos_v2 v left join V_FARMACI_CONTENITORI_UOM CFC on v.codice_farmaco_somministrato = cfc.amp_code
      WHERE extension = '`+reparto+`'`,
        [],
        { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });    
    
        const rs = result.resultSet;
        let row;
        let riga = 1;
        worksheetPAZ.cell(riga,1).string('STRUTTURA').style(stylePAZ);
        worksheetPAZ.cell(riga,2).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZ.cell(riga,3).string('REPARTO_GIURIDICO').style(stylePAZ);
        worksheetPAZ.cell(riga,4).string('ID_PEOPLE').style(stylePAZ);
        worksheetPAZ.cell(riga,5).string('NOSOLOGICO').style(stylePAZ);
        worksheetPAZ.cell(riga,6).string('DATA_INIZIO_PRESCRIZIONE').style(stylePAZ);
        worksheetPAZ.cell(riga,7).string('DATA_FINE_PRESCRIZIONE').style(stylePAZ);
        worksheetPAZ.cell(riga,8).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
        worksheetPAZ.cell(riga,9).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
        worksheetPAZ.cell(riga,10).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
        worksheetPAZ.cell(riga,11).string('CODICE_FARMACO_SOMMINISTRATO').style(stylePAZ);
        worksheetPAZ.cell(riga,12).string('DESCRIZIONE_FARMACTO_SOMMINISTRATO').style(stylePAZ);        
        worksheetPAZ.cell(riga,13).string('UNITA_DI_MISURA').style(stylePAZ);
        worksheetPAZ.cell(riga,14).string('QUANTITA').style(stylePAZ);
        worksheetPAZ.cell(riga,15).string('STATO').style(stylePAZ);
        worksheetPAZ.cell(riga,16).string('DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA').style(stylePAZ);
        worksheetPAZ.cell(riga,17).string('DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA').style(stylePAZ);
        worksheetPAZ.cell(riga,18).string('ROUTE_DESC').style(stylePAZ);
        worksheetPAZ.cell(riga,19).string('TIPO_FORNITURA').style(stylePAZ);
        worksheetPAZ.cell(riga,20).string('TOT_PA').style(stylePAZ);
        worksheetPAZ.cell(riga,21).string('UNITA_RIFERIMENTO_PA').style(stylePAZ);
        worksheetPAZ.cell(riga,22).string('DESCRIZIONE_ESTESA_CONTENITORE').style(stylePAZ);
        worksheetPAZ.cell(riga,23).string('DESCRIZIONE_FORMA_FARMACEUTICA').style(stylePAZ);

  
  
        riga++;
  
  
        while ((row = await rs.getRow())) {
          //console.log(riga);  
          //console.log(row);
          //console.log(row.ISTITUTO);
          worksheetPAZ.cell(riga,1).string(row.STRUTTURA).style(stylePAZ);
          worksheetPAZ.cell(riga,2).string(row.REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZ.cell(riga,3).string(row.REPARTO_GIURIDICO).style(stylePAZ);
          worksheetPAZ.cell(riga,4).string(row.ID_PEOPLE).style(stylePAZ);
          worksheetPAZ.cell(riga,5).string(row.EXTENSION).style(stylePAZ);
          worksheetPAZ.cell(riga,6).string(row.DATA_INIZIO_PRESCRIZIONE).style(stylePAZ);
          worksheetPAZ.cell(riga,7).string(row.DATA_FINE_PRESCRIZIONE).style(stylePAZ);
          worksheetPAZ.cell(riga,8).string(row.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
          worksheetPAZ.cell(riga,9).string(row.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
          worksheetPAZ.cell(riga,10).string(row.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
          worksheetPAZ.cell(riga,11).string(row.CODICE_FARMACO_SOMMINISTRATO).style(stylePAZ);
          worksheetPAZ.cell(riga,12).string(row.DESCRIZIONE_FARMACTO_SOMMINISTRATO).style(stylePAZ);
          worksheetPAZ.cell(riga,13).string(row.UNITA_DI_MISURA).style(stylePAZ);
          worksheetPAZ.cell(riga,14).string(row.QUANTITA).style(stylePAZ);
          worksheetPAZ.cell(riga,15).string(row.STATO).style(stylePAZ);
          worksheetPAZ.cell(riga,16).string(row.DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA).style(stylePAZ);
          worksheetPAZ.cell(riga,17).string(row.DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA).style(stylePAZ);
          worksheetPAZ.cell(riga,18).string(row.ROUTE_DESC).style(stylePAZ);
          worksheetPAZ.cell(riga,19).string(row.TIPO_FORNITURA).style(stylePAZ);
          worksheetPAZ.cell(riga,20).string(row.TOT_PA).style(stylePAZ);
          worksheetPAZ.cell(riga,21).string(row.UNITA_RIFERIMENTO_PA).style(stylePAZ);
          worksheetPAZ.cell(riga,22).string(row.DESCRIZIONE_ESTESA_CONTENITORE).style(stylePAZ);
          worksheetPAZ.cell(riga,23).string(row.DESCRIZIONE_FORMA_FARMACEUTICA).style(stylePAZ);
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

  var workbook =            new excel.Workbook();
  var workbookPAZ =         new excel.Workbook();
  var worksheetPAZ =        workbookPAZ.addWorksheet('FARMACI_PAZIENTE');

  var style = workbook.createStyle({font: {color: '#000000',size: 10}});
  var stylePAZ = workbookPAZ.createStyle({font: {color: '#000000',size: 10}});
  
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
      
      result7 = await connection.execute(
        `select CASE WHEN oo.class_code = 'SRV' THEN 'codice_reparto_giuridico' WHEN oo.class_code = 'HOUS' THEN 'codice_reparto_assistenziale' ELSE null END tipologia_reparto 
      from CIS4C_DM.org_id oi join cis4c_dm.org_organization oo on oo.id = oi.owner_id_org and oi.concept_id_orgcnpt = 40000
      where oi.extension = 'OSR_OSR_00210050'`,
        [],
        { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });

      const rs7 = result7.resultSet;
      let row7;
      let tipologiaReparto = '';

        while ((row7 = await rs7.getRow())) {
          tipologiaReparto = row7.TIPOLOGIA_REPARTO
          console.log(tipologiaReparto);  
          }


       if (dati.funzione === 'pazienti'){
        //console.log("1:"+new Date().toString());

        

        /*result = await connection.execute(
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
          sostituibilita,
          CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,
          CASE WHEN (NVL(codice_farmaco_somministrato,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_farmaco_somministrato,'')) END codice_farmaco_somministrato,
          CASE WHEN (NVL(descrizione_farmacto_somministrato,'')) is NULL then ' ' ELSE TO_CHAR(NVL(descrizione_farmacto_somministrato,'')) END descrizione_farmacto_somministrato,
          CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
          CASE WHEN (NVL(quantita,'')) is NULL then ' ' ELSE TO_CHAR(NVL(quantita,'')) END quantita,
          CASE WHEN (NVL((qty_arr),'')) is NULL then ' ' ELSE rtrim(to_char(NVL((qty_arr),'') , 'FM999999999999990.99'), '.') END qty_arrotondata,
          to_char(stato) stato,
          to_char(data_inizio_somministrazione_pianificata) data_inizio_somministrazione_pianificata,
          CASE WHEN (NVL(data_inizio_somministrazione_efettuata,'')) is NULL then ' ' ELSE TO_CHAR(NVL(data_inizio_somministrazione_efettuata,'')) END data_inizio_somministrazione_efettuata,
          to_char(route_desc) route_desc,
          decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale,codice_reparto_giuridico),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario
      FROM
          V_SOMM_PAZ_WARD WHERE data_inizio_somministrazione_pianificata between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+`','DD/MM/YYYY') + (86399/86400) and (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
          console.log("2:"+new Date().toString());*/

          /* 2024/08/09 AF ottimizzazione query, aggiunto filtro dinamico sulla tipologia di reparto ass/giu, abbattimento costo da 11069 a 1381 */
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
            sostituibilita,
            CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,
            CASE WHEN (NVL(codice_farmaco_somministrato,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_farmaco_somministrato,'')) END codice_farmaco_somministrato,
            CASE WHEN (NVL(descrizione_farmacto_somministrato,'')) is NULL then ' ' ELSE TO_CHAR(NVL(descrizione_farmacto_somministrato,'')) END descrizione_farmacto_somministrato,
            CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
            CASE WHEN (NVL(quantita,'')) is NULL then ' ' ELSE TO_CHAR(NVL(quantita,'')) END quantita,
            CASE WHEN (NVL((qty_arr),'')) is NULL then ' ' ELSE rtrim(to_char(NVL((qty_arr),'') , 'FM999999999999990.99'), '.') END qty_arrotondata,
            to_char(stato) stato,
            to_char(data_inizio_somministrazione_pianificata) data_inizio_somministrazione_pianificata,
            CASE WHEN (NVL(data_inizio_somministrazione_efettuata,'')) is NULL then ' ' ELSE TO_CHAR(NVL(data_inizio_somministrazione_efettuata,'')) END data_inizio_somministrazione_efettuata,
            to_char(route_desc) route_desc,
            decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale,codice_reparto_giuridico),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario,
            CASE WHEN (NVL(sum_num_strength_val,'')) is NULL then ' ' ELSE TO_CHAR(NVL(sum_num_strength_val,'')) END tot_pa,
            CASE WHEN (NVL(CODE_UOM,'')) is NULL then ' ' ELSE TO_CHAR(NVL(CODE_UOM,'')) END unita_riferimento_pa,
            CASE WHEN (NVL(DESCRIZIONE_ESTESA_CONTENITORE,'')) is NULL then ' ' ELSE TO_CHAR(NVL(DESCRIZIONE_ESTESA_CONTENITORE,'')) END DESCRIZIONE_ESTESA_CONTENITORE,
            CASE WHEN (NVL(DESCRIZIONE_FORMA_FARMACEUTICA,'')) is NULL then ' ' ELSE TO_CHAR(NVL(DESCRIZIONE_FORMA_FARMACEUTICA,'')) END DESCRIZIONE_FORMA_FARMACEUTICA
        FROM
            V_SOMM_PAZ_WARD v left join V_FARMACI_CONTENITORI_UOM CFC on v.codice_farmaco_somministrato = cfc.amp_code WHERE data_inizio_somministrazione_pianificata between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+`','DD/MM/YYYY') + (86399/86400) and (`+tipologiaReparto+` = '`+dati.unitCode+`')`,
            [],
            { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
            //console.log("2:"+new Date().toString());


      const rs = result.resultSet;
      let row;
      let riga = 1;

      worksheetPAZ.cell(riga,1).string('PERIODO_DA').style(stylePAZ);
      worksheetPAZ.cell(riga,2).string('PERIODO_AL').style(stylePAZ);
      worksheetPAZ.cell(riga,3).string('STRUTTURA').style(stylePAZ);
      worksheetPAZ.cell(riga,4).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,5).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,6).string('CODICE_REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,7).string('REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,8).string('ID_PEOPLE').style(stylePAZ);
      worksheetPAZ.cell(riga,9).string('NOSOLOGICO').style(stylePAZ);
      worksheetPAZ.cell(riga,10).string('DATA_INIZIO_PRESCRIZIONE').style(stylePAZ);
      worksheetPAZ.cell(riga,11).string('DATA_FINE_PRESCRIZIONE').style(stylePAZ);
      worksheetPAZ.cell(riga,12).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,13).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,14).string('SOSTITUIBILITA').style(stylePAZ);
      worksheetPAZ.cell(riga,15).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
      worksheetPAZ.cell(riga,16).string('CODICE_FARMACO_SOMMINISTRATO').style(stylePAZ);
      worksheetPAZ.cell(riga,17).string('DESCRIZIONE_FARMACTO_SOMMINISTRATO').style(stylePAZ);        
      worksheetPAZ.cell(riga,18).string('UNITA_DI_MISURA').style(stylePAZ);
      worksheetPAZ.cell(riga,19).string('QUANTITA').style(stylePAZ);
      worksheetPAZ.cell(riga,20).string('QTY_ARROTONDATA').style(stylePAZ);      
      worksheetPAZ.cell(riga,21).string('STATO').style(stylePAZ);
      worksheetPAZ.cell(riga,22).string('DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA').style(stylePAZ);
      worksheetPAZ.cell(riga,23).string('DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA').style(stylePAZ);
      worksheetPAZ.cell(riga,24).string('ROUTE_DESC').style(stylePAZ);
      worksheetPAZ.cell(riga,25).string('IN_PRONTUARIO').style(stylePAZ);
      worksheetPAZ.cell(riga,26).string('TOT_PA').style(stylePAZ);
      worksheetPAZ.cell(riga,27).string('UNITA_RIFERIMENTO_PA').style(stylePAZ);
      worksheetPAZ.cell(riga,28).string('DESCRIZIONE_ESTESA_CONTENITORE').style(stylePAZ);
      worksheetPAZ.cell(riga,29).string('DESCRIZIONE_FORMA_FARMACEUTICA').style(stylePAZ);

      riga++;

      //console.log("3:"+new Date().toString());

      while ((row = await rs.getRow())) {
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
        worksheetPAZ.cell(riga,8).string(row.ID_PEOPLE).style(stylePAZ);
        worksheetPAZ.cell(riga,9).string(row.EXTENSION).style(stylePAZ);
        worksheetPAZ.cell(riga,10).string(row.DATA_INIZIO_PRESCRIZIONE).style(stylePAZ);
        worksheetPAZ.cell(riga,11).string(row.DATA_FINE_PRESCRIZIONE).style(stylePAZ);
        worksheetPAZ.cell(riga,12).string(row.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,13).string(row.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,14).string(row.SOSTITUIBILITA).style(stylePAZ);
        worksheetPAZ.cell(riga,15).string(row.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
        worksheetPAZ.cell(riga,16).string(row.CODICE_FARMACO_SOMMINISTRATO).style(stylePAZ);
        worksheetPAZ.cell(riga,17).string(row.DESCRIZIONE_FARMACTO_SOMMINISTRATO).style(stylePAZ);
        worksheetPAZ.cell(riga,18).string(row.UNITA_DI_MISURA).style(stylePAZ);
        worksheetPAZ.cell(riga,19).string(row.QUANTITA).style(stylePAZ);
        worksheetPAZ.cell(riga,20).string(row.QTY_ARROTONDATA).style(stylePAZ);
        worksheetPAZ.cell(riga,21).string(row.STATO).style(stylePAZ);
        worksheetPAZ.cell(riga,22).string(row.DATA_INIZIO_SOMMINISTRAZIONE_PIANIFICATA).style(stylePAZ);
        worksheetPAZ.cell(riga,23).string(row.DATA_INIZIO_SOMMINISTRAZIONE_EFETTUATA).style(stylePAZ);
        worksheetPAZ.cell(riga,24).string(row.ROUTE_DESC).style(stylePAZ);
        worksheetPAZ.cell(riga,25).string(row.IN_PRONTUARIO).style(stylePAZ);
        worksheetPAZ.cell(riga,26).string(row.TOT_PA).style(stylePAZ);
        worksheetPAZ.cell(riga,27).string(row.UNITA_RIFERIMENTO_PA).style(stylePAZ);
        worksheetPAZ.cell(riga,28).string(row.DESCRIZIONE_ESTESA_CONTENITORE).style(stylePAZ);
        worksheetPAZ.cell(riga,29).string(row.DESCRIZIONE_FORMA_FARMACEUTICA).style(stylePAZ);
        riga++;
      }
      //console.log("4:"+new Date().toString());

        //workbook.write('statistiche.xlsx', res);
        workbookPAZ.write(dati.funzione+" "+dati.unitCode+" "+ date + "-" + month + "-" + year+" ore " + hour+"-" + minutes+".xlsx", res);
        await rs.close();

      }
      
      //timeFinale

      if (dati.funzione === 'carrello'){
        console.log("dati.listaLetti:"+dati.listaLetti);
        console.log("dati.timeFinale:"+dati.timeFinale);
        var worksheetPAZbisogno = workbookPAZ.addWorksheet('AL_BISOGNO');

        /*result2 = await connection.execute(
          `select t.*, decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario from (
            WITH appoggio as (
            select * from v_somm_presc_ward
            ) 
            select 
              struttura,
              CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
              CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
              codice_farmaco_prescritto,
              descrizione_farmacto_prescritto,
              sostituibilita,
              CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
              CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
              atc_code, 
              CASE WHEN (NVL(sum(qty_arr),'')) is NULL then ' ' ELSE rtrim(to_char(NVL(sum(qty_arr),'') , 'FM999999999999990.99'), '.') END qty_arrotondata
              from appoggio 
            where nome_stanza in (`+dati.listaLetti+`) and appoggio.planned_start between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+` `+dati.timeFinale+`','DD/MM/YYYY HH24:MI') and (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')
            group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_farmaco_prescritto,descrizione_farmacto_prescritto,sostituibilita,unita_di_misura,forma_farmaceutica_prescritta,atc_code 
            order by descrizione_farmacto_prescritto) t`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
          */

          result2 = await connection.execute(
            `select t.*, decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario from (
              WITH appoggio as (
              select * from v_somm_presc_ward
              ) 
              select 
                struttura,
                CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
                CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
                codice_farmaco_prescritto,
                descrizione_farmacto_prescritto,
                sostituibilita,
                CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
                CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
                atc_code, 
                CASE WHEN (NVL(sum(qty_arr),'')) is NULL then ' ' ELSE rtrim(to_char(NVL(sum(qty_arr),'') , 'FM999999999999990.99'), '.') END qty_arrotondata
                from appoggio 
              where nome_stanza in (`+dati.listaLetti+`) and appoggio.planned_start between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+` `+dati.timeFinale+`','DD/MM/YYYY HH24:MI') and (`+tipologiaReparto+` = '`+dati.unitCode+`')
              group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_farmaco_prescritto,descrizione_farmacto_prescritto,sostituibilita,unita_di_misura,forma_farmaceutica_prescritta,atc_code 
              order by descrizione_farmacto_prescritto) t`,
            [],
            { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });

        const rs2 = result2.resultSet;
        let row2;
        let riga = 1;

        worksheetPAZ.cell(riga,1).string('PERIODO_DA').style(stylePAZ);
        worksheetPAZ.cell(riga,2).string('PERIODO_AL').style(stylePAZ);
        worksheetPAZ.cell(riga,3).string('STRUTTURA').style(stylePAZ);
        worksheetPAZ.cell(riga,4).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZ.cell(riga,5).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZ.cell(riga,6).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
        worksheetPAZ.cell(riga,7).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
        worksheetPAZ.cell(riga,8).string('SOSTITUIBILITA').style(stylePAZ);
        worksheetPAZ.cell(riga,9).string('UNITA_DI_MISURA').style(stylePAZ);
        worksheetPAZ.cell(riga,10).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
        worksheetPAZ.cell(riga,11).string('ATC_CODE').style(stylePAZ);        
        worksheetPAZ.cell(riga,12).string('QTY_ARROTONDATA').style(stylePAZ);
        worksheetPAZ.cell(riga,13).string('IN_PRONTUARIO').style(stylePAZ);
        

        riga++;


        while ((row2 = await rs2.getRow())) {
          //console.log(riga);  
          //console.log(row);
          //console.log(row.ISTITUTO);
          worksheetPAZ.cell(riga,1).string(dati.dataIniziale).style(stylePAZ);
          worksheetPAZ.cell(riga,2).string(dati.dataFinale).style(stylePAZ);
          worksheetPAZ.cell(riga,3).string(row2.STRUTTURA).style(stylePAZ);
          worksheetPAZ.cell(riga,4).string(row2.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZ.cell(riga,5).string(row2.REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZ.cell(riga,6).string(row2.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
          worksheetPAZ.cell(riga,7).string(row2.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
          worksheetPAZ.cell(riga,8).string(row2.SOSTITUIBILITA).style(stylePAZ);
          worksheetPAZ.cell(riga,9).string(row2.UNITA_DI_MISURA).style(stylePAZ);
          worksheetPAZ.cell(riga,10).string(row2.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
          worksheetPAZ.cell(riga,11).string(row2.ATC_CODE).style(stylePAZ);
          worksheetPAZ.cell(riga,12).string(row2.QTY_ARROTONDATA).style(stylePAZ);
          worksheetPAZ.cell(riga,13).string(row2.IN_PRONTUARIO).style(stylePAZ);

          riga++;
        }


        /* Per le prescrizioni al bisogno */
        result4 = await connection.execute(
          `select t.*, decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario from (
            WITH appoggio as (
              select * from v_somm_presc_ward_req
              ) 
              select 
                struttura,
                CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
                CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
                codice_farmaco_prescritto,
                descrizione_farmacto_prescritto,
                sostituibilita,
                CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
                atc_code 
                from appoggio 
              where nome_stanza in (`+dati.listaLetti+`) and (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')
              group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_farmaco_prescritto,descrizione_farmacto_prescritto,sostituibilita,forma_farmaceutica_prescritta,atc_code
              order by descrizione_farmacto_prescritto) t`,
          [],
          { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });


        const rs4 = result4.resultSet;
        let row4;
        riga = 1;

        worksheetPAZbisogno.cell(riga,1).string('PERIODO_DA').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,2).string('PERIODO_AL').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,3).string('STRUTTURA').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,4).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,5).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,6).string('ATC_CODE').style(stylePAZ);        
        worksheetPAZbisogno.cell(riga,7).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,8).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,9).string('SOSTITUIBILITA').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,10).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
        worksheetPAZbisogno.cell(riga,11).string('IN_PRONTUARIO').style(stylePAZ);
        

        riga++;


        while ((row4 = await rs4.getRow())) {
          //console.log(riga);  
          //console.log(row);
          //console.log(row.ISTITUTO);
          worksheetPAZbisogno.cell(riga,1).string(dati.dataIniziale).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,2).string(dati.dataFinale).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,3).string(row4.STRUTTURA).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,4).string(row4.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,5).string(row4.REPARTO_ASSISTENZIALE).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,6).string(row4.ATC_CODE).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,7).string(row4.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,8).string(row4.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,9).string(row4.SOSTITUIBILITA).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,10).string(row4.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
          worksheetPAZbisogno.cell(riga,11).string(row4.IN_PRONTUARIO).style(stylePAZ);

          riga++;
        }
        /* FINE prescrizionia al bisogno */

        //workbook.write('statistiche.xlsx', res);
        workbookPAZ.write(dati.funzione+" "+dati.unitCode+" "+ date + "-" + month + "-" + year+" ore " + hour+"-" + minutes+".xlsx", res);
        await rs2.close();
        await rs4.close();

      }
  
  
      if (dati.funzione === 'farmacia'){
        var worksheetPAZbisogno = workbookPAZ.addWorksheet('AL_BISOGNO');
    
        /*
          result3 = await connection.execute(
            `select t.*, decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale,codice_reparto_giuridico),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario from (
              WITH appoggio as (
              select * from v_somm_presc_ward
              ) 
              select 
                struttura,
                CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
                CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
                codice_reparto_giuridico,
                reparto_giuridico,
                codice_farmaco_prescritto,
                sostituibilita,
                descrizione_farmacto_prescritto,
                CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
                CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
                atc_code, 
                CASE WHEN (NVL(sum(qty_arr),'')) is NULL then ' ' ELSE rtrim(to_char(NVL(sum(qty_arr),'') , 'FM999999999999990.99'), '.') END qty_arrotondata
                from appoggio 
              where appoggio.planned_start between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+`','DD/MM/YYYY') + (86399/86400) and (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')
              group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_reparto_giuridico,reparto_giuridico,codice_farmaco_prescritto,descrizione_farmacto_prescritto,sostituibilita,unita_di_misura,forma_farmaceutica_prescritta,atc_code
              order by descrizione_farmacto_prescritto) t`,
            [],
            { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
            */
           
            result3 = await connection.execute(
              `select t.*, decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale,codice_reparto_giuridico),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario from (
                WITH appoggio as (
                select * from v_somm_presc_ward
                ) 
                select 
                  struttura,
                  CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
                  CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
                  codice_reparto_giuridico,
                  reparto_giuridico,
                  codice_farmaco_prescritto,
                  sostituibilita,
                  descrizione_farmacto_prescritto,
                  CASE WHEN (NVL(unita_di_misura,'')) is NULL then ' ' ELSE TO_CHAR(NVL(unita_di_misura,'')) END unita_di_misura,
                  CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
                  atc_code, 
                  CASE WHEN (NVL(sum(qty_arr),'')) is NULL then ' ' ELSE rtrim(to_char(NVL(sum(qty_arr),'') , 'FM999999999999990.99'), '.') END qty_arrotondata
                  from appoggio 
                where appoggio.planned_start between to_date('`+dati.dataIniziale+`','DD/MM/YYYY') and to_date('`+dati.dataFinale+`','DD/MM/YYYY') + (86399/86400) and (`+tipologiaReparto+` = '`+dati.unitCode+`')
                group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_reparto_giuridico,reparto_giuridico,codice_farmaco_prescritto,descrizione_farmacto_prescritto,sostituibilita,unita_di_misura,forma_farmaceutica_prescritta,atc_code
                order by descrizione_farmacto_prescritto) t`,
              [],
              { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });

  
      const rs3 = result3.resultSet;
      let row3;
      let riga = 1;

      worksheetPAZ.cell(riga,1).string('PERIODO_DA').style(stylePAZ);
      worksheetPAZ.cell(riga,2).string('PERIODO_AL').style(stylePAZ);
      worksheetPAZ.cell(riga,3).string('STRUTTURA').style(stylePAZ);
      worksheetPAZ.cell(riga,4).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,5).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZ.cell(riga,6).string('CODICE_REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,7).string('REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZ.cell(riga,8).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,9).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
      worksheetPAZ.cell(riga,10).string('SOSTITUIBILITA').style(stylePAZ);
      worksheetPAZ.cell(riga,11).string('UNITA_DI_MISURA').style(stylePAZ);
      worksheetPAZ.cell(riga,12).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
      worksheetPAZ.cell(riga,13).string('ATC_CODE').style(stylePAZ);        
      worksheetPAZ.cell(riga,14).string('QTY_ARROTONDATA').style(stylePAZ);
      worksheetPAZ.cell(riga,15).string('IN_PRONTUARIO').style(stylePAZ);

      riga++;


      while ((row3 = await rs3.getRow())) {
        //console.log(riga);  
        //console.log(row);
        //console.log(row.ISTITUTO);
        worksheetPAZ.cell(riga,1).string(dati.dataIniziale).style(stylePAZ);
        worksheetPAZ.cell(riga,2).string(dati.dataFinale).style(stylePAZ);
        worksheetPAZ.cell(riga,3).string(row3.STRUTTURA).style(stylePAZ);
        worksheetPAZ.cell(riga,4).string(row3.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZ.cell(riga,5).string(row3.REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZ.cell(riga,6).string(row3.CODICE_REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZ.cell(riga,7).string(row3.REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZ.cell(riga,8).string(row3.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,9).string(row3.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
        worksheetPAZ.cell(riga,10).string(row3.SOSTITUIBILITA).style(stylePAZ);
        worksheetPAZ.cell(riga,11).string(row3.UNITA_DI_MISURA).style(stylePAZ);
        worksheetPAZ.cell(riga,12).string(row3.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
        worksheetPAZ.cell(riga,13).string(row3.ATC_CODE).style(stylePAZ);
        worksheetPAZ.cell(riga,14).string(row3.QTY_ARROTONDATA).style(stylePAZ);
        worksheetPAZ.cell(riga,15).string(row3.IN_PRONTUARIO).style(stylePAZ);

        riga++;
      }


      /* Per le prescrizioni al bisogno */
      result5 = await connection.execute(
        `select t.*, decode(farmacoinprontuario(codice_farmaco_prescritto,codice_reparto_assistenziale),1,'In prontuario', 0, 'Fuori Prontuario', -1, 'Errore') in_prontuario from (
          WITH appoggio as (
            select * from v_somm_presc_ward_req
            ) 
            select 
              struttura,
              CASE WHEN (NVL(codice_reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(codice_reparto_assistenziale,'')) END codice_reparto_assistenziale,
              CASE WHEN (NVL(reparto_assistenziale,'')) is NULL then ' ' ELSE TO_CHAR(NVL(reparto_assistenziale,'')) END reparto_assistenziale,
              codice_reparto_giuridico,
              reparto_giuridico,
              codice_farmaco_prescritto,
              descrizione_farmacto_prescritto,
              sostituibilita,
              CASE WHEN (NVL(forma_farmaceutica_prescritta,'')) is NULL then ' ' ELSE TO_CHAR(NVL(forma_farmaceutica_prescritta,'')) END forma_farmaceutica_prescritta,                
              atc_code
              from appoggio 
            where 
            (codice_reparto_assistenziale = '`+dati.unitCode+`' OR codice_reparto_giuridico = '`+dati.unitCode+`')
            group by struttura,codice_reparto_assistenziale,reparto_assistenziale,codice_reparto_giuridico,reparto_giuridico,codice_farmaco_prescritto,descrizione_farmacto_prescritto,sostituibilita,forma_farmaceutica_prescritta,atc_code
            order by descrizione_farmacto_prescritto) t`,
        [],
        { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });


      const rs5 = result5.resultSet;
      let row5;
      riga = 1;

      worksheetPAZbisogno.cell(riga,1).string('PERIODO_DA').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,2).string('PERIODO_AL').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,3).string('STRUTTURA').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,4).string('CODICE_REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,5).string('REPARTO_ASSISTENZIALE').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,6).string('CODICE_REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,7).string('REPARTO_GIURIDICO').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,8).string('ATC_CODE').style(stylePAZ);        
      worksheetPAZbisogno.cell(riga,9).string('CODICE_FARMACO_PRESCRITTO').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,10).string('DESCRIZIONE_FARMACTO_PRESCRITTO').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,11).string('SOSTITUIBILITA').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,12).string('FORMA_FARMACEUTICA_PRESCRITTA').style(stylePAZ);
      worksheetPAZbisogno.cell(riga,13).string('IN_PRONTUARIO').style(stylePAZ);
      

      riga++;


      while ((row5 = await rs5.getRow())) {
        //console.log(riga);  
        //console.log(row);
        //console.log(row.ISTITUTO);
        worksheetPAZbisogno.cell(riga,1).string(dati.dataIniziale).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,2).string(dati.dataFinale).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,3).string(row5.STRUTTURA).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,4).string(row5.CODICE_REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,5).string(row5.REPARTO_ASSISTENZIALE).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,6).string(row5.CODICE_REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,7).string(row5.REPARTO_GIURIDICO).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,8).string(row5.ATC_CODE).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,9).string(row5.CODICE_FARMACO_PRESCRITTO).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,10).string(row5.DESCRIZIONE_FARMACTO_PRESCRITTO).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,11).string(row5.SOSTITUIBILITA).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,12).string(row5.FORMA_FARMACEUTICA_PRESCRITTA).style(stylePAZ);
        worksheetPAZbisogno.cell(riga,13).string(row5.IN_PRONTUARIO).style(stylePAZ);

        riga++;
      }
      /* FINE prescrizionia al bisogno */


      //workbook.write('statistiche.xlsx', res);
      workbookPAZ.write(dati.funzione+" "+dati.unitCode+" "+ date + "-" + month + "-" + year+" ore " + hour+"-" + minutes+".xlsx", res);
      await rs3.close();
      await rs5.close();

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
  res.send('Report4C is up!');
});


app.engine('html', require('ejs').renderFile);
app.set('view engine', 'html');

app.get('/ward', (req, res) => {
  
  console.log("Richiesta ricevuta per report reparto:");
  console.log(req.query);

  const unitCode = req.query.unitCode;
  const idUser = req.query.idUser;

  recuperaListaLetti(unitCode, res).then(function(result){
    console.log("listaLetti");
    console.log(result)
    res.render(path.join(__dirname, '/', 'ward.html'), {unitCode: unitCode, idUser: idUser, listaLetti: result });
  });

  //generaReportTerapia(req.params.wsd1+'/'+req.params.wsd2, res);
  //res.sendFile(__dirname + "/index.html");




});

app.get('/wardreport', (req, res) => {
  
    console.log("Richiesta ricevuta per wardreport:");
    console.log(req.query);
  
    const unitCode = req.query.unitCode;
    const idUser = req.query.idUser;
  
    //generaReportTerapia(req.params.wsd1+'/'+req.params.wsd2, res);
    //res.sendFile(__dirname + "/index.html");
    //res.render(path.join(__dirname, '/', 'ward.html'), {unitCode: unitCode, idUser: idUser });
  
    req.query.listaLetti = req.query.listaLetti.replace(/"/g, "'");
    console.log(req.query.listaLetti);
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
