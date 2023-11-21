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

const app = express();
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

const port = 3000;

//console.log( process.env);
let EMAIL_PASSWORD = process.env.EMAIL_PASSWORD || '/var/secret/EMAIL_PASSWORD';
if (!path.isAbsolute(EMAIL_PASSWORD)) { EMAIL_PASSWORD = path.resolve(__dirname, EMAIL_PASSWORD); }
let hasEMAIL_PASSWORD = fs.existsSync(EMAIL_PASSWORD);
app.locals.hasEMAIL_PASSWORD = hasEMAIL_PASSWORD;

let DB_PASSWORD = process.env.DB_PASSWORD || '/var/secret/DB_PASSWORD';
if (!path.isAbsolute(DB_PASSWORD)) { DB_PASSWORD = path.resolve(__dirname, DB_PASSWORD); }
let hasDB_PASSWORD = fs.existsSync(DB_PASSWORD);
app.locals.hasDB_PASSWORD = hasDB_PASSWORD;

app.listen(port, () => console.log(`STATS is listening on port ${port}!`))

console.log("TEST STARTED");

if(hasEMAIL_PASSWORD && hasDB_PASSWORD){
    fs.readFile(EMAIL_PASSWORD, "utf8", function (err, contents) {
        if (err) {
            console.error('secret not found');
            console.error('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
            console.log(contents);
        }
    });
    
    fs.readFile(DB_PASSWORD, "utf8", function (err, contents) {
        if (err) {
            console.error('secret not found');
            console.error('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
            console.log(contents);
        }
    });
}else{
    console.log("Please check your secret configuration. Variable or bind not setted.");
}



  