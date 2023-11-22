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



app.listen(port, () => console.log(`STATS is listening on port ${port}!`))

console.log("TEST STARTED");

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

app.get('/health', function(request, response) {
    if( healthy ) {
      response.status(200);
    } else {
      response.status(500);
    }
  });  