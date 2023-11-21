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

const app = express();
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

const port = 3000;

//console.log( process.env);
let secretFile = process.env.SECRET_FILE || '/var/secret/secret.txt';
if (!path.isAbsolute(secretFile)) { secretFile = path.resolve(__dirname, secretFile); }
let hasSecret = fs.existsSync(secretFile);
console.log("hasSecret:"+hasSecret);
app.locals.hasSecret = hasSecret;

app.listen(port, () => console.log(`STATS is listening on port ${port}!`))

console.log("TEST STARTED");

fs.readFile(secretFile, function (err, contents) {
    if (err) {
      console.error('secret not found');
      console.log(response.render('error', {'msg': JSON.stringify(err, null, 4)}));
    } else {
      console.log(response.render('secrets', {'secret': contents}));
    }
  });