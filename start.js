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

let app = express();
app.use(express.static(__dirname + '/public'));
app.use(bodyParser.urlencoded({
  extended: true
}));

app.set('views', __dirname + '/views');
app.set('view engine', 'ejs');
app.set('port', process.env.PORT || 8080);

//app.use(bodyParser.urlencoded({ extended: false }));
//app.use(bodyParser.json());

const port = 3000;

//console.log( process.env);
let secretFile = process.env.SECRET_FILE || '/var/secret/secret.txt';
if (!path.isAbsolute(secretFile)) { secretFile = path.resolve(__dirname, secretFile); }
let hasSecret = fs.existsSync(secretFile);
console.log("hasSecret:"+hasSecret);
app.locals.hasSecret = hasSecret;

app.listen(port, () => console.log(`STATS is listening on port ${port}!`))

console.log("TEST STARTED");

/*
  SECRETS URLS/FUNCTIONS
 */
  if (hasSecret) {
    app.get('/secrets', function (request, response) {
      fs.readFile(secretFile, function (err, contents) {
        if (err) {
          console.error('secret not found');
          response.render('error', {'msg': JSON.stringify(err, null, 4)});
        } else {
          response.render('secrets', {'secret': contents});
        }
      });
    });
  }