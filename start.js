// Require library
var excel = require('excel4node');
const oracledb = require('oracledb');
const nodemailer = require('nodemailer');
fs = require('fs');
var CronJob = require('cron').CronJob;
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

const port = 3000;

let secretFile = process.env.SECRET_FILE || '/var/secret/secret.txt';

app.listen(port, () => console.log(`STATS is listening on port ${port}!`))

console.log("TEST STARTED");
console.log("secretFile:"+secretFile);