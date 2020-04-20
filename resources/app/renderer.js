const electron = require('electron');
const process = require('process');
const path = require('path');
const { dialog, app } = require('electron').remote;
const BrowserWindow = electron.remote.BrowserWindow;
const WIN = electron.remote.getCurrentWindow();
const fs = require('fs');
const jsforce = require('jsforce');
const readline = require('readline');
const createDOMPurify = require('dompurify');
const Cryptr = require('cryptr');
const cryptr = new Cryptr('ThisIsHowIGo');
const { JSDOM } = require('jsdom');
const windows = (new JSDOM('')).window;
const DOMPurify = createDOMPurify(windows);
const Imap = require('imap'), inspect = require('util').inspect;
const imaps = require('imap-simple');
const XLSX = require('xlsx');
window.$ = window.jQuery = require('jquery'); // not sure if you need this at all
window.Bootstrap = require('bootstrap');
window.Popper = require('popper.js');

// Load Custom Objects
const page = require('./Objects/page');
const forms = require('./Objects/forms');
const helpers = require('./Objects/helper');
const salesforce = require('./Objects/salesforce');


// Load sales force login
page.loadSFLoginPage();

// 
$('#sfsignin').on('click', function (e) {
  e.preventDefault();
  helpers.functions.loginToSF();
});

$('#sfreset').on('click', function (e) {
  e.preventDefault();
});