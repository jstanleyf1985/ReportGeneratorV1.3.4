exports.functions = {
  filterCreds: function (cred) {
    cred = String(cred);
    cred = cred.trim();
    cred = DOMPurify.sanitize(cred);
    return cred;
  },
  checkValid: function (username, password) {
    if (
      username != undefined &&
      username != null &&
      username != '' &&
      password != undefined &&
      password != null &&
      password != '') {
      return true;
    } else {
      return false;
    }
  },
  checkIfFileExists: function (pathString) {
    try {
      let filePresent = true;
      let file = fs.statSync(path.join(__dirname + '/..' + pathString));
      return filePresent;
    } catch (err) {
      // Handle missing file by examining error message
      if (!String(err).indexOf('no such file or directory') != -1) {
        let filePresent = false;
        return filePresent;
      }
    }
  },
  copyDailyCalls: function (dailyCallString) {
    try {
      let string = dailyCallString;

      let copyToClipboard = string => {
        let el = document.createElement('textarea');  // Create a <textarea> element
        el.value = string;                                 // Set its value to the string that you want copied
        el.setAttribute('readonly', '');                // Make it readonly to be tamper-proof
        el.style.position = 'absolute';
        el.style.left = '-9999px';                      // Move outside the screen to make it invisible
        document.body.appendChild(el);                  // Append the <textarea> element to the HTML document
        let selected =
          document.getSelection().rangeCount > 0        // Check if there is any content selected previously
            ? document.getSelection().getRangeAt(0)     // Store selection if found
            : false;                                    // Mark as false to know no selection existed before
        el.select();                                    // Select the <textarea> content
        document.execCommand('copy');                   // Copy - only works as a result of a user action (e.g. click events)
        document.body.removeChild(el);                  // Remove the <textarea> element
        if (selected) {                                 // If a selection existed before copying
          document.getSelection().removeAllRanges();    // Unselect everything on the HTML document
          document.getSelection().addRange(selected);   // Restore the original selection
        }
      };

      copyToClipboard(string);

      // Used to keep popup from disappearing immediately
      $('#getDailyCallsReport').focus();
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  copyDailyCases: function (dailyCasesString) {
    try {
      let string = dailyCasesString;

      let copyToClipboard = string => {
        let el = document.createElement('textarea');  // Create a <textarea> element
        el.value = string;                                 // Set its value to the string that you want copied
        el.setAttribute('readonly', '');                // Make it readonly to be tamper-proof
        el.style.position = 'absolute';
        el.style.left = '-9999px';                      // Move outside the screen to make it invisible
        document.body.appendChild(el);                  // Append the <textarea> element to the HTML document
        let selected =
          document.getSelection().rangeCount > 0        // Check if there is any content selected previously
            ? document.getSelection().getRangeAt(0)     // Store selection if found
            : false;                                    // Mark as false to know no selection existed before
        el.select();                                    // Select the <textarea> content
        document.execCommand('copy');                   // Copy - only works as a result of a user action (e.g. click events)
        document.body.removeChild(el);                  // Remove the <textarea> element
        if (selected) {                                 // If a selection existed before copying
          document.getSelection().removeAllRanges();    // Unselect everything on the HTML document
          document.getSelection().addRange(selected);   // Restore the original selection
        }
      };

      copyToClipboard(string);

      // Used to keep popup from disappearing immediately
      $('#getDailyCasesReport').focus();
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  findExcelLocations: function (document) {
    let ezGeneralObj = {
      location: false,
      call: false,
      callback: false,
      callbackExists: false,
      abandoned: false,
      wait: false,
      cbtime: false
    }

    let ezSuiteObj = {
      location: false,
      call: false,
      callback: false,
      callbackExists: false,
      abandoned: false,
      wait: false,
      cbtime: false
    }

    let ezPremierObj = {
      location: false,
      call: false,
      callback: false,
      callbackExists: false,
      abandoned: false,
      wait: false,
      cbtime: false
    }

    let voicemailObj = {
      location: false,
      voicemailExists: false,
      voicemails: false
    }

    // Find the Key name locations, use key name locations as a reference location to obtain required data
    for (let doc in document) {
      // Get location for EzGeneralSupport
      if (ezGeneralObj.callbackExists == false) {
        if (document[doc].v == 'EzGeneralSupport') {
          ezGeneralObj.location = String(doc);
          ezGeneralObj.callbackExists = true;

          // Check for ezGenCBExists,
          let locLetter = String((ezGeneralObj.location).substr(0, 1));
          let locNum = parseInt((ezGeneralObj.location).substr(1, (ezGeneralObj.location).length));
          let locationCBCheck = locLetter.concat(String(locNum + 5));
          let locationCall = `C${locNum + 2}`;
          let locationAban = `M${locNum + 2}`;
          let locationWait = `U${locNum + 2}`;
          if (typeof document[locationCBCheck] == 'undefined') {
            ezGeneralObj.call = document[locationCall].v;
            ezGeneralObj.abandoned = document[locationAban].v;
            ezGeneralObj.wait = document[locationWait].v;
            ezGeneralObj.callback = '0';
            ezGeneralObj.cbtime = '0:00'
          } else {
            // Check if equal to 'Callback' if so, assign relevant data
            if (document[locationCBCheck].v == 'Callback') {
              let locationCallback = `C${locNum + 5}`;
              let locationCBTime = `U${locNum + 5}`;
              ezGeneralObj.call = document[locationCall].v;
              ezGeneralObj.abandoned = document[locationAban].v;
              ezGeneralObj.wait = document[locationWait].v;
              ezGeneralObj.callback = document[locationCallback].v;
              ezGeneralObj.cbtime = document[locationCBTime].v;
              ezGeneralObj.callbackExists = true;
            }
          }
        }
      }

      // Get location for EzSuiteSupport
      if (ezSuiteObj.callbackExists == false) {
        if (document[doc].v == 'EzSuiteSupport') {
          ezSuiteObj.location = String(doc);
          ezSuiteObj.callbackExists = true;

          // Check for ezGenCBExists,
          let locLetter = String((ezSuiteObj.location).substr(0, 1));
          let locNum = parseInt((ezSuiteObj.location).substr(1, (ezSuiteObj.location).length));
          let locationCBCheck = locLetter.concat(String(locNum + 5));
          let locationCall = `C${locNum + 2}`;
          let locationAban = `M${locNum + 2}`;
          let locationWait = `U${locNum + 2}`;
          if (typeof document[locationCBCheck] == 'undefined') {
            ezSuiteObj.call = document[locationCall].v;
            ezSuiteObj.abandoned = document[locationAban].v;
            ezSuiteObj.wait = document[locationWait].v;
            ezSuiteObj.callback = '0';
            ezSuiteObj.cbtime = '0:00'
          } else {
            // Check if equal to 'Callback' if so, assign relevant data
            if (document[locationCBCheck].v == 'Callback') {
              let locationCallback = `C${locNum + 5}`;
              let locationCBTime = `U${locNum + 5}`;
              ezSuiteObj.call = document[locationCall].v;
              ezSuiteObj.abandoned = document[locationAban].v;
              ezSuiteObj.wait = document[locationWait].v;
              ezSuiteObj.callback = document[locationCallback].v;
              ezSuiteObj.cbtime = document[locationCBTime].v;
              ezSuiteObj.callbackExists = true;
            }
          }
        }
      }

      // Get location for EzPremierSupport
      if (ezPremierObj.callbackExists == false) {
        if (document[doc].v == 'EzPremierSupport') {
          ezPremierObj.location = String(doc);
          ezPremierObj.callbackExists = true;

          // Check for ezGenCBExists,
          let locLetter = String((ezPremierObj.location).substr(0, 1));
          let locNum = parseInt((ezPremierObj.location).substr(1, (ezPremierObj.location).length));
          let locationCBCheck = locLetter.concat(String(locNum + 5));
          let locationCall = `C${locNum + 2}`;
          let locationAban = `M${locNum + 2}`;
          let locationWait = `U${locNum + 2}`;
          if (typeof document[locationCBCheck] == 'undefined') {
            ezPremierObj.call = document[locationCall].v;
            ezPremierObj.abandoned = document[locationAban].v;
            ezPremierObj.wait = document[locationWait].v;
            ezPremierObj.callback = '0';
            ezPremierObj.cbtime = '0:00'
          } else {
            // Check if equal to 'Callback' if so, assign relevant data
            if (document[locationCBCheck].v == 'Callback') {
              let locationCallback = `C${locNum + 5}`;
              let locationCBTime = `U${locNum + 5}`;
              ezPremierObj.call = document[locationCall].v;
              ezPremierObj.abandoned = document[locationAban].v;
              ezPremierObj.wait = document[locationWait].v;
              ezPremierObj.callback = document[locationCallback].v;
              ezPremierObj.cbtime = document[locationCBTime].v;
              ezPremierObj.callbackExists = true;
            }
          }
        }
      }

      // Get location for voicemails
      if (voicemailObj.voicemailExists == false) {
        if (document[doc].v == 'zFlowOuts') {
          voicemailObj.location = String(doc);
          voicemailObj.voicemailExists = true;
          let locLetter = String((voicemailObj.location).substr(0, 1));
          let locNum = parseInt((voicemailObj.location).substr(1, (voicemailObj.location).length));
          let locCallCheck = locLetter.concat(String(locNum + 2));
          if (typeof document[locCallCheck] == 'undefined') {
            voicemailObj.voicemails = '0';
          } else {
            let voicemailsLoc = `Y${locNum + 2}`;
            voicemailObj.voicemails = document[voicemailsLoc].v;
          }
        } else {
          // No voicemails, set to 0
          voicemailObj.voicemails = '0';
        }

      }
    }

    // Return an object of keyname: objects
    let reportValuesObj = { ezGeneral: ezGeneralObj, ezSuite: ezSuiteObj, ezPremier: ezPremierObj, voicemails: voicemailObj };
    return reportValuesObj;
  },
  getEmailFromConfig: function () {
    try {
      let emailConfigExists = this.checkIfFileExists('/Configs/EmailConfig.txt');
      let userEmailObj = { username: '', password: '', checked: false };
      if (emailConfigExists == true) {
        // Exists, now read it and extract username and password
        let emailConfigContents = fs.readFileSync(path.join(__dirname + '/../Configs/EmailConfig.txt'), 'utf-8');
        // Filter contents
        emailConfigContents = DOMPurify.sanitize(emailConfigContents);

        // Create indexes to cut get values with
        let userIndex = emailConfigContents.indexOf('USERNAME: "');
        let userStopIndex = emailConfigContents.indexOf('\n', userIndex);
        let pwdIndex = emailConfigContents.indexOf('PASSWORD: "');
        let pwdStopIndex = emailConfigContents.indexOf('\n', pwdIndex);
        let checkedIndex = emailConfigContents.indexOf('CHECKED: "');
        let checkedStopIndex = emailConfigContents.indexOf('\n', checkedIndex);

        // Check if username or password is invalid, if so just return blank values due to invalid config file
        if (userIndex == -1 || pwdIndex == -1 || checkedIndex == -1 || emailConfigContents.length < 5) {
          userEmailObj = { username: '', password: '', checked: false };
        } else {
          // Obtain values
          userEmailObj.username = emailConfigContents.slice(userIndex, userStopIndex);
          userEmailObj.username = String(userEmailObj.username).replace('USERNAME: ', '');
          userEmailObj.username = String(userEmailObj.username).replace(/"/gi, '');
          userEmailObj.password = emailConfigContents.slice(pwdIndex, pwdStopIndex);
          userEmailObj.password = String(userEmailObj.password).replace('PASSWORD: ', '');
          userEmailObj.password = String(userEmailObj.password).replace(/"/gi, '');
          userEmailObj.checked = emailConfigContents.slice(checkedIndex, checkedStopIndex);
          userEmailObj.checked = String(userEmailObj.checked).replace('CHECKED: ', '');
          userEmailObj.checked = String(userEmailObj.checked).replace(/"/gi, '');
        }
        return userEmailObj;
      } else {
        return userEmailObj;
      }
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  getTodayUNIX: function () {
    try {
      let today = new Date();
      let day = today.getDate();
      let month = today.getMonth();
      let months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      let year = today.getFullYear();
      let todayUNIX = Date.parse(`${months[month]} ${day}, ${year}`);
      todayUNIX = String(todayUNIX);
      todayUNIX = todayUNIX.substr(0, todayUNIX.length - 3);
      todayUNIX = parseInt(todayUNIX);
      return todayUNIX;
    } catch (err) { }

  },
  getUnix: function (dateObj) {
    try {
      let day = dateObj.getDate();
      let month = dateObj.getMonth();
      let months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      let year = dateObj.getFullYear();
      let unixdate = Date.parse(`${months[month]} ${day}, ${year}`);
      unixdate = String(unixdate);
      unixdate = unixdate.substr(0, unixdate.length - 3);
      unixdate = parseInt(unixdate);
      return unixdate;
    } catch (err) { }
  },
  getDayOfYear: function () {
    try {
      let now = new Date();
      let start = new Date(now.getFullYear(), 0, 0);
      let diff = (now - start) + ((start.getTimezoneOffset() - now.getTimezoneOffset()) * 60 * 1000);
      let oneDay = 1000 * 60 * 60 * 24;
      let day = Math.floor(diff / oneDay);
      return day;
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  getDateString: function () {
    try {
      let monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      let today = new Date();
      let todayDay = String(today.getDate());
      let todayMonth = String(monthNames[today.getMonth()]);
      let todayYear = String(today.getFullYear());
      let dateString = `${todayMonth} ${todayDay}, ${todayYear}`;
      return dateString;
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  getDateStringYesterday: function () {
    try {
      let today = new Date();
      let todayUnixGMT = parseInt(today.getTime());
      //let offset = parseInt(today.getTimezoneOffset());
      let yesterday = todayUnixGMT;
      /* if (offset != 0) {
        let offSetMilliseconds = parseInt(offset * 60000);

        if (offset > 0) {
          // Subtract offset
          yesterday = todayUnixGMT - offSetMilliseconds;
        } else {
          // Add offset
          yesterday = todayUnixGMT + offSetMilliseconds;
        }
      } else {
        yesterday = todayUnixGMT;
      }
      */
      let yesterdayDate = yesterday - (24 * (60 * 60000));
      yesterdayDate = new Date(yesterdayDate);

      let monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      let yesterdayDay = String(yesterdayDate.getDate());
      let yesterdayMonth = String(monthNames[yesterdayDate.getMonth()]);
      let yesterdayYear = String(yesterdayDate.getFullYear());
      let dateString = `${yesterdayMonth} ${yesterdayDay}, ${yesterdayYear}`;
      return dateString;
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  getTodayMMDDYYYY: function () {
    try {
      let today = new Date();
      let dd = today.getDate();
      let mm = today.getMonth() + 1; //January is 0!

      let yyyy = today.getFullYear();
      if (dd < 10) {
        dd = '0' + dd;
      }
      if (mm < 10) {
        mm = '0' + mm;
      }
      today = mm + '/' + dd + '/' + yyyy;
      return today;
    } catch (err) { dialog.showErrorBox('Error', String(err)); }

  },
  loadingPage: function () {
    // Used to display loading page while fetching data, after logging into email successfully
    $('#main').empty();
    $('#main').prepend(forms.loading);
  },
  loginToSF: function () {
    try {
      // Show uploading message
      $('#sfSignInErr').removeClass('invisible redText greenText').addClass('blueText visible').text('Logging Into SalesForce . . . .');

      // Proceed to login to sales force
      let sfUsername = this.filterCreds(String($('#sfusername').val()));
      let sfPassword = this.filterCreds(String($('#sfpassword').val()));
      let sfChecked = this.filterCreds(String($('#sfChecked').is(':checked')));

      let conn = new jsforce.Connection();
      conn.login(sfUsername, sfPassword, function (err, userInfo) {
        if (err) {
          let errMSG = String(err);
          //dialog.showErrorBox('Error', String(err));
          if (errMSG.indexOf('INVALID_LOGIN') != -1) {
            errMSG = 'Invalid Login Credentials';
          } else if (errMSG.indexOf('getaddrinfo') != -1) {
            errMSG = 'SalesForce Unreachable / Network Error';
          }

          // Set err message
          $('#sfSignInErr').removeClass('invisible blueText greenText').addClass('redText visible').text(`${errMSG}`);
        } else {
          // No errors, proceed to obtaining report information and saving the contents
          $('#main').empty();
          $('#main').prepend(forms.emailForm);

          // If CHECKED is true, save credentials, else delete config file for sales force
          if (sfChecked == true || sfChecked == 'true') {
            try {
              let configData = `USERNAME: "${sfUsername}"\nPASSWORD: "${sfPassword}"\nCHECKED: "${sfChecked}"`;
              fs.writeFileSync(path.join(__dirname + '/../Configs/LoginConfig.txt'), configData);

              // Load email login section, pass sf login connection along, to be used after email parsing
              page.loadEmailLoginPage(conn);
            } catch (err) { dialog.showErrorBox('Error', String(err)); }
          } else {
            // Delete sales force configuration file
            try {
              let configData = `USERNAME: "${sfUsername}"\nPASSWORD: "${sfPassword}"\nCHECKED: "${sfChecked}"`;
              fs.writeFileSync(path.join(__dirname + '/../Configs/LoginConfig.txt'), configData);

              // Load email login section, pass sf login connection along, to be used after email parsing
              page.loadEmailLoginPage(conn);

              // Deletes login config if one exists, if not checked
              fs.unlinkSync(path.join(__dirname + '/../Configs/LoginConfig.txt'));
            } catch (err) {/* No one cares if it doesn't delete a file that doesn't exist */ }
          }
        }
      });
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  parseEmailAttachments: function (emailUsername, emailPW, atTheTurnString, atTheTurnHTML, inTheClubhouseString, inTheClubhouseHTML, reportTotObj, totalCasesOpen) {
    try {
      let thisContext = this;
      // Remove existing mail attachments before creating new ones // this avoids parsing old files and cleans up storage space
      try {
        mailAttachmentsList = fs.readdirSync(path.join(__dirname + '/../MailAttachments'));
        for (let attachment in mailAttachmentsList) {
          fs.unlinkSync(path.join(__dirname + '/../MailAttachments/' + mailAttachmentsList[attachment]));
        }
      } catch (err) { /* Do nothing, errors will be silent */ }


      // Everything complete here, load final page
      var attachConfig = {
        imap: {
          user: emailUsername,
          password: emailPW,
          host: 'outlook.office365.com',
          port: 993,
          tls: true,
          connTimeout: 5000,
          authTimeout: 5000,
          socketTimeout: 5000
        }
      }

      imaps.connect(attachConfig).then(function (connection) {

        connection.openBox('INBOX').then(function () {

          // Fetch emails from today/yesterday
          let today = new Date();
          let searchSubject = 'Scheduled Report: SupportDesk';
          today = today.toISOString();
          today = today.slice(0, today.indexOf('T'));
          today = today.concat('T00:00:00.999Z');
          let searchCriteria = ['ALL', ['SINCE', today], ['SUBJECT', searchSubject]];
          let fetchOptions = { bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'], struct: true };

          // retrieve only the headers of the messages
          return connection.search(searchCriteria, fetchOptions);
        }).then(function (messages) {
          var attachments = [];

          messages.forEach(function (message) {
            var parts = imaps.getParts(message.attributes.struct);
            attachments = attachments.concat(parts.filter(function (part) {
              return part.disposition && part.disposition.type.toUpperCase() === 'ATTACHMENT';
            }).map(function (part) {
              // retrieve the attachments only of the messages with attachments
              return connection.getPartData(message, part)
                .then(function (partData) {
                  return {
                    filename: part.disposition.params.filename,
                    data: partData
                  };

                });
            }));
          });

          return Promise.all(attachments);
        }).then(function (attachments) {
          try {
            // Create variables that will hold the data requested from the attachments
            // ********* REPORTS ARE IN CHICAGO TIME, use lazy man's method of checking if reports are in the desired time range
            let todayUNIX = thisContext.getTodayUNIX();
            let today12AMUnix = todayUNIX - 18000;
            let yesterday12AMUnix = today12AMUnix - 86400;
            let yEndDayReportMin = yesterday12AMUnix + 72000;
            let yEndDayReportMax = yesterday12AMUnix + 86400;
            let middayReportMin = today12AMUnix + 39600;  // 11AM
            let middayReportMax = today12AMUnix + 54000;  // 4PM
            let enddayReportMin = today12AMUnix + 72000;  // 8PM
            let enddayReportMax = today12AMUnix + 86400;  // 11:59PM
            let yEndDayReport = false;
            let middayReport = false;
            let enddayReport = false;

            // Loop through files and parse data, put results in middayReportData or enddayReportData based on print date/time variables in data
            for (let file in attachments) {
              let filename = String(attachments[file].filename);

              if (filename.indexOf('Distribution Queue Performance') != -1) {
                // Match exists, save file data to file
                if (attachments[file].hasOwnProperty('data')) {
                  fs.writeFileSync(path.join(__dirname + `/../MailAttachments/DistributionQueuePerformance${file}.xls`), attachments[file].data);
                  let workbook = XLSX.readFile(path.join(__dirname + `/../MailAttachments/DistributionQueuePerformance${file}.xls`));
                  let report = workbook.Sheets.Sheet1;
                  // Create midday or endday report properties using data
                  for (let item in report) {
                    if (String(report[item].v).indexOf('Print Date:') != -1) {
                      let baseHours = 0;
                      let dateTimeData = String(report[item].v).split(',');
                      let dateTimeYearDate = String(dateTimeData[2]).split(' ');
                      let dateTimeMonthDay = String(dateTimeData[1]);
                      // Filter blank values out of the dateTimeMonthDay array due to excessive blank spaces present in the report data
                      dateTimeYearDate = dateTimeYearDate.filter(function (element) {
                        return element != '';
                      });
                      let dateTimeYear = String(dateTimeYearDate[0]);
                      let dateTimeTime = String(dateTimeYearDate[1]);

                      // Convert HH:MM:SS AM/PM into seconds and add to date.parse time to get actual time
                      if (dateTimeTime.indexOf('PM') != -1) { baseHours = 12; }

                      // Convert time data into seconds to add to reportDateTime total
                      let HHMMSS = String(dateTimeTime.substr(0, dateTimeTime.length - 2));
                      HHMMSS = dateTimeTime.split(':');

                      let HH2S = parseInt(HHMMSS[0]);
                      let MM2S = parseInt(HHMMSS[1]);
                      let SS2S = parseInt(HHMMSS[2]);

                      // Append data to dateTimeTime to create a complete string that can be used to successfully parse it to seconds
                      let reportDTUNIX = Date.parse(`${dateTimeMonthDay}, ${dateTimeYear} ${HH2S + baseHours}:${MM2S}:${SS2S}`);

                      reportDTUNIX = String(reportDTUNIX);
                      reportDTUNIX = reportDTUNIX.substr(0, reportDTUNIX.length - 3);
                      reportDTUNIX = parseInt(reportDTUNIX);
                      reportDTUNIX = reportDTUNIX - 18000; // Adjust to eastern from UTC

                      if (reportDTUNIX > yEndDayReportMin && reportDTUNIX < yEndDayReportMax) {
                        // Determine locations of key values, then grab data from said locations
                        yEndDayReport = thisContext.findExcelLocations(report);
                      } else if (reportDTUNIX > middayReportMin && reportDTUNIX < middayReportMax) {
                        // Determine locations of key values, then grab data from said locations
                        middayReport = thisContext.findExcelLocations(report);
                      } else if (reportDTUNIX > enddayReportMin && reportDTUNIX < enddayReportMax) {
                        // Determine locations of key values, then grab data from said locations
                        enddayReport = thisContext.findExcelLocations(report);
                      }
                    }
                  }
                }
              } else {
                // File doesn't exist yet (either ran too early or wasn't sent out, set a variable that tracks these for use later)
              }
            }

            /* Only 3 reports should exist that are important, yesterday enddayreport, today midday report and today's enddayreport */
            // Create middayreport abandoned percentage and add into string
            if (middayReport != false) {
              let middayGenCallBack = typeof middayReport['ezGeneral'].callback == 'undefined' ? 0 : parseInt(middayReport['ezGeneral'].callback);
              let middayTurnCallBack = typeof middayReport['ezSuite'].callback == 'undefined' ? 0 : parseInt(middayReport['ezSuite'].callback);

              let middayCallTotal = parseInt(middayReport['ezGeneral'].call) + parseInt(middayReport['ezSuite'].call);
              let middayCallBackTotal = parseInt(middayGenCallBack + middayTurnCallBack);
              let middayAbanTotal = parseInt(middayReport['ezGeneral'].abandoned) + parseInt(middayReport['ezSuite'].abandoned);
              let middayCombinedTotals = middayCallTotal + middayCallBackTotal + middayAbanTotal;
              let middayAbanP = Math.round((middayAbanTotal / (middayCallTotal + middayCallBackTotal + middayAbanTotal)) * 100).toFixed(0);

              atTheTurnString = atTheTurnString.replace('XX', middayAbanP);
              atTheTurnString = atTheTurnString.replace('[Daily Calls Column H] %', `${middayAbanP}%`);
              atTheTurnString = atTheTurnString.replace('[Daily Calls Column E]', middayAbanTotal);
              atTheTurnString = atTheTurnString.replace('[Daily Calls Column F]', middayCombinedTotals);

              atTheTurnHTML = atTheTurnHTML.replace('XX', middayAbanP);
              atTheTurnHTML = atTheTurnHTML.replace('[Daily Calls Column H] %', `${middayAbanP}%`);
              atTheTurnHTML = atTheTurnHTML.replace('[Daily Calls Column E]', middayAbanTotal);
              atTheTurnHTML = atTheTurnHTML.replace('[Daily Calls Column F]', middayCombinedTotals);
            }

            if (enddayReport != false) {
              // Create enddayreport abandoned percentage and add into string
              let enddayGenCallBack = typeof enddayReport['ezGeneral'].callback == 'undefined' ? 0 : parseInt(enddayReport['ezGeneral'].callback);
              let enddayTurnCallBack = typeof enddayReport['ezSuite'].callback == 'undefined' ? 0 : parseInt(enddayReport['ezSuite'].callback);

              let enddayCallTotal = parseInt(enddayReport['ezGeneral'].call) + parseInt(enddayReport['ezSuite'].call);
              let enddayCallBackTotal = parseInt(enddayGenCallBack + enddayTurnCallBack);
              let enddayAbanTotal = parseInt(enddayReport['ezGeneral'].abandoned) + parseInt(enddayReport['ezSuite'].abandoned);
              let enddayCombinedTotals = enddayCallTotal + enddayCallBackTotal + enddayAbanTotal;
              let enddayAbanP = Math.round((enddayAbanTotal / (enddayCallTotal + enddayCallBackTotal + enddayAbanTotal)) * 100).toFixed(0);

              inTheClubhouseString = inTheClubhouseString.replace('XX', enddayAbanP);
              inTheClubhouseString = inTheClubhouseString.replace('[Daily Calls Column H] %', `${enddayAbanP}%`);
              inTheClubhouseString = inTheClubhouseString.replace('[Daily Calls Column E]', enddayAbanTotal);
              inTheClubhouseString = inTheClubhouseString.replace('[Daily Calls Column F]', enddayCombinedTotals);

              inTheClubhouseHTML = inTheClubhouseHTML.replace('XX', enddayAbanP);
              inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Calls Column H] %', `${enddayAbanP}%`);
              inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Calls Column E]', enddayAbanTotal);
              inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Calls Column F]', enddayCombinedTotals);
            }


            // Create dailyCases2020 String and dailyCalls2020 String to paste into excel reports
            let dayOfYear = parseInt(thisContext.getDayOfYear());
            let dateMMDDYYYY = thisContext.getTodayMMDDYYYY();
            let todayIndex = new Date().getDay();
            let daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
            let todayDOW = daysOfWeek[todayIndex];

            let repNum = dayOfYear;
            let repNum2 = dayOfYear + 1; // Used for daily calls report since it begins at 1
            let dailyCasesC = reportTotObj.createdByRep;
            let dailyCasesD = reportTotObj.stillOpenToday;
            let dailyCasesE = `=(C${repNum2}-D${repNum2})`;
            let dailyCasesF = `=(E${repNum2}/C${repNum2})`;
            let dailyCasesG = reportTotObj.casesClosedByRep;
            let dailyCasesH = reportTotObj.totalCasesOpen;
            let dailyCasesI = `=(H${repNum2}-H${repNum2 - 1})`;

            let dailyCallsC = `=M${repNum2}+U${repNum2}`;
            let dailyCallsD = `=N${repNum2}+V${repNum2}`;
            let dailyCallsE = `=O${repNum2}+W${repNum2}`;
            let dailyCallsF = `=SUM(C${repNum2}:E${repNum2})`;
            let dailyCallsG = `=Q${repNum2}+Y${repNum2}`;
            let dailyCallsH = `=E${repNum2}/F${repNum2}`;
            let dailyCallsI = `=C${repNum2}/F${repNum2}`;
            let dailyCallsJ = `=D${repNum2}/G${repNum2}`;
            let dailyCallsK = `=AVERAGE(S${repNum2},AA${repNum2})`;
            let dailyCallsL = `=AVERAGE(T${repNum2},AB${repNum2})`;
            let dailyCallsM = typeof enddayReport.ezSuite == 'undefined' ? 0 : `${enddayReport.ezSuite.call}`;
            let dailyCallsN = typeof enddayReport.ezSuite == 'undefined' ? 0 : `${enddayReport.ezSuite.callback}`;
            let dailyCallsO = typeof enddayReport.ezSuite == 'undefined' ? 0 : `${enddayReport.ezSuite.abandoned}`;
            let dailyCallsP = `=M${repNum2}+N${repNum2}+O${repNum2}`;
            let dailyCallsQ = `=O${repNum2}/P${repNum2}`;
            let dailyCallsR = `=M${repNum2}/P${repNum2}`;
            let dailyCallsS = typeof enddayReport.ezSuite == 'undefined' ? 0 : `${enddayReport.ezSuite.wait}`;
            let dailyCallsT = typeof enddayReport.ezSuite == 'undefined' ? 0 : `${enddayReport.ezSuite.cbtime}`;
            let dailyCallsU = typeof enddayReport.ezGeneral == 'undefined' ? 0 : `${enddayReport.ezGeneral.call}`;
            let dailyCallsV = typeof enddayReport.ezGeneral == 'undefined' ? 0 : `${enddayReport.ezGeneral.callback}`;
            let dailyCallsW = typeof enddayReport.ezGeneral == 'undefined' ? 0 : `${enddayReport.ezGeneral.abandoned}`;
            let dailyCallsX = `=U${repNum2}+V${repNum2}+W${repNum2}`;
            let dailyCallsY = `=W${repNum2}/X${repNum2}`;
            let dailyCallsZ = `=U${repNum2}/X${repNum2}`;
            let dailyCallsAA = typeof enddayReport.ezGeneral == 'undefined' ? 0 : `${enddayReport.ezGeneral.wait}`;
            let dailyCallsAB = typeof enddayReport.ezGeneral == 'undefined' ? 0 : `${enddayReport.ezGeneral.cbtime}`;
            let voicemailTotal = typeof enddayReport.voicemails == 'undefined' ? 0 : `${enddayReport.voicemails.voicemails}`;

            // New daily calls structure was added by bob, new values are below for only those values that have changed
            let dailyCallsCNew = parseInt(dailyCallsU); // ezGenCallsAnswered
            let dailyCallsDNew = parseInt(dailyCallsM); // ezSuiteCallsAnswered
            let dailyCallsENew = parseInt(dailyCallsV); // ezGenCallbacks
            let dailyCallsFNew = parseInt(dailyCallsN); // ezSuiteCallbacks
            let dailyCallsGNew = parseInt(dailyCallsO) + parseInt(dailyCallsW); // Total abandoned calls
            let dailyCallsHNew = parseInt(voicemailTotal); // voicemail total
            let dailyCallsINew = parseInt(dailyCallsM) + parseInt(dailyCallsN) + parseInt(dailyCallsO) + parseInt(dailyCallsU) + parseInt(dailyCallsV) + parseInt(dailyCallsW);
            let dailyCallsJNew = dailyCallsGNew / dailyCallsINew;// abandoned rate%
            let dailyCallsKNew = (dailyCallsCNew + dailyCallsDNew) / dailyCallsINew;
            let dailyCallsLNew = dailyCallsHNew / dailyCallsINew;

            // Calculated separately as it has to be parsed HH:MM:SS format
            let dailyCallsMSuite = String(dailyCallsS).split(':');
            let dailyCallsMezGen = String(dailyCallsAA).split(':');
            let dailyCallsMSuiteHours = dailyCallsMSuite[0] == '' ? '0' : dailyCallsMSuite[0];
            let dailyCallsMezGenHours = dailyCallsMezGen[0] == '' ? '0' : dailyCallsMezGen[0];
            let dailyCallsMSuiteMins = dailyCallsMSuite[1] == '' ? '0' : dailyCallsMSuite[1];
            let dailyCallsMezGenMins = dailyCallsMezGen[1] == '' ? '0' : dailyCallsMezGen[1];
            let dailyCallsMSuiteSecs = dailyCallsMSuite[2] == '' ? '0' : dailyCallsMSuite[2];
            let dailyCallsMezGenSecs = dailyCallsMezGen[2] == '' ? '0' : dailyCallsMezGen[2];
            let mFinalHrs = '00';
            let mFinalMins = '00';
            let mFinalSecs = '00';

            // Parse to int before mathing
            if (dailyCallsMSuiteHours != '0' || dailyCallsMezGenHours != '0') {
              // Hours exists, sum them
              mFinalHrs = (parseInt(dailyCallsMSuiteHours) + parseInt(dailyCallsMezGenHours)) / 2;
              mFinalHrs = parseFloat(mFinalHrs);
              mFinalHrs = Math.ceil(mFinalHrs);
              mFinalHrs = mFinalHrs > 0 && mFinalHrs < 9 ? `0${mFinalHrs}` : mFinalHrs;
            }

            if (dailyCallsMSuiteMins != '0' || dailyCallsMezGenMins != '0') {
              // Mins exist, sum them
              mFinalMins = (parseInt(dailyCallsMSuiteMins) + parseInt(dailyCallsMezGenMins)) / 2;
              mFinalMins = parseFloat(mFinalMins);
              mFinalMins = Math.ceil(mFinalMins);
              mFinalMins = mFinalMins > 0 && mFinalMins < 9 ? `0${mFinalMins}` : mFinalMins;
            }

            if (dailyCallsMSuiteSecs != '0' || dailyCallsMezGenSecs != '0') {
              // Mins exist, sum them
              mFinalSecs = (parseInt(dailyCallsMSuiteSecs) + parseInt(dailyCallsMezGenSecs)) / 2;
              mFinalSecs = parseFloat(mFinalSecs);
              mFinalSecs = Math.ceil(mFinalSecs);
              mFinalSecs = mFinalSecs > 0 && mFinalSecs < 9 ? `0${mFinalSecs}` : mFinalSecs;
            }

            // Determine if each is below 9, if so add 0 into string before HH MM or SS to create pattern HH:MM:SS
            let dailyCallsMNew = `${mFinalHrs}:${mFinalMins}:${mFinalSecs}`;


            let dailyCasesString = `${dateMMDDYYYY}\t${todayDOW}\t${dailyCasesC}\t${dailyCasesD}\t${dailyCasesE}\t${dailyCasesF}\t${dailyCasesG}\t${dailyCasesH}\t${dailyCasesI}`;
            let dailyCallsStringOld /* Unused now */ = `${dateMMDDYYYY}\t${todayDOW}\t${dailyCallsC}\t${dailyCallsD}\t${dailyCallsE}\t${dailyCallsF}\t${dailyCallsG}\t${dailyCallsH}\t${dailyCallsI}\t${dailyCallsJ}\t${dailyCallsK}\t${dailyCallsL}\t${dailyCallsM}\t${dailyCallsN}\t${dailyCallsO}\t${dailyCallsP}\t${dailyCallsQ}\t${dailyCallsR}\t${dailyCallsS}\t${dailyCallsT}\t${dailyCallsU}\t${dailyCallsV}\t${dailyCallsW}\t${dailyCallsX}\t${dailyCallsY}\t${dailyCallsZ}\t${dailyCallsAA}\t${dailyCallsAB}\t`;
            let dailyCallsString = `${dateMMDDYYYY}\t${todayDOW}\t${dailyCallsCNew}\t${dailyCallsDNew}\t${dailyCallsENew}\t${dailyCallsFNew}\t${dailyCallsGNew}\t${dailyCallsHNew}\t${dailyCallsINew}\t${dailyCallsJNew}\t${dailyCallsKNew}\t${dailyCallsLNew}\t${dailyCallsMNew}`;

            // Lastly build [I] value into into HTML and text In The Clubhouse strings
            thisContext.parseYesterdayEmailForIValue(emailUsername, emailPW, atTheTurnString, atTheTurnHTML, inTheClubhouseString, inTheClubhouseHTML, dailyCasesString, dailyCallsString, totalCasesOpen);

          } catch (err) { dialog.showErrorBox('Error', String(err)); }

        });
      });


    } catch (err) { dialog.showErrorBox('Error', String(err)); }

  },
  parseYesterdayEmailForIValue: function (emailUsername, emailPW, atTheTurnString, atTheTurnHTML, inTheClubhouseString, inTheClubhouseHTML, dailyCasesString, dailyCallsString, totalCasesOpen) {
    try {
      let thisContext = this;

      // Delete folder contents before populating mail messages
      try {
        mailYesterdayList = fs.readdirSync(path.join(__dirname + '/../Mail'));
        if (mailYesterdayList.length > 0) {
          for (let mailItem in mailYesterdayList) {
            fs.unlinkSync(path.join(__dirname + '/../Mail/' + mailYesterdayList[mailItem]));
          }
        }
      } catch (err) { /* Do nothing, errors will be silent */ }

      // Process mail messages for yesterday based on Subject
      let dateString = thisContext.getDateStringYesterday();

      let imap = new Imap({
        user: emailUsername,
        password: emailPW,
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
        connTimeout: 5000,
        authTimeout: 5000,
        socketTimeout: 5000
      });


      function openInbox(cb) {
        imap.openBox('INBOX', true, cb);
      }
      imap.once('ready', function () {
        openInbox(function (err, box) {
          if (err) throw err;
          imap.search(['ALL', ['SINCE', `${dateString}`], ['SUBJECT', 'In the Clubhouse']], function (err, results) {
            if (err) throw err;
            //var f = imap.fetch(results, { bodies: '' });
            var f = imap.fetch(results, { bodies: 'TEXT', struct: true });

            f.on('message', function (msg, seqno) {
              //console.log('Message #%d', seqno);
              var prefix = '(#' + seqno + ') ';
              msg.on('body', function (stream, info) {
                //console.log(prefix + 'Body');
                stream.pipe(fs.createWriteStream(path.join(__dirname + '/../Mail/Message#' + seqno + '.txt')));
              });
              msg.once('attributes', function (attrs) {
                //console.log(prefix + 'Attributes: %s', inspect(attrs, false, 8));
              });
              msg.once('end', function () {
                //console.log(prefix + 'Finished');
              });
            });
            f.once('error', function (err) {
              // Doesn't fire on invalid credentials
            });
            f.once('end', function () {
              imap.end();
            });
          });
        });
      });
      imap.once('error', function (err) {
        try {
          let errMSG = String(err);
          if (errMSG.indexOf('LOGIN failed') != -1) {
            $('#emailSignInErr').removeClass('blueText greenText invisible').addClass('redText visible').text('Failed to Authenticate');
            imapErrThrown = true;
          } else {
            $('#emailSignInErr').removeClass('blueText greenText invisible').addClass('redText visible').text(`${errMSG}`);
            imapErrThrown = true;
          }
        } catch (err) { dialog.showErrorBox('Error', String(err)); }

      });
      imap.once('end', function () {
        try {
          let mailDir = fs.readdirSync(path.join(__dirname + '/../Mail'));
          let message = '[I]';
          let messageArray = ["[I]"]; // In case of extra "In the Clubhouse" being found, use the last instance
          for (let mailItem in mailDir) {
            let mailMessage = fs.readFileSync(path.join(__dirname + `/../Mail/${mailDir[mailItem]}`), 'utf-8');
            let mailMessagePurified = DOMPurify.sanitize(mailMessage);
            let messageStart = mailMessagePurified.indexOf('total cases closed.');
            let messageEnd = mailMessagePurified.indexOf('cases remain open');
            if (messageStart != -1 && messageEnd != -1) {
              message = mailMessagePurified.slice(messageStart, messageEnd);
              messageArray.push(message);
            }
          }


          // Remove HTML and trim results
          let lastMessage = messageArray.length - 1;
          let I = "[I]";
          messageArray[lastMessage] = String(messageArray[lastMessage]).replace(/<b>/gi, '');
          messageArray[lastMessage] = String(messageArray[lastMessage]).replace(/<\/b>/gi, '');
          messageArray[lastMessage] = String(messageArray[lastMessage]).replace(/<br>/gi, '');
          messageArray[lastMessage] = String(messageArray[lastMessage]).replace('total cases closed.', '');
          messageArray[lastMessage] = String(messageArray[lastMessage]).trim();

          // If value was found return yesterday's I value, else return [I] string, return as string either way
          totCasesYesterday = messageArray.length == 1 ? "[I]" : messageArray[lastMessage];

          // Parse I, making sure both H and I are numbers
          totCasesYesterday = parseInt(totCasesYesterday);
          totCasesToday = parseInt(totalCasesOpen);

          if (typeof totCasesYesterday == 'number' && typeof totCasesToday == 'number' && (isNaN(totCasesToday) == false && isNaN(totCasesYesterday) == false)) {
            I = totCasesToday - totCasesYesterday;
          }

          // Replace final [I] value (stitched in a hotfix) in HTML and Text versions of the InTheClubhouse Strings
          inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Cases Column I]', I);
          inTheClubhouseString = inTheClubhouseString.replace('[Daily Cases Column I]', I);

          // Build last page (GUI for copying generated data)
          page.loadFinalPage(atTheTurnString, atTheTurnHTML, inTheClubhouseString, inTheClubhouseHTML, dailyCasesString, dailyCallsString);


        } catch (err) { dialog.showErrorBox('Error', String(err)); }
      });
      imap.connect();
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  parseEmailToday: function (emailUsername, emailPW, emailChecked, sfconn) {
    try {
      // Call this here because down the page 'this' changes context and cannot be used easily
      let thisContext = this;

      // Update Message and proceed
      $('#emailSignInErr').removeClass('redText blueText invisible').addClass('greenText visible').text('Successfully Authenticated!');

      // If checkbox is checked on email page, save data, else erase config
      if (emailChecked == true) {
        try {
          // Save config file if emailChecked == true
          let emailLoginStr = `USERNAME: "${emailUsername}"\nPASSWORD: "${emailPW}"\nCHECKED: "${emailChecked}"\n`;
          fs.writeFileSync(path.join(__dirname + '/../Configs/EmailConfig.txt'), emailLoginStr);
        } catch (err) {/* do nothing */ }
      } else {
        try {
          fs.unlinkSync(path.join(__dirname + '/../Configs/EmailConfig.txt'));
        } catch (err) {/* do nothing */ }
      }

      // get contents of folder and delete all files before beginning to obtain a clean slate
      try {
        let mailFolderContents = fs.readdirSync(path.join(__dirname + '/../Mail'));
        for (let mailFolderItem of mailFolderContents) {
          fs.unlinkSync(path.join(__dirname + '/../Mail/' + mailFolderItem));
        }
      } catch (err) {
        if (String(err).indexOf('no such file or directory') == -1) {
          dialog.showErrorBox('Error', String(err));
        } else {
          return false;
        }
      }

      // Proceed to create txt files of email contents for today's date
      let dateString = thisContext.getDateString();
      let imapErrThrown = false;

      // Define Imap object that will be used for fetching mail
      var imap = new Imap({
        user: emailUsername,
        password: emailPW,
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
        connTimeout: 5000,
        authTimeout: 5000,
        socketTimeout: 5000
      });


      function openInbox(cb) {
        imap.openBox('INBOX', true, cb);
      }
      imap.once('ready', function () {
        openInbox(function (err, box) {
          if (err) throw err;
          imap.search(['ALL', ['SINCE', `${dateString}`]], function (err, results) {
            if (err) throw err;
            //var f = imap.fetch(results, { bodies: '' });
            var f = imap.fetch(results, { bodies: 'HEADER.FIELDS (FROM TO SUBJECT DATE)', struct: true });

            f.on('message', function (msg, seqno) {
              //console.log('Message #%d', seqno);
              var prefix = '(#' + seqno + ') ';
              msg.on('body', function (stream, info) {
                //console.log(prefix + 'Body');
                stream.pipe(fs.createWriteStream(path.join(__dirname + '/../Mail/Message#' + seqno + '.txt')));
              });
              msg.once('attributes', function (attrs) {
                //console.log(prefix + 'Attributes: %s', inspect(attrs, false, 8));
              });
              msg.once('end', function () {
                //console.log(prefix + 'Finished');
              });
            });
            f.once('error', function (err) {
              // Doesn't fire on invalid credentials
            });
            f.once('end', function () {
              imap.end();
            });
          });
        });
      });
      imap.once('error', function (err) {
        try {
          let errMSG = String(err);
          if (errMSG.indexOf('LOGIN failed') != -1) {
            $('#emailSignInErr').removeClass('blueText greenText invisible').addClass('redText visible').text('Failed to Authenticate');
            imapErrThrown = true;
          } else {
            $('#emailSignInErr').removeClass('blueText greenText invisible').addClass('redText visible').text(`${errMSG}`);
            imapErrThrown = true;
          }
        } catch (err) { dialog.showErrorBox('Error', String(err)); }

      });
      imap.once('end', function () {
        try {
          if (imapErrThrown == false) {
            thisContext.loadingPage(); // Displays loading page
            let mailList = fs.readdirSync(path.join(__dirname + '/../Mail'));
            let p1String = '';
            let p2String = '';
            let p3String = '';
            let deptDownString = '';
            let dailyDivotsText = '';
            let dailyDivotsHTML = '';
            let noteworthyText = '';
            let noteworthyHTML = '';
            let deptDownArray = ['Course Down', 'Department Down', 'Credit Card Down', 'CC Down', 'Dept Down', 'Server Down', 'No Power', 'Outage'];
            for (let mailItem in mailList) {
              let mailSubject = fs.readFileSync(path.join(__dirname + `/../Mail/${mailList[mailItem]}`), 'utf-8');
              mailSubject = DOMPurify.sanitize(mailSubject);
              // Ignore emails generated by slack
              if (mailSubject.indexOf('From: "Slack" <no-reply@slack.com>') == -1) {
                if (mailSubject.indexOf('Subject:') != -1) {
                  let subjectIndex = mailSubject.indexOf('Subject:');
                  let nextNewLineIndex = mailSubject.indexOf('\n', subjectIndex);
                  let extractedSubject = mailSubject.slice(subjectIndex, nextNewLineIndex);
                  let subject = (extractedSubject.replace('Subject:', '')).trim();

                  // Look for P incidents and append them to their proper strings
                  if (subject.indexOf('P1') != -1 && p1String.indexOf(subject) == -1 || subject.indexOf('p1') != -1 && p1String.String.indexOf(subject) != -1) { p1String = p1String.concat(`• ${subject}\n`); }
                  if (subject.indexOf('P2') != -1 && p2String.indexOf(subject) == -1 || subject.indexOf('p2') != -1 && p2String.String.indexOf(subject) != -1) { p2String = p2String.concat(`• ${subject}\n`); }
                  if (subject.indexOf('P3') != -1 && p3String.indexOf(subject) == -1 || subject.indexOf('p3') != -1 && p3String.String.indexOf(subject) != -1) { p3String = p3String.concat(`• ${subject}\n`); }

                  /* Make sure deptdown messages don't already appear as a p1,p2 or p3 message */
                  for (let msg in deptDownArray) {
                    if (subject.indexOf(deptDownArray[msg]) != -1) { deptDownString = deptDownString.concat(`• ${subject}\n`); }
                  }
                }
              }
            }

            // Check for null, undefined or blank values for daily divot string
            if (p1String == '' && p2String == '' && p3String == '') {
              dailyDivotsText = `\nDaily Divots:\n- None to report.\n\n`;
              dailyDivotsHTML = `<br>Daily Divots:<br>- None to report.<br><br>`;
            } else if (p1String != '' || p2String != '' || p3String != '') {
              // Something exists, write dailyDivots using variables
              dailyDivotsText = `\nDaily Divots:\n${p1String}${p2String}${p3String}\n`;
              dailyDivotsHTML = `<br>Daily Divots:<br><b>${p1String}${p2String}${p3String}</b><br>`;
            }

            // Check for null, undefined or blank values for noteworthy cases string
            if (deptDownString == '') {
              noteworthyText = `Noteworthy Cases:\n- None to report.\n\n`;
              noteworthyHTML = `Noteworthy Cases:<br>- None to report.<br><br>`;
            } else {
              // Something exists, write noteworthy using variables
              noteworthyText = `Noteworthy Cases:\n${deptDownString}\n`;
              noteworthyHTML = `Noteworthy Cases:<br><b>${deptDownString}</b><br>`;
            }

            // Generate data using mail P subjects along with report data, then show copy text screen to copy data
            let createdByRep = salesforce.reports.getCreatedByRep(sfconn);
            let stillOpenToday = salesforce.reports.getStillOpenToday(sfconn);
            let casesClosedByRep = salesforce.reports.getCasesClosedByRep(sfconn);
            let totalCasesOpen = salesforce.reports.getTotalCasesOpen(sfconn);
            let totalCasesOpenYesterday = salesforce.reports.getTotalCasesOpenYesterday(sfconn);
            let sfReportTotals = { createdByRep: 0, stillOpenToday: 0, casesClosedByRep: 0, totalCasesOpen: 0, totalCasesOpenYesterday: 0 };

            createdByRep.then(function () {
              sfReportTotals.createdByRep = parseInt(createdByRep._55.factMap['T!T'].aggregates[2].value);
              stillOpenToday.then(function () {
                sfReportTotals.stillOpenToday = parseInt(stillOpenToday._55.factMap['T!T'].aggregates[1].value);
                casesClosedByRep.then(function () {
                  sfReportTotals.casesClosedByRep = parseInt(casesClosedByRep._55.factMap['T!T'].aggregates[0].value);
                  totalCasesOpen.then(function () {
                    totalCasesOpenYesterday.then(function () {
                      sfReportTotals.totalCasesOpen = parseInt(totalCasesOpen._55.factMap['T!T'].aggregates[0].value);
                      sfReportTotals.totalCasesOpenYesterday = parseInt(totalCasesOpenYesterday._55.factMap['T!T'].aggregates[0].value);

                      let C = sfReportTotals.createdByRep;
                      let D = sfReportTotals.stillOpenToday;
                      let E = C - D;
                      let F = (Math.round((E / C) * 100)).toFixed(0);
                      let G = sfReportTotals.casesClosedByRep;
                      let H = sfReportTotals.totalCasesOpen;
                      // sfReportTotals Object has been filled with values, now create reports using those values and save them
                      let atTheTurnHTML = `Team,<br><br>At the turn today we’ve closed <b>${sfReportTotals.createdByRep - sfReportTotals.stillOpenToday}</b> of <b>${sfReportTotals.createdByRep}</b> cases - <b>${Math.round(((sfReportTotals.createdByRep - sfReportTotals.stillOpenToday) / sfReportTotals.createdByRep) * 100)}%</b> SDC<br><br>Abandoned Rate – <b>[Daily Calls Column H] %</b> ( <b>[Daily Calls Column E]</b> Abandoned / <b>[Daily Calls Column F]</b> total calls)<br>${dailyDivotsHTML}`;
                      let atTheTurnText = `Team,\n\nAt the turn today we’ve closed ${sfReportTotals.createdByRep - sfReportTotals.stillOpenToday} of ${sfReportTotals.createdByRep} cases - ${Math.round(((sfReportTotals.createdByRep - sfReportTotals.stillOpenToday) / sfReportTotals.createdByRep) * 100)}% SDC\n\nAbandoned Rate – [Daily Calls Column H] % ( [Daily Calls Column E] Abandoned / [Daily Calls Column F] total calls)\n${dailyDivotsText}`;
                      let inTheClubhouseHTML = `Team,<br><br>Daily Case Metrics: <b>[Daily Cases Column E]</b> / <b>[Daily Cases Column C]</b> cases closed today (<b>[Daily Cases Column F]%</b>), <b>[Daily Cases Column G]</b> total cases closed.<br><br><b>[Daily Cases Column H]</b> cases remain open for a net change of <b>[Daily Cases Column I]</b>.<br><br>Abandoned Rate – <b>[Daily Calls Column H] %</b> ( <b>[Daily Calls Column E]</b> Abandoned / <b>[Daily Calls Column F]</b> total calls)<br>${dailyDivotsHTML}${noteworthyHTML}`;
                      let inTheClubhouseText = `Team,\n\nDaily Case Metrics: [Daily Cases Column E] / [Daily Cases Column C] cases closed today ([Daily Cases Column F]%), [Daily Cases Column G] total cases closed.\n\n[Daily Cases Column H] cases remain open for a net change of [Daily Cases Column I].\n\nAbandoned Rate – [Daily Calls Column H] % ( [Daily Calls Column E] Abandoned / [Daily Calls Column F] total calls)\n${dailyDivotsText}${noteworthyText}`;
                      // Assign in the clubhouse values to the string passed to the parseEmailAttachments function
                      inTheClubhouseText = inTheClubhouseText.replace('[Daily Cases Column E]', E)
                      inTheClubhouseText = inTheClubhouseText.replace('[Daily Cases Column C]', C);
                      inTheClubhouseText = inTheClubhouseText.replace('[Daily Cases Column F]', F);
                      inTheClubhouseText = inTheClubhouseText.replace('[Daily Cases Column G]', G);
                      inTheClubhouseText = inTheClubhouseText.replace('[Daily Cases Column H]', H);

                      inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Cases Column E]', E)
                      inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Cases Column C]', C);
                      inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Cases Column F]', F);
                      inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Cases Column G]', G);
                      inTheClubhouseHTML = inTheClubhouseHTML.replace('[Daily Cases Column H]', H);

                      // Fetch attachments and parse them for ICStats data
                      thisContext.parseEmailAttachments(emailUsername, emailPW, atTheTurnText, atTheTurnHTML, inTheClubhouseText, inTheClubhouseHTML, sfReportTotals, H);
                    });


                  });
                });
              });
            });
          }
        } catch (err) { dialog.showErrorBox('Error', String(err)); }
      });
      imap.connect();

    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  },
  returnFileContents: function (pathString) {
    try {
      let fileContents = fs.readFileSync(path.join(__dirname + '/..' + pathString), 'utf-8');
      return fileContents;
    } catch (err) { dialog.showErrorBox('Error', String(err)); }
  }
}