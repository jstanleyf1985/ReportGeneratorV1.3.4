const forms = require('./forms');
const helpers = require('./helper');

exports.loadSFLoginPage = function () {
  try {
    $('#main').empty();
    $('#main').prepend(forms.salesforceform);

    // Check if sales force config file exists
    let configExists = helpers.functions.checkIfFileExists('/Configs/LoginConfig.txt');
    // If exists read it and pre-load data into login field
    if (configExists == true) {
      let configContents = helpers.functions.returnFileContents('/Configs/LoginConfig.txt');

      // Read line contents and parse, store login credentials in loginCreds object
      let loginCreds = { USERNAME: '', PASSWORD: '', CHECKED: '' };
      let readLoginFileInterface = readline.createInterface({
        input: fs.createReadStream(path.join(__dirname + '/../Configs/LoginConfig.txt'))
      });

      readLoginFileInterface.on('line', function (line) {
        line = line.trim();
        line = DOMPurify.sanitize(line);

        if (line.indexOf('USERNAME: "') != -1) {
          line = line.replace('USERNAME: "', '');
          line = line.replace('"', '');
          loginCreds.USERNAME = line;
        }

        if (line.indexOf('PASSWORD: "') != -1) {
          line = line.replace('PASSWORD: "', '');
          line = line.replace('"', '');
          loginCreds.PASSWORD = line;
        }

        if (line.indexOf('CHECKED: "', '') != -1) {
          line = line.replace('CHECKED: "', '');
          line = line.replace('"', '');
          loginCreds.CHECKED = line;
        }
      }).on('close', function () {
        // Check if login credentials are valid
        loginValid = helpers.functions.checkValid(loginCreds.USERNAME, loginCreds.PASSWORD);

        // If valid, autofill contents, else delete the configuration file and proceed without one
        if (loginValid) {
          $('#sfusername').val(`${loginCreds.USERNAME}`);
          $('#sfpassword').val(`${loginCreds.PASSWORD}`);

          if (loginCreds.CHECKED == 'true') {
            $('#sfChecked').prop('checked', true);
          }

          // Update message saying loaded from config
          $('#sfSignInErr').removeClass('ininvisible redText greenText').addClass('blueText invisible').text(`Loaded from config file successfully`);
        } else {
          // Invalid data in login file, delete file and refresh login form
          let invalidConfigExists = helpers.functions.checkIfFileExists('/Configs/LoginConfig.txt');
          if (invalidConfigExists) {
            try {
              fs.unlinkSync(path.join(__dirname + '/../Configs/LoginConfig.txt'));
            } catch (err) {/* dialog.showErrorBox('Error', String(err)); */ }

            $('#main').empty();
            $('#main').prepend(forms.salesforceform);
          }
        }
      });
    }

  } catch (err) { dialog.showErrorBox('Error', String(err)); }
}
exports.loadEmailLoginPage = function (sfconnection) {
  $('#main').empty();
  $('#main').prepend(forms.emailform);
  // Prefill data if a config file exists
  let loginObj = helpers.functions.getEmailFromConfig();

  // Update Message on front end for loaded by config
  if (loginObj.username != '' && loginObj.password != '') {
    $('#emailSignInErr').removeClass('redText greenText ininvisible').addClass('blueText invisible').text('Successfully loaded from Config!');
  }

  /* Prefill in credentials using loginObj, if marked checked, save data to config file */
  $('#emailusername').val(`${loginObj.username}`);
  $('#emailpassword').val(`${loginObj.password}`);
  if (loginObj.checked == 'true') {
    $('#emailChecked').prop('checked', true);
  }

  // Generate report on click, an extra step is required for in the clubhouse reports
  $('#emailsignin').on('click', function (e) {
    e.preventDefault();

    // Get changed(or not) username and password
    let emailUser = $('#emailusername').val();
    let emailPass = $('#emailpassword').val();
    let emailChecked = $('#emailChecked').prop('checked');


    // Trim and sanitize before sending to parseEmailToday
    if (emailUser == null || emailUser == undefined) { emailUser = '' }
    if (emailPass == null || emailPass == undefined) { emailPass = '' }

    emailUser = DOMPurify.sanitize(emailUser);
    emailPass = DOMPurify.sanitize(emailPass);

    emailUser = emailUser.trim();
    emailPass = emailPass.trim();

    helpers.functions.parseEmailToday(emailUser, emailPass, emailChecked, sfconnection);
  });
}
exports.loadFinalPage = function (atTheTurnString, atTheTurnHTML, inTheClubhouseString, inTheClubhouseHTML, dailyCasesString, dailyCallsString) {
  $('#main').empty();
  $('#main').prepend(forms.final);

  function copyToClip(str) {
    function listener(e) {
      e.clipboardData.setData("text/html", str);
      e.clipboardData.setData("text/plain", str);
      e.preventDefault();
    }
    document.addEventListener("copy", listener);
    document.execCommand("copy");
    document.removeEventListener("copy", listener);
  };

  // CopyToClipboard Function that handles copying text/html
  $('#getTurnReport').on('click', function (e) {
    e.preventDefault();
    copyToClip(atTheTurnHTML);

    // Update message for user
    $('#finalInfo').empty().prepend('Last Copied: <b>At the Turn(formatted)<b>').removeClass('invisible');
  });

  $('#getTurnReportText').on('click', function (e) {
    e.preventDefault();
    copyToClip(atTheTurnString);

    // Update message for user
    $('#finalInfo').empty().prepend('Last Copied: <b>At the Turn(text)<b>').removeClass('invisible');
  });

  $('#getClubhouseReport').on('click', function (e) {
    e.preventDefault();
    copyToClip(inTheClubhouseHTML);

    // Update message for user
    $('#finalInfo').empty().prepend('Last Copied: <b>In the Clubhouse(formatted)<b>').removeClass('invisible');
  });

  $('#getClubhouseReportText').on('click', function (e) {
    e.preventDefault();
    copyToClip(inTheClubhouseString);

    // Update message for user
    $('#finalInfo').empty().prepend('Last Copied: <b>In the Clubhouse(text)<b>').removeClass('invisible');
  });

  $('#getDailyCasesReport').on('click', function (e) {
    e.preventDefault();
    helpers.functions.copyDailyCases(dailyCasesString);

    // Update message for user
    $('#finalInfo').empty().prepend('Last Copied: <b>Daily Cases<b>').removeClass('invisible');
  });

  $('#getDailyCallsReport').on('click', function (e) {
    e.preventDefault();
    helpers.functions.copyDailyCalls(dailyCallsString);

    // Update message for user
    $('#finalInfo').empty().prepend('Last Copied: <b>Daily Calls<b>').removeClass('invisible');
  });

  $('#copyToggle').on('click', function (e) {
    e.preventDefault();

    if ($('#copyToggle').hasClass('formattedCopy')) {
      $('#getTurnReport, #getClubhouseReport').addClass('hidden');
      $('#getTurnReportText, #getClubhouseReportText').removeClass('hidden');
      $('#copyToggle').text('Currently In: [Plain Text Copy Mode]').removeClass('formattedCopy').addClass('plainTextCopy');
    } else {
      $('#getTurnReport, #getClubhouseReport').removeClass('hidden');
      $('#getTurnReportText, #getClubhouseReportText').addClass('hidden');
      $('#copyToggle').text('Currently In: [Formatted Copy Mode]').addClass('formattedCopy').removeClass('plainTextCopy');
    }

  });
}