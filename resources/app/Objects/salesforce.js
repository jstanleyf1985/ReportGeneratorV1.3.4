const helpers = require('./helper');
exports.sfConfig = {
  // Report IDs
  casesCreatedByRep: '00O70000004y85sEAA',
  casesStillOpenToday: '00O70000004y66KEAQ',
  casesClosedByRep: '00O70000004y65gEAA',
  totalCasesOpen: '00O0g000005PBYhEAO',
  totalCasesOpenYesterday: '00O0g000005Uu4hEAC'
}

exports.reports = {
  getCreatedByRep: function (sfconn) {
    try {
      let reportId = '00O70000004y85sEAA';
      let report = sfconn.analytics.report(reportId);
      let createdByRep = report.execute({ details: true }, function (err, result) {
        if (err) { dialog.showErrorBox('Error', String(err)); }
      });
      return createdByRep;
    } catch (err) {
      dialog.showErrorBox('Error', String(err));
    }
  },
  getStillOpenToday: function (sfconn) {
    try {
      let reportId = '00O70000004y66KEAQ';
      let report = sfconn.analytics.report(reportId);
      let stillOpenToday = report.execute({ details: true }, function (err, result) {
        if (err) { dialog.showErrorBox('Error', String(err)); }
      });
      return stillOpenToday;
    } catch (err) {
      dialog.showErrorBox('Error', String(err));
    }
  },
  getCasesClosedByRep: function (sfconn) {
    try {
      let reportId = '00O70000004y65gEAA';
      let report = sfconn.analytics.report(reportId);
      let casesClosedByRep = report.execute({ details: true }, function (err, result) {
        if (err) { dialog.showErrorBox('Error', String(err)); }
      });
      return casesClosedByRep;
    } catch (err) {
      dialog.showErrorBox('Error', String(err));
    }
  },
  getTotalCasesOpen: function (sfconn) {
    try {
      let reportId = '00O0g000005PBYhEAO';
      let report = sfconn.analytics.report(reportId);
      let totalCasesOpen = report.execute({ details: true }, function (err, result) {
        if (err) { dialog.showErrorBox('Error', String(err)); }
      });
      return totalCasesOpen;
    } catch (err) {
      dialog.showErrorBox('Error', String(err));
    }
  },
  getTotalCasesOpenYesterday: function (sfconn) {
    try {
      let reportId = '00O0g000005UxlPEAS';
      let todayDate = new Date();
      let todayParsed = Date.parse(todayDate);
      todayParsed = parseInt(todayParsed);
      todayParsed = todayParsed - 86400000; // - 29 hours accounting for UTC
      todayParsed = new Date(todayParsed);
      todayParsed = todayParsed.toISOString();
      todayParsed = todayParsed.substr(0, todayParsed.indexOf('T'));
      todayParsed = todayParsed.concat(' 00:00:00');
      let metadata = {
        reportMetadata: {
          standardDateFilter: {
            column: 'CREATED_DATEONLY',
            durationValue: 'CUSTOM',
            endDate: todayParsed,
            startDate: '1970-01-01 00:00:00'
          }
        }
      }
      let report = sfconn.analytics.report(reportId);
      let totalCasesOpenYesterday = report.execute({ details: true, metadata: metadata }, function (err, result) {
        if (err) { dialog.showErrorBox('Error', String(err)); }
      });
      return totalCasesOpenYesterday;
    } catch (err) {
      dialog.showErrorBox('Error', String(err));
    }
  },
}