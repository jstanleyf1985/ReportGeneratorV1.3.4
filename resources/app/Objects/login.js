exports.credentials = {
  getLogin: function() {
    // Check if file exists, if not create one, if it exists parse it
    try {
      var returnVal = ['USERNAME', 'PASSWORD', false];
      let readLoginFileInterface = readline.createInterface({
        input: fs.createReadStream(path.join(__dirname + '/../Configs/LoginConfig.txt'))
      });

      readLoginFileInterface.on('line', function(line) {
        line = line.trim();
        line = DOMPurify.sanitize(line);

        if(line.indexOf('USERNAME: "') != -1) {
          line = line.replace('USERNAME: "','');
          line = line.replace('"', '');
          returnVal[0] = line;
        }

        if(line.indexOf('PASSWORD: "') != -1) {
          line = line.replace('PASSWORD: "', '');
          line = line.replace('"', '');
          returnVal[1] = line;
        }
      }).on('close', function() {
        if(returnVal[0] != undefined && returnVal[0] != null && returnVal[0] != '' && returnVal[1] != undefined && returnVal[1] != null && returnVal[1] != '') {
          // Valid
          returnVal[2] = true;
        }
      });
    } catch(err) {
      dialog.showErrorBox('Error', String(err));
    }
  }
}