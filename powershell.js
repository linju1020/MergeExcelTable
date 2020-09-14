function PowerShell() {
  this.BrowseForFolder = async function (Title) {
    var promise = new Promise(function (resolve, reject) {
      // ... some code

      var psScript = `(new-object -COM 'Shell.Application').BrowseForFolder(0,'${Title}',529,0).self.path`;

      var spawn = require('child_process').spawn;
      var child = spawn('powershell', [psScript]);
      child.stdout.on('data', function (data) {
        //console.log('Powershell Data: ' + data);
        if (data.length > 0) {
          resolve(data.toString().replace('\r', '').replace('\n', ''));
        }
      });
      child.stderr.on('data', function (data) {
        //this script block will get the output of the PS script
        console.log('Powershell Errors: ' + data);
        reject(null);
      });
      child.on('exit', function () {
        //console.log('Powershell Script finished');
      });
      child.stdin.end(); //end
    });
    return promise;
  };
}

//new PowerShell().BrowseForFolder();
module.exports = PowerShell;