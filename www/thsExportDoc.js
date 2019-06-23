var exec = require('cordova/exec');

exports.exportDoc = function (arg0, success, error) {
    exec(success, error, 'thsExportDoc', 'exportDoc', [arg0]);
};
