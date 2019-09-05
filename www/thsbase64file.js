var exec = require('cordova/exec');

module.exports = {
    saveBase64ToFile: function (base64Str,path,successCallback, errorCallback) {
        cordova.exec(
            successCallback,
            errorCallback,
            "thsbase64file", "saveBase64ToFile", [base64Str,path]);
    }
};
