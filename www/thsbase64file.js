var exec = require('cordova/exec');

module.exports = {
    saveBase64ToFile: function (successCallback, errorCallback) {
        cordova.exec(
            successCallback,
            errorCallback,
            "thsbase64file", "saveBase64ToFile", []);
    }
};
