var exec = require('cordova/exec');

module.exports = {
    showApns: function (successCallback, errorCallback) {
        cordova.exec(
            successCallback,
            errorCallback,
            "thsbase64file", "saveBase64ToFile", []);
    }
};
