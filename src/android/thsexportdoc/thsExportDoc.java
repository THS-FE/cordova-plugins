package cn.com.ths.exportdoc.thsexportdoc;

import org.apache.cordova.CordovaPlugin;
import org.apache.cordova.CallbackContext;

import org.json.JSONArray;
import org.json.JSONException;

import  cn.com.ths.exportdoc.thsexportdoc.utils.ExportWordUtils;

/**
 * 导出word
 */
public class thsExportDoc extends CordovaPlugin {

    @Override
    public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {
        if (action.equals("exportDoc")) {
            String message = args.getString(0);
            // ExportWordUtils.getInstance().exportWord("",cordova.getActivity());
            try{
                ExportWordUtils.getInstance().exportWord("",cordova.getActivity(),callbackContext);
                //this.exportDoc(message, callbackContext);
                callbackContext.success(message);
            }catch (Exception e){
                callbackContext.error(e.getMessage());
            }

            return true;
        }
        return false;
    }

    private void exportDoc(String message, CallbackContext callbackContext) {
        if (message != null && message.length() > 0) {
            callbackContext.success(message);
        } else {
            callbackContext.error("Expected one non-empty string argument.");
        }
    }
}
