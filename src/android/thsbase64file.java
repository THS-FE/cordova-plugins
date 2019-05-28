package thsbase64file;

import org.apache.cordova.CordovaPlugin;
import org.apache.cordova.CallbackContext;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import android.content.Intent;
import android.provider.Settings;
/**
 * liuyx
 */
public class thsbase64file extends CordovaPlugin {
    
    @Override
    public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {
        if (action.equals("saveBase64ToFile")) {
            this.saveBase64ToFile(callbackContext);
            return true;
        }
        return false;
    }
    //展示系统APN设置
    private void saveBase64ToFile(CallbackContext callbackContext) {
        try {
            // Intent intent = new Intent(Settings.ACTION_APN_SETTINGS);
            // cordova.getActivity().startActivity(intent);
            callbackContext.success("成功");
        }catch (Exception exception){
            callbackContext.error("发生异常了");
        }
    }
}
