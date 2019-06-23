package cn.com.ths.exportdoc.thsexportdoc;

import android.Manifest;
import android.content.pm.PackageManager;

import org.apache.cordova.CordovaInterface;
import org.apache.cordova.CordovaPlugin;
import org.apache.cordova.CallbackContext;

import org.apache.cordova.CordovaWebView;
import org.apache.cordova.PermissionHelper;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import  cn.com.ths.exportdoc.thsexportdoc.utils.ExportWordUtils;

/**
 * 导出word
 */
public class thsExportDoc extends CordovaPlugin {
    /**
     * 权限列表
     */
    private String[] locPerArr = new String[] {
            Manifest.permission.WRITE_EXTERNAL_STORAGE //存储权限
    };
    private CallbackContext callbackContext; //插件调用回调
    private String docData; //生成word的数据，从ts/js传入

    @Override
    public void onRequestPermissionResult(int requestCode,
                                          String[] permissions, int[] grantResults) throws JSONException {
        // TODO Auto-generated method stub
        for (int r : grantResults) {
            if (r == PackageManager.PERMISSION_DENIED) {
                return;
            }
        }
        //授权完成后执行文档生成
        exportWord();
    }
    /**
     * 检查权限并申请
     */
    private void promtForExport() {
        for (int i = 0, len = locPerArr.length; i < len; i++) {
            if (!PermissionHelper.hasPermission(this, locPerArr[i])) {
                PermissionHelper.requestPermission(this, i, locPerArr[i]);
                return;
            }
        }
        exportWord();
    }
    private void exportWord(){
        if(callbackContext!=null&&docData!=null){
            try{
                ExportWordUtils.getInstance().exportWord(docData,cordova.getActivity(),callbackContext);
            }catch (Exception e){
                callbackContext.error(e.getMessage());
            }

        }
    }
    @Override
    public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {
        if(this.callbackContext==null){
            this.callbackContext=callbackContext;
        }
        if (action.equals("exportDoc")) {
            docData= args.getString(0);
            promtForExport();
            return true;
        }
        return false;
    }

}
