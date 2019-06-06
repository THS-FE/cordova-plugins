package thsbase64file;

import org.apache.cordova.CordovaPlugin;
import org.apache.cordova.CallbackContext;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import android.content.Intent;
import android.os.Environment;
import android.provider.Settings;
import android.util.Log;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Base64;

/**
 * liuyx
 */
public class thsbase64file extends CordovaPlugin {
    
    @Override
    public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {
        if (action.equals("saveBase64ToFile")) {
            String base64Str = args.getString(0);
            String  path = args.getString(1);
            this.saveBase64ToFile(callbackContext,base64Str,path);
            return true;
        }
        return false;
    }
    //保存base64 到文件
    private void saveBase64ToFile(CallbackContext callbackContext,String base64Str,String path) {
        try {
            decoderBase64File(base64Str,path);
            callbackContext.success("成功");
        }catch (Exception exception){
            callbackContext.error("发生异常了");
        }
    }

    /**
     * base64 专生成文件
     * @param base64Str  base64字符
     * @param savePath 保存路径，传入后会在根SD卡根目录下创建文件test/test.png
     * @throws Exception
     */
    public  void decoderBase64File(String base64Str,String savePath) throws Exception {
        savePath=getSDPath()+"/" +savePath;
        File file = new File(savePath);
        if (!file.exists()) {
            Log.d("TestFile", "Create the file:" + savePath);
            file.getParentFile().mkdirs();
            file.createNewFile();
        }
        byte[] buffer =android.util.Base64.decode(base64Str, android.util.Base64.DEFAULT);
        FileOutputStream out = new FileOutputStream(savePath);
        out.write(buffer);
        out.close();
    }

    /**
     * 获取SD卡路径
     * @return 返回路径字符串
     */
    public String getSDPath(){
        File sdDir = null;
        boolean sdCardExist = Environment.getExternalStorageState()
                .equals(android.os.Environment.MEDIA_MOUNTED);//判断sd卡是否存在
        if(sdCardExist)
        {
            sdDir = Environment.getExternalStorageDirectory();//获取跟目录
        }
        return sdDir.toString();
    }
}
