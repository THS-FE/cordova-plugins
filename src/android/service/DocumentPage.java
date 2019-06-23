package cn.com.ths.exportdoc.thsexportdoc.service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocumentPage {


    //文档体参数
    private Map<String,Object> paramMap=new HashMap<String,Object>();
    //是否更新目录
    private boolean updateCatalog=false;
    //文档体图片
    private Map<String,List<DocImage>> images=new HashMap<String,List<DocImage>>();
    //文档体文件对象
    private Map<String,List<DocOleObject>> oleObjects = new HashMap<String,List<DocOleObject>>();
    //需要进行上下标转换的特殊字符集
    private List<DocSuperAndSub> superAndSubs=new ArrayList<DocSuperAndSub>();
    //是否在线预览
    private boolean onlineBrowse=false;

    public Map<String, Object> getParamMap() {
        return paramMap;
    }
    public void setParamMap(Map<String, Object> paramMap) {
        this.paramMap = paramMap;
    }
    public Map<String, List<DocImage>> getImages() {
        return images;
    }
    public void setImages(Map<String, List<DocImage>> images) {
        this.images = images;
    }
    public boolean isUpdateCatalog() {
        return updateCatalog;
    }
    public void setUpdateCatalog(boolean updateCatalog) {
        this.updateCatalog = updateCatalog;
    }
    public boolean isOnlineBrowse() {
        return onlineBrowse;
    }
    public void setOnlineBrowse(boolean onlineBrowse) {
        this.onlineBrowse = onlineBrowse;
    }
    public List<DocSuperAndSub> getSuperAndSubs() {
        return superAndSubs;
    }
    public void setSuperAndSubs(List<DocSuperAndSub> superAndSubs) {
        this.superAndSubs = superAndSubs;
    }
    /**
     * @return the oleObjects
     */
    public Map<String, List<DocOleObject>> getOleObjects() {
        return oleObjects;
    }
    /**
     * @param oleObjects the oleObjects to set
     */
    public void setOleObjects(Map<String, List<DocOleObject>> oleObjects) {
        this.oleObjects = oleObjects;
    }
}
