package cn.com.ths.exportdoc.thsexportdoc.service;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocComplexTable {


    //表格体参数
    private Map<String,Object> paramMap=new HashMap<String,Object>();
    //表格体图片
    private Map<String,List<DocImage>> images=new HashMap<String,List<DocImage>>();
    //文件对象
    private Map<String,List<DocOleObject>> oleObjects = new HashMap<String,List<DocOleObject>>();
    //复杂表格map
    private Map<String,DocTable> innerTableMap = new HashMap<String,DocTable>();


    public Map<String, Object> getParamMap() {
        return paramMap;
    }
    public void setParamMap(Map<String, Object> paramMap) {
        this.paramMap = paramMap;
    }
    public Map<String, DocTable> getInnerTableMap() {
        return innerTableMap;
    }
    public void setInnerTableMap(Map<String, DocTable> innerTableMap) {
        this.innerTableMap = innerTableMap;
    }
    public Map<String, List<DocImage>> getImages() {
        return images;
    }
    public void setImages(Map<String, List<DocImage>> images) {
        this.images = images;
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
