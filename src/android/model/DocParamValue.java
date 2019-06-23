package cn.com.ths.exportdoc.thsexportdoc.model;

import java.io.Serializable;

import cn.com.ths.exportdoc.thsexportdoc.enums.DocParamTypeEnum;

public class DocParamValue implements Serializable {

    /**
     *
     */
    private static final long serialVersionUID = 1L;
    //参数类型
    private DocParamTypeEnum type;
    //类型名称
    private String className;
    //标题级别
    private Integer titleLevel;
    //参数值
    private Object value;

    private DocParamValue() {

    }
    /**
     * 回调类的构造函数
     * @param type
     * @param className
     * @param value
     */
    public DocParamValue(DocParamTypeEnum type, String className, Object value) {
        this.type = type;
        this.className = className;
        this.value = value;
    }
    /**
     * 标题的构造函数
     * @param type
     * @param titleLevel
     * @param value
     */
    public DocParamValue(DocParamTypeEnum type, Integer titleLevel, Object value) {
        this.type = type;
        this.titleLevel = titleLevel;
        this.value = value;
    }

    public DocParamTypeEnum getType() {
        return type;
    }
    public void setType(DocParamTypeEnum type) {
        this.type = type;
    }
    public String getClassName() {
        return className;
    }
    public void setClassName(String className) {
        this.className = className;
    }
    public Integer getTitleLevel() {
        return titleLevel;
    }
    public void setTitleLevel(Integer titleLevel) {
        this.titleLevel = titleLevel;
    }
    public Object getValue() {
        return value;
    }
    public void setValue(Object value) {
        this.value = value;
    }



}
