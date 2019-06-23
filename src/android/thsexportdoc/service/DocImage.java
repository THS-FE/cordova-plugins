package cn.com.ths.exportdoc.thsexportdoc.service;

import java.io.InputStream;
import java.util.Map;

public class DocImage {

    private String path;//图片路径
    private InputStream inputStream;//图片流
    private Double width;//图片宽度
    private Double height;//图片高度
    private Double left;//图片左边距
    private Double top;//图片上边距
    //图片环绕类型,0:嵌入型;1:上下型;2:四周型;3:无环绕类型;4:紧密型;5:穿越型;6:衬于文字下方;7:衬于文字上方
    private Integer wrapType;
    private String legendContent;//图片图例内容
    private Map<String,Object> lengendFont;//图例字体样式
    public DocImage(InputStream inputStream, Double width, Double height){
        this.inputStream=inputStream;
        this.width=width;
        this.height=height;
    }

    public DocImage(InputStream inputStream, Double width, Double height, Integer wrapType){
        this.inputStream=inputStream;
        this.width=width;
        this.height=height;
        this.wrapType=wrapType;
    }

    public DocImage(String path, Double width, Double height){
        this.path=path;
        this.width=width;
        this.height=height;
    }

    public DocImage(String path, Double width, Double height, Integer wrapType){
        this.path=path;
        this.width=width;
        this.height=height;
        this.wrapType=wrapType;
    }

    public DocImage(String path, Double width, Double height, Double left, Double top){
        this.path=path;
        this.width=width;
        this.height=height;
        this.left=left;
        this.top=top;
    }

    public DocImage(String path, Double width, Double height, Double left, Double top, Integer wrapType){
        this.path=path;
        this.width=width;
        this.height=height;
        this.left=left;
        this.top=top;
        this.wrapType=wrapType;
    }

    public DocImage(InputStream inputStream, Double width, Double height, Double left, Double top){
        this.inputStream=inputStream;
        this.width=width;
        this.height=height;
        this.left=left;
        this.top=top;
    }

    public DocImage(InputStream inputStream, Double width, Double height, Double left, Double top, Integer wrapType){
        this.inputStream=inputStream;
        this.width=width;
        this.height=height;
        this.left=left;
        this.top=top;
        this.wrapType=wrapType;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public Double getWidth() {
        return width;
    }

    public void setWidth(Double width) {
        this.width = width;
    }

    public Double getHeight() {
        return height;
    }

    public void setHeight(Double height) {
        this.height = height;
    }

    public Double getLeft() {
        return left;
    }

    public void setLeft(Double left) {
        this.left = left;
    }

    public Double getTop() {
        return top;
    }

    public void setTop(Double top) {
        this.top = top;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public Integer getWrapType() {
        return wrapType;
    }

    public void setWrapType(Integer wrapType) {
        this.wrapType = wrapType;
    }


    public String getLegendContent() {
        return legendContent;
    }

    public void setLegendContent(String legendContent) {
        this.legendContent = legendContent;
    }

    public Map<String, Object> getLengendFont() {
        return lengendFont;
    }

    public void setLengendFont(Map<String, Object> lengendFont) {
        this.lengendFont = lengendFont;
    }

}
