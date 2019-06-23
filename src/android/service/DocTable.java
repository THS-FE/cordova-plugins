package cn.com.ths.exportdoc.thsexportdoc.service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocTable {


    //数据行的开始行
    private int[] dataRowStarts;
    //需要合并的列
    private int[] mergeColumns;
    //表格体参数
    private Map<String,Object> paramMap=new HashMap<String,Object>();
    //表格体图片
    private Map<String,List<DocImage>> images=new HashMap<String,List<DocImage>>();
    //动态行数据
    private List<Map<String,Object>> rowList = new ArrayList<Map<String,Object>>();

    /**
     * 追加一行数据
     * @param row
     */
    public void addRow(Map<String,Object> row){
        rowList.add(row);
    }
    /**
     * 追加多行数据
     * @param rows
     */
    public void addRows(List<Map<String,Object>> rows){
        rowList.addAll(rows);
    }

    public int[] getDataRowStarts() {
        return dataRowStarts;
    }

    public void setDataRowStarts(int[] dataRowStarts) {
        this.dataRowStarts = dataRowStarts;
    }

    public List<Map<String, Object>> getRowList() {
        return rowList;
    }

    public void setRowList(List<Map<String, Object>> rowList) {
        this.rowList = rowList;
    }
    public Map<String, Object> getParamMap() {
        return paramMap;
    }
    public void setParamMap(Map<String, Object> paramMap) {
        this.paramMap = paramMap;
    }
    public int[] getMergeColumns() {
        return mergeColumns;
    }
    public void setMergeColumns(int[] mergeColumns) {
        this.mergeColumns = mergeColumns;
    }
    public Map<String, List<DocImage>> getImages() {
        return images;
    }
    public void setImages(Map<String, List<DocImage>> images) {
        this.images = images;
    }


}
