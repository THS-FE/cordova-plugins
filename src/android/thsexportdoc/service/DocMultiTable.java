package cn.com.ths.exportdoc.thsexportdoc.service;

import java.util.ArrayList;
import java.util.List;


/**
 * 动态多表对象，针对主子模板情况下，子模版数据的封装对象
 * @author wangjp
 * @date   2018年7月20日
 */
public class DocMultiTable {


    //子模版中为单一表或者单独
    @Deprecated
    private List<DocDynamicTable.InnerDocTable> multiTable=new ArrayList<DocDynamicTable.InnerDocTable>();

    private List<DocComplexTable> multiComplexTable = new ArrayList<DocComplexTable>();
    private String modelPath;

    @Deprecated
    public DocDynamicTable.InnerDocTable newDocTable(){
        DocDynamicTable.InnerDocTable newDocTable=new DocDynamicTable().new InnerDocTable();
        multiTable.add(newDocTable);
        return newDocTable;
    }
    @Deprecated
    public void add(DocDynamicTable.InnerDocTable docTable){
        multiTable.add(docTable);
    }
    @Deprecated
    public List<DocDynamicTable.InnerDocTable> getMultiTable() {
        return multiTable;
    }
    @Deprecated
    public void setMultiTable(List<DocDynamicTable.InnerDocTable> multiTable) {
        this.multiTable = multiTable;
    }

    public String getModelPath() {
        return modelPath;
    }

    public void setModelPath(String modelPath) {
        this.modelPath = modelPath;
    }

    public DocComplexTable newDocComplexTable() {
        DocComplexTable complexTable = new DocComplexTable();
        multiComplexTable.add(complexTable);
        return complexTable;
    }

    public void add(DocComplexTable complexTable){

        multiComplexTable.add(complexTable);
    }

    public List<DocComplexTable> getMultiComplexTable() {
        return multiComplexTable;
    }

    public void setMultiComplexTable(List<DocComplexTable> multiComplexTable) {
        this.multiComplexTable = multiComplexTable;
    }



}
