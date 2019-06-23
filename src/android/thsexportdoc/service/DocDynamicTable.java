package cn.com.ths.exportdoc.thsexportdoc.service;

public class DocDynamicTable {


    /**
     *
     */
    private static final long serialVersionUID = -4393573016834228674L;

    //动态单表
    private InnerDocTable docTable=new InnerDocTable();
    //动态多表
    private InnerDocMultiTable multiTable=new InnerDocMultiTable();

    public InnerDocTable getDocTable() {
        return docTable;
    }

    public void setDocTable(InnerDocTable docTable) {
        this.docTable = docTable;
    }

    public InnerDocMultiTable getMultiTable() {
        return multiTable;
    }

    public void setMultiTable(InnerDocMultiTable multiTable) {
        this.multiTable = multiTable;
    }

    /**
     *
     * @author 兼容历史项目
     * @date   2018年7月20日
     */
    public class InnerDocTable extends DocTable{

    }

    /**
     *
     * @author 兼容历史项目
     * @date   2018年7月20日
     */
    public class  InnerDocMultiTable extends DocMultiTable{

    }

}
