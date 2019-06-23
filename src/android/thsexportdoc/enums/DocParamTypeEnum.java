package cn.com.ths.exportdoc.thsexportdoc.enums;

public enum DocParamTypeEnum {

    TITLE("TITLE", "标题"),
    REPLACECLASS("REPLACECLASS","替换回调CLASS");

    private String code;
    private String desc;

    DocParamTypeEnum(String code, String desc){
        this.code = code;
        this.desc = desc;
    }

    /**
     * @return the code
     */
    public String getCode() {
        return code;
    }
    /**
     * @param code the code to set
     */
    public void setCode(String code) {
        this.code = code;
    }
    /**
     * @return the desc
     */
    public String getDesc() {
        return desc;
    }
    /**
     * @param desc the desc to set
     */
    public void setDesc(String desc) {
        this.desc = desc;
    }

}
