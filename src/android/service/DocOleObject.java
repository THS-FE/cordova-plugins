package cn.com.ths.exportdoc.thsexportdoc.service;

public class DocOleObject {
    //文件路径
    private String filePath;

    public DocOleObject(String filePath) {
        this.filePath = filePath;
    }

    /**
     * @return the filePath
     */
    public String getFilePath() {
        return filePath;
    }

    /**
     * @param filePath the filePath to set
     */
    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
}
