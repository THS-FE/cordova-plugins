package cn.com.ths.exportdoc.thsexportdoc.service;
/**
 * 封装上下标内容的类
 * @author wangjp
 *
 */
public class DocSuperAndSub {


    private String allWord;//所有的内容，包含左侧的单词和右侧的上下标
    private String baseWord;//左侧的单词
    private String superscript;//右侧的上标
    private String subscript;//右侧的下标

    public String getAllWord() {
        return allWord;
    }
    public void setAllWord(String allWord) {
        this.allWord = allWord;
    }
    public String getBaseWord() {
        return baseWord;
    }
    public void setBaseWord(String baseWord) {
        this.baseWord = baseWord;
    }
    public String getSuperscript() {
        return superscript;
    }
    public void setSuperscript(String superscript) {
        this.superscript = superscript;
    }
    public String getSubscript() {
        return subscript;
    }
    public void setSubscript(String subscript) {
        this.subscript = subscript;
    }



}
