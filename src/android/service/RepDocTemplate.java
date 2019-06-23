package cn.com.ths.exportdoc.thsexportdoc.service;

import android.content.Context;
import android.os.Environment;
import android.util.Log;

import com.aspose.words.Body;
import com.aspose.words.Cell;
import com.aspose.words.CellCollection;
import com.aspose.words.CellMerge;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.aspose.words.Font;
import com.aspose.words.FormField;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeImporter;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Row;
import com.aspose.words.RowCollection;
import com.aspose.words.Run;
import com.aspose.words.SaveFormat;
import com.aspose.words.Shape;
import com.aspose.words.Style;
import com.aspose.words.Table;
import com.aspose.words.WrapType;

import org.apache.cordova.CallbackContext;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import cn.com.ths.exportdoc.thsexportdoc.enums.DocParamTypeEnum;
import cn.com.ths.exportdoc.thsexportdoc.model.DocParamValue;
import cn.com.ths.exportdoc.thsexportdoc.utils.StringUtils;

public class RepDocTemplate {

    private boolean isLicense = false;
    private Context ctx;
    //初始化日志

    //用户初始化License
    public RepDocTemplate(Context ctx) {
        this.ctx = ctx;
    }



    /**
     * 根据模板和数据导出word
     * @param documentPage  文档体参数对象
     * @param paramDynamicTableMap  动态表Map
     * @param bookmarks  循环表格书签集合
     * @param fileName   下载显示的名称
     * @param templatePath 模板地址
     * @param isTransToPDF 是否转为pdf
     * @throws Exception
     */
    public  void exportWord(DocumentPage documentPage,
                            Map<String,DocDynamicTable> paramDynamicTableMap,
                            List<String> bookmarks,
                            String fileName,
                            String templatePath,
                            boolean isTransToPDF,CallbackContext callbackContext) throws Exception {
        //移动模版文件到指定目录
        //templatePath = cpTemplate(templatePath);

        File checkFile = new File(templatePath);
        if(!checkFile.exists()){
            Log.e("提示","模板文件不存在!");
            callbackContext.error("模板文件不存在!");
            return;
        }
        System.out.println("Word模板路径:" + templatePath);

        Document doc = new Document(templatePath);
        doc=writeWordWithTemplete(documentPage,paramDynamicTableMap,bookmarks,doc);
        //是否在浏览器直接打开
        boolean onlineBrowse=false;
        if(documentPage!=null && documentPage.isOnlineBrowse())
            onlineBrowse=true;
        //获取后缀名
        String suffix = StringUtils.getFilenameExtension(templatePath);
        //导出
        exportWord(doc,fileName,suffix,onlineBrowse);
    }

    /**
     * 将doucument导出word文件
     * @param document 对象
     * @param fileName 文件名称
     * @param suffix 后缀名
     * @param onlineBrowse 是否在线预览
     * @throws Exception
     */
    public void exportWord(Document document, String fileName, String suffix, boolean onlineBrowse) throws Exception {
        if(document == null) {
            throw new Exception("document is null,导出异常");
        }
        if(StringUtils.isNull(suffix)) {
            throw new Exception("后缀名不存在,导出异常");
        }
        if(!suffix.startsWith(".")) {
            suffix = "."+suffix;
        }
        ByteArrayOutputStream outbyte = null;
        ByteArrayInputStream inbyte = null;
        OutputStream out = null;
        try{
            //初始化字节输出流
            outbyte = new ByteArrayOutputStream(1024);
            String extension="";
            //保存修改的内容
            if(StringUtils.endsWithIgnoreCase(suffix, ".pdf")){
                document.save(outbyte, SaveFormat.PDF);
                extension="pdf";
            }else if(StringUtils.endsWithIgnoreCase(suffix, ".docx")){
                document.save(outbyte, SaveFormat.DOCX);
                extension="docx";
            }else{
                document.save(outbyte, SaveFormat.DOC);
                extension="doc";
            }
            // 字节数据流转换为字节输入流
            inbyte = new ByteArrayInputStream(outbyte.toByteArray());
            byte[] buf = new byte[1024];
            int len = 0;
            // Environment.getExternalStorageDirectory().getAbsolutePath()+ File.separator + "deviceId/";

            out = new FileOutputStream(new File(getSDPath()+"/"+fileName+"."+extension));
            while ((len = inbyte.read(buf)) > 0) {
                out.write(buf, 0, len);
            }
            out.flush();
            out.close();
            // this.sendToBrowser(inbyte, fileName, extension, onlineBrowse);
        }catch (Exception e) {
            Log.e("提示",e.getMessage(), e);
        } finally {
            if (null != outbyte) {
                try {
                    outbyte.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (null != inbyte) {
                try {
                    inbyte.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    public String getSDPath(){
        File sdDir = null;
        boolean sdCardExist = Environment.getExternalStorageState()
                .equals(android.os.Environment.MEDIA_MOUNTED);//判断sd卡是否存在
        if(sdCardExist)
        {
            sdDir = Environment.getExternalStorageDirectory();//获取跟目录
        }
        return sdDir.toString();
    }
    /**
     * 处理word模板并将word保存到指定路径
     * @param documentPage 文档体参数对象
     * @param paramDynamicTableMap 动态表Map
     * @param bookmarks 循环表格书签集合
     * @param fileName 下载显示的名称
     * @param templatePath 模板地址
     * @param isTransToPDF 是否转为PDF
     * @param destPath word保存的路径
     * @return
     * @throws Exception
     */
    public void  saveWord(DocumentPage documentPage,
                          Map<String,DocDynamicTable> paramDynamicTableMap,
                          List<String> bookmarks,
                          String fileName,
                          String templatePath,
                          boolean isTransToPDF,
                          String destPath) throws Exception {

        File checkFile = new File(templatePath);
        if(!checkFile.exists()){
            Log.e("提示","模板文件不存在!");
            return;
        }
        if(!isLicense){
            Log.e("提示","许可证已失效!");
            return;
        }
        System.out.println("Word模板路径:" + templatePath);

        Document doc = new Document(templatePath);
        doc=writeWordWithTemplete(documentPage,paramDynamicTableMap,bookmarks,doc);
        String suffix = StringUtils.getFilenameExtension(templatePath);
        saveWord(doc,fileName,suffix,isTransToPDF,destPath);
    }

    /**
     * 将word对象保存到本地
     * @param document
     * @param fileName 文件名称
     * @param isTransToPDF 是否转为pdf
     * @param destPath word保存的路径（不包含文件名）
     * @throws Exception
     */
    public void saveWord(Document document, String fileName, String suffix, boolean isTransToPDF, String destPath) throws Exception {
        //保存修改的内容
        if(isTransToPDF){
            document.save(destPath+"/"+fileName+".pdf", SaveFormat.PDF);
        }else if(".docx".endsWith(suffix)){
            document.save(destPath+"/"+fileName+".docx", SaveFormat.DOCX);
        }else{
            document.save(destPath+"/"+fileName+".doc", SaveFormat.DOC);
        }
    }
    /**
     *
     * @param documentPage 文档体参数对象
     * @param paramDynamicTableMap 动态表Map
     * @param bookmarks 循环表格书签集合
     * @param templatePath 模板路径
     * @return
     * @throws Exception
     */
    public Document writeWordWithTemplete(DocumentPage documentPage,
                                          Map<String,DocDynamicTable> paramDynamicTableMap,
                                          List<String> bookmarks,
                                          String templatePath) throws Exception {

        File checkFile = new File(templatePath);
        if(!checkFile.exists()){
            Log.e("提示","模板文件不存在!");
            return null;
        }
        if(!isLicense){
            Log.e("提示","许可证已失效!");
            return null;
        }
        Document doc=new Document(templatePath);
        return writeWordWithTemplete(documentPage,paramDynamicTableMap,bookmarks,doc);
    }

    /**
     * 向模板doc中填充数据
     * @param documentPage
     * @param paramDynamicTableMap
     * @param bookmarks 循环表格书签集合
     * @param doc
     * @return
     * @throws Exception
     */
    public Document writeWordWithTemplete(DocumentPage documentPage,
                                          Map<String,DocDynamicTable> paramDynamicTableMap,
                                          List<String> bookmarks,
                                          Document doc) throws Exception {

        DocumentBuilder builder = new DocumentBuilder(doc);
        //处理动态表的数据
        if(paramDynamicTableMap!=null){
            this.replaceDynamicTable(doc,builder,paramDynamicTableMap);
        }
        //处理循环动态表格
        if(paramDynamicTableMap!=null){
            this.replaceDynamicMultiTable(doc,builder,paramDynamicTableMap,bookmarks);
        }

        //处理文档对象
        if(documentPage!=null){
            //处理文档体参数
            if(documentPage.getParamMap()!=null && documentPage.getParamMap().size()>0){
                this.replaceDocParam(doc,builder, documentPage.getParamMap());
            }
            //处理文档体图片
            if(documentPage.getImages()!=null && documentPage.getImages().size()>0){
                this.insertImages(builder, documentPage.getImages());
            }
            //处理文档体文件对象
            if(documentPage.getOleObjects()!=null && documentPage.getOleObjects().size()>0) {
                this.insertOleObjects(builder, documentPage.getOleObjects());
            }
            //处理文档体内的上标和下标
            if(documentPage.getSuperAndSubs()!=null && documentPage.getSuperAndSubs().size()>0){
                this.replaceSuperAndSub(doc,documentPage.getSuperAndSubs());
            }
            //更新目录
            if(documentPage.isUpdateCatalog()){
                doc.updateFields();
            }

        }
        return doc;
    }

    /**
     * 替换整个文档体的上下标
     * @param doc
     * @param superAndSubs
     * @throws Exception
     */
    private void replaceSuperAndSub(Document doc,
                                    List<DocSuperAndSub> superAndSubs) throws Exception {

        for(DocSuperAndSub docSuperAndSub:superAndSubs){
            String allWord=docSuperAndSub.getAllWord();
            String baseWord=docSuperAndSub.getBaseWord();
            if(!StringUtils.isNull(allWord) && !StringUtils.isNull(baseWord)){
                Pattern compile = Pattern.compile(allWord);
                doc.getRange().replace(compile, new SupAndSubRepCallback(docSuperAndSub), true);
            }
        }
    }

    /**
     * 替换整个document的参数
     * @param doc
     * @param builder
     * @param dataMap
     * @throws Exception
     */
    private void replaceDocParam(Document doc, DocumentBuilder builder, Map<String, Object> dataMap) throws Exception {
        //遍历要替换的内容
        if(dataMap!=null){
            Iterator<String> keys = dataMap.keySet().iterator();
            while (keys.hasNext()) {
                String key=keys.next();
                //替换占位符的内容
                this.replace(doc,builder, key, dataMap.get(key));
            }
        }
    }

    /**
     * 处理动态单表格的参数
     * @param doc
     * @param builder
     * @param tableRowData
     * @throws Exception
     */
    @SuppressWarnings("rawtypes")
    private void replaceDynamicTable(Document doc, DocumentBuilder builder, Map<String,DocDynamicTable> tableRowData) throws Exception {
        //替换需要动态行的表格
        NodeCollection tables = doc.getChildNodes(NodeType.TABLE, true);
        for(int i=0;i<tables.getCount();i++){
            Table currentTable=(Table) tables.get(i);
            RowCollection currentRows = currentTable.getRows();
            int currentRowNum = 0;
            int totalRowNum = currentRows.getCount();
            //循环所有的行，当前行数到达行底部的时候结束循环；因为循环时会插入数据行，所以当前表格的行数是在一直变化的
            while(currentRowNum<totalRowNum) {
                Row currentRow = currentRows.get(currentRowNum);
                String rowTitle=this.getTile(currentRow);
                DocDynamicTable tableRow=tableRowData.get(rowTitle);
                if(tableRow==null) {
                    currentRowNum = currentRowNum+1;
                    continue;
                }
                currentRow.remove();
                DocTable docTable = tableRow.getDocTable();
                if(docTable.getDataRowStarts() == null || docTable.getDataRowStarts().length<=0) {
                    throw new Exception("未设置数据的开始行信息");
                }
                if(docTable.getRowList()==null || docTable.getRowList().size()<=0) {
                    //将空表格的占位符替换为空
                    currentRowNum=this.replaceNull(docTable, currentRows,currentRowNum);
                }else{
                    //处理动态表格的数据
                    currentRowNum=this.insertRows(builder, currentRows, docTable,currentRowNum);
                }
                //获取当前的总行数
                totalRowNum = currentRows.getCount();
            }
        }
    }

    /**
     * 处理动态多表格的数据
     * @param doc
     * @param builder
     * @param paramDynamicTableMap
     * @param bookmarks 标签集合
     * @throws Exception
     */
    @SuppressWarnings("deprecation")
    private void replaceDynamicMultiTable(Document doc,
                                          DocumentBuilder builder,
                                          Map<String, DocDynamicTable> paramDynamicTableMap,
                                          List<String> bookmarks) throws Exception {
        if(bookmarks==null || bookmarks.size()<=0)
            return;
        //遍历标签集合，查找需要进行子模版替换的标签
        for(String bookmark:bookmarks){
            //移动到标签位置
            builder.moveToBookmark(bookmark);
            //获取设置的模板数据对象
            DocDynamicTable dynamicTable=paramDynamicTableMap.get(bookmark);
            if(dynamicTable==null)
                continue;
            DocMultiTable multiTable=dynamicTable.getMultiTable();
            if(multiTable.getMultiTable()!=null && multiTable.getMultiTable().size()>0) {
                //旧的动态多表处理逻辑，兼容老项目
                this.multiTable4Old(doc, multiTable, builder);
            }else if(multiTable.getMultiComplexTable()!=null && multiTable.getMultiComplexTable().size()>0) {
                //新的动态多表处理逻辑
                this.multiTable4New(doc, multiTable, builder);
            }
        }
    }

    /**
     * 旧的动态多表处理方式
     * @throws Exception
     */
    @SuppressWarnings({ "deprecation", "rawtypes" })
    private void multiTable4Old(Document doc, DocMultiTable multiTable, DocumentBuilder builder) throws Exception {
        List<DocDynamicTable.InnerDocTable> docTables=multiTable.getMultiTable();
        Document childDoc=new Document(multiTable.getModelPath());
        //复制子模板，后续每循环一次都要复制一个
        Document copyDoc=childDoc.deepClone();
        //上一节点，初始位置为标签位置，没插入一个表格，更新该值为表格，以保证下一次插入在表格下方
        Node previousNode=builder.getCurrentNode().getParentNode();
        for(DocTable docTable:docTables){
            DocumentBuilder copyBuilder=new DocumentBuilder(copyDoc);
            //处理表体参数
            this.replaceDocParam(copyDoc, copyBuilder, docTable.getParamMap());
            //如果存在动态行数据，进行动态行添加
            if(docTable.getRowList()!=null && docTable.getRowList().size()>0){
                NodeCollection tables = copyDoc.getChildNodes(NodeType.TABLE, true);
                if(tables!=null && tables.getCount()>0) {
                    Table currentTable=(Table) tables.get(0);
                    //插入动态表数据
                    if(docTable.getRowList().size()>0){
                        this.insertRows(copyBuilder, currentTable.getRows(), docTable,0);
                    }
                }
            }
            //处理表体图片
            if(docTable.getImages()!=null && docTable.getImages().size()>0) {
                this.insertImages(copyBuilder, docTable.getImages());
            }
            //将另一个文档的元素导入本文档，需要先定义一个导入器
            NodeImporter importer = new NodeImporter(copyDoc, doc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            //获取子模板body元素中所有的Paragraph和Table，然后将这些内容导入父文档中
            NodeCollection childNodes = copyDoc.getChildNodes(NodeType.BODY,true);
            for(int i=0;i<childNodes.getCount();i++){
                Body body=(Body) childNodes.get(i);
                for(int j=0;j<body.getChildNodes().getCount();j++){
                    Node destNode=importer.importNode(body.getChildNodes().get(j), true);
                    //获取标签位置所在的body，然后将元素插在该标签的后面
                    builder.getCurrentNode().getParentNode().getParentNode().insertAfter(destNode,previousNode);
                    previousNode=destNode;
                }
            }
            //复制一个新的子模板
            copyDoc=childDoc.deepClone();
        }
    }

    /**
     * 新的动态多表处理方式
     * @param doc 文档对象
     * @param multiTable 多表数据对象
     * @param builder 文档builder对象
     * @throws Exception
     */
    @SuppressWarnings("rawtypes")
    private void multiTable4New(Document doc, DocMultiTable multiTable, DocumentBuilder builder) throws Exception {
        List<DocComplexTable> multiComplexTables = multiTable.getMultiComplexTable();
        Document childDoc=new Document(multiTable.getModelPath());
        //复制子模板，后续每循环一次都要复制一个
        Document copyDoc=childDoc.deepClone();
        //上一节点，初始位置为标签位置，每插入一个子模版元素，就更新该值
        Node previousNode=builder.getCurrentParagraph();
        //获取标签位置的body对象，子模版的元素都要插入在该body对象中
        Body curBody = (Body) builder.getCurrentSection().getChildNodes(NodeType.BODY,true).get(0);
        for(DocComplexTable complexTable:multiComplexTables){
            DocumentBuilder copyBuilder=new DocumentBuilder(copyDoc);
            //处理表体参数
            this.replaceDocParam(copyDoc, copyBuilder, complexTable.getParamMap());
            //如果存在动态行数据，进行动态行添加
            if(complexTable.getInnerTableMap()!=null && complexTable.getInnerTableMap().size()>0){
                this.multiTable4NewDataHandler(copyDoc, copyBuilder, complexTable.getInnerTableMap());
            }
            //处理表体图片
            if(complexTable.getImages()!=null && complexTable.getImages().size()>0) {
                this.insertImages(copyBuilder, complexTable.getImages());
            }
            //插入文件对象
            if(complexTable.getOleObjects()!=null && complexTable.getOleObjects().size()>0) {
                this.insertOleObjects(copyBuilder, complexTable.getOleObjects());
            }
            //将另一个文档的元素导入本文档，需要先定义一个导入器
            NodeImporter importer = new NodeImporter(copyDoc, doc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            //获取子模板body元素中所有的Paragraph和Table，然后将这些内容导入父文档中
            NodeCollection childNodes = copyDoc.getChildNodes(NodeType.BODY,true);
            for(int i=0;i<childNodes.getCount();i++){
                Body body=(Body) childNodes.get(i);
                for(int j=0;j<body.getChildNodes().getCount();j++){
                    //通过导入器获取子模版的元素，将元素插入到curbody中，并更新前置元素previousNode的值
                    Node destNode=importer.importNode(body.getChildNodes().get(j), true);
                    curBody.insertAfter(destNode, previousNode);
                    previousNode = destNode;
                }
            }
            //复制一个新的子模板
            copyDoc=childDoc.deepClone();
        }
    }

    /**
     * 多表新的处理方式核心处理逻辑
     * @param copyDoc
     * @param copyBuilder
     * @param innerTableMap
     * @throws Exception
     */
    @SuppressWarnings("rawtypes")
    private void multiTable4NewDataHandler(Document copyDoc, DocumentBuilder copyBuilder, Map<String, DocTable> innerTableMap) throws Exception {
        NodeCollection tables = copyDoc.getChildNodes(NodeType.TABLE, true);
        for(int i=0;i<tables.getCount();i++){
            Table currentTable=(Table) tables.get(i);
            RowCollection currentRows = currentTable.getRows();
            int currentRowNum = 0;
            int totalRowNum = currentRows.getCount();
            //循环所有的行，当前行数到达行底部的时候结束循环；因为循环时会插入数据行，所以当前表格的行数是在一直变化的
            while(currentRowNum<totalRowNum) {
                Row currentRow = currentRows.get(currentRowNum);
                String rowTitle=this.getTile(currentRow);
                DocTable innerTable=innerTableMap.get(rowTitle);
                if(innerTable==null) {
                    currentRowNum = currentRowNum+1;
                    continue;
                }
                if(innerTable.getDataRowStarts() == null || innerTable.getDataRowStarts().length<=0) {
                    throw new Exception("未设置数据开始行信息");
                }
                currentRow.remove();
                if(innerTable.getRowList()==null || innerTable.getRowList().size()<=0) {
                    //将空表格的占位符替换为空
                    currentRowNum=this.replaceNull(innerTable, currentRows,currentRowNum);
                }else{
                    //处理动态表格的数据
                    currentRowNum=this.insertRows(copyBuilder, currentRows, innerTable,currentRowNum);
                }
                //获取当前的总行数
                totalRowNum = currentRows.getCount();
            }
        }
    }

    /**
     * 替换某个节点中的标识符
     * @param node
     * @param builder
     * @param key
     * @param value
     * @throws Exception
     */
    private void replace(Node node, DocumentBuilder builder, String key, Object value) throws Exception {
        //如果参数值是字符串，则采用字符串的方式进行替换
        if(value instanceof DocParamValue) {
            DocParamValue paramValue = (DocParamValue) value;
            //如果替换值是类，就通过回调类来处理
            if(DocParamTypeEnum.REPLACECLASS.equals(paramValue.getType())) {
                replaceClassParam(node,key,paramValue);
            }else if(DocParamTypeEnum.TITLE.equals(paramValue.getType())) {
                //如果替换值是标题，用特殊的标题处理
                replaceTilteParam(node,key,paramValue);
            }
        }else{
            replaceStrParam(node,builder,key,value);
        }
    }

    /**
     * 替换string
     * @param node
     * @param builder
     * @param key
     * @param value
     * @throws Exception
     */
    private void replaceStrParam(Node node, DocumentBuilder builder, String key, Object value) throws Exception {
        String valueStr = String.valueOf(value);
        //对显示值得修改
        if(StringUtils.isNull(valueStr)){
            valueStr = "";
        }
        //如果值中包含换行符，则需要将值中的换行符统一处理为"\r\n"
        if(valueStr.contains("\r\n")){
            Pattern compile = Pattern.compile("\\$"+key+"\\$");
            node.getRange().replace(compile, new DocReplaceCallback(valueStr), false);
        }else{
            //替换占位符,此种方式下，替换的value值中不能包含换行符，否则会报错
            node.getRange().replace("$"+key+"$", valueStr, true, false);
        }
        for(FormField field:node.getRange().getFormFields()){
            if(key.equalsIgnoreCase(field.getName())){
                builder.moveToBookmark(field.getName());
                builder.write(valueStr);
                field.remove();
            }
        }
    }

    /**
     * 通过回调类来替换参数
     * @param node
     * @param key
     * @param paramValue
     * @throws Exception
     */
    @SuppressWarnings({ "unchecked", "rawtypes" })
    private void replaceClassParam(Node node, String key, DocParamValue paramValue) throws Exception {
        String className = paramValue.getClassName();
        Object value = paramValue.getValue();
        if(StringUtils.isNull(className)) {
            throw new Exception("反射类为空，替换出现异常！");
        }
        Class clazz = Class.forName(className);
        Constructor constructor = clazz.getConstructor(Object.class);
        Pattern compile = Pattern.compile("\\$"+key+"\\$");
        node.getRange().replace(compile, (IReplacingCallback)constructor.newInstance(value), false);
    }

    /**
     * 替换为标题
     * @param node
     * @param key
     * @throws Exception
     */
    @SuppressWarnings({ "unused", "unchecked" })
    private void replaceTilteParam(Node node, String key, DocParamValue paramValue) throws Exception {
        Object value = paramValue.getValue();
        Integer titleLevel = paramValue.getTitleLevel();
        Pattern compile = Pattern.compile("\\$"+key+"\\$");
        node.getRange().replace(compile, new TitleParamCallback(titleLevel,value),false);
    }
    /**
     * 获取首行文本内容
     * @param row
     * @return
     */
    private String getTile(Row row){
        String title=row.getText();
        if(!StringUtils.isNull(title) && title.startsWith("#")){
            title=title.split("#")[1];
        }
        return title;
    }

    /**
     * 表格插入多行数据的方法
     * @param builder
     * @param currentRows
     * @param docTable
     * @throws Exception
     */
    @SuppressWarnings("unused")
    private int insertRows(DocumentBuilder builder, RowCollection currentRows, DocTable docTable, int currentRowNum) throws Exception {
        int pageSize=100;
        //获取动态行信息
        List<Map<String,Object>> rowList=docTable.getRowList();
        //获取数据的开始行
        int[] startRows= docTable.getDataRowStarts();
        //获取需要合并的列
        int[] mergeColumns=docTable.getMergeColumns();
        Row[] models=new Row[startRows.length];
        //封装每次折行的开始行
        for(int m=0;m<startRows.length;m++){
            models[m]=currentRows.get(currentRowNum+startRows[m]-1);
        }
        for(int m=0;m<models.length;m++){
            //为了解决动态表格大数据量导出的问题,设置为每100条创建一个子模板进行替换，同一个模板如果数据量过大时，replace的效率特别低
            int page=0;
            Row currentRow=models[m];
            Document childDoc=null;//子模版
            RowCollection childRows=null;//子模版数据行数
            int childTableRow=0;//子模版数据游标
            DocumentBuilder childBuilder=null;
            Map<String,Row> rowMap=new HashMap<String,Row>();
            Row childNewRow=null;//子模版临时新创建行
            Row childCurrentRow=null;//子模版临时当前行
            Row childpreviousRow=null;//子模版临时前置行
            int totalSize= pageSize;
            for(int j=0;j<rowList.size();j++){
                //每pageSize条数据创建一个子doc
                if(j%pageSize==0){
                    childTableRow=0;
                    //计算该子模版下的可替换的总行数
                    totalSize=((rowList.size()-page*pageSize)>=pageSize)?pageSize:rowList.size()%pageSize;
                    childDoc=new Document();
                    childBuilder=new DocumentBuilder(childDoc);
                    Table startTable = new Table(childDoc);
                    //两个模板之间同步元素时，需要定义importor
                    NodeImporter importer = new NodeImporter(currentRow.getDocument(),childDoc , ImportFormatMode.KEEP_SOURCE_FORMATTING);
                    childNewRow=(Row) importer.importNode(currentRow, true);
                    childCurrentRow=(Row) childNewRow.deepClone(true);
                    childRows = startTable.getRows();
                    childRows.add(childCurrentRow);
                    page++;
                }
                rowMap=this.handleRow(childBuilder, childpreviousRow, childCurrentRow, childNewRow, childRows, rowList.get(j), mergeColumns, totalSize-1, 1, childTableRow, 0);
                childTableRow++;
                childpreviousRow=rowMap.get("previousRow");
                childCurrentRow=rowMap.get("currentRow");
                childNewRow=rowMap.get("newRow");
                if((j+1)%pageSize==0 || j==rowList.size()-1){
                    int diversity=pageSize;
                    if(j==rowList.size()-1){
                        diversity=(rowList.size()-1)%pageSize;
                    }
                    NodeImporter importer = new NodeImporter(childDoc, builder.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);
                    for(int n=0;n<childRows.getCount();n++){
                        Row temp=childRows.get(n);
                        Row destRow = (Row) importer.importNode(temp, true);
                        if(m==0){
                            currentRows.insert(currentRowNum+startRows[m]+(j-totalSize)+n, destRow);
                        }else{
                            currentRows.insert(currentRowNum+startRows[m]+rowList.size()+(j-totalSize)+n-1,destRow);
                        }
                    }
                }
            }
            currentRow.remove();
        }
        //获取遍历后的当前数据行行数，初始行+模板数据开始行（相对）+数据行数*标题折行数
        currentRowNum = currentRowNum+startRows[startRows.length-1]-1+rowList.size()*startRows.length;
        return currentRowNum;
    }
    /**
     * 处理每行的数据
     * @param builder
     * @param previousRow 前置行
     * @param currentRow 当前行
     * @param newRow 新复制的行
     * @param currentRows 当前table的row对象集合
     * @param data 每行的数据
     * @param mergeColumns 需要合并的列
     * @param totalRowSize 数据总行数
     * @param startRowNum 开始行数
     * @param curRowNum 当前行数
     * @param wrapNum 表格的折行数
     * @throws Exception
     */
    private Map<String,Row> handleRow(
            DocumentBuilder builder,
            Row previousRow,
            Row currentRow,
            Row newRow,
            RowCollection currentRows,
            Map<String,Object> data,
            int[] mergeColumns,
            int totalRowSize,
            int startRowNum,
            int curRowNum,
            int wrapNum) throws Exception {

        for(Map.Entry<String,Object> entry:data.entrySet()){
            this.replace(currentRow,builder, entry.getKey(), entry.getValue());
            for(Field field:currentRow.getRange().getFields()){
                field.remove();
            }
        }
        //合并列
        if(curRowNum>0){
            CellCollection currentRowcells = currentRow.getCells();
            CellCollection previousRowcells=previousRow.getCells();
            for(int n=0;n<currentRowcells.getCount();n++){
                if(mergeColumns!=null){
                    //循环合并列的参数，如果存在该列，则进行合并
                    for(int column:mergeColumns){
                        if(n==column-1){
                            Cell currentCell=currentRowcells.get(n);
                            Cell previousCell=previousRowcells.get(n);
                            if(currentCell.getText().equals(previousCell.getText())){
                                currentCell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
                                if(curRowNum==1){
                                    previousCell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
                                }
                            }else{
                                currentCell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
                            }
                        }
                    }
                }
            }
        }
        //更新前置行的值
        previousRow=currentRow;
        //插入模板行
        if(curRowNum<totalRowSize){
            currentRows.insert(startRowNum+curRowNum, newRow);
            currentRow=currentRows.get(startRowNum+curRowNum);
            newRow=(Row) newRow.deepClone(true);
        }
        Map<String,Row> rowMap=new HashMap<String,Row>();
        rowMap.put("currentRow", currentRow);
        rowMap.put("newRow", newRow);
        rowMap.put("previousRow", previousRow);
        return rowMap;
    }
    /**
     * 插入图片
     * @param builder
     * @param images
     * @throws Exception
     */
    private void insertImages(DocumentBuilder builder, Map<String,List<DocImage>> images) throws Exception {

        for(Map.Entry<String,List<DocImage>> entry:images.entrySet()){
            String bookmark=entry.getKey();
            boolean isMove = builder.moveToBookmark(bookmark,true,false);
            ParagraphFormat paragraphFormat = builder.getCurrentParagraph().getParagraphFormat();
            if(isMove){
                List<DocImage> imageList=entry.getValue();
                for(DocImage image:imageList){
                    String imgPath=image.getPath();
                    InputStream in=image.getInputStream();
                    Double width=image.getWidth();
                    Double height=image.getHeight();
                    Double left=image.getLeft();
                    Double top=image.getTop();
                    Integer wrapType=image.getWrapType();
                    if(width==null){
                        width=-1.0;
                    }
                    if(height==null){
                        height=-1.0;
                    }
                    if(left==null){
                        left=0.0;
                    }
                    if(top==null){
                        top=0.0;
                    }
                    Boolean behindText = null;
                    if(wrapType==null){
                        wrapType= WrapType.INLINE;
                    }else if(wrapType==6) {
                        wrapType = WrapType.NONE;
                        behindText = true;
                    }else if(wrapType == 7) {
                        wrapType = WrapType.NONE;
                        behindText = false;
                    }
                    Shape insertImage = null;
                    if(!StringUtils.isNull(imgPath)){
                        insertImage = builder.insertImage(imgPath, 0, left, 0, top, width, height, wrapType);
                    }else if(null!=in){
                        insertImage = builder.insertImage(in, 0, left, 0, top, width, height, wrapType);
                    }
                    //处理图片衬于文字上方和衬于文字下方两种方式
                    if(behindText!=null) {
                        insertImage.setBehindText(behindText);
                    }
                    //如果图例内容不为空，则插入图例内容
                    if(!StringUtils.isNull(image.getLegendContent())){
                        //创建段落
                        Paragraph paragraph = builder.insertParagraph();
                        builder.moveTo(paragraph);
                        paragraph.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
                        paragraph.getParagraphFormat().setSpaceBefore(5d);
                        //创建文字块
                        Run run=new Run(builder.getDocument());
                        paragraph.appendChild(run);
                        run.setText(image.getLegendContent());
                        if(image.getLengendFont()==null){
                            //默认设置为14
                            run.getFont().setSize(14d);
                        }else{
                            this.copyFont(image.getLengendFont(), run.getFont());
                        }
                        Paragraph nextParagraph = builder.insertParagraph();
                        nextParagraph.getParagraphFormat().setAlignment(paragraphFormat.getAlignment());
                        nextParagraph.getParagraphFormat().setSpaceBefore(paragraphFormat.getSpaceBefore());
                        builder.moveTo(nextParagraph);
                    }
                }
            }
        }
    }

    /**
     * 插入文档对象
     * @param builder
     * @param objects
     * @throws Exception
     */
    private void insertOleObjects(DocumentBuilder builder, Map<String,List<DocOleObject>> objects) throws Exception {
        for(Map.Entry<String,List<DocOleObject>> entry:objects.entrySet()){
            String bookmark=entry.getKey();
            boolean isMove = builder.moveToBookmark(bookmark,true,false);
            if(isMove) {
                List<DocOleObject> objectList = entry.getValue();
                for(DocOleObject object : objectList) {
                    String filePath = object.getFilePath();
                    if(!StringUtils.isNull(filePath)) {
                        builder.insertOleObject(filePath, false, true,null);
                    }
                }
            }
        }


    }
    /**
     * 处理设置的字体格式
     * @param srcFont
     * @param targetFont
     * @throws Exception
     */
    private void copyFont(Map<String,Object> srcFont, Font targetFont) throws Exception {
        if(srcFont.get("AllCaps")!=null){
            targetFont.setAllCaps(Boolean.valueOf(String.valueOf(srcFont.get("allCaps"))));
        }
        if(srcFont.get("Bidi")!=null){
            targetFont.setBidi(Boolean.valueOf(String.valueOf(srcFont.get("Bidi"))));
        }
        if(srcFont.get("Bold")!=null){
            targetFont.setBold(Boolean.valueOf(String.valueOf(srcFont.get("Bold"))));
        }
        if(srcFont.get("BoldBi")!=null){
            targetFont.setBoldBi(Boolean.valueOf(String.valueOf(srcFont.get("BoldBi"))));
        }
//        if(srcFont.get("Color")!=null){
//            targetFont.setColor((Color)srcFont.get("Color"));
//        }
        if(srcFont.get("ComplexScript")!=null){
            targetFont.setComplexScript(Boolean.valueOf(String.valueOf(srcFont.get("ComplexScript"))));
        }
        if(srcFont.get("DoubleStrikeThrough")!=null){
            targetFont.setDoubleStrikeThrough(Boolean.valueOf(String.valueOf(srcFont.get("DoubleStrikeThrough"))));
        }
        if(srcFont.get("Emboss")!=null){
            targetFont.setEmboss(Boolean.valueOf(String.valueOf(srcFont.get("Emboss"))));
        }
        if(srcFont.get("Engrave")!=null){
            targetFont.setEngrave(Boolean.valueOf(String.valueOf(srcFont.get("Engrave"))));
        }
        if(srcFont.get("Hidden")!=null){
            targetFont.setHidden(Boolean.valueOf(String.valueOf(srcFont.get("Hidden"))));
        }
//        if(srcFont.get("HighlightColor")!=null){
//            targetFont.setHighlightColor((Color)srcFont.get("HighlightColor"));
//        }
        if(srcFont.get("Italic")!=null){
            targetFont.setItalic(Boolean.valueOf(String.valueOf(srcFont.get("Italic"))));
        }
        if(srcFont.get("ItalicBi")!=null){
            targetFont.setItalicBi(Boolean.valueOf(String.valueOf(srcFont.get("ItalicBi"))));
        }
        if(srcFont.get("Kerning")!=null){
            targetFont.setKerning(Double.valueOf(String.valueOf(srcFont.get("Kerning"))));
        }
        if(srcFont.get("LocaleId")!=null){
            targetFont.setLocaleId(Integer.valueOf(String.valueOf(srcFont.get("LocaleId"))));
        }
        if(srcFont.get("LocaleIdBi")!=null){
            targetFont.setLocaleIdBi(Integer.valueOf(String.valueOf(srcFont.get("LocaleIdBi"))));
        }
        if(srcFont.get("LocaleIdFarEast")!=null){
            targetFont.setLocaleIdFarEast(Integer.valueOf(String.valueOf(srcFont.get("LocaleIdFarEast"))));
        }
        if(srcFont.get("Name")!=null){
            targetFont.setName(String.valueOf(srcFont.get("Name")));
        }
        if(srcFont.get("NameAscii")!=null){
            targetFont.setNameAscii(String.valueOf(srcFont.get("NameAscii")));
        }
        if(srcFont.get("NameBi")!=null){
            targetFont.setNameBi(String.valueOf(srcFont.get("NameBi")));
        }
        if(srcFont.get("NameFarEast")!=null){
            targetFont.setNameFarEast(String.valueOf(srcFont.get("NameFarEast")));
        }
        if(srcFont.get("NameOther")!=null){
            targetFont.setNameOther(String.valueOf(srcFont.get("NameOther")));
        }
        if(srcFont.get("NoProofing")!=null){
            targetFont.setNoProofing(Boolean.valueOf(String.valueOf(srcFont.get("NoProofing"))));
        }
        if(srcFont.get("NoProofing")!=null){
            targetFont.setNoProofing(Boolean.valueOf(String.valueOf(srcFont.get("NoProofing"))));
        }
        if(srcFont.get("Outline")!=null){
            targetFont.setOutline(Boolean.valueOf(String.valueOf(srcFont.get("Outline"))));
        }
        if(srcFont.get("Position")!=null){
            targetFont.setPosition(Double.valueOf(String.valueOf(srcFont.get("Position"))));
        }
        if(srcFont.get("Scaling")!=null){
            targetFont.setScaling(Integer.valueOf(String.valueOf(srcFont.get("Scaling"))));
        }
        if(srcFont.get("Shadow")!=null){
            targetFont.setShadow(Boolean.valueOf(String.valueOf(srcFont.get("Shadow"))));
        }
        if(srcFont.get("Size")!=null){
            targetFont.setSize(Double.valueOf(String.valueOf(srcFont.get("Size"))));
        }
        if(srcFont.get("SizeBi")!=null){
            targetFont.setSizeBi(Double.valueOf(String.valueOf(srcFont.get("SizeBi"))));
        }
        if(srcFont.get("SmallCaps")!=null){
            targetFont.setSmallCaps(Boolean.valueOf(String.valueOf(srcFont.get("SmallCaps"))));
        }
        if(srcFont.get("Spacing")!=null){
            targetFont.setSpacing(Double.valueOf(String.valueOf(srcFont.get("Spacing"))));
        }
        if(srcFont.get("StrikeThrough")!=null){
            targetFont.setStrikeThrough(Boolean.valueOf(String.valueOf(srcFont.get("StrikeThrough"))));
        }
        if(srcFont.get("Subscript")!=null){
            targetFont.setSubscript(Boolean.valueOf(String.valueOf(srcFont.get("Subscript"))));
        }
        if(srcFont.get("Superscript")!=null){
            targetFont.setSuperscript(Boolean.valueOf(String.valueOf(srcFont.get("Superscript"))));
        }
        if(srcFont.get("Style")!=null){
            targetFont.setStyle((Style)srcFont.get("Style"));
        }
        if(srcFont.get("StyleIdentifier")!=null){
            targetFont.setStyleIdentifier(Integer.valueOf(String.valueOf(srcFont.get("StyleIdentifier"))));
        }
        if(srcFont.get("StylName")!=null){
            targetFont.setStyleName(String.valueOf(srcFont.get("StylName")));
        }
        if(srcFont.get("TextEffect")!=null){
            targetFont.setTextEffect(Integer.valueOf(String.valueOf(srcFont.get("TextEffect"))));
        }
        if(srcFont.get("Underline")!=null){
            targetFont.setUnderline(Integer.valueOf(String.valueOf(srcFont.get("Underline"))));
        }
        if(srcFont.get("UnderlineColor")!=null){
            targetFont.setUnderlineColor((Color)srcFont.get("UnderlineColor"));
        }
    }



    //将word文档转换为pdf
    private InputStream wordToPdf(Document doc){
        ByteArrayOutputStream outbyte = null;
        ByteArrayInputStream inbyte = null;
        try{
            //初始化字节输出流
            outbyte = new ByteArrayOutputStream(1024);
            doc.save(outbyte, SaveFormat.PDF);
            // 字节数据流转换为字节输入流
            inbyte = new ByteArrayInputStream(outbyte.toByteArray());
        }catch (Exception e) {
            Log.e("提示",e.getMessage(), e);
        } finally {
            if (null != outbyte) {
                try {
                    outbyte.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (null != inbyte) {
                try {
                    inbyte.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return inbyte;
    }

    /**
     * 将WORD转换为pdf
     * @param srcPath
     * @return
     * @throws Exception
     */
    public InputStream wordToPdf(String srcPath) throws Exception {
        Document doc=new Document(srcPath);
        return this.wordToPdf(doc);
    }

    /**
     * 将word转为pdf，并且按照目标路径进行保存
     * @param srcPath 原始文件地址
     * @param destPath 目标地址
     */
    public void wordToPdf(String srcPath, String destPath) {
        try {
            Document doc=new Document(srcPath);
            doc.save(destPath, SaveFormat.PDF);
        } catch (Exception e) {
            System.out.println("word文件转换失败");
            e.printStackTrace();
        }
    }

    /**
     * 将word转换为PDF
     * @param pdfIn
     * @return
     * @throws Exception
     */
    public InputStream wordToPdf(InputStream pdfIn) throws Exception {
        Document doc=new Document(pdfIn);
        return this.wordToPdf(doc);
    }

    /**
     * 将空表格的占位符替换为空
     * @throws Exception
     */
    private  int replaceNull(DocTable docTable, RowCollection currentRows, int currentRowNum) throws Exception {
        int[] starts = docTable.getDataRowStarts();
        int tempRowNum = 0;
        for(int temp:starts){
            tempRowNum = currentRowNum+temp-1;
            Row currentRow=currentRows.get(tempRowNum);
            CellCollection cells = currentRow.getCells();
            for(int m=0;m<cells.getCount();m++){
                Cell cell = cells.get(m);
                cell.getRange().replace(Pattern.compile("\\$.*\\$"), "");
            }
        }
        return tempRowNum;
    }


}

/**
 * 主要用于解决替换内容中包含有换行符的情况
 * @author wangjp
 *
 */
class DocReplaceCallback implements IReplacingCallback {

    private String value;
    public DocReplaceCallback(String value){
        this.value=value;
    }

    @Override
    public int replacing(ReplacingArgs e) throws Exception {
        //将待替换的值用换行符分割为数组
        String[] values=value.split("\r\n");
        Node matchNode = e.getMatchNode();
        Document doc = (Document)matchNode.getDocument();
        DocumentBuilder builder = new DocumentBuilder(doc);
        //移动到待替换的节点
        builder.moveTo(matchNode);
        for(int i=0;i<values.length;i++){
            if(i==0){
                builder.write(values[i]);
            }else{
                builder.writeln();
                builder.write(values[i]);
            }
        }
        return ReplaceAction.REPLACE;
    }
}

/**
 * 主要用于对特殊字符进行上下标处理
 * @author wangjp
 *
 */
class SupAndSubRepCallback implements IReplacingCallback {

    private DocSuperAndSub superAndSub;
    public SupAndSubRepCallback(DocSuperAndSub superAndSub){
        this.superAndSub=superAndSub;
    }

    @Override
    public int replacing(ReplacingArgs e) throws Exception {
        String matchStr=superAndSub.getAllWord();
        //匹配的节点是Run节点，为dom对象的最小单位
        Node matchNode = e.getMatchNode();
        if(!matchNode.getText().contains(matchStr))
            return ReplaceAction.SKIP;
        Run matchRun=(Run)matchNode;
        //获取当前的段落
        Paragraph curPara = matchRun.getParentParagraph();
        /*
         * 修改上下标的时候用builder对象插入节点时会有问题，此处采用其他方式处理，
         * 此处的处理逻辑时，将上下标所在的Run节点文字，按照匹配字符切分，切分后每一单位字符串各自创建一个run对象，上下标文字单独创建run对象，
         * 因为上下标的效果是以run对象为单位进行设置的，所以最后的文字效果是以多个run节点拼成的。
         */
        String allStr=matchRun.getText();
        String[] alls=allStr.split(matchStr);
        Run lastRun=matchRun;
        if(alls.length==0){
            //插入左侧单词
            matchRun.setText(superAndSub.getBaseWord());
            //处理上下标
            lastRun=handler(curPara,matchRun,lastRun);
        }else{
            for(int i=0;i<alls.length;i++){
                String tempStr=alls[i];
                if(i==0){
                    matchRun.setText(tempStr);
                }else{
                    //用matchRun克隆生成Run节点是为了获取原来文本的格式
                    Run tempRun=(Run) matchRun.deepClone(true);
                    tempRun.setText(tempStr);
                    curPara.insertAfter(tempRun, lastRun);
                    lastRun=tempRun;
                }
                if(i<alls.length-1 || alls.length==1){
                    //插入左侧单词
                    Run baseRun = (Run) matchRun.deepClone(true);
                    baseRun.setText(superAndSub.getBaseWord());
                    curPara.insertAfter(baseRun, lastRun);
                    lastRun=baseRun;
                    //处理上下标
                    lastRun=handler(curPara,matchRun,lastRun);
                }
            }
        }
        return ReplaceAction.SKIP;
    }

    /**
     * 处理上下标
     * @throws Exception
     */
    public Run handler(Paragraph curPara, Run matchRun, Run lastRun) throws Exception {
        //处理上标
        if(!StringUtils.isNull(superAndSub.getSuperscript())){
            Run superRun=(Run) matchRun.deepClone(true);
            superRun.setText(superAndSub.getSuperscript());
            superRun.getFont().setSuperscript(true);
            curPara.insertAfter(superRun, lastRun);
            lastRun=superRun;
        }
        //处理下标
        if(!StringUtils.isNull(superAndSub.getSubscript())){
            Run subRun=(Run) matchRun.deepClone(true);
            subRun.setText(superAndSub.getSubscript());
            subRun.getFont().setSubscript(true);
            curPara.insertAfter(subRun, lastRun);
            lastRun=subRun;
        }
        return lastRun;
    }
}

/**
 *
 * 处理标题
 * @author wangjp
 * @date   2018年11月12日
 */
class TitleParamCallback implements IReplacingCallback {
    //标题值
    private Object value;
    //标题级别
    private Integer titleLevel;
    public TitleParamCallback(Integer titleLevel, Object value) {
        this.value = value;
        //设置标题默认是一级
        if(titleLevel == null) {
            titleLevel = 1;
        }
        this.titleLevel = titleLevel;
    }

    @SuppressWarnings("unchecked")
    @Override
    public int replacing(ReplacingArgs e) throws Exception {
        Node matchNode = e.getMatchNode();
        Document doc = (Document)matchNode.getDocument();
        DocumentBuilder builder = new DocumentBuilder(doc);
        //移动到待替换的节点
        builder.moveTo(matchNode);
        //如果值是list，就循环插入段落，设置标题
        if(value instanceof List) {
            List<DocParamValue> titleList = (List<DocParamValue>) value;
            for(int i = 0;i<titleList.size();i++) {
                DocParamValue paramValue = titleList.get(i);
                Integer curTitleLevel = paramValue.getTitleLevel();
                if(curTitleLevel == null) {
                    curTitleLevel = titleLevel;
                }
                String curValue = (String) paramValue.getValue();
                if(i>0) {
                    Paragraph insertParagraph = builder.insertParagraph();
                    builder.moveTo(insertParagraph);
                }
                writeTitle(builder,curTitleLevel,curValue);
            }
        }else if(value instanceof String) {
            String valueStr = (String) value;
            writeTitle(builder,titleLevel,valueStr);
        }
        //删除原来空白的标题行
        return ReplaceAction.REPLACE;
    }

    /**
     * 写标题
     * @param builder
     * @param titleLevel
     * @param value
     * @throws Exception
     */
    private void writeTitle(DocumentBuilder builder, int titleLevel, String value) throws Exception {
        //计算标题的字体大小，最大值为24，最小值为12
        int maxSize = 24;
        int curFontSize = maxSize-2*(titleLevel-1);
        if(curFontSize<12) {
            curFontSize = 12;
        }
        //设置大纲级别
        builder.getParagraphFormat().setOutlineLevel(titleLevel-1);
        //设置段后距离
        builder.getParagraphFormat().setSpaceAfter(6);
        //设置加粗
        builder.getFont().setBold(true);
        //设置字体大小
        builder.getFont().setSize(curFontSize);
        //写入值
        builder.write(value);
    }
}
