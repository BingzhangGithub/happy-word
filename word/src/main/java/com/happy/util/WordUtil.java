package com.happy.util;

import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import com.happy.bean.*;
import com.sun.org.slf4j.internal.Logger;
import com.sun.org.slf4j.internal.LoggerFactory;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.jsoup.nodes.Node;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
* word util
* @author Luck
* @date 2023年06月20日 20:21
*/
public class WordUtil {

    private static final Logger log = LoggerFactory.getLogger(WordUtil.class);

    /**
     * 英文公制单位
     * 表示1英寸的英文公制单位
     */
    private static final Double EMU_UNIT = 914400D;
    /**
     * 像素单位
     * 表示1英寸的像素单位
     */
    private static final Double PIXEL_UNIT = 72D;
    /**
     * WORD 单元格的宽度单位
     * 表示1英寸为1440 twips
     */
    private static final Double WORD_CELL_WIDTH_TWIPS_UNIT = 1440D;

    /**
     * @param ctGraphicalObject 图片数据
     * @param deskFileName      图片描述
     * @param width             宽
     * @param height            高
     * @param leftOffset        水平偏移 left
     * @param topOffset         垂直偏移 top
     * @param behind            文字上方，文字下方
     * @return 浮动属性对象
     */
    private static CTAnchor getAnchorWithGraphic(CTGraphicalObject ctGraphicalObject,
                                                 String deskFileName, int width, int height,
                                                 int leftOffset, int topOffset, boolean behind) {
        String anchorXml =
                "<wp:anchor xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                        + "simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"" + ((behind) ? 1 : 0) + "\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
                        + "<wp:simplePos x=\"0\" y=\"0\"/>"
                        + "<wp:positionH relativeFrom=\"column\">"
                        + "<wp:posOffset>" + leftOffset + "</wp:posOffset>"
                        + "</wp:positionH>"
                        + "<wp:positionV relativeFrom=\"paragraph\">"
                        + "<wp:posOffset>" + topOffset + "</wp:posOffset>" +
                        "</wp:positionV>"
                        + "<wp:extent cx=\"" + width + "\" cy=\"" + height + "\"/>"
                        + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
                        + "<wp:wrapNone/>"
                        + "<wp:docPr id=\"1\" name=\"Drawing 0\" descr=\"" + deskFileName + "\"/><wp:cNvGraphicFramePr/>"
                        + "</wp:anchor>";

        CTDrawing drawing;
        CTAnchor anchor = null;
        try {
            drawing = CTDrawing.Factory.parse(anchorXml);
            anchor = drawing.getAnchorArray(0);
            anchor.setGraphic(ctGraphicalObject);
        } catch (XmlException e) {
            log.error("error:", e);
        }
        return anchor;
    }

    /**
     * 插入图片信息
     *
     * @param imageInputStreamMap 图片对象map
     * @param ctBookmark          段落书签名称
     * @param xwpfParagraph       段落信息
     * @throws Exception 异常
     */
    private static void insertImages(Map<String, List<WordImageBean>> imageInputStreamMap, CTBookmark ctBookmark, XWPFParagraph xwpfParagraph) throws Exception {
        if (ObjectUtil.isEmpty(imageInputStreamMap)) {
            return;
        }
        String ctBookmarkName = ctBookmark.getName();
        List<WordImageBean> wordImageBeanList = imageInputStreamMap.get(ctBookmarkName);
        String paraXml = xwpfParagraph.getCTP().xmlText();
        if (ObjectUtil.isEmpty(wordImageBeanList)) {
            if (!imageInputStreamMap.containsKey(ctBookmarkName)){
                return;
            }
            // 图片数据无值时置空当前书签值
            String regex = "<w:bookmarkStart w:id=\"..{0,1}\" w:name=\"" + ctBookmarkName + "\"/>(.*?)<w:bookmarkEnd w:id=\"..{0,1}\"/>";
            List<String> bookmarkXmlList = PatternUtil.matcherToList(regex, paraXml);
            if (ObjectUtil.isNotEmpty(bookmarkXmlList)){
                for (String bookmarkXml : bookmarkXmlList) {
                    String replace = bookmarkXml.replaceAll("<w:t.*?>(.*?)</w:t>", "<w:t></w:t>");
                    paraXml = paraXml.replaceFirst(bookmarkXml, replace);
                }
            }
            xwpfParagraph.getCTP().set(XmlObject.Factory.parse(paraXml));
            return;
        }
        String bookmarkRegex = "<w:bookmarkStart w:id=\"..{0,1}\" w:name=\"" + ctBookmarkName + "\"/>(.*?)<w:bookmarkEnd w:id=\"..{0,1}\"/>";
        List<String> bookmarkXmlList = PatternUtil.matcherToList(bookmarkRegex, paraXml);
        int bookmarkStartIndex = -1;
        int bookmarkEndIndex = -1;
        if (ObjectUtil.isNotEmpty(bookmarkXmlList)){
            for (String bookmarkXml : bookmarkXmlList) {
                bookmarkStartIndex = paraXml.indexOf(bookmarkXml);
                List<String> bookmarkGroups = PatternUtil.matcherToList("<w:bookmarkEnd w:id=\"..{0,1}\"/>", bookmarkXml);
                if (ObjectUtil.isNotEmpty(bookmarkGroups)) {
                    bookmarkEndIndex = paraXml.indexOf(bookmarkGroups.get(0));
                }
            }
        }
        String regex = "<w:r>(.*?)</w:r>";
        List<String> runXmlList = PatternUtil.matcherToList(regex, paraXml);
        int tempBeginIndex = -1;
        int tempEndIndex = -1;
        if (ObjectUtil.isNotEmpty(runXmlList)){
            int formIndex = 0;
            for (int i = 0; i < runXmlList.size(); i++) {
                String runXml = runXmlList.get(i);
                int runIndex = paraXml.indexOf(runXml, formIndex);
                if (bookmarkStartIndex < runIndex) {
                    if (tempBeginIndex == -1) {
                        tempBeginIndex = i;
                    }
                    if (bookmarkEndIndex < runIndex){
                        tempEndIndex = i - 1;
                        break;
                    } else if((i == runXmlList.size()-1)){
                        tempEndIndex = i;
                    }
                }
                formIndex = formIndex + runXml.length();
            }
        }
        XWPFRun run;
        if(tempBeginIndex == -1){
            run = xwpfParagraph.createRun();
        }else {
            if (tempEndIndex > tempBeginIndex) {
                for (int i = tempBeginIndex; i <= tempEndIndex; i++) {
                    xwpfParagraph.removeRun(tempBeginIndex);
                }
            } else {
                xwpfParagraph.removeRun(tempBeginIndex);
            }
            run = xwpfParagraph.insertNewRun(tempBeginIndex);
        }
        int leftOffset = 0;
        int imageNumber = 0;
        for (WordImageBean wordImageBean : wordImageBeanList) {
            // 定义图片类型
            int picType = XWPFDocument.PICTURE_TYPE_JPEG;
            if (wordImageBean.getName().toLowerCase().endsWith(".png")) {
                picType = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (wordImageBean.getName().toLowerCase().endsWith(".gif")) {
                picType = XWPFDocument.PICTURE_TYPE_GIF;
            }
            // 处理图片宽高
            int width = 100;
            int height = 70;
            //fileName为鼠标在word里选择图片时，图片显示的名字，Width，Height则为像素单元，根据实际需要的大小进行调整即可。
            File imageFile = wordImageBean.getImageFile();
            try (FileInputStream inputStream = new FileInputStream(imageFile)) {
                run.addPicture(inputStream, picType, wordImageBean.getName(),
                        Units.toEMU(Double.parseDouble(width + "")),
                        Units.toEMU(Double.parseDouble(height + "")));
            }
            //设置图片显示属性 0 显示文字 1 嵌入式方式插入图片 2 浮于文字上方 3 浮于文字下方
            if (wordImageBean.getBehind() > 1) {
                if (imageNumber > 0) {
                    leftOffset += width;
                }
                CTDrawing drawing = run.getCTR().getDrawingArray(0);
                CTGraphicalObject geographical = drawing.getInlineArray(0).getGraphic();
                Random random = new Random();
                //拿到新插⼊的图⽚替换添加CTAnchor 设置浮动属性 删除inline属性
                int number = random.nextInt(999) + 1;
                //相对当前段落位置及偏移
                CTAnchor anchor1 = getAnchorWithGraphic(geographical, "Seal" + number, Units.toEMU(width), Units.toEMU(height), Units.toEMU(leftOffset), Units.toEMU(0), wordImageBean.getBehind() == 3);
                //添加浮动属性
                drawing.setAnchorArray(new CTAnchor[]{anchor1});
                //删除⾏内属性//添加签名图⽚
                drawing.removeInline(0);
            }
            imageNumber++;
        }
    }

    /**
     * 解析word书签
     *
     * @param inputStream 输入word文件流
     * @param wordBean    word操作对象
     * @return 解析后的文件流
     * @throws IOException IO异常
     */
    public static ByteArrayOutputStream parseBookmarkOfWord(InputStream inputStream, WordBean wordBean) throws Exception {
        Map<String, String> bookmarkMap = wordBean.getBookmarkMap();
        Map<String, List<WordImageBean>> imageInputStreamMap = wordBean.getImageInputStreamMap();
        ByteArrayOutputStream byteArrayOutputStream = null;
        ByteArrayOutputStream imageByteOutputStream = null;
        XWPFDocument htXwpfDocument;
        try {
            htXwpfDocument = new XWPFDocument(inputStream);
            // 获取word完整的xml结构内容
            String htXwpfDocumentXmlText = htXwpfDocument.getDocument().xmlText();
            // 替换普通书签
            htXwpfDocumentXmlText = replaceTextBookmark(htXwpfDocumentXmlText, bookmarkMap);
            // 重置word中的文件xml信息
            htXwpfDocument.getDocument().set(XmlObject.Factory.parse(htXwpfDocumentXmlText));
            //替换页眉页脚
            for (int i = 0; i < htXwpfDocument.getFooterList().size(); i++) {
                XWPFFooter htXwpfFooterXml = htXwpfDocument.getFooterList().get(i);
                String htXwpfFooterXmlString = htXwpfFooterXml._getHdrFtr().xmlText();
                htXwpfFooterXmlString = replaceTextBookmark(htXwpfFooterXmlString, bookmarkMap);
                htXwpfFooterXml._getHdrFtr().set(XmlObject.Factory.parse(htXwpfFooterXmlString));
            }
            // 替换图片信息
            if (ObjectUtil.isNotEmpty(imageInputStreamMap)) {
                imageByteOutputStream = new ByteArrayOutputStream();
                htXwpfDocument.write(imageByteOutputStream);
                htXwpfDocument = new XWPFDocument(new ByteArrayInputStream(imageByteOutputStream.toByteArray()));
                //替换图片
                replaceImageBookmark(htXwpfDocument, imageInputStreamMap);
                FileUtil.closeStream(imageByteOutputStream);
            }

            // 替换明细表数据
            replaceBmTableData(htXwpfDocument, wordBean.getBmBeanTableDataMap());
            byteArrayOutputStream = new ByteArrayOutputStream();
            htXwpfDocument.write(byteArrayOutputStream);
            htXwpfDocument.close();
            return byteArrayOutputStream;
        } finally {
            if(ObjectUtil.isNotEmpty(imageInputStreamMap)){
                for (List<WordImageBean> imageBeans : imageInputStreamMap.values()) {
                    if (ObjectUtil.isNotEmpty(imageBeans)){
                        imageBeans.forEach((imageBean)->FileUtil.deleteFile("delete wordImageFile error", imageBean.getImageFile()));
                    }
                }
            }
            if (ObjectUtil.isNotEmpty(wordBean.getBmBeanTableDataMap())){
                for (WordBmBean bmBean : wordBean.getBmBeanTableDataMap().values()) {
                    Map<Integer, WordBmBean> bamBeanMap = bmBean.getBamBeanMap();
                    if (ObjectUtil.isNotEmpty(bamBeanMap)){
                        for (WordBmBean wordBmBean : bamBeanMap.values()) {
                            if (ObjectUtil.isNotEmpty(wordBmBean)) {
                                List<WordImageBean> imageBeans = wordBmBean.getImageList();
                                if (ObjectUtil.isNotEmpty(imageBeans)) {
                                    imageBeans.forEach((imageBean) -> FileUtil.deleteFile("delete wordImageFile error", imageBean.getImageFile()));
                                }
                            }
                        }
                    }
                }
            }
            FileUtil.closeStream(byteArrayOutputStream, imageByteOutputStream);
        }
    }

    /**
     * 替换书签表格数据
     * @param htXwpfDocument word对象
     * @param bmBeanMap 书签Map
     * @throws Exception 异常
     */
    private static void replaceBmTableData(XWPFDocument htXwpfDocument, Map<String, WordBmBean> bmBeanMap) throws Exception{
        // 替换表格字段、表格图片以及表格富文本数据
        if (ObjectUtil.isEmpty(bmBeanMap)) {
            return;
        }
        // 根文件
        String rootXml = htXwpfDocument.getDocument().xmlText();
        // 读取表格
        List<String> tableGroups = PatternUtil.matcherToList("<w:tbl>([\\s\\S]*?)</w:tbl>", rootXml);
        for (String tableGroup : tableGroups) {
            // 校验是否有嵌套表格  有则不处理当前表格数据
            int lastIndexOfTable = tableGroup.lastIndexOf("<w:tbl>");
            if (lastIndexOfTable > 0) {
                continue;
            }
            List<String> rowGroups = PatternUtil.matcherToList("<w:tr.*?>([\\s\\S]*?)</w:tr>", tableGroup);
            for (String rowGroup : rowGroups) {
                // 声明当前行需要追加的最大值
                long currentRowMaxIndex = 0L;
                // 记录每个单元格的书签和cellXml
                List<Map<String, String>> cellBookmarkXmlMapList = new ArrayList<>();
                // 记录当前行的整体cell xml数据
                StringBuilder currentRowCellXml = new StringBuilder();
                // 记录cell和书签信息，并记录书签的最大值信息
                List<String> cellGroups = PatternUtil.matcherToList("<w:tc>([\\s\\S]*?)</w:tc>", rowGroup);
                for (String cellGroup : cellGroups) {
                    currentRowCellXml.append(cellGroup);
                    String bookmarkName = "";
                    List<String> bookmarkGroups = PatternUtil.matcherToList("<w:bookmarkStart .*?/>", cellGroup);
                    for (String bookmarkGroup : bookmarkGroups) {
                        bookmarkName = getPropValueOfClosingTag("w:bookmarkStart", bookmarkGroup, "w:name");
                        if ("_GoBack".equals(bookmarkName) || "".equals(bookmarkName)){
                            continue;
                        }
                        // 处理明细书签对应的row index
                        if (bmBeanMap.containsKey(bookmarkName)) {
                            WordBmBean wordBmBean = bmBeanMap.get(bookmarkName);
                            if (wordBmBean.isTableData() && ObjectUtil.isNotEmpty(wordBmBean.getBamBeanMap())){
                                Set<Integer> integers;
                                integers = wordBmBean.getBamBeanMap().keySet();
                                if (ObjectUtil.isNotEmpty(integers)){
                                    for (Integer index : integers) {
                                        currentRowMaxIndex = currentRowMaxIndex >= index ? currentRowMaxIndex : index;
                                    }
                                }
                            }
                        }
                        break;
                    }
                    // 当前单元格的书签和cellXml信息
                    Map<String, String> cellBookmarkXmlMap = new HashMap<>(8);
                    cellBookmarkXmlMap.put("bookmarkName", bookmarkName);
                    cellBookmarkXmlMap.put("cellXml", cellGroup);
                    cellBookmarkXmlMapList.add(cellBookmarkXmlMap);
                }
                // 新增所有行的最终数据
                StringBuilder targetXml = new StringBuilder();
                boolean initSerialNumber = false;
                if (currentRowMaxIndex == 0) {
                    currentRowMaxIndex = 1;
                    initSerialNumber = true;
                }
                // 逐行增加数据
                for (int i = 0; i < currentRowMaxIndex; i++) {
                    // 当前行的所有cell数据
                    StringBuilder targetRowCellXml = new StringBuilder();
                    for (Map<String, String> cellBookmarkXmlMap : cellBookmarkXmlMapList) {
                        String bookmarkName = cellBookmarkXmlMap.get("bookmarkName");
                        String cellXml = cellBookmarkXmlMap.get("cellXml");
                        if (ObjectUtil.isEmpty(bookmarkName) || !bmBeanMap.containsKey(bookmarkName) || !bmBeanMap.get(bookmarkName).isTableData()) {
                            // 保留第一行数据
                            if (i == 0) {
                                cellXml = verticalMergeCell(true, cellXml);
                                targetRowCellXml.append(cellXml);
                            } else {
                                cellXml = verticalMergeCell(false, cellXml);
                                // 置空后续行数据
                                cellXml = replaceStringCellValue(cellXml, "");
                                targetRowCellXml.append(cellXml);
                            }
                        } else {
                            if (bmBeanMap.containsKey(bookmarkName) && bmBeanMap.get(bookmarkName).isTableData()){
                                // 如果是明细字段数据
                                WordBmBean wordBmBean = bmBeanMap.get(bookmarkName);
                                if (ObjectUtil.isNotEmpty(wordBmBean)){
                                    if (wordBmBean.isImageType()){
                                        Map<Integer, WordBmBean> imageBmBeanMap = wordBmBean.getBamBeanMap();
                                        if (ObjectUtil.isEmpty(imageBmBeanMap)) {
                                            cellXml = replaceStringCellValue(cellXml, "");
                                            targetRowCellXml.append(cellXml);
                                            continue;
                                        }
                                        WordBmBean imageWordBmBean = imageBmBeanMap.get(i + 1);
                                        if (ObjectUtil.isEmpty(imageWordBmBean)) {
                                            cellXml = replaceStringCellValue(cellXml, "");
                                            targetRowCellXml.append(cellXml);
                                            continue;
                                        }
                                        // 当前逻辑处理图片型数据
                                        StringBuilder imageXml = new StringBuilder();
                                        for (WordImageBean wordImage : imageWordBmBean.getImageList()) {
                                            // 获取到cell的宽度
                                            Double cellWidth = Double.valueOf(getCellWidth(cellXml));
                                            // 获取图片型xml数据
                                            imageXml.append(getImageXml(wordImage, cellWidth, htXwpfDocument));
                                        }
                                        // 处理当前cellXml 生成具体的cell数据
                                        cellXml = replaceStringCellAsImageCellXml(cellXml, imageXml.toString());
                                        // 追加给当前行的xml
                                        targetRowCellXml.append(cellXml);
                                    } else {
                                        Map<Integer, WordBmBean> textBmBeanMap = wordBmBean.getBamBeanMap();
                                        StringBuilder builder = new StringBuilder();
                                        if (WordFieldType.SERIAL_NUMBER == wordBmBean.getFiledType()){
                                            if (!initSerialNumber){
                                                cellXml = replaceStringCellValue(cellXml, (i + 1) + "");
                                                // 追加给当前行的xml
                                                targetRowCellXml.append(cellXml);
                                                continue;
                                            }
                                        }
                                        if (ObjectUtil.isEmpty(textBmBeanMap)){
                                            cellXml = replaceStringCellValue(cellXml, builder.toString());
                                            // 追加给当前行的xml
                                            targetRowCellXml.append(cellXml);
                                            continue;
                                        }
                                        WordBmBean wordTexts = textBmBeanMap.get(i + 1);
                                        if (ObjectUtil.isNotEmpty(wordTexts)) {
                                            for (String value : wordTexts.getValues()) {
                                                builder.append(value == null ? "" : value);
                                            }
                                        }
                                        // 处理字符型数据
                                        // 处理当前cellXml 生成具体的cell文本数据
                                        cellXml = replaceStringCellValue(cellXml, builder.toString());
                                        // 追加给当前行的xml
                                        targetRowCellXml.append(cellXml);
                                    }
                                }else {
                                    cellXml = replaceStringCellValue(cellXml, "");
                                    targetRowCellXml.append(cellXml);
                                }
                            }
                        }
                    }
                    // 替换当前行的cell数据为新的cell数据，得到新的行数据
                    String newRowXml = rowGroup.replace(currentRowCellXml, targetRowCellXml);
                    // 追加在最终的xml信息中
                    targetXml.append(newRowXml);
                }
                // 替换行数据
                if(ObjectUtil.isEmpty(targetXml)){
                    continue;
                }
                // 替换当前行数据为最终的cell数据
                rootXml = rootXml.replace(rowGroup, targetXml);
            }
        }
        htXwpfDocument.getDocument().set(XmlObject.Factory.parse(rootXml));
    }

    /**
     * 替换图片
     *
     * @param xwpfDocument        word文档
     * @param imageInputStreamMap 图片输入流
     */
    private static void replaceImageBookmark(XWPFDocument xwpfDocument, Map<String, List<WordImageBean>> imageInputStreamMap) {
        if (ObjectUtil.isEmpty(imageInputStreamMap)) {
            return;
        }

        log.warn("--ConvertToDocxTest----插入图片--replaceInPara()--imgInfoList = " + imageInputStreamMap);
        try {
            //纯文字，不含表格 处理插入图片
            insertImagesIntoWord(xwpfDocument.getParagraphs(), imageInputStreamMap);

            //处理表格中的图片
            List<XWPFTable> xwpfTableList = xwpfDocument.getTables();
            for (XWPFTable xwpfTable : xwpfTableList) {
                List<XWPFTableRow> rows = xwpfTable.getRows();
                for (XWPFTableRow xwpfTableRow : rows) {
                    List<XWPFTableCell> cells = xwpfTableRow.getTableCells();
                    for (XWPFTableCell xwpfTableCell : cells) {
                        insertImagesIntoWord(xwpfTableCell.getParagraphs(), imageInputStreamMap);
                    }
                }
            }
        } catch (FileNotFoundException e) {
            log.error("--ConvertToDocxTest----插入图片--replaceInPara()--FileNotFoundException异常：" + e.getMessage(), e);
        } catch (IOException e) {
            log.error("--ConvertToDocxTest----插入图片--replaceInPara()--IOException异常：" + e.getMessage(), e);
        } catch (Exception e) {
            log.error("--ConvertToDocxTest----插入图片--replaceInPara()--异常：" + e.getMessage(), e);
        }
    }

    /**
     * fill picture into word file , notice : This operation points to non-tabular data.
     * @param xwpfDocument word Obj
     * @param imageInputStreamMap image inputStream
     * @throws Exception some exception, forgive me, I don't know what exception this operation will throw.
     */
    private static void insertImagesIntoWord(List<XWPFParagraph> xwpfDocument, Map<String, List<WordImageBean>> imageInputStreamMap) throws Exception {
        for (XWPFParagraph xwpfParagraph : xwpfDocument) {
            CTP ctp = xwpfParagraph.getCTP();
            for (int dwI = 0; dwI < ctp.sizeOfBookmarkStartArray(); dwI++) {
                CTBookmark bookmark = ctp.getBookmarkStartArray(dwI);
                //处理图片
                insertImages(imageInputStreamMap, bookmark, xwpfParagraph);
            }
        }
    }

    /**
     * 替换文本书签方法
     *
     * @param xwpfDocumentXmlText 需要填充的文件流xml
     * @param bookmarkMap         书签内容 key: 书签名称， value: 书签内容
     * @return word的xml文本对象
     */
    private static String replaceTextBookmark(String xwpfDocumentXmlText, Map<String, String> bookmarkMap) {
        if (ObjectUtil.isEmpty(bookmarkMap)) {
            return xwpfDocumentXmlText;
        }
        String regEx = "<w:bookmarkStart .*?/>";
        Pattern pattern = Pattern.compile(regEx);
        Matcher matcher = pattern.matcher(xwpfDocumentXmlText);
        while (matcher.find()) {
            String groupStr = matcher.group();
            String bookmarkId = getPropValueOfClosingTag("w:bookmarkStart", groupStr, "w:id");
            String bookmarkName = getPropValueOfClosingTag("w:bookmarkStart", groupStr, "w:name");
            String value = bookmarkMap.get(bookmarkName);
            if (value == null) {
                continue;
            }
            String wholeBookmarkXmlRegEx = "<w:bookmarkStart w:id=\"" + bookmarkId + "\" w:name=\"" + bookmarkName + "\"/>.*?<w:bookmarkEnd w:id=\"" + bookmarkId + "\"/>";
            Pattern bookmarkXmlPattern = Pattern.compile(wholeBookmarkXmlRegEx);
            Matcher bookmarkXmlMatcher = bookmarkXmlPattern.matcher(xwpfDocumentXmlText);
            if (bookmarkXmlMatcher.find()) {
                String bookmarkPstr = bookmarkXmlMatcher.group();

                String endBookmarkXmlRegEx = "<w:bookmarkStart w:id=\"" + bookmarkId + "\" w:name=\"" + bookmarkName + "\"/>(<w:bookmarkStart w:id=\".*\" w:name=\"_GoBack\"/>)?<w:bookmarkEnd w:id=\"" + bookmarkId + "\"/>";
                Pattern endBookmarkXmlPattern = Pattern.compile(endBookmarkXmlRegEx);
                Matcher endBookmarkXmlMatcher = endBookmarkXmlPattern.matcher(bookmarkPstr);
                if (endBookmarkXmlMatcher.find()) {
                    // 为最后的书签标志增加字体内容
                    xwpfDocumentXmlText = xwpfDocumentXmlText.replace("<w:bookmarkEnd w:id=\"" + bookmarkId + "\"/>", "<w:r><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:hAnsi=\"仿宋_GB2312\" w:cs=\"仿宋_GB2312\" w:hint=\"eastAsia\"/><w:szCs w:val=\"32\"/></w:rPr><w:t>" + value + "</w:t></w:r><w:bookmarkEnd w:id=\"" + bookmarkId + "\"/>");
                } else {
                    // 替换书签内容
                    String regEx2 = "<w:t>.*?</w:t>";
                    Pattern pattern2 = Pattern.compile(regEx2);
                    Matcher matcher2 = pattern2.matcher(bookmarkPstr);
                    boolean hasMatched = false;
                    while (matcher2.find()) {
                        hasMatched = true;
                        String bv = matcher2.group();
                        bookmarkPstr = bookmarkPstr.replaceAll(bv, "<w:t>" + value + "</w:t>");
                        value = "";
                    }
                    if (!hasMatched) {

                        String regEx4 = "<w:t .*?</w:t>";
                        Pattern pattern4 = Pattern.compile(regEx4);
                        Matcher matcher4 = pattern4.matcher(bookmarkPstr);
                        while (matcher4.find()) {
                            String bv = matcher4.group();
                            bookmarkPstr = bookmarkPstr.replaceAll(bv, "<w:t>" + value + "</w:t>");
                            value = "";
                        }
                    }
                    xwpfDocumentXmlText = xwpfDocumentXmlText.replace(bookmarkXmlMatcher.group(), bookmarkPstr);
                }
            }
        }
        return xwpfDocumentXmlText;
    }
    
    /**
     * 根据cell的Xml片段获取cell宽度数据
     * @param cellXml cellXml片段
     * @return String的cell宽度数据
     */
    private static String getCellWidth(String cellXml){
        String cellWidthRegEx = "<w:tcW .*?/>";
        Pattern cellWidthPattern = Pattern.compile(cellWidthRegEx);
        Matcher cellWidthMatcher = cellWidthPattern.matcher(cellXml);
        // 给定默认cell宽度，2130为一行四列的列宽
        String defaultCellWidth = "2130";
        if (cellWidthMatcher.find()) {
            String cellWidthGroup = cellWidthMatcher.group();
            String w = getPropValueOfClosingTag("w:tcW",cellWidthGroup, "w:w");
            defaultCellWidth = w == null ? defaultCellWidth : w;
        }
        return defaultCellWidth;
    }

    /**
     * 获取image的尺寸
     * @param wordImage word图片Bean
     * @param cellWidth 单元格宽度
     * @param htXwpfDocument doc数据
     * @return image 的xml数据
     * @throws Exception 异常
     */
    public static String getImageXml(WordImageBean wordImage, Double cellWidth, XWPFDocument htXwpfDocument) throws Exception {
        String imageName = wordImage.getName();
        Double imageWidthEmu = getImageWidthEmu(cellWidth);
        BufferedImage image = ImageIO.read(wordImage.getImageFile());
        double imageWidth = image.getWidth();
        double imageHeight = image.getHeight();
        int pictureType = WordPictureUtil.findPictureType(FileUtil.getFileType(wordImage.getImageFile()));
        String imageRefId = htXwpfDocument.addPictureData(Files.newInputStream(wordImage.getImageFile().toPath()), pictureType);
        int imageId = htXwpfDocument.getNextPicNameNumber(pictureType);
        Double imageHeightEmu = getImageHeightEmu(cellWidth, imageWidth, imageHeight);
        return generateImageXml(imageWidthEmu + "", imageHeightEmu + "", imageName, imageRefId, imageId);
    }

    /**
     * 获取图片的宽度
     * 单位EMU (English Metric Unit) 英文公制单位
     * 1EMU等于1/914400英寸，用于处理图片的宽高属性（cx， cy）
     * @param cellWidth 图片的宽度
     * @return 图片英文公制单位宽度
     */
    private static Double getImageWidthEmu(Double cellWidth){
        Double cellWidthInch = cellWidth / WORD_CELL_WIDTH_TWIPS_UNIT;
        return cellWidthInch * EMU_UNIT;
    }

    /**
     * 获取图片高度
     * 单位EMU  参考方法 getImageWidthEmu(Double cellWidth)
     * @param cellWidth 单元格宽度
     * @param imageWidth 图片宽度
     * @param imageHeight 图片高度
     * @return 图片高度英制公制单位
     */
    private static Double getImageHeightEmu(Double cellWidth, Double imageWidth, Double imageHeight){
        Double cellWidthInch = cellWidth / WORD_CELL_WIDTH_TWIPS_UNIT;
        Double imageWidthInch = imageWidth / PIXEL_UNIT;
        Double scaling = imageWidthInch/cellWidthInch;
        Double imageHeightInch = imageHeight / PIXEL_UNIT;
        Double cellHeightInch = imageHeightInch / scaling;
        // 图片高度单位
        return cellHeightInch * EMU_UNIT;
    }

    /**
     * 匹配单个闭合标签元素属性值
     * @param element 匹配元素
     * @param source 来源字符串
     * @param attr 字符属性
     * @return 属性值
     */
    private static String getPropValueOfClosingTag(String element, String source, String attr) {
        if (ObjectUtil.isEmpty(source)){
            return "";
        }
        source = source.replaceAll("/>", " />");
        source = source.replaceAll(">", " >");
        List<String> result = new ArrayList<>();
        String reg = "<" + element + "[^<>]*?\\s" + attr + "=['\"]?(.*?)['\"]?\\s.*?>";
        Matcher m = Pattern.compile(reg).matcher(source);
        while (m.find()) {
            String r = m.group(1);
            result.add(r);
        }
        return result.size() > 0 ? result.get(0) : "";
    }

    /**
     * 解析html字体加粗问题
     *
     * @param node         元素节点
     * @param xmlstr       xml字符串
     * @param parentCssMap Css映射Map
     * @return 解析后的字符串
     */
    public static String getWordXml(Node node, String xmlstr, Map<String, String> parentCssMap) {
        List<Node> elements = node.childNodes();
        for (Node element : elements) {
            //样式传递控制
            Map<String, String> cssMap = new HashMap(parentCssMap);
            // 样式加粗处理
            if ("b".equalsIgnoreCase(node.nodeName()) || "strong".equalsIgnoreCase(node.nodeName())) {
                cssMap.put("bold", "true");
            }
            getTextStyle(node, cssMap);
            getStyle(element.attr("style"), cssMap);
            if ("#text".equalsIgnoreCase(element.nodeName())) {
                // 转换为word的xml数据
                xmlstr += transToDocxXml(cssMap, element.toString());
            }
            xmlstr = getWordXml(element, xmlstr, cssMap);
        }
        return xmlstr;
    }

    /**
     * 获取段落居中居左居右样式
     *
     * @param node   node元素
     * @param cssMap 样式Map
     */
    public static void getTextStyle(Node node, Map<String, String> cssMap) {
        String nodeStyle = StrUtil.toStringOrNull(node.attributes().get("style"));
        // 居右标志
        String rightMark = "text-align:right";
        // 居中标志
        String centerMark = "text-align:center";
        // 首行缩进标志
        String textIndent = "text-indent";
        // 字号
        String fontSize = "font-size";
        // 行高
        String lineHeight = "line-height";
        if ((rightMark).equals(nodeStyle)) {
            //居右
            cssMap.put("align", "right");
        } else if ((centerMark).equals(nodeStyle)) {
            //居中
            cssMap.put("align", "center");
        }
        if (nodeStyle.contains(textIndent)) {
            //首行缩进
            cssMap.put(textIndent, "true");
        }
        //字体大小
        if (nodeStyle.contains(fontSize)) {
            String tempStr = nodeStyle.substring(nodeStyle.indexOf("font-size") + "font-size".length());
            tempStr = tempStr.replace(":", "");
            tempStr = tempStr.replace("px", "");
            tempStr = tempStr.trim();
            cssMap.put("font_size", tempStr);
        }
        //段落行距
        if (nodeStyle.contains(lineHeight)) {
            String tempStr = nodeStyle.substring(nodeStyle.indexOf("line-height") + "line-height".length());
            tempStr = tempStr.replace(":", "");
            tempStr = tempStr.trim();
            cssMap.put("line_height", tempStr);
        }
    }

    /**
     * 获取字体样式
     *
     * @param cssStyle 样式字符串
     * @param cssMap   样式Map
     */
    public static void getStyle(String cssStyle, Map<String, String> cssMap) {
        String cssName = "font-family";
        if (cssStyle.contains(cssName)) {
            String tempStr = cssStyle.substring(cssStyle.indexOf(cssName) + cssName.length());
            String separator = ";";
            if (tempStr.contains(separator)) {
                tempStr = tempStr.substring(0, tempStr.indexOf(";"));
            }
            tempStr = tempStr.replace(":", "");
            tempStr = tempStr.trim();
            cssMap.put("font", tempStr);
        }
    }

    /**
     * 增加一个内嵌的图片数据
     * @param imageWidthEmu 图片宽度
     * @param imageHeightEmu 图片高度
     * @param imageNameDesc 图片描述，接入图片名称
     * @param imageRefId 图片引用id
     * @return imageXml数据
     */
    public static String generateImageXml(String imageWidthEmu, String imageHeightEmu, String imageNameDesc, String imageRefId, int imageId){
        String imageRpr = "<w:p><w:pPr><w:rPr><w:rFonts w:hint=\"default\"/><w:vertAlign w:val=\"baseline\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"default\"/><w:vertAlign w:val=\"baseline\"/></w:rPr>";
        String drawingXml = "<w:drawing><wp:inline distT=\"114300\" distB=\"114300\" distL=\"114300\" distR=\"114300\"><wp:extent cx=\""+imageWidthEmu+"\" cy=\""+imageHeightEmu+"\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:docPr id=\""+imageId+"\" name=\""+"图片 "+imageId+"\" descr=\""+imageNameDesc+"\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\""+imageId+"\" name=\"图片 "+imageId+"\" descr=\""+imageNameDesc+"\"/><pic:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=\""+imageRefId+"\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\""+imageWidthEmu+"\" cy=\""+imageHeightEmu+"\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>";
        String endMark = "</w:r></w:p>";
        return imageRpr + drawingXml +endMark;
    }

    /**
     * 用于html多行文本替换书签带字体样式
     *
     * @param cssMap cssMap
     * @param text   待处理文本
     * @return 处理后的文本
     */
    private static String transToDocxXml(Map<String, String> cssMap, String text) {
        String xmlStr = "";
        int fontSizeValue = 32;
        text = toHtmlMode(text);
        String fontSize = "font_size";
        if (cssMap.get(fontSize) != null) {
            fontSizeValue = NumberUtil.parseInt(cssMap.get("font_size"), -1) * 2;
            int tempFontSize = 24;
            if (fontSizeValue < tempFontSize) {
                fontSizeValue = 32;
            }
        }
        //处理段落样式
        String textIndent = "text_indent";
        String align = "align";
        String lineHeight = "line_height";
        if (cssMap.get(textIndent) != null || cssMap.get(align) != null || cssMap.get(lineHeight) != null) {
            xmlStr += "<w:pPr>";
            if (cssMap.get(lineHeight) != null) {
                float lineHeightValue = Float.parseFloat(cssMap.get("line_height"));
                int line = NumberUtil.parseInt(lineHeightValue * 240 + "", -1);
                xmlStr += "<w:spacing w:line=\"" + line + "\" w:lineRule=\"auto\"/>";
            }
            if (cssMap.get(textIndent) != null) {
                xmlStr += "<w:ind w:right=\"210\" w:firstLineChars=\"200\" w:firstLine=\"420\"/>";
            }
            String tempAlignRight = "right";
            String tempAlignCenter = "center";
            if ((tempAlignRight).equals(cssMap.get(align))) {
                xmlStr += "<w:jc w:val=\"right\"/>";
            } else if ((tempAlignCenter).equals(cssMap.get(align))) {
                xmlStr += "<w:jc w:val=\"center\"/>";
            } else if (cssMap.get(textIndent) != null) {
                xmlStr += "<w:jc w:val=\"left\"/>";
            }
            xmlStr += "</w:pPr>";
        }

        xmlStr += "<w:r w:rsidRPr=\"00422646\">";
        xmlStr += "<w:rPr>";
        String font = "font";
        if (cssMap.get(font) != null) {
            String tempFontStr = cssMap.get("font");
            String fzFont = "FZXiaoBiaoSong-B05S";
            String fsFont = "FangSong_GB2312";
            String shFont = "SimHei";
            String ssFont = "SimSun";
            String ktFont = "KaiTi";
            String fontStr;
            //字体翻译
            if (fzFont.equalsIgnoreCase(tempFontStr)) {
                fontStr = "方正小标宋简体";
            } else if (fsFont.equalsIgnoreCase(tempFontStr)) {
                fontStr = "仿宋_GB2312";
            } else if (shFont.equalsIgnoreCase(tempFontStr)) {
                fontStr = "黑体";
            } else if (ssFont.equalsIgnoreCase(tempFontStr)) {
                fontStr = "宋体";
            } else if (ktFont.equalsIgnoreCase(tempFontStr)) {
                fontStr = "楷体";
            } else {
                fontStr = "仿宋_GB2312";
            }
            xmlStr += "<w:rFonts w:ascii=\"" + fontStr + "\" w:eastAsia=\"" + fontStr + "\" w:hAnsi=\"" + "微软雅黑" + "\" w:hint=\"eastAsia\"/>";
        } else {
            xmlStr += "<w:rFonts w:ascii=\"仿宋_GB2312\" w:eastAsia=\"仿宋_GB2312\" w:hAnsi=\"" + "微软雅黑" + "\" w:hint=\"eastAsia\"/>";
        }
        xmlStr += "<w:sz w:val=\"" + fontSizeValue + "\"/>";
        xmlStr += "<w:szCs w:val=\"" + fontSizeValue + "\"/>";
        String bold = "bold";
        if (cssMap.get(bold) != null) {
            xmlStr += "<w:b/>";
        }
        xmlStr += "</w:rPr>";
        String tempSeparator = " ";
        if (text.contains(tempSeparator)) {
            xmlStr += "<w:t xml:space=\"preserve\">" + text + "</w:t>";
        } else {
            xmlStr += "<w:t>" + text + "</w:t>";
        }
        xmlStr += "</w:r>";
        return xmlStr;
    }

    /**
     * 解析富文本特殊字符替换
     * @param s 富文本数据
     * @return 解析后数据
     */
    public static String toHtmlMode(String s) {
        return s.replaceAll("''", "'").
                replaceAll("&nbsp;", " ").
                replaceAll("&lt;", "<").
                replaceAll("&gt;", ">").
                replaceAll("&quot;", '"' + "")
                .replaceAll("<br>", "\n");
    }

    

    /**
     * 正则替换是时，转义特殊符号
     *
     * @return 替换后的数据
     */
    public static String replaceMatcher(String matcherGroup) {
        if ("".equals(matcherGroup)) {
            return matcherGroup;
        }
        matcherGroup = matcherGroup.replaceAll("\\(", "（");
        matcherGroup = matcherGroup.replaceAll("\\)", "）");
        matcherGroup = matcherGroup.replaceAll("\\{", "");
        matcherGroup = matcherGroup.replaceAll("}", "");
        matcherGroup = matcherGroup.replaceAll("\\[", "");
        matcherGroup = matcherGroup.replaceAll("]", "");

        return matcherGroup;
    }

    /**
     * 替换当前的单元格中文本数据
     * @param cellXml 当前单元格xml
     * @param value 待替换的数据
     */
    public static String replaceStringCellValue(String cellXml, String value) {
        // 基于run标签进行每个文本数据处理，每个run下面可能存在文本或者图片
        List<String> runGroups = PatternUtil.matcherToList("<w:r>(.*?)</w:r>", cellXml);
        if (ObjectUtil.isNotEmpty(runGroups)){
            for (int i = 0; i < runGroups.size(); i++) {
                List<String> textGroups = PatternUtil.matcherToList("<w:t.*?>(.*?)</w:t>", runGroups.get(i));
                if (ObjectUtil.isNotEmpty(textGroups)){
                    // 当前数据可能存在多个t标签，只给第一个t 标签赋值，其他t标签置空
                    for (String textGroup : textGroups) {
                        if (i == 0) {
                            cellXml = cellXml.replaceFirst(textGroup, "<w:t>" + value + "</w:t>");
                        } else {
                            cellXml = cellXml.replace(textGroup, "<w:t></w:t>");
                        }
                    }
                }
                List<String> drawingGroups = PatternUtil.matcherToList("<w:drawing.*?>(.*?)</w:drawing>", runGroups.get(i));
                if (ObjectUtil.isNotEmpty(drawingGroups)) {
                    for (String drawingGroup : drawingGroups) {
                        if (i == 0) {
                            cellXml = cellXml.replaceFirst(drawingGroup, "<w:t>" + value + "</w:t>");
                        }else {
                            cellXml = cellXml.replace(drawingGroup, "<w:t></w:t>");
                        }
                    }
                }
                List<String> closeTextGroups = PatternUtil.matcherToList("(<w:t\\b[^<]*?\\/>?)", runGroups.get(i));
                if (ObjectUtil.isNotEmpty(closeTextGroups)) {
                    for (String closeTextGroup : closeTextGroups) {
                        if (i == 0) {
                            cellXml = cellXml.replaceFirst(closeTextGroup, "<w:t>" + value + "</w:t>");
                        }
                    }
                }
            }
        }
        return cellXml;
    }


    /**
     * 将字符型的xml单元格数据替换为图片型的xml单元格数据
     * @param cellXml 原始单元格xml
     * @param imageXml 图片型xml数据
     */
    public static String replaceStringCellAsImageCellXml(String cellXml, String imageXml){
        // 匹配当前cellXml下的段落标签数据
        String textRegEx = "<w:p.*?>(.*?)</w:p>";
        Pattern textPattern = Pattern.compile(textRegEx);
        Matcher textMatcher = textPattern.matcher(cellXml);
        int flag = 0;
        while (textMatcher.find()) {
            String textGroup = textMatcher.group();
            // 当前数据可能存在多个t标签，第一个t 修改为图片数据，其他t标签置空
            if (flag > 0) {
                cellXml = cellXml.replace(textGroup, "");
            } else {
                cellXml = cellXml.replace(textGroup, imageXml);
            }
            flag++;
        }
        return cellXml;
    }

    /**
     * 垂直合并
     * @param isStart 是否为开始单元格
     * @param cellXml 当前操作的cell xml数据
     * @return 解析后的cellCml数据
     */
    public static String verticalMergeCell(boolean isStart, String cellXml){
        if (ObjectUtil.isEmpty(cellXml)) {
            return cellXml;
        }
        // 单元格属性根标签
        String cellPrRootLabel = "<w:tcPr>";
        if (isStart) {
            String startPr = "<w:tcPr><w:vMerge w:val=\"restart\"/>";
            cellXml = cellXml.replace(cellPrRootLabel, startPr);
        } else {
            String continuePr = "<w:tcPr><w:vMerge w:val=\"continue\"/>";
            cellXml = cellXml.replace(cellPrRootLabel, continuePr);
        }
        return cellXml;
    }
}
