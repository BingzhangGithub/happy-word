package com.happy.util;

import com.happy.bean.BmNonTableBean;
import com.happy.bean.BmTableBean;
import com.happy.bean.ValueType;
import lombok.NonNull;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlObject;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
* word util
* @author Luck
* @date 2023年06月20日 20:21
*/public class WordUtil {

    public static void convertToImage(){

    }
    /**
     * 1. 读取行，记录行
     * 2. 读取当前行的单元格中的每一个数据的书签 获取最大值并新增对应的行数
     * 3. 按照书签和行数进行取值处理数据，含有图片和文本
     * 4. 替换行数据，循环执行下一行
     * 5. 重置会word解析数据，写出文件
     * 6. 可能使用到正则标签： "<w:tbl>(?<content>.*?)</w:tbl>"   "<w:tbl>(?:(?!<w:tbl>).)*?</w:tbl>"
     * @param htXwpfDocument word 对象
     * @param bmTableMap word复制
     * @throws Exception 异常数据
     */
    public static void parseTableBmData(XWPFDocument htXwpfDocument, Map<String, BmTableBean> bmTableMap) throws Exception {
        // 替换表格字段、表格图片以及表格富文本数据
        if (bmTableMap == null) {
            return;
        }
        // 根文件
        String rootXml = htXwpfDocument.getDocument().xmlText();
        // 读取表格
        Matcher tableMatcher = getMatcherByRegEx("<w:tbl>(.*?)</w:tbl>", rootXml);
        while (tableMatcher.find()) {
            String tableGroup = tableMatcher.group();
            // 校验是否有嵌套表格  有则不处理当前表格数据
            int lastIndexOfTable = tableGroup.lastIndexOf("<w:tbl>");
            if (lastIndexOfTable > 0) {
                continue;
            }
            Matcher rowMatcher = getMatcherByRegEx("<w:tr>(.*?)</w:tr>", tableGroup);
            while (rowMatcher.find()) {
                // 声明当前行需要追加的最大值
                long currentRowMaxIndex = 0L;
                String rowGroup = rowMatcher.group();
                Matcher cellMatcher = getMatcherByRegEx("<w:tc>(.*?)</w:tc>", rowGroup);
                // 记录每个单元格的书签和cellXml
                List<Map<String, String>> cellBookmarkXmlMapList = new ArrayList<>();
                // 记录当前行的整体cell xml数据
                StringBuilder currentRowCellXml = new StringBuilder();
                // 记录cell和书签信息，并记录书签的最大值信息
                while (cellMatcher.find()) {
                    String cellGroup = cellMatcher.group();
                    currentRowCellXml.append(cellGroup);
                    Matcher bookmarkMatcher = getMatcherByRegEx("<w:bookmarkStart .*?/>", cellGroup);
                    String bookmarkName = "";
                    if (bookmarkMatcher.find()) {
                        String bookmarkGroup = bookmarkMatcher.group();
                        bookmarkName = getPrValueOfXmlTag( bookmarkGroup, "w:name");
                        // 处理明细书签对应的row index
                        if (bmTableMap.containsKey(bookmarkName)) {
                            BmTableBean bmTableBean = bmTableMap.get(bookmarkName);
                            int index = bmTableBean.getTotalRow();
                            currentRowMaxIndex = currentRowMaxIndex >= index ? currentRowMaxIndex : index;
                        }
                    }
                    // 当前单元格的书签和cellXml信息
                    Map<String, String> cellBookmarkXmlMap = new HashMap<>(8);
                    cellBookmarkXmlMap.put("bookmarkName", bookmarkName);
                    cellBookmarkXmlMap.put("cellXml", cellGroup);
                    cellBookmarkXmlMapList.add(cellBookmarkXmlMap);
                }
                // 新增所有行的最终数据
                StringBuilder targetXml = new StringBuilder();
                // 逐行增加数据
                for (int i = 0; i < currentRowMaxIndex; i++) {
                    // 当前行的所有cell数据
                    StringBuilder targetRowCellXml = new StringBuilder();
                    for (Map<String, String> cellBookmarkXmlMap : cellBookmarkXmlMapList) {
                        String bookmarkName = cellBookmarkXmlMap.get("bookmarkName");
                        String cellXml = cellBookmarkXmlMap.get("cellXml");
                        if (bookmarkName == null || !bmTableMap.containsKey(bookmarkName)) {
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
                            // 如果是明细字段数据
                            BmTableBean bmTableBean = bmTableMap.get(bookmarkName);
                            BmNonTableBean bmDataBean = bmTableBean.getRowData().get(i + 1);
                            ValueType valueType = bmDataBean.getValueType();
                            if (ValueType.TEXT == valueType) {
                                String value = bmDataBean.getTextValue();
                                value = value == null ? "" : value;
                                // 处理字符型数据
                                // 处理当前cellXml 生成具体的cell文本数据
                                cellXml = replaceStringCellValue(cellXml, value);
                                // 追加给当前行的xml
                                targetRowCellXml.append(cellXml);
                            } else {
                                List<File> imageFiles = bmDataBean.getImageValue();
                                if (imageFiles == null) {
                                    cellXml = replaceStringCellValue(cellXml, "");
                                    targetRowCellXml.append(cellXml);
                                    continue;
                                }
                                // 当前逻辑处理图片型数据
                                StringBuilder imageXml = new StringBuilder();
                                for (File file : imageFiles) {
                                    // 获取到cell的宽度
                                    Double cellWidth = Double.valueOf(getCellWidth(cellXml));
                                    // 获取图片型xml数据
                                    imageXml.append(getImageXml(file, cellWidth, htXwpfDocument));
                                }
                                // 处理当前cellXml 生成具体的cell数据
                                cellXml = replaceStringCellAsImageCellXml(cellXml, imageXml.toString());
                                // 追加给当前行的xml
                                targetRowCellXml.append(cellXml);
                            }
                        }
                    }
                    // 替换当前行的cell数据为新的cell数据，得到新的行数据
                    String newRowXml = rowGroup.replace(currentRowCellXml, targetRowCellXml);
                    // 追加在最终的xml信息中
                    targetXml.append(newRowXml);
                }
                // 替换行数据
                if (targetXml.length() != 0) {
                    continue;
                }
                // 替换当前行数据为最终的cell数据
                rootXml = rootXml.replace(rowGroup, targetXml);
            }
        }
        htXwpfDocument.getDocument().set(XmlObject.Factory.parse(rootXml));
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
        String cellWidth = "";
        while (cellWidthMatcher.find()) {
            String cellWidthGroup = cellWidthMatcher.group();
            String w = getPrValueOfXmlClosingTag("w:tcW",cellWidthGroup, "w:w");
            cellWidth =  w == null? "2500" : w;
            break;
        }
        return cellWidth;
    }

    /**
     * 获取image的尺寸
     * @param file word图片Bean
     * @param cellWidth 单元格宽度
     * @param htXwpfDocument doc数据
     * @return image 的xml数据
     * @throws Exception IO异常
     */
    public static String getImageXml(File file, Double cellWidth, XWPFDocument htXwpfDocument) throws Exception {
        String imageName = file.getName();
        Double imageWidthEmu = getImageWidthEmu(cellWidth);
        BufferedImage image = ImageIO.read(file);
        double imageWidth = image.getWidth();
        double imageHeight = image.getHeight();
        String imageRefId = htXwpfDocument.addPictureData(Files.newInputStream(file.toPath()), 3);
        int imageId = htXwpfDocument.getNextPicNameNumber(3);
        Double imageHeightEmu = getImageHeightEmu(cellWidth, imageWidth, imageHeight);
        return generateImageXml(imageWidthEmu + "", imageHeightEmu + "", imageName, imageRefId, imageId);
    }

    private static Double getImageWidthEmu(Double cellWidth){
        Double imageEmuUnit = 914400D;
        Double cellUnit = 1440D;
        Double cellWidthInch = cellWidth / cellUnit;
        return cellWidthInch * imageEmuUnit;
    }

    private static Double getImageHeightEmu(Double cellWidth, Double imageWidth, Double imageHeight){
        Double imageUnit = 72D;
        Double imageEmuUnit = 914400D;
        Double cellUnit = 1440D;
        Double cellWidthInch = cellWidth / cellUnit;
        Double imageWidthInch = imageWidth / imageUnit;
        Double scaling = imageWidthInch/cellWidthInch;
        Double imageHeightInch = imageHeight / imageUnit;
        Double cellHeightInch = imageHeightInch / scaling;
        // 图片高度单位
        return cellHeightInch * imageEmuUnit;
    }

    /**
     * 增加一个内嵌的图片数据
     * @param imageWidthEmu 图片宽度
     * @param imageHeightEmu 图片高度
     * @param imageNameDesc 图片描述，接入图片名称
     * @param imageRefId 图片引用id
     * @return 生成的图片字符串
     */
    public static String generateImageXml(String imageWidthEmu, String imageHeightEmu, String imageNameDesc, String imageRefId, int imageId){
        String imageRpr = "<w:p><w:pPr><w:rPr><w:rFonts w:hint=\"default\"/><w:vertAlign w:val=\"baseline\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"default\"/><w:vertAlign w:val=\"baseline\"/></w:rPr>";
        String drawingXml = "<w:drawing><wp:inline distT=\"114300\" distB=\"114300\" distL=\"114300\" distR=\"114300\"><wp:extent cx=\""+imageWidthEmu+"\" cy=\""+imageHeightEmu+"\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:docPr id=\""+imageId+"\" name=\""+"图片 "+imageId+"\" descr=\""+imageNameDesc+"\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\""+imageId+"\" name=\"图片 "+imageId+"\" descr=\""+imageNameDesc+"\"/><pic:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=\""+imageRefId+"\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\""+imageWidthEmu+"\" cy=\""+imageHeightEmu+"\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>";
        String endMark = "</w:r></w:p>";
        return imageRpr + drawingXml +endMark;
    }

    /**
     * 获取输入文本对应正则表达式的匹配对象
     * @param regEx 正则表达式
     * @param input 输入字符串
     * @return matcher对象
     */
    public static Matcher getMatcherByRegEx(@NonNull String regEx, String input){
        Pattern tablePattern = Pattern.compile(regEx);
        return tablePattern.matcher(input);
    }

    /**
     * 根据开始标志匹配书签属性
     *
     * @param source 来源的xml信息
     * @param attr   匹配属性
     * @return 匹配结果
     */
    private static String getPrValueOfXmlTag(String source, String attr) {
        String element = "w:bookmarkStart";
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
     * 根据标签名称匹配标签中的属性
     * @param element 元素
     * @param source 来源数据必须是单一闭合标签内容
     * @param attr 属性
     * @return 字符串
     */
    private static String getPrValueOfXmlClosingTag(String element, String source, String attr) {
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
     * 垂直合并
     * @param isStart 是否为开始单元格
     * @param cellXml 当前操作的cell xml数据
     * @return 解析后的cellCml数据
     */
    public static String verticalMergeCell(boolean isStart, String cellXml){
        if (cellXml == null) {
            return null;
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

    /**
     * 替换当前的单元格中文本数据
     * @param cellXml 当前单元格xml
     * @param value 待替换的数据
     */
    public static String replaceStringCellValue(String cellXml, String value) {
        String textRegEx = "<w:r>(.*?)</w:r>";
        Pattern textPattern = Pattern.compile(textRegEx);
        Matcher textMatcher = textPattern.matcher(cellXml);
        int flag = 0;
        while (textMatcher.find()) {
            String textGroup = textMatcher.group();
            // 当前数据可能存在多个t标签，只给第一个t 标签赋值，其他t标签置空
            if (flag > 0) {
                cellXml = cellXml.replace(textGroup, "<w:r><w:rPr><w:rFonts w:hint=\"default\"/><w:vertAlign w:val=\"baseline\"/><w:lang w:eastAsia=\"zh-Hans\"/></w:rPr><w:t></w:t></w:r>");
            } else {
                cellXml = cellXml.replace(textGroup, "<w:r><w:rPr><w:rFonts w:hint=\"default\"/><w:vertAlign w:val=\"baseline\"/><w:lang w:eastAsia=\"zh-Hans\"/></w:rPr><w:t>" + value + "</w:t></w:r>");
            }
            flag++;
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
        String textRegEx = "<w:p>(.*?)</w:p>";
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
}
