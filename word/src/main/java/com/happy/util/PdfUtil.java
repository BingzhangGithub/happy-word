package com.happy.util;

import cn.hutool.core.util.ObjectUtil;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.Loader;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * this is a desc for this class
 *
 * @author bing.zhang
 * @date 2023年08月15日 14:03
 */
public class PdfUtil {

    /**
     * 将PDF文件转化为图片
     * @param pdfFile pdf 输入文件
     * @param fileOutputPath 输出路径
     * @return 图片的文件列表
     * @throws IOException IO异常
     */
    public static List<File> convertToImage(File pdfFile, String fileOutputPath) throws IOException {
        if (pdfFile == null || ObjectUtil.isEmpty(fileOutputPath)) {
            throw new IOException("File is Empty.");
        }
        PDDocument document = Loader.loadPDF(pdfFile);
        int numberOfPages = document.getNumberOfPages();
        if (numberOfPages <= 0){
            throw new RuntimeException("File is Empty! Please check your file. filePath:" + pdfFile.getAbsolutePath());
        }

        PDFRenderer pdfRenderer = new PDFRenderer(document);
        List<File> imageFiles = new ArrayList<>();
        for (int i = 0; i < numberOfPages; i++) {
            BufferedImage bim  = pdfRenderer.renderImage(i);
            File tempFile = File.createTempFile("temp", ".png", new File(fileOutputPath));
            ImageIO.write(bim, "PNG", tempFile);
            System.out.println(tempFile.getAbsolutePath());
            imageFiles.add(tempFile);
        }
        document.close();
        return imageFiles;
    }
}
