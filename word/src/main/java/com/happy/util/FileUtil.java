package com.happy.util;


import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;

import cn.hutool.core.util.ObjectUtil;
import com.happy.bean.ImageType;
import com.sun.org.slf4j.internal.Logger;
import com.sun.org.slf4j.internal.LoggerFactory;
import java.io.*;
import java.nio.file.NoSuchFileException;

/**
 * this is a desc for this class
 *
 * @author bing.zhang
 * @date 2023年08月15日 14:02
 */
public class FileUtil {

    private static final Logger log = LoggerFactory.getLogger(FileUtil.class);


    /**
     * 获取byte[]
     *
     * @param file      文件
     * @param imageType imageType类型
     */
    public static byte[] getByteByImageFile(File file, ImageType imageType) {
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage bufferImg;
        try {
            if (file != null) {
                InputStream inputStream = new FileInputStream(file);
                bufferImg = ImageIO.read(inputStream);
                ImageIO.write(bufferImg, imageType.toString(), byteArrayOut);
            }
        } catch (Exception e) {
            log.error("err", e);
        }
        return byteArrayOut.toByteArray();
    }

    /**
     * 创建临时文件
     *
     * @param inputStream 文件流
     * @param name        文件名
     * @param ext         扩展名
     * @return file
     * @throws IOException io异常
     */
    public static File createTmpFile(InputStream inputStream, String name, String ext) throws IOException {
        FileOutputStream fos = null;
        try {
            // 文件名称长度最小不得小于3， 最大不得超过200
            int nameMinLength = 3;
            int nameMaxLength = 200;
            if (ObjectUtil.isEmpty(name) || name.length() < nameMinLength) {
                name = "temp";
            }else if(name.length() > nameMaxLength) {
                name = name.substring(0, nameMaxLength);
            }

            File tmpFile = File.createTempFile(name, '.' + ext);
            tmpFile.deleteOnExit();
            fos = new FileOutputStream(tmpFile);
            int read = 0;
            byte[] bytes = new byte[1024 * 100];
            while ((read = inputStream.read(bytes)) != -1) {
                fos.write(bytes, 0, read);
            }
            fos.flush();
            return tmpFile;
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    log.error(e.getMessage(), e);
                }
            }
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    log.error(e.getMessage(), e);
                }
            }
        }
    }

    /**
     * 关闭流
     *
     * @param closeableStream 可关闭的流对象
     */
    public static void closeStream(Closeable ... closeableStream){
        if (ObjectUtil.isNotEmpty(closeableStream)) {
            for (Closeable closeable : closeableStream) {
                if (closeable != null) {
                    try {
                        closeable.close();
                    } catch (IOException e) {
                        log.error("流关闭失败", e);
                    }
                }
            }
        }
    }

    public static void closeStream(Closeable is) {
        try {
            if (is != null) {
                if (is instanceof InputStream) {
                    is.close();
                }
            }
        } catch (Exception e) {
            log.error("closeStream error", e);
        }
    }

    /**
     * 删除文件
     *
     * @param errorLocation 错误位置：Class.method
     * @param files         文件对象
     */
    public static void deleteFile(String errorLocation, File... files) {
        if (ObjectUtil.isNotEmpty(files)) {
            for (File file : files) {
                if (file != null) {
                    try {
                        file.delete();
                    } catch (Exception e) {
                        log.error("临时文件删除异常，功能调用位置：" + errorLocation, e);
                    }
                }
            }
        }
    }


    /**
     * 获取文件名称没有后缀
     *
     * @param file 文件对象
     * @return 附件格式的枚举类型
     */
    public static String getFileName(File file) throws NoSuchFileException {
        if (file == null || !file.isFile()) {
            throw new NoSuchFileException("file不是文件");
        }
        String fileName = file.getName();
        fileName = fileName.substring(0, fileName.lastIndexOf("."));
        if (ObjectUtil.isEmpty(fileName)) {
            return "";
        }
        return fileName;
    }

    /**
     * 获取文件后缀名称, 没有点
     *
     * @param file 文件对象
     * @return 附件格式的枚举类型
     */
    public static String getFileType(File file) {
        if (!file.isFile()) {
            return null;
        }
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
        if (ObjectUtil.isEmpty(fileType)) {
            return null;
        }
        return fileType;
    }

    /**
     * 获取文件后缀名称，带点
     *
     * @param file 文件对象
     * @return 附件格式的枚举类型
     */
    public static String getFileTypeWithPoint(File file) throws IOException {
        if (file == null || !file.isFile()) {
            throw new IOException("参数file不是文件");
        }
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (ObjectUtil.isEmpty(fileType)) {
            throw new IOException("文件类型为空！");
        }
        return fileType;
    }
}
