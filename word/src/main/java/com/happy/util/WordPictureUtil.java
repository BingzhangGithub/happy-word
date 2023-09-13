package com.happy.util;

import cn.hutool.core.util.ObjectUtil;
import com.happy.bean.ImageTypeConstant;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.Serializable;

/**
 * word 转化图片帮助类
 *
 * @author hxd
 * @date 2022年11月09日 16:20
 */
public class WordPictureUtil implements Serializable {

    private static final long serialVersionUID = -8073919425397661062L;

    /**
     * 根据图片类型，取得对应的图片类型代码。
     *
     * @param picType 图片格式
     * @return XWPFDocument.PICTURE_TYPE
     */
    public static int findPictureType(String picType) {
        if (ObjectUtil.isNotEmpty(picType)) {
            switch (picType) {
                case ImageTypeConstant.PICTURE_TYPE_PNG:
                    return XWPFDocument.PICTURE_TYPE_PNG;
                case ImageTypeConstant.PICTURE_TYPE_DIB:
                    return XWPFDocument.PICTURE_TYPE_DIB;
                case ImageTypeConstant.PICTURE_TYPE_EMF:
                    return XWPFDocument.PICTURE_TYPE_EMF;
                case ImageTypeConstant.PICTURE_TYPE_JPG:
                case ImageTypeConstant.PICTURE_TYPE_JPEG:
                    return XWPFDocument.PICTURE_TYPE_JPEG;
                case ImageTypeConstant.PICTURE_TYPE_WMF:
                    return XWPFDocument.PICTURE_TYPE_WMF;
                default:
                    return XWPFDocument.PICTURE_TYPE_PICT;
            }
        }
        return XWPFDocument.PICTURE_TYPE_PICT;
    }
}
