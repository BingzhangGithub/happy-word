package com.happy.bean;

import lombok.Data;

import java.io.File;
import java.util.List;

/**
 * this is a desc for this class
 *
 * @author bing.zhang
 * @date 2023年06月25日 15:17
 */
@Data
public class BmNonTableBean {
    /**
     * 区分数据类型
     */
    private ValueType valueType;
    /**
     * 文本类型数据存储
     */
    private String textValue;
    /**
     * 图片类型数据存储
     */
    private List<File> imageValue;
    /**
     *  富文本数据
     */
    private String richTextValue;
}
