package com.happy.bean;

import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * word书签Bean对象
 * 用于书签替换的bean
 *
 * @author bing.zhang
 * @date 2023年03月30日 13:39
 */
@Data
public class WordBean implements Serializable {
    private static final long serialVersionUID = 9092630739278007295L;
    /**
     * 书签名称，文本型映射关系
     */
    private Map<String, String> bookmarkMap;
    /**
     * 书签名称，图片型对象映射关系
     */
    Map<String, List<WordImageBean>> imageInputStreamMap;

    /**
     * 书签名称，富文本型对象映射关系
     */
    Map<String, WordBmBean> bmBeanTableDataMap;

    public WordBean() {
    }

    public WordBean(Map<String, String> bookmarkMap, Map<String, List<WordImageBean>> imageInputStreamMap, Map<String, WordBmBean> bmBeanTableDataMap) {
        this.bookmarkMap = bookmarkMap;
        this.imageInputStreamMap = imageInputStreamMap;
        this.bmBeanTableDataMap = bmBeanTableDataMap;
    }
}
