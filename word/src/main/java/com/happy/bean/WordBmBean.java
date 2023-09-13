package com.happy.bean;

import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * word书签bean
 *
 * @author bing.zhang
 * @date 2023年06月29日 21:33
 */
@Data
public class WordBmBean implements Serializable {
    private static final long serialVersionUID = 2138223004131346168L;

    /**
     * 字段类型：用于校验是否为文本，用于拼接文本处理
     */
    private WordFieldType filedType;
    /**
     * 显示设置
     */
    private String showSettingsId;
    /**
     * 是否是表格数据：是则在Map中存取数据
     */
    boolean isTableData;

    /**
     * 文本类型数据存储,或者为图片型的id集合
     */
    private List<String> values;

    /**
     * 当前字段对应的图片文件数据
     */
    private List<WordImageBean> imageList;

    /**
     * 表格数据存储集合
     */
    private Map<Integer, WordBmBean> bamBeanMap;

    public WordBmBean() {
    }

    public WordBmBean(List<String> values) {
        this.values = values;
    }

    public WordBmBean(WordFieldType filedType, List<String> values) {
        this.filedType = filedType;
        this.values = values;
    }

    public WordBmBean(WordFieldType filedType, String showSettingsId) {
        this.filedType = filedType;
        this.showSettingsId = showSettingsId;
    }

    /**
     * 是否为图片类型
     * @return 布尔值
     */
    public boolean isImageType(){
        return WordFieldType.IMAGE == filedType;
    }
}
