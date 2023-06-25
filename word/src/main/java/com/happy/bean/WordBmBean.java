package com.happy.bean;

import lombok.Data;

import java.util.Map;

/**
 * this is a desc for this class
 *
 * @author bing.zhang
 * @date 2023年06月25日 15:06
 */
@Data
public class WordBmBean {
    /**
     * 主表数据，用于对普通数据替换
     */
    private Map<String, BmNonTableBean> bmNonTableMap;
    /**
     * 表格数据，用于对表格扩展
     */
    private Map<String, BmTableBean> bmTableMap;
}
