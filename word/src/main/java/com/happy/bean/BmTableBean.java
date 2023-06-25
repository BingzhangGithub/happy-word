package com.happy.bean;

import lombok.Data;

import java.util.Map;

/**
 * this is a desc for this class
 *
 * @author bing.zhang
 * @date 2023年06月25日 15:17
 */
@Data
public class BmTableBean {
    /**
     * 当前书签数据对应的总行数
     */
    private Integer totalRow;
    /**
     * 当前书签对应的行数据
     */
     private Map<Integer, BmNonTableBean> rowData;

}
