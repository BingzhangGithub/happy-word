package com.happy.bean;

import lombok.Data;

import java.io.File;
import java.io.Serializable;

/**
 * word图片对象
 *
 * @author bing.zhang
 * @date 2023年03月29日 15:52
 */
@Data
public class WordImageBean implements Serializable {
    private static final long serialVersionUID = 1271796091592148202L;
    /**
     * 图片名称，需要带后缀
     */
    private String name;
    /**
     * 设置图片显示属性 0 显示文字 1 嵌入式方式插入图片 2 浮于文字上方 3 浮于文字下方
     * 默认为0
     */
    private int behind;
    /**
     * 图片文件 id 集合
     */
    private Long imageFileId;
    /**
     * 图片文件
     */
    private File imageFile;

    public WordImageBean() {
    }

    public WordImageBean(String name, int behind, File imageFile) {
        this.name = name;
        this.behind = behind;
        this.imageFile = imageFile;
    }

    public WordImageBean(Long imageFileId) {
        this.imageFileId = imageFileId;
    }
}
