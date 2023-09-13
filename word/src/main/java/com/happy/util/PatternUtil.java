package com.happy.util;

import cn.hutool.core.util.ObjectUtil;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author hxd
 * @date 2022年11月24日 11:43
 */
public class PatternUtil implements Serializable {

    private static final long serialVersionUID = 1833865966655402399L;

    /**
     * 正则匹配 返回list
     *
     * @param regex   正则规则
     * @param content 匹配内容
     * @return list
     */
    public static List<String> matcherToList(String regex, String content) {
        if (ObjectUtil.isNotEmpty(regex) && ObjectUtil.isNotEmpty(content)) {
            Pattern p1 = Pattern.compile(regex);
            Matcher matcher = p1.matcher(content);
            List<String> list = new ArrayList<>();
            while (matcher.find()) {
                list.add(matcher.group(0));
            }
            return list;
        }
        return Collections.emptyList();
    }

    /**
     * 判断内容是否符合规则
     *
     * @param regex   正则
     * @param content 内容
     * @return true 符合， false不符合
     */
    public static boolean matcherHasRule(String regex, String content) {
        if (ObjectUtil.isNotEmpty(regex) && ObjectUtil.isNotEmpty(content)) {
            return content.matches(regex);
        }
        return false;
    }
}
