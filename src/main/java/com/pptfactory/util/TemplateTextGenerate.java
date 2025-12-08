package com.pptfactory.util;

/*
 * 生成模板文字
 * 
 * 如果模板文字长度 > 目标长度，则截取
 * 如果模板文字长度 < 目标长度，则重复模板文字直到达到目标长度
 * 
 * @param targetLength 目标长度（原始文字的字数）
 * @return 生成后的模板文字
 */
public class TemplateTextGenerate {
    
    private static final String TEMPLATE_TEXT = "模板文字模板文字模板文字模板文字模板文字模板文字";

    public static String generateTemplateText(int targetLength) {
        if (targetLength <= 0) {
            return "";
        }
        if (TEMPLATE_TEXT.length() >= targetLength) {
            return TEMPLATE_TEXT.substring(0, targetLength);
        }
        StringBuilder sb = new StringBuilder();
        while (sb.length() < targetLength) {
            int remaining = targetLength - sb.length();
            if (remaining >= TEMPLATE_TEXT.length()) {
                sb.append(TEMPLATE_TEXT);
            } else {
                sb.append(TEMPLATE_TEXT.substring(0, remaining));
            }
        }
        return sb.toString();
    }   
}
