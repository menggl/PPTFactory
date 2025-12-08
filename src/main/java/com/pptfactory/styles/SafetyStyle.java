package com.pptfactory.styles;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import java.awt.Color;

/**
 * 安全生产风格策略类
 * 
 * 针对安全生产、安全培训类PPT优化的风格，特点：
 * - 使用醒目的安全配色（红色、橙色、黄色等警示色）
 * - 使用较大的字体，确保清晰可读
 * - 使用加粗标题，突出安全重要性
 * - 使用适中的间距，便于阅读和理解
 */
public class SafetyStyle implements StyleStrategy {
    
    // 字体大小配置（安全培训需要清晰易读）
    private static final double MAIN_TITLE_SIZE = 52.0;  // 主标题字体大小（较大，突出重要性）
    private static final double TITLE_SIZE = 46.0;       // 普通标题字体大小
    private static final double SUBTITLE_SIZE = 26.0;    // 副标题字体大小
    private static final double CONTENT_SIZE = 20.0;     // 正文字体大小
    private static final double BULLET_SIZE = 18.0;      // 要点列表字体大小
    
    // 颜色配置（使用安全警示色系）
    private static final Color TITLE_COLOR = new Color(220, 20, 60);      // 深红色标题（警示色）
    private static final Color SUBTITLE_COLOR = new Color(255, 140, 0);   // 橙色副标题（警告色）
    private static final Color BULLET_COLOR = new Color(51, 51, 51);      // 深灰色要点（确保可读性）
    
    // 间距配置
    private static final double BULLET_SPACING = 14.0;   // 要点列表段间距
    
    @Override
    public void applyTitleStyle(XSLFTextParagraph paragraph, boolean isMainTitle) {
        // 安全生产类PPT的标题使用醒目的红色，加粗显示，突出安全重要性
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE);
            run.setBold(true);  // 标题加粗，强调重要性
            run.setFontColor(TITLE_COLOR);  // 使用警示红色
        }
    }
    
    @Override
    public void applySubtitleStyle(XSLFTextParagraph paragraph) {
        // 使用橙色作为副标题颜色，起到警告提示作用
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(SUBTITLE_SIZE);
            run.setBold(true);  // 副标题也加粗
            run.setFontColor(SUBTITLE_COLOR);  // 使用警告橙色
        }
    }
    
    @Override
    public void applyContentStyle(XSLFTextParagraph paragraph) {
        // 正文使用标准黑色，确保良好的可读性
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(CONTENT_SIZE);
            run.setBold(false);
            // 使用默认黑色（不设置颜色）
        }
    }
    
    @Override
    public void applyBulletStyle(XSLFTextParagraph paragraph) {
        // 要点列表使用深灰色，字体稍小但清晰可读
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(BULLET_SIZE);
            run.setBold(false);
            run.setFontColor(BULLET_COLOR);  // 使用深灰色
        }
    }
    
    @Override
    public double getBulletSpacing() {
        return BULLET_SPACING;
    }
    
    @Override
    public double getTitleFontSize(boolean isMainTitle) {
        return isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE;
    }
    
    @Override
    public double getSubtitleFontSize() {
        return SUBTITLE_SIZE;
    }
    
    @Override
    public double getContentFontSize() {
        return CONTENT_SIZE;
    }
    
    @Override
    public double getBulletFontSize() {
        return BULLET_SIZE;
    }
}

