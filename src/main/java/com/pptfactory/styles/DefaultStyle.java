package com.pptfactory.styles;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * 默认风格策略类
 * 
 * 实现了通用的默认风格，适合大多数PPT场景。
 * 使用标准的字体大小和颜色配置。
 */
public class DefaultStyle implements StyleStrategy {
    
    // 字体大小配置（单位：点）
    private static final double MAIN_TITLE_SIZE = 48.0;  // 主标题字体大小
    private static final double TITLE_SIZE = 44.0;       // 普通标题字体大小
    private static final double SUBTITLE_SIZE = 24.0;    // 副标题字体大小
    private static final double CONTENT_SIZE = 18.0;     // 正文字体大小
    private static final double BULLET_SIZE = 16.0;      // 要点列表字体大小
    
    // 间距配置（单位：点）
    private static final double BULLET_SPACING = 12.0;   // 要点列表段间距
    
    @Override
    public void applyTitleStyle(XSLFTextParagraph paragraph, boolean isMainTitle) {
        // 应用标题样式
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE);
            run.setBold(true);
        }
    }
    
    @Override
    public void applySubtitleStyle(XSLFTextParagraph paragraph) {
        // 应用副标题样式
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(SUBTITLE_SIZE);
            run.setBold(false);
        }
    }
    
    @Override
    public void applyContentStyle(XSLFTextParagraph paragraph) {
        // 应用正文样式
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(CONTENT_SIZE);
            run.setBold(false);
        }
    }
    
    @Override
    public void applyBulletStyle(XSLFTextParagraph paragraph) {
        // 应用要点列表样式
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(BULLET_SIZE);
            run.setBold(false);
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

