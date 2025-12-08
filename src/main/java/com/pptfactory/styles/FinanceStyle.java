package com.pptfactory.styles;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import java.awt.Color;

/**
 * 金融风格策略类
 * 
 * 实现了适合金融类PPT的风格，使用专业的配色和清晰的排版。
 */
public class FinanceStyle implements StyleStrategy {
    
    private static final double MAIN_TITLE_SIZE = 48.0;
    private static final double TITLE_SIZE = 44.0;
    private static final double SUBTITLE_SIZE = 24.0;
    private static final double CONTENT_SIZE = 18.0;
    private static final double BULLET_SIZE = 16.0;
    private static final double BULLET_SPACING = 12.0;
    
    // 金融风格配色（深蓝色系，专业稳重）
    private static final Color TITLE_COLOR = new Color(0, 51, 102);  // 深蓝色
    private static final Color SUBTITLE_COLOR = new Color(0, 102, 204);  // 蓝色
    
    @Override
    public void applyTitleStyle(XSLFTextParagraph paragraph, boolean isMainTitle) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE);
            run.setBold(true);
            run.setFontColor(TITLE_COLOR);
        }
    }
    
    @Override
    public void applySubtitleStyle(XSLFTextParagraph paragraph) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(SUBTITLE_SIZE);
            run.setBold(false);
            run.setFontColor(SUBTITLE_COLOR);
        }
    }
    
    @Override
    public void applyContentStyle(XSLFTextParagraph paragraph) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(CONTENT_SIZE);
            run.setBold(false);
        }
    }
    
    @Override
    public void applyBulletStyle(XSLFTextParagraph paragraph) {
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

