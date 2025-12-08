package com.pptfactory.styles;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * 数学风格策略类
 * 
 * 实现了适合数学类PPT的风格，使用等宽字体和清晰的排版。
 */
public class MathStyle implements StyleStrategy {
    
    private static final double MAIN_TITLE_SIZE = 48.0;
    private static final double TITLE_SIZE = 44.0;
    private static final double SUBTITLE_SIZE = 24.0;
    private static final double CONTENT_SIZE = 18.0;
    private static final double BULLET_SIZE = 16.0;
    private static final double BULLET_SPACING = 12.0;
    
    @Override
    public void applyTitleStyle(XSLFTextParagraph paragraph, boolean isMainTitle) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE);
            run.setBold(true);
            run.setFontFamily("Times New Roman"); // 数学公式常用字体
        }
    }
    
    @Override
    public void applySubtitleStyle(XSLFTextParagraph paragraph) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(SUBTITLE_SIZE);
            run.setBold(false);
            run.setFontFamily("Times New Roman");
        }
    }
    
    @Override
    public void applyContentStyle(XSLFTextParagraph paragraph) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(CONTENT_SIZE);
            run.setBold(false);
            run.setFontFamily("Times New Roman");
        }
    }
    
    @Override
    public void applyBulletStyle(XSLFTextParagraph paragraph) {
        for (XSLFTextRun run : paragraph.getTextRuns()) {
            run.setFontSize(BULLET_SIZE);
            run.setBold(false);
            run.setFontFamily("Times New Roman");
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

