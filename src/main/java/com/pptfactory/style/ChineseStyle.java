package com.pptfactory.style;

import com.aspose.slides.IPortion;
import com.aspose.slides.NullableBool;
import com.aspose.slides.FillType;

/**
 * 中文风格策略类
 * 
 * 实现了适合中文PPT的风格，使用中文字体和适合中文阅读的间距。
 */
public class ChineseStyle implements StyleStrategy {
    
    private static final double MAIN_TITLE_SIZE = 48.0;
    private static final double TITLE_SIZE = 44.0;
    private static final double SUBTITLE_SIZE = 24.0;
    private static final double CONTENT_SIZE = 18.0;
    private static final double BULLET_SIZE = 16.0;
    private static final double BULLET_SPACING = 12.0;
    
    @Override
    public void applyTitleStyle(IPortion portion, boolean isMainTitle) {
        portion.getPortionFormat().setFontHeight((float)(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE));
        // portion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().setLatinFont(new com.aspose.slides.FontData("微软雅黑"));
    }
    
    @Override
    public void applySubtitleStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)SUBTITLE_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().setLatinFont(new com.aspose.slides.FontData("微软雅黑"));
    }
    
    @Override
    public void applyContentStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)CONTENT_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().setLatinFont(new com.aspose.slides.FontData("微软雅黑"));
    }
    
    @Override
    public void applyBulletStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)BULLET_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().setLatinFont(new com.aspose.slides.FontData("微软雅黑"));
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

