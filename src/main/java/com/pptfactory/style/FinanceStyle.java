package com.pptfactory.style;

import com.aspose.slides.IPortion;
import com.aspose.slides.NullableBool;
import com.aspose.slides.FillType;
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
    public void applyTitleStyle(IPortion portion, boolean isMainTitle) {
        portion.getPortionFormat().setFontHeight((float)(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE));
        // portion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);
    }
    
    @Override
    public void applySubtitleStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)SUBTITLE_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(SUBTITLE_COLOR);
    }
    
    @Override
    public void applyContentStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)CONTENT_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new Color(0, 0, 0));
    }
    
    @Override
    public void applyBulletStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)BULLET_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new Color(0, 0, 0));
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

