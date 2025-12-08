package com.pptfactory.style;

import com.aspose.slides.IPortion;
import com.aspose.slides.NullableBool;
import com.aspose.slides.FillType;
import java.awt.Color;

/**
 * 默认风格策略类
 * 
 * 实现了通用的默认风格，适合大多数PPT场景。
 * 使用标准的字体大小和颜色配置。
 * 
 * 注意：已改为使用 Aspose.Slides API
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
    
    // 颜色配置
    private static final Color TITLE_COLOR = new Color(0, 0, 0);        // 黑色标题
    private static final Color SUBTITLE_COLOR = new Color(64, 64, 64);  // 深灰色副标题
    private static final Color CONTENT_COLOR = new Color(0, 0, 0);      // 黑色正文
    private static final Color BULLET_COLOR = new Color(0, 0, 0);       // 黑色要点
    
    @Override
    public void applyTitleStyle(IPortion portion, boolean isMainTitle) {
        // 应用标题样式
        portion.getPortionFormat().setFontHeight((float)(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE));
        // portion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);
    }
    
    @Override
    public void applySubtitleStyle(IPortion portion) {
        // 应用副标题样式
        portion.getPortionFormat().setFontHeight((float)SUBTITLE_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(SUBTITLE_COLOR);
    }
    
    @Override
    public void applyContentStyle(IPortion portion) {
        // 应用正文样式
        portion.getPortionFormat().setFontHeight((float)CONTENT_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(CONTENT_COLOR);
    }
    
    @Override
    public void applyBulletStyle(IPortion portion) {
        // 应用要点列表样式
        portion.getPortionFormat().setFontHeight((float)BULLET_SIZE);
        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(BULLET_COLOR);
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

