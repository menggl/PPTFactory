package com.pptfactory.style;

import com.aspose.slides.IPortion;
import com.aspose.slides.FillType;
import java.awt.Color;

/**
 * classic_image_text_16 风格策略类
 * 
 * 从《1.2 安全生产方针政策.pptx》第16页提取的风格。
 * 此风格类保留了源PPT页面的样式特征。
 */
public class ImageText16Style implements StyleStrategy {

    // 字体大小配置（从源PPT第16页提取）
    private static final double MAIN_TITLE_SIZE = 30.0;
    private static final double TITLE_SIZE = 30.0;
    private static final double SUBTITLE_SIZE = 21.0;
    private static final double CONTENT_SIZE = 18.0;
    private static final double BULLET_SIZE = 18.0;

    // 颜色配置（从源PPT第16页提取）
    private static final Color TITLE_COLOR = new Color(23, 69, 108);
    private static final Color CONTENT_COLOR = new Color(255, 255, 255);

    // 间距配置
    private static final double BULLET_SPACING = 12.0;

    @Override
    public void applyTitleStyle(IPortion portion, boolean isMainTitle) {
        portion.getPortionFormat().setFontHeight((float)(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE));
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);
    }

    @Override
    public void applySubtitleStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)SUBTITLE_SIZE);
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);
    }

    @Override
    public void applyContentStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)CONTENT_SIZE);
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(CONTENT_COLOR);
    }

    @Override
    public void applyBulletStyle(IPortion portion) {
        portion.getPortionFormat().setFontHeight((float)BULLET_SIZE);
        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(CONTENT_COLOR);
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
