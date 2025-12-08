package com.pptfactory.style;

import com.aspose.slides.IPortion;

/**
 * 风格策略接口
 * 
 * 定义了风格策略的接口，所有具体的风格策略类都需要实现此接口。
 * 风格策略类负责控制PPT的视觉效果：颜色、字体、间距等。
 * 
 * 设计模式：策略模式（Strategy Pattern）
 * - 模板引擎保持统一，所有风格通过策略类实现
 * - 模板文件决定"页面长什么样、有哪些占位符、布局怎么摆"
 * - 风格策略类决定"颜色、字体、图片风格、行距间距等视觉效果"
 * 
 * 注意：已改为使用 Aspose.Slides API
 */
public interface StyleStrategy {
    
    /**
     * 应用标题样式
     * 
     * @param portion 文本部分对象（Aspose.Slides IPortion）
     * @param isMainTitle 是否为主标题（标题页的大标题）
     */
    void applyTitleStyle(IPortion portion, boolean isMainTitle);
    
    /**
     * 应用副标题样式
     * 
     * @param portion 文本部分对象（Aspose.Slides IPortion）
     */
    void applySubtitleStyle(IPortion portion);
    
    /**
     * 应用正文样式
     * 
     * @param portion 文本部分对象（Aspose.Slides IPortion）
     */
    void applyContentStyle(IPortion portion);
    
    /**
     * 应用要点列表样式
     * 
     * @param portion 文本部分对象（Aspose.Slides IPortion）
     */
    void applyBulletStyle(IPortion portion);
    
    /**
     * 获取要点列表的段间距（单位：点）
     * 
     * @return 段间距值（点）
     */
    double getBulletSpacing();
    
    /**
     * 获取标题字体大小（单位：点）
     * 
     * @param isMainTitle 是否为主标题
     * @return 字体大小（点）
     */
    double getTitleFontSize(boolean isMainTitle);
    
    /**
     * 获取副标题字体大小（单位：点）
     * 
     * @return 字体大小（点）
     */
    double getSubtitleFontSize();
    
    /**
     * 获取正文字体大小（单位：点）
     * 
     * @return 字体大小（点）
     */
    double getContentFontSize();
    
    /**
     * 获取要点列表字体大小（单位：点）
     * 
     * @return 字体大小（点）
     */
    double getBulletFontSize();
}

