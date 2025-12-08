package com.pptfactory.style;

import com.aspose.slides.Presentation;
import com.aspose.slides.ISlide;
import com.aspose.slides.IAutoShape;
import com.aspose.slides.ITextFrame;
import com.aspose.slides.IParagraph;
import com.aspose.slides.IPortion;
import com.aspose.slides.Paragraph;
import com.aspose.slides.Portion;
import java.util.Locale;

/**
 * StyleStrategy 单步调试测试类
 * 
 * 用于测试风格策略类的功能，包括：
 * - 不同风格策略的创建
 * - 标题样式应用
 * - 正文样式应用
 * - 字体大小和颜色设置
 * 
 * 使用方法：
 * 1. 确保使用Java 17或更高版本运行（项目要求Java 17+）
 * 2. 在IDE中打开此文件
 * 3. 在main方法中设置断点
 * 4. 以Debug模式运行
 * 5. 单步调试查看样式应用的每个步骤
 * 
 * IDE配置说明：
 * - IntelliJ IDEA: File -> Project Structure -> Project -> SDK 选择 Java 17
 * - Eclipse: Window -> Preferences -> Java -> Installed JREs 添加 Java 17
 */
public class TestStyleStrategy {
    
    static {
        Locale.setDefault(Locale.US);
        checkJavaVersion();
    }
    
    private static void checkJavaVersion() {
        String javaVersion = System.getProperty("java.version");
        try {
            int majorVersion = getJavaMajorVersion();
            if (majorVersion < 17) {
                System.err.println("错误：此项目需要Java 17或更高版本，当前版本: " + majorVersion);
                System.err.println("请配置IDE使用Java 17运行此测试");
                System.exit(1);
            }
        } catch (Exception e) {
            System.err.println("警告：无法检查Java版本: " + e.getMessage());
        }
    }
    
    private static int getJavaMajorVersion() {
        String version = System.getProperty("java.version");
        if (version.startsWith("1.")) {
            version = version.substring(2, 3);
        } else {
            int dot = version.indexOf(".");
            if (dot != -1) {
                version = version.substring(0, dot);
            }
        }
        return Integer.parseInt(version);
    }
    
    public static void main(String[] args) {
        System.out.println("=== StyleStrategy 单步调试测试 ===");
        
        try {
            // 创建一个测试用的PPT
            Presentation presentation = new Presentation();
            ISlide slide = presentation.getSlides().addEmptySlide(presentation.getLayoutSlides().get_Item(0));
            
            // 创建测试文本框
            IAutoShape textShape = slide.getShapes().addAutoShape(
                com.aspose.slides.ShapeType.Rectangle, 100, 100, 500, 100);
            ITextFrame textFrame = textShape.getTextFrame();
            textFrame.getParagraphs().clear();
            
            // 测试1: SafetyStyle 风格策略
            System.out.println("\n[测试1] SafetyStyle 风格策略...");
            StyleStrategy safetyStyle = new SafetyStyle();
            
            // 测试主标题样式
            System.out.println("  测试主标题样式...");
            IParagraph titlePara = new Paragraph();
            textFrame.getParagraphs().add(titlePara);
            IPortion titlePortion = new Portion();
            titlePortion.setText("安全生产主标题");
            titlePara.getPortions().add(titlePortion);
            
            // 设置断点：查看样式应用前后的变化
            safetyStyle.applyTitleStyle(titlePortion, true);
            System.out.println("  ✓ 主标题样式应用完成");
            System.out.println("    字体大小: " + titlePortion.getPortionFormat().getFontHeight());
            
            // 测试副标题样式
            System.out.println("  测试副标题样式...");
            IParagraph subtitlePara = new Paragraph();
            textFrame.getParagraphs().add(subtitlePara);
            IPortion subtitlePortion = new Portion();
            subtitlePortion.setText("副标题文本");
            subtitlePara.getPortions().add(subtitlePortion);
            
            safetyStyle.applySubtitleStyle(subtitlePortion);
            System.out.println("  ✓ 副标题样式应用完成");
            System.out.println("    字体大小: " + subtitlePortion.getPortionFormat().getFontHeight());
            
            // 测试正文样式
            System.out.println("  测试正文样式...");
            IParagraph contentPara = new Paragraph();
            textFrame.getParagraphs().add(contentPara);
            IPortion contentPortion = new Portion();
            contentPortion.setText("这是正文内容，用于测试正文样式应用。");
            contentPara.getPortions().add(contentPortion);
            
            safetyStyle.applyContentStyle(contentPortion);
            System.out.println("  ✓ 正文样式应用完成");
            System.out.println("    字体大小: " + contentPortion.getPortionFormat().getFontHeight());
            
            // 测试要点列表样式
            System.out.println("  测试要点列表样式...");
            IParagraph bulletPara = new Paragraph();
            textFrame.getParagraphs().add(bulletPara);
            IPortion bulletPortion = new Portion();
            bulletPortion.setText("要点1：安全第一");
            bulletPara.getPortions().add(bulletPortion);
            
            safetyStyle.applyBulletStyle(bulletPortion);
            System.out.println("  ✓ 要点列表样式应用完成");
            System.out.println("    字体大小: " + bulletPortion.getPortionFormat().getFontHeight());
            System.out.println("    段间距: " + safetyStyle.getBulletSpacing());
            
            // 测试2: DefaultStyle 风格策略
            System.out.println("\n[测试2] DefaultStyle 风格策略...");
            StyleStrategy defaultStyle = new DefaultStyle();
            
            System.out.println("  主标题字体大小: " + defaultStyle.getTitleFontSize(true));
            System.out.println("  副标题字体大小: " + defaultStyle.getSubtitleFontSize());
            System.out.println("  正文字体大小: " + defaultStyle.getContentFontSize());
            System.out.println("  要点字体大小: " + defaultStyle.getBulletFontSize());
            
            // 保存测试结果
            String outputFile = "test_output_style.pptx";
            presentation.save(outputFile, com.aspose.slides.SaveFormat.Pptx);
            System.out.println("\n✓ 测试结果已保存: " + outputFile);
            
            presentation.dispose();
            
            System.out.println("\n=== 所有测试完成 ===");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
