package com.pptfactory.template.engine;

import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import java.io.File;
import java.util.Locale;

/**
 * 水印移除功能单步调试测试类
 * 
 * 用于测试水印移除功能，包括：
 * - XML方式移除水印
 * - 水印检测
 * - 水印文本框删除
 * 
 * 使用方法：
 * 1. 确保使用Java 17或更高版本运行（项目要求Java 17+）
 * 2. 在IDE中打开此文件
 * 3. 在main方法中设置断点
 * 4. 以Debug模式运行
 * 5. 单步调试查看水印移除的每个步骤
 * 
 * IDE配置说明：
 * - IntelliJ IDEA: File -> Project Structure -> Project -> SDK 选择 Java 17
 * - Eclipse: Window -> Preferences -> Java -> Installed JREs 添加 Java 17
 */
public class TestWatermarkRemoval {
    
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
        System.out.println("=== 水印移除功能单步调试测试 ===");
        
        try {
            // 测试文件路径
            String testFile = "templates/master_template.pptx";
            String outputFile = "test_output_no_watermark.pptx";
            
            // 检查测试文件是否存在
            File file = new File(testFile);
            if (!file.exists()) {
                System.err.println("错误：测试文件不存在: " + testFile);
                System.err.println("请先运行模板提取流程生成 master_template.pptx");
                return;
            }
            
            // 测试1: 加载包含水印的PPT
            System.out.println("\n[测试1] 加载包含水印的PPT...");
            Presentation presentation = new Presentation(testFile);
            System.out.println("✓ PPT加载成功，共 " + presentation.getSlides().size() + " 张幻灯片");
            
            // 设置断点：检查加载的PPT内容
            
            // 测试2: 使用反射调用私有方法进行水印移除
            // 注意：由于removeWatermarksFromXML是私有方法，我们需要通过PPTTemplateEngine来测试
            System.out.println("\n[测试2] 创建模板引擎并触发水印移除...");
            
            // 先保存一个副本用于测试
            String tempFile = "test_temp_with_watermark.pptx";
            presentation.save(tempFile, SaveFormat.Pptx);
            presentation.dispose();
            
            // 创建模板引擎（会自动触发水印移除）
            com.pptfactory.style.StyleStrategy styleStrategy = new com.pptfactory.style.DefaultStyle();
            PPTTemplateEngine engine = new PPTTemplateEngine(tempFile, styleStrategy);
            
            // 保存移除水印后的文件
            engine.save(outputFile);
            System.out.println("✓ 水印移除完成，结果保存到: " + outputFile);
            
            // 测试3: 验证水印是否已移除
            System.out.println("\n[测试3] 验证水印是否已移除...");
            Presentation resultPresentation = new Presentation(outputFile);
            System.out.println("✓ 验证文件加载成功，共 " + resultPresentation.getSlides().size() + " 张幻灯片");
            
            // 设置断点：检查结果文件内容，确认水印已移除
            
            resultPresentation.dispose();
            engine.close();
            
            // 清理临时文件
            new File(tempFile).delete();
            
            System.out.println("\n=== 所有测试完成 ===");
            System.out.println("请手动打开 " + outputFile + " 检查水印是否已完全移除");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
