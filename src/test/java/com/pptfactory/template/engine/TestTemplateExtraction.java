package com.pptfactory.template.engine;

import com.pptfactory.template.extractor.TemplateExtractor;
import java.io.File;
import java.util.Locale;

/**
 * 模板提取功能单步调试测试类
 * 
 * 用于测试模板提取功能，包括：
 * - 从源PPT提取模板幻灯片
 * - 模板文字替换
 * - 模板图片替换
 * - master_template.pptx生成
 * 
 * 使用方法：
 * 1. 确保使用Java 17或更高版本运行（项目要求Java 17+）
 * 2. 在IDE中打开此文件
 * 3. 在main方法中设置断点
 * 4. 以Debug模式运行
 * 5. 单步调试查看模板提取的每个步骤
 * 
 * IDE配置说明：
 * - IntelliJ IDEA: File -> Project Structure -> Project -> SDK 选择 Java 17
 * - Eclipse: Window -> Preferences -> Java -> Installed JREs 添加 Java 17
 */
public class TestTemplateExtraction {
    
    static {
        Locale.setDefault(Locale.US);
        
        // 检查Java版本
        String javaVersion = System.getProperty("java.version");
        System.out.println("当前Java版本: " + javaVersion);
        
        // 检查是否为Java 17或更高版本
        try {
            int majorVersion = getJavaMajorVersion();
            if (majorVersion < 17) {
                System.err.println("错误：此项目需要Java 17或更高版本，当前版本: " + majorVersion);
                System.err.println("请配置IDE使用Java 17运行此测试");
                System.err.println("Java版本检查: " + javaVersion);
                System.exit(1);
            }
        } catch (Exception e) {
            System.err.println("警告：无法检查Java版本: " + e.getMessage());
        }
    }
    
    /**
     * 获取Java主版本号
     */
    private static int getJavaMajorVersion() {
        String version = System.getProperty("java.version");
        if (version.startsWith("1.")) {
            // Java 8及以下版本格式: 1.8.0_291
            version = version.substring(2, 3);
        } else {
            // Java 9及以上版本格式: 17.0.17
            int dot = version.indexOf(".");
            if (dot != -1) {
                version = version.substring(0, dot);
            }
        }
        return Integer.parseInt(version);
    }
    
    public static void main(String[] args) {
        System.out.println("=== 模板提取功能单步调试测试 ===");
        
        try {
            // 检查源文件是否存在
            String sourceFile = "test/1.2 安全生产方针政策.pptx";
            File source = new File(sourceFile);
            if (!source.exists()) {
                System.err.println("错误：源文件不存在: " + sourceFile);
                return;
            }
            
            // 测试1: 使用独立的模板提取工具提取模板
            System.out.println("\n[测试1] 使用模板提取工具提取模板...");
            System.out.println("  源文件: " + sourceFile);
            
            String templateFile = "templates/master_template.pptx";
            
            // 如果master_template.pptx已存在，先备份
            File masterTemplate = new File(templateFile);
            if (masterTemplate.exists()) {
                String backupFile = templateFile + ".backup";
                System.out.println("  备份现有模板文件到: " + backupFile);
                masterTemplate.renameTo(new File(backupFile));
            }
            
            // 设置断点：在TemplateExtractor.extractTemplate方法中单步调试模板提取流程
            // 从第5页到倒数第2页提取模板（endPage=-1表示倒数第2页）
            TemplateExtractor.extractTemplate(sourceFile, templateFile, 5, -1);
            System.out.println("✓ 模板提取完成");
            
            // 测试2: 验证模板文件是否生成
            System.out.println("\n[测试2] 验证模板文件...");
            if (masterTemplate.exists()) {
                long fileSize = masterTemplate.length();
                System.out.println("✓ 模板文件已生成: " + templateFile);
                System.out.println("  文件大小: " + fileSize + " 字节");
                
                // 设置断点：检查模板文件内容
                
                // 加载模板文件验证内容
                com.aspose.slides.Presentation templatePresentation = 
                    new com.aspose.slides.Presentation(templateFile);
                int slideCount = templatePresentation.getSlides().size();
                System.out.println("  模板幻灯片数量: " + slideCount);
                templatePresentation.dispose();
            } else {
                System.err.println("✗ 模板文件未生成");
            }
            
            // 测试3: 验证模板文字替换
            System.out.println("\n[测试3] 验证模板文字替换...");
            // 设置断点：检查模板文件中的文本内容是否为"模板文字模板文字模板文字"
            System.out.println("  请手动打开 " + templateFile + " 检查文本是否已替换为'模板文字'");
            
            // 测试4: 验证模板图片替换
            System.out.println("\n[测试4] 验证模板图片替换...");
            System.out.println("  请手动打开 " + templateFile + " 检查图片是否已替换为'No Image'占位符");
            
            System.out.println("\n=== 所有测试完成 ===");
            System.out.println("请手动检查 " + templateFile + " 的内容是否符合预期");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
