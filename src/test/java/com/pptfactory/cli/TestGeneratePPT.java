package com.pptfactory.cli;

import java.io.File;
import java.util.Locale;

/**
 * GeneratePPT CLI 单步调试测试类
 * 
 * 用于测试CLI入口的功能，包括：
 * - 命令行参数解析
 * - 输入文件加载
 * - 风格策略选择
 * - 模板文件选择
 * - 完整的PPT生成流程
 * 
 * 使用方法：
 * 1. 确保使用Java 17或更高版本运行（项目要求Java 17+）
 * 2. 在IDE中打开此文件
 * 3. 在main方法中设置断点
 * 4. 以Debug模式运行
 * 5. 单步调试查看CLI处理的每个步骤
 * 
 * IDE配置说明：
 * - IntelliJ IDEA: File -> Project Structure -> Project -> SDK 选择 Java 17
 * - Eclipse: Window -> Preferences -> Java -> Installed JREs 添加 Java 17
 */
public class TestGeneratePPT {
    
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
        System.out.println("=== GeneratePPT CLI 单步调试测试 ===");
        
        try {
            // 测试1: 模拟命令行参数
            System.out.println("\n[测试1] 模拟命令行参数...");
            String[] testArgs = {
                "examples/safety_slides_extended.json",
                "-o", "test_output_cli.pptx",
                "--style", "safety",
                "--template", "safety"
            };
            
            System.out.println("  命令行参数:");
            for (int i = 0; i < testArgs.length; i++) {
                System.out.println("    [" + i + "] " + testArgs[i]);
            }
            
            // 设置断点：在GeneratePPT.main方法中单步调试参数解析
            
            // 测试2: 调用主函数
            System.out.println("\n[测试2] 调用GeneratePPT.main...");
            GeneratePPT.main(testArgs);
            
            // 测试3: 验证输出文件
            System.out.println("\n[测试3] 验证输出文件...");
            File outputFile = new File("test_output_cli.pptx");
            if (outputFile.exists()) {
                long fileSize = outputFile.length();
                System.out.println("✓ 输出文件已生成: " + outputFile.getName());
                System.out.println("  文件大小: " + fileSize + " 字节");
            } else {
                System.err.println("✗ 输出文件未生成");
            }
            
            System.out.println("\n=== 所有测试完成 ===");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
