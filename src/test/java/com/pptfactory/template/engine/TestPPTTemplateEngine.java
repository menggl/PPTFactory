package com.pptfactory.template.engine;

import com.pptfactory.style.StyleStrategy;
import com.pptfactory.style.SafetyStyle;
import com.aspose.slides.Presentation;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;

import java.io.File;
import java.util.Map;
import java.util.Locale;

/**
 * PPTTemplateEngine 单步调试测试类
 * 
 * 用于测试模板引擎的核心功能，包括：
 * - 模板加载
 * - 风格策略应用
 * - PPT渲染
 * - 水印移除
 * 
 * 使用方法：
 * 1. 确保使用Java 17或更高版本运行（项目要求Java 17+）
 * 2. 在IDE中打开此文件
 * 3. 在main方法中设置断点
 * 4. 以Debug模式运行
 * 5. 单步调试查看每个步骤的执行情况
 * 
 * IDE配置说明：
 * - IntelliJ IDEA: File -> Project Structure -> Project -> SDK 选择 Java 17
 * - Eclipse: Window -> Preferences -> Java -> Installed JREs 添加 Java 17
 */
public class TestPPTTemplateEngine {
    
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
        System.out.println("=== PPTTemplateEngine 单步调试测试 ===");
        
        try {
            // 测试1: 创建模板引擎实例
            System.out.println("\n[测试1] 创建模板引擎实例...");
            String templateFile = "templates/master_template.pptx";
            StyleStrategy styleStrategy = new SafetyStyle();
            
            // 设置断点：检查模板文件是否存在
            File template = new File(templateFile);
            if (!template.exists()) {
                System.err.println("错误：模板文件不存在: " + templateFile);
                System.err.println("请确保已运行模板提取流程生成 master_template.pptx");
                return;
            }
            
            PPTTemplateEngine engine = new PPTTemplateEngine(templateFile, styleStrategy);
            System.out.println("✓ 模板引擎创建成功");
            
            // 测试2: 加载测试数据
            System.out.println("\n[测试2] 加载测试数据...");
            String testDataFile = "examples/safety_slides_extended.json";
            ObjectMapper mapper = new ObjectMapper();
            Map<String, Object> slidesData = mapper.readValue(
                new File(testDataFile),
                new TypeReference<Map<String, Object>>() {}
            );
            System.out.println("✓ 测试数据加载成功，共 " + slidesData.size() + " 个字段");
            
            // 设置断点：查看slidesData的内容结构
            System.out.println("数据字段: " + slidesData.keySet());
            
            // 测试3: 渲染PPT
            System.out.println("\n[测试3] 渲染PPT...");
            // 设置断点：在renderFromJson方法内部单步调试
            engine.renderFromJson(slidesData);
            System.out.println("✓ PPT渲染完成");
            
            // 测试4: 保存PPT
            System.out.println("\n[测试4] 保存PPT...");
            String outputFile = "test_output_engine.pptx";
            engine.save(outputFile);
            System.out.println("✓ PPT保存成功: " + outputFile);
            
            // 测试5: 关闭引擎
            System.out.println("\n[测试5] 关闭引擎...");
            engine.close();
            System.out.println("✓ 引擎关闭成功");
            
            System.out.println("\n=== 所有测试完成 ===");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
