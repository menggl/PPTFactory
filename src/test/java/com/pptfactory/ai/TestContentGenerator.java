package com.pptfactory.ai;

import java.util.Map;
import java.util.List;
import java.util.Locale;

/**
 * ContentGenerator 单步调试测试类
 * 
 * 用于测试内容生成器的功能，包括：
 * - 内容生成器初始化
 * - 用户输入处理
 * - 结构化内容生成
 * 
 * 使用方法：
 * 1. 在IDE中打开此文件
 * 2. 在main方法中设置断点
 * 3. 以Debug模式运行
 * 4. 单步调试查看内容生成的每个步骤
 */
public class TestContentGenerator {
    
    static {
        Locale.setDefault(Locale.US);
    }
    
    public static void main(String[] args) {
        System.out.println("=== ContentGenerator 单步调试测试 ===");
        
        try {
            // 测试1: 创建内容生成器实例
            System.out.println("\n[测试1] 创建内容生成器实例...");
            String modelName = "gpt-4";
            ContentGenerator generator = new ContentGenerator(modelName);
            System.out.println("✓ 内容生成器创建成功，模型: " + modelName);
            
            // 设置断点：检查生成器状态
            
            // 测试2: 生成PPT内容
            System.out.println("\n[测试2] 生成PPT内容...");
            String userInput = "请生成一个关于安全生产的PPT，包含标题页和3个内容页";
            System.out.println("  用户输入: " + userInput);
            
            // 设置断点：在generateSlides方法中单步调试
            Map<String, Object> result = generator.generateSlides(userInput);
            System.out.println("✓ 内容生成完成");
            
            // 测试3: 验证生成的内容结构
            System.out.println("\n[测试3] 验证生成的内容结构...");
            Object slidesObj = result.get("slides");
            if (slidesObj instanceof List) {
                @SuppressWarnings("unchecked")
                List<Map<String, Object>> slides = (List<Map<String, Object>>) slidesObj;
                System.out.println("  幻灯片数量: " + slides.size());
                
                for (int i = 0; i < slides.size(); i++) {
                    Map<String, Object> slide = slides.get(i);
                    System.out.println("  幻灯片 " + (i + 1) + ":");
                    System.out.println("    布局: " + slide.get("layout"));
                    System.out.println("    标题: " + slide.get("title"));
                    
                    // 设置断点：查看每个幻灯片的内容
                }
            }
            
            System.out.println("\n=== 所有测试完成 ===");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
