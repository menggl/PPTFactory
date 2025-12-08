package com.pptfactory.ai;

import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Locale;

/**
 * LayoutClassifier 单步调试测试类
 * 
 * 用于测试布局分类器的功能，包括：
 * - 布局分类器初始化
 * - 单个内容布局分类
 * - 批量幻灯片布局分类
 * 
 * 使用方法：
 * 1. 在IDE中打开此文件
 * 2. 在main方法中设置断点
 * 3. 以Debug模式运行
 * 4. 单步调试查看布局分类的每个步骤
 */
public class TestLayoutClassifier {
    
    static {
        Locale.setDefault(Locale.US);
    }
    
    public static void main(String[] args) {
        System.out.println("=== LayoutClassifier 单步调试测试 ===");
        
        try {
            // 测试1: 创建布局分类器实例
            System.out.println("\n[测试1] 创建布局分类器实例...");
            LayoutClassifier classifier = new LayoutClassifier();
            System.out.println("✓ 布局分类器创建成功");
            
            // 设置断点：检查分类器状态
            
            // 测试2: 测试标题页布局分类
            System.out.println("\n[测试2] 测试标题页布局分类...");
            Map<String, Object> titlePageContent = new HashMap<>();
            titlePageContent.put("title", "安全生产培训");
            titlePageContent.put("subtitle", "2024年度");
            
            // 设置断点：在classifyLayout方法中单步调试
            String layout1 = classifier.classifyLayout(titlePageContent);
            System.out.println("  内容: " + titlePageContent);
            System.out.println("  分类结果: " + layout1);
            System.out.println("  ✓ 标题页布局分类完成");
            
            // 测试3: 测试内容页布局分类
            System.out.println("\n[测试3] 测试内容页布局分类...");
            Map<String, Object> contentPageContent = new HashMap<>();
            contentPageContent.put("title", "安全规定");
            contentPageContent.put("bullets", Arrays.asList("规定1", "规定2", "规定3"));
            
            String layout2 = classifier.classifyLayout(contentPageContent);
            System.out.println("  内容: " + contentPageContent);
            System.out.println("  分类结果: " + layout2);
            System.out.println("  ✓ 内容页布局分类完成");
            
            // 测试4: 测试图片+文字布局分类
            System.out.println("\n[测试4] 测试图片+文字布局分类...");
            Map<String, Object> imageTextContent = new HashMap<>();
            imageTextContent.put("title", "安全设备");
            imageTextContent.put("text", "这是安全设备的说明");
            imageTextContent.put("image_path", "/path/to/image.jpg");
            
            String layout3 = classifier.classifyLayout(imageTextContent);
            System.out.println("  内容: " + imageTextContent);
            System.out.println("  分类结果: " + layout3);
            System.out.println("  ✓ 图片+文字布局分类完成");
            
            // 测试5: 批量分类幻灯片
            System.out.println("\n[测试5] 批量分类幻灯片...");
            Map<String, Object> slidesData = new HashMap<>();
            List<Map<String, Object>> slides = new ArrayList<>();
            
            Map<String, Object> slide1 = new HashMap<>();
            slide1.put("title", "第一页");
            slide1.put("subtitle", "副标题");
            slides.add(slide1);
            
            Map<String, Object> slide2 = new HashMap<>();
            slide2.put("title", "第二页");
            slide2.put("bullets", Arrays.asList("要点1", "要点2"));
            slides.add(slide2);
            
            slidesData.put("slides", slides);
            
            // 设置断点：在autoClassifySlides方法中单步调试
            Map<String, Object> result = classifier.autoClassifySlides(slidesData);
            System.out.println("  输入幻灯片数量: " + slides.size());
            
            @SuppressWarnings("unchecked")
            List<Map<String, Object>> resultSlides = (List<Map<String, Object>>) result.get("slides");
            System.out.println("  输出幻灯片数量: " + resultSlides.size());
            
            for (int i = 0; i < resultSlides.size(); i++) {
                Map<String, Object> slide = resultSlides.get(i);
                System.out.println("  幻灯片 " + (i + 1) + " 布局: " + slide.get("layout"));
            }
            
            System.out.println("  ✓ 批量分类完成");
            
            System.out.println("\n=== 所有测试完成 ===");
            
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
