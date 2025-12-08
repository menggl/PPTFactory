package com.pptfactory.lm;

import java.util.*;

/**
 * 布局分类器类
 * 
 * 负责根据内容特征自动选择合适的PPT布局类型。
 * 可以使用大模型、规则引擎或机器学习模型来实现分类逻辑。
 */
public class LayoutClassifier {
    
    private List<String> availableLayouts;
    
    /**
     * 初始化布局分类器
     */
    public LayoutClassifier() {
        // 可用的布局类型
        this.availableLayouts = Arrays.asList(
            "title_page",      // 标题页
            "content_page",    // 内容页（标题+要点列表）
            "two_column",      // 两列布局
            "image_with_text"  // 图片+文字布局
        );
    }
    
    /**
     * 根据内容特征分类布局类型
     * 
     * @param content 内容Map，包含title、text等字段
     * @return 布局类型名称
     */
    public String classifyLayout(Map<String, Object> content) {
        // 简单的规则分类逻辑
        // TODO: 实际实现中可以使用大模型或更复杂的分类算法
        
        // 如果包含图片路径，使用图片+文字布局
        if (content.containsKey("image_path") && content.get("image_path") != null) {
            return "image_with_text";
        }
        
        // 如果包含左右两列内容，使用两列布局
        if (content.containsKey("left_content") || content.containsKey("right_content")) {
            return "two_column";
        }
        
        // 如果只有标题和副标题，使用标题页布局
        if (content.containsKey("title") && content.containsKey("subtitle") && 
            !content.containsKey("bullets")) {
            return "title_page";
        }
        
        // 默认使用内容页布局
        return "content_page";
    }
    
    /**
     * 自动为所有幻灯片分类布局类型
     * 
     * @param slidesData 包含slides数组的Map，可能没有layout字段
     * @return 添加了layout字段的slides数据
     */
    @SuppressWarnings("unchecked")
    public Map<String, Object> autoClassifySlides(Map<String, Object> slidesData) {
        Map<String, Object> result = new HashMap<>();
        List<Map<String, Object>> slides = new ArrayList<>();
        
        Object slidesObj = slidesData.get("slides");
        if (slidesObj instanceof List) {
            for (Object slideObj : (List<?>) slidesObj) {
                if (slideObj instanceof Map) {
                    Map<String, Object> slide = new HashMap<>((Map<String, Object>) slideObj);
                    // 如果已经有layout字段，保留原值
                    if (!slide.containsKey("layout")) {
                        slide.put("layout", classifyLayout(slide));
                    }
                    slides.add(slide);
                }
            }
        }
        
        result.put("slides", slides);
        return result;
    }
}

