package com.pptfactory.util;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.xmlbeans.XmlObject;

public class CleanShapeUtil {
    // 所有文本跟这些不相同的shape都要清除
    private static final Map<Integer,List<String>> textMap = new HashMap<>();
    static{
        textMap.put(0, Arrays.asList("一、煤矿安全生产核心方针", "安全第一", "预防为主", "综合治理"));
        textMap.put(1, Arrays.asList("一、煤矿安全生产核心方针",
            "1.“安全第一”：优先保障生命安全", 
            "在煤矿生产全流程中，必须将人员安全置于产量、效率之上",
            "在煤矿生产的整个过程里，人员安全的重要性要远超产量和效率。对采煤机司机而言，当操作中发现顶底板冒顶预兆、瓦斯超限等安全隐患时，需立即停机处理，严禁“冒险作业”“带病开机”。",
            "对应岗位操作中的“紧急停机情形”",
            "在岗位操作方面，当遇到顶底板有冒顶预兆等情况时，采煤机司机必须立即停机。这体现了“安全第一”的方针，将人员安全放在首位，避免因继续作业而导致安全事故。"));
    }

    private static final Set<String> shapeIdSet = new HashSet<>();
    
    public static void main(String[] args) {
        cleanUnuseShapes("/Users/menggl/workspace/PPTFactory/templates/master_template.pptx");
    }

    public static void cleanUnuseShapes(String path) {
        try (FileInputStream fis = new FileInputStream(path);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {

            int page = 0;
            for (XSLFSlide slide : ppt.getSlides()) {
                shapeIdSet.clear(); // 每页重置，避免跨页去重
                deleteShapesFromSlide(slide, page);
                page++;
            }
            
            // 保存文件
            // String outputPath = path.replace(".pptx", "_cleaned.pptx");
            try (FileOutputStream fos = new FileOutputStream(path)) {
                ppt.write(fos);
            }
            System.out.println("清理完成，输出文件: " + path);
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 从 slide 中删除需要删除的形状（包括 GroupShape 中的子形状）
     */
    private static void deleteShapesFromSlide(XSLFSlide slide, int page) {
        List<XSLFShape> shapesToRemove = new ArrayList<>();
        
        for (XSLFShape shape : slide.getShapes()) {
            // 处理 GroupShape 中的子形状
            if (shape instanceof XSLFGroupShape) {
                deleteShapesFromGroup((XSLFGroupShape) shape, page);
            } else {
                // 检查是否需要删除
                if (shouldDeleteShape(shape, page)) {
                    shapesToRemove.add(shape);
                }
            }
        }
        
        // 删除收集到的形状
        for (XSLFShape shape : shapesToRemove) {
            slide.removeShape(shape);
        }
    }
    
    /**
     * 从 GroupShape 中删除需要删除的子形状
     */
    private static void deleteShapesFromGroup(XSLFGroupShape groupShape, int page) {
        List<XSLFShape> shapesToRemove = new ArrayList<>();
        
        for (XSLFShape child : groupShape.getShapes()) {
            // 如果子形状也是 GroupShape，递归处理
            if (child instanceof XSLFGroupShape) {
                deleteShapesFromGroup((XSLFGroupShape) child, page);
            } else {
                // 检查是否需要删除
                if (shouldDeleteShape(child, page)) {
                    shapesToRemove.add(child);
                }
            }
        }
        
        // 删除收集到的子形状
        for (XSLFShape shape : shapesToRemove) {
            groupShape.removeShape(shape);
        }
    }
    
    /**
     * 判断形状是否应该删除
     */
    private static boolean shouldDeleteShape(XSLFShape shape, int page) {
        // 如果 shape 是 GroupShape，不删除（GroupShape 本身不删除，只删除其中的子形状）
        if (shape instanceof XSLFGroupShape) {
            return false;
        }
        
        // 如果形状之前遍历过，则忽略
        String shapeId = extractXmlId(shape);
        if (shapeId == null || shapeIdSet.contains(shapeId)) return false;
        shapeIdSet.add(shapeId);
        
        // 如果 shape 的名称包含 SmartArt （智能图形/图表）或 组合，则忽略
        String name = shape.getShapeName();
        if (name != null && (name.contains("SmartArt") || name.contains("组合"))) {
            return false;
        }

        if (shape instanceof XSLFTextShape ts) {
            if (ts.getPlaceholder() != null) return false; //过滤母版/布局占位符

            // 如果一个文本框里包含多个不同的小文本框内容，且 shape 名称像自动生成的
            List<XSLFTextParagraph> paragraphs = ts.getTextParagraphs();
            if (paragraphs != null && paragraphs.size() > 1 && 
                shape.getShapeName() != null && shape.getShapeName().matches("文本框 \\d+")) {
                // 这是 GroupShape 的汇总文本，忽略
                return false;
            }

            String txt = ts.getText();
            if(txt == null || txt.trim().isEmpty()) return false;
            System.out.println("遍历出的文本: ("+page+"):" + txt);

            List<String> textList = textMap.get(page);
            if(textList == null || textList.isEmpty()) return false;

            boolean contains = false;
            for (String text : textList) {
                if(txt.equals(text)) {
                    contains = true;
                    break;
                }
            }
            if(!contains){
                System.out.println("需要删除的shape: " + shape.getShapeName()+", page: " + page + ", txt: " + txt);
            }
            return !contains;
        }
        return false;
    }
    /**
     * 从XML中提取形状ID
     */
    private static String extractXmlId(XSLFShape shape) {
        try {
            XmlObject xmlObj = shape.getXmlObject();
            String xml = xmlObj.xmlText();
            
            // 查找 id="xxx" 或 id='xxx'
            java.util.regex.Pattern pattern = 
                java.util.regex.Pattern.compile("id\\s*=\\s*[\"']([^\"']+)[\"']");
            java.util.regex.Matcher matcher = pattern.matcher(xml);
            
            if (matcher.find()) {
                return matcher.group(1);
            }
        } catch (Exception e) {
            // 忽略异常
        }
        return null;
    }
}
