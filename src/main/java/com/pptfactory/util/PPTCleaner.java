package com.pptfactory.util;

import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.*;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class PPTCleaner {

    public static void main(String[] args) throws Exception {
        cleanPPTX("/Users/menggl/workspace/PPTFactory/master_template.pptx", 
        "/Users/menggl/workspace/PPTFactory/master_template_clean.pptx");
        System.out.println("PPT清理完成！");
    }
    
    /**
     * 清理整个PPT文件
     */
    public static void cleanPPTX(String inputPath, String outputPath) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputPath);
             XMLSlideShow ppt = new XMLSlideShow(fis);
             FileOutputStream fos = new FileOutputStream(outputPath)) {
            
            // 1. 清理所有幻灯片
            for (XSLFSlide slide : ppt.getSlides()) {
                cleanSlide(slide);
            }
            
            // 2. 清理所有母版
            for (XSLFSlideMaster master : ppt.getSlideMasters()) {
                cleanMaster(master);
            }
            
            // 3. 清理所有版式
            for (XSLFSlideMaster master : ppt.getSlideMasters()) {
                for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                    cleanLayout(layout);
                }
            }
            
            // 4. 清理未使用的图片等资源
            cleanUnusedResources(ppt);
            
            // 保存
            ppt.write(fos);
            System.out.println("PPT清理完成！");
        }
    }
    
    /**
     * 清理单个幻灯片
     */
    private static void cleanSlide(XSLFSlide slide) {
        System.out.println("清理幻灯片: " + slide.getSlideName());
        
        // 收集需要删除的形状
        List<XSLFShape> shapesToRemove = new ArrayList<>();
        
        for (XSLFShape shape : slide.getShapes()) {
            if (shouldRemoveShape(shape)) {
                shapesToRemove.add(shape);
                System.out.println("  标记删除: " + getShapeInfo(shape));
            }
        }
        
        // 批量删除
        for (XSLFShape shape : shapesToRemove) {
            try {
                slide.removeShape(shape);
                System.out.println("  已删除: " + getShapeInfo(shape));
            } catch (Exception e) {
                System.err.println("  删除失败: " + e.getMessage());
            }
        }
        
        // 清理分组形状内的子形状
        cleanGroupShapes(slide);
    }
    
    /**
     * 判断形状是否需要删除
     */
    private static boolean shouldRemoveShape(XSLFShape shape) {
        // 1. 隐藏的形状
        if (isHiddenShape(shape)) {
            return true;
        }
        
        // 2. 空的文本框
        if (shape instanceof XSLFTextShape) {
            XSLFTextShape textShape = (XSLFTextShape) shape;
            String text = textShape.getText();
            if (text == null || text.trim().isEmpty()) {
                return true;
            }
        }
        
        // 3. 占位符（可选删除）
        if (shape instanceof XSLFTextShape) {
            XSLFTextShape textShape = (XSLFTextShape) shape;
            if (textShape.isPlaceholder()) {
                // 可以根据需要决定是否删除占位符
                String text = textShape.getText();
                if (text == null || text.trim().isEmpty()) {
                    return true; // 删除空的占位符
                }
            }
        }
        
        // 4. 其他无效形状
        if (isInvalidShape(shape)) {
            return true;
        }
        
        return false;
    }
    
    /**
     * 检查形状是否隐藏
     */
    private static boolean isHiddenShape(XSLFShape shape) {
        try {
            // 方法1：检查 visible 属性
            if (shape instanceof XSLFSimpleShape) {
                XSLFSimpleShape simpleShape = (XSLFSimpleShape) shape;
                // POI的XSLFSimpleShape未公开isHidden，可以从底层XML属性判断
                // 形状的<... hidden="1"/>属性 表示为隐藏
                org.apache.xmlbeans.XmlObject xmlObj = simpleShape.getXmlObject();
                if (xmlObj != null) {
                    String xml = xmlObj.xmlText();
                    if (xml.contains("hidden=\"1\"") || xml.contains("hidden='1'")) {
                        return true;
                    }
                }
            }
            // 方法2：检查透明度（完全透明视为隐藏）
            if (shape instanceof XSLFSimpleShape) {
                XSLFSimpleShape simpleShape = (XSLFSimpleShape) shape;
                // getFillPaint() may not be visible; try getFillColor() as a workaround
                Color color = simpleShape.getFillColor();
                if (color != null && color.getAlpha() == 0) {
                    return true; // 完全透明
                }
            }
        } catch (Exception e) {
            // 忽略异常
        }
        return false;
    }
    
    /**
     * 检查是否为无效形状
     */
    private static boolean isInvalidShape(XSLFShape shape) {
        try {
            // 1. 宽度或高度为0
            Rectangle2D bounds = shape.getAnchor();
            if (bounds != null && (bounds.getWidth() <= 0 || bounds.getHeight() <= 0)) {
                return true;
            }
            
            // 2. 位置在画布之外
            if (bounds != null && (bounds.getX() < -10000 || bounds.getY() < -10000)) {
                return true;
            }
            
            // 3. 损坏的图片
            if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape pic = (XSLFPictureShape) shape;
                XSLFPictureData picData = pic.getPictureData();
                if (picData == null || picData.getData() == null || picData.getData().length == 0) {
                    return true;
                }
            }
            
        } catch (Exception e) {
            // 如果获取属性时出错，可能形状已损坏
            return true;
        }
        return false;
    }
    
    /**
     * 清理分组形状
     */
    private static void cleanGroupShapes(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFGroupShape) {
                XSLFGroupShape group = (XSLFGroupShape) shape;
                cleanGroupShape(group);
            }
        }
    }
    
    private static void cleanGroupShape(XSLFGroupShape group) {
        List<XSLFShape> shapesToRemove = new ArrayList<>();
        
        for (XSLFShape shape : group.getShapes()) {
            if (shouldRemoveShape(shape)) {
                shapesToRemove.add(shape);
            }
            
            // 递归清理嵌套分组
            if (shape instanceof XSLFGroupShape) {
                cleanGroupShape((XSLFGroupShape) shape);
            }
        }
        
        // 从分组中删除
        for (XSLFShape shape : shapesToRemove) {
            try {
                group.removeShape(shape);
            } catch (Exception e) {
                // 忽略删除失败
            }
        }
    }
    
    /**
     * 清理母版
     */
    private static void cleanMaster(XSLFSlideMaster master) {
        System.out.println("清理母版: " + master.getXmlObject().toString());
        cleanShapesInContainer(master);
    }
    
    /**
     * 清理版式
     */
    private static void cleanLayout(XSLFSlideLayout layout) {
        System.out.println("清理版式: " + layout.getName());
        cleanShapesInContainer(layout);
    }
    
    /**
     * 通用清理容器中的形状
     */
    private static void cleanShapesInContainer(ShapeContainer<?, ?> container) {
        List<XSLFShape> shapesToRemove = new ArrayList<>();

        for (Object obj : container.getShapes()) {
            if (obj instanceof XSLFShape) {
                XSLFShape shape = (XSLFShape) obj;
                if (shouldRemoveShape(shape)) {
                    shapesToRemove.add(shape);
                }
            }
        }
        
        for (XSLFShape shape : shapesToRemove) {
            try {
                if (container instanceof XSLFSlide) {
                    ((XSLFSlide) container).removeShape(shape);
                } else if (container instanceof XSLFGroupShape) {
                    ((XSLFGroupShape) container).removeShape(shape);
                } else if (container instanceof XSLFSlideLayout) {
                    ((XSLFSlideLayout) container).removeShape(shape);
                } else if (container instanceof XSLFSlideMaster) {
                    ((XSLFSlideMaster) container).removeShape(shape);
                } else {
                    // 类型不匹配不再强转调用，兼容性更高，仅记录警告
                    System.err.println("Warning: Unable to remove shape from unknown container type: " + container.getClass().getName());
                }
            } catch (Exception e) {
                // 忽略异常但打印警告信息，便于后续排查
                System.err.println("Warning: Failed to remove shape: " + e.getMessage());
            }
        }
    }
    
    /**
     * 清理未使用的资源
     */
    private static void cleanUnusedResources(XMLSlideShow ppt) {
        // Apache POI 没有直接方法，需要手动处理
        // 或者使用底层XML操作
    }
    
    /**
     * 获取形状信息
     */
    private static String getShapeInfo(XSLFShape shape) {
        try {
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                return String.format("文本框[%s] 内容: '%s'", 
                    textShape.isPlaceholder() ? "占位符" : "普通",
                    textShape.getText());
            } else if (shape instanceof XSLFPictureShape) {
                return "图片形状";
            } else if (shape instanceof XSLFGroupShape) {
                return "分组形状";
            } else {
                return shape.getClass().getSimpleName();
            }
        } catch (Exception e) {
            return "未知形状";
        }
    }
}
