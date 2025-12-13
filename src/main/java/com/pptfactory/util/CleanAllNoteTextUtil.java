package com.pptfactory.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * 遍历PPTX文件中的所有文本
 */
public class CleanAllNoteTextUtil {
    public static void main(String[] args) {
        cleanAllNoteText("/Users/menggl/workspace/PPTFactory/templates/master_template.pptx");
    }

    public static void cleanAllNoteText(String filename) {
        try (FileInputStream fis = new FileInputStream(filename)) {
            XMLSlideShow ppt = new XMLSlideShow(fis);
            int slideIndex = 0;
            int totalNotesDeleted = 0;
            int totalSlideShapesDeleted = 0;
            
            for (XSLFSlide slide : ppt.getSlides()) {
                System.out.println("处理幻灯片索引: " + slideIndex);
                
                // 1. 删除备注页中的所有内容
                try {
                    XSLFNotes notes = slide.getNotes();
                    if (notes != null) {
                        List<XSLFShape> notesShapesToRemove = new ArrayList<>();
                        for (XSLFShape noteShape : notes.getShapes()) {
                            if (noteShape instanceof XSLFTextShape) {
                                XSLFTextShape noteTextShape = (XSLFTextShape) noteShape;
                                String noteText = noteTextShape.getText();
                                if (noteText != null && !noteText.trim().isEmpty()) {
                                    System.out.println("  发现备注文本: " + (noteText.length() > 50 ? noteText.substring(0, 50) + "..." : noteText));
                                    notesShapesToRemove.add(noteShape);
                                }
                            } else {
                                // 删除备注页中的所有形状（包括非文本形状）
                                notesShapesToRemove.add(noteShape);
                            }
                        }
                        
                        // 删除备注页中的形状
                        for (XSLFShape shape : notesShapesToRemove) {
                            try {
                                notes.removeShape(shape);
                                totalNotesDeleted++;
                                System.out.println("  ✓ 已删除备注页中的形状");
                            } catch (Exception e) {
                                System.err.println("  ✗ 删除备注页形状失败: " + e.getMessage());
                            }
                        }
                    }
                } catch (Exception e) {
                    System.err.println("  访问备注页时出错: " + e.getMessage());
                }
                
                // // 2. 删除幻灯片上包含"详细描述"等关键词的备注文本
                // List<XSLFShape> shapesToRemove = new ArrayList<>();
                
                // for (XSLFShape shape : slide.getShapes()) {
                //     if (shape instanceof XSLFTextShape) {
                //         XSLFTextShape textShape = (XSLFTextShape) shape;
                //         String text = textShape.getText();
                        
                //         if (text != null && !text.trim().isEmpty()) {
                //             String lowerText = text.toLowerCase();
                //             // 检查是否包含备注相关的关键词
                //             if (lowerText.contains("详细描述") || lowerText.contains("备注") || 
                //                 lowerText.contains("提示") || lowerText.contains("说明") ||
                //                 lowerText.contains("注意") || lowerText.contains("建议") ||
                //                 lowerText.contains("讲解") || lowerText.contains("补充")) {
                //                 shapesToRemove.add(shape);
                //                 System.out.println("  标记删除备注文本: " + (text.length() > 50 ? text.substring(0, 50) + "..." : text));
                //             }
                //         }
                //     }
                // }
                
                // // 批量删除收集到的形状，使用slide.removeShape()确保从底层XML中删除
                // for (XSLFShape shape : shapesToRemove) {
                //     try {
                //         slide.removeShape(shape);
                //         totalSlideShapesDeleted++;
                //         System.out.println("  ✓ 已删除幻灯片上的备注形状");
                //     } catch (Exception e) {
                //         System.err.println("  ✗ 删除失败: " + e.getMessage());
                //         e.printStackTrace();
                //     }
                // }
                
                System.out.println("--------------------------------\n");
                slideIndex++;
            }
            
            System.out.println("\n删除完成！");
            System.out.println("总计删除备注页形状: " + totalNotesDeleted + " 个");
            System.out.println("总计删除幻灯片备注形状: " + totalSlideShapesDeleted + " 个");
            
            ppt.write(new FileOutputStream(filename));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
