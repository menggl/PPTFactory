package com.pptfactory.util;

import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.Notes;
import java.io.FileInputStream;
import java.util.List;

public class Test {
    
    public static void main(String[] args) {
        String pptPath = "/Users/menggl/workspace/PPTFactory/templates/type_purchase/master_template.pptx";
        
        try (FileInputStream fis = new FileInputStream(pptPath);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {
            
            // 获取所有幻灯片
            List<XSLFSlide> slides = ppt.getSlides();
            
            for (int i = 0; i < slides.size(); i++) {
                XSLFSlide slide = slides.get(i);
                int slideNumber = i + 1;
                
                // 获取备注页
                XSLFNotes notes = slide.getNotes();
                
                if (notes != null) {
                    // 提取备注文本
                    StringBuilder notesText = new StringBuilder();
                    for (XSLFTextShape shape : notes.getPlaceholders()) {
                        if (shape.getText() != null && !shape.getText().trim().isEmpty()) {
                            notesText.append(shape.getText()).append("\n");
                        }
                    }
                    
                    if (notesText.length() > 0) {
                        System.out.println("幻灯片 " + slideNumber + " 的备注:");
                        System.out.println(notesText.toString());
                        System.out.println("------------------------");
                    }
                }
            }
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
