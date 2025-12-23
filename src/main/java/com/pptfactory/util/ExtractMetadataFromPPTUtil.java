package com.pptfactory.util;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 提取PPTX中每页的备注信息（JSON格式Metadata）并保存为独立的JSON文件
 */
public class ExtractMetadataFromPPTUtil {

    public static void main(String[] args) {
        // 默认PPT路径
        String pptPath = "/Users/menggl/workspace/PPTFactory/templates/type_purchase/master_template.pptx";
        
        // 如果提供了参数，则使用参数路径
        if (args.length > 0) {
            pptPath = args[0];
        }

        extractMetadata(pptPath);
    }

    public static void extractMetadata(String pptPath) {
        System.out.println("开始从PPT提取元数据: " + pptPath);
        File pptFile = new File(pptPath);
        if (!pptFile.exists()) {
            System.err.println("文件不存在: " + pptPath);
            return;
        }

        // 确定输出目录: 同级目录下的 metadata 文件夹
        Path parentDir = pptFile.toPath().getParent();
        Path metadataDir = parentDir.resolve("metadata");

        try {
            // 创建 metadata 目录
            if (!Files.exists(metadataDir)) {
                Files.createDirectories(metadataDir);
                System.out.println("创建目录: " + metadataDir);
            }

            try (FileInputStream fis = new FileInputStream(pptFile);
                 XMLSlideShow ppt = new XMLSlideShow(fis)) {

                List<XSLFSlide> slides = ppt.getSlides();
                System.out.println("共找到 " + slides.size() + " 页幻灯片");

                ObjectMapper mapper = new ObjectMapper();

                for (int i = 0; i < slides.size(); i++) {
                    XSLFSlide slide = slides.get(i);
                    int pageIndex = i + 1;
                    
                    // 获取备注信息
                    String notesText = getNotesText(slide);
                    
                    if (notesText == null || notesText.trim().isEmpty()) {
                        System.out.println("第 " + pageIndex + " 页没有备注信息，跳过");
                        continue;
                    }

                    try {
                        // 解析JSON
                        JsonNode rootNode = mapper.readTree(notesText);
                        
                        if (rootNode.isObject()) {
                            ObjectNode objectNode = (ObjectNode) rootNode;
                            
                            // 构造 template_id (Txxx)
                            String templateId = String.format("T%03d", pageIndex);
                            
                            // 更新字段
                            objectNode.put("template_id", templateId);
                            objectNode.put("page_index", pageIndex);
                            
                            // 保存文件
                            String fileName = templateId + ".json";
                            Path jsonPath = metadataDir.resolve(fileName);
                            
                            // 写入JSON文件 (格式化输出)
                            String jsonOutput = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(objectNode);
                            Files.write(jsonPath, jsonOutput.getBytes(StandardCharsets.UTF_8));
                            
                            System.out.println("已保存: " + fileName);
                        } else {
                            System.err.println("第 " + pageIndex + " 页备注不是JSON对象: " + notesText);
                        }
                        
                    } catch (Exception e) {
                        System.err.println("第 " + pageIndex + " 页备注解析JSON失败: " + e.getMessage());
                        System.err.println("原始内容: " + notesText);
                    }
                }
            }
            
            System.out.println("提取完成！");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getNotesText(XSLFSlide slide) {
        XSLFNotes notes = slide.getNotes();
        if (notes == null) {
            return null;
        }
        
        // 收集所有段落的文本
        StringBuilder sb = new StringBuilder();
        for (List<XSLFTextParagraph> paragraphs : notes.getTextParagraphs()) {
            for (XSLFTextParagraph p : paragraphs) {
                String text = p.getText();
                if (text != null && !text.trim().isEmpty()) {
                    sb.append(text).append("\n");
                }
            }
        }
        
        // 尝试另一种方式：如果上面的方式获取不到，可能备注在其它文本框里
        // 通常 notes.getTextParagraphs() 是对的，它是 List<List<XSLFTextParagraph>>
        
        return sb.toString().trim();
    }
}
