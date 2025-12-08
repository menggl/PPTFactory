package com.pptfactory.util;

import com.aspose.slides.ISlide;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.aspose.slides.exceptions.FileNotFoundException;
import java.io.File;
import java.util.Locale;

public class CopyPPTSlideUtil {
    static {
        Locale.setDefault(Locale.US);
    }

    public static void main(String[] args) {
        System.out.println("=== 模板提取功能单步调试测试 ===");
        try {
            String sourceFile = "test/1.2 安全生产方针政策.pptx";
            String templateFile = "templates/master_template.pptx";
            File source = new File(sourceFile);
            if (!source.exists()) {
                System.err.println("错误：源文件不存在: " + sourceFile);
                return;
            }
            extractTemplate(sourceFile, templateFile, 5, -1);
        } catch (Exception e) {
            System.err.println("测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    /**
     * 从源PPT文件提取模板，生成 master_template.pptx
     * 
     * @param sourceFile 源PPT文件路径（如 "1.2 安全生产方针政策.pptx"）
     * @param outputFile 输出模板文件路径（如 "templates/master_template.pptx"）
     * @param startPage 开始页码（从1开始，例如第5页）
     * @param endPage 结束页码（从1开始，例如倒数第2页，传入-1表示倒数第2页）
     * @throws IOException 如果文件操作失败
     */
    public static void extractTemplate(String sourceFile, String outputFile, int startPage, int endPage) throws Exception {
        System.out.println("=== 开始提取模板 ===");
        System.out.println("源文件: " + sourceFile);
        System.out.println("输出文件: " + outputFile);
        
        // 检查源文件是否存在
        File source = new File(sourceFile);
        if (!source.exists()) {
            throw new FileNotFoundException("源文件不存在: " + sourceFile);
        }
        
        // 创建输出目录
        File output = new File(outputFile);
        if (output.getParentFile() != null && !output.getParentFile().exists()) {
            output.getParentFile().mkdirs();
        }
        
        // 加载源PPT
        Presentation sourcePresentation = null;
        try {
            sourcePresentation = new Presentation(sourceFile);
            int totalSlides = sourcePresentation.getSlides().size();
            System.out.println("源PPT共有 " + totalSlides + " 页");
            
            // 确定实际结束页码
            int actualEndPage = endPage;
            if (endPage == -1 || endPage > totalSlides) {
                actualEndPage = totalSlides - 1; // 倒数第2页
            }
            
            if (startPage < 1 || startPage > totalSlides || actualEndPage < startPage) {
                throw new IllegalArgumentException("页码范围无效: " + startPage + " 到 " + actualEndPage);
            }
            
            System.out.println("提取范围: 第" + startPage + "页到第" + actualEndPage + "页（共" + (actualEndPage - startPage + 1) + "页）");
            
            // 创建模板PPT
            Presentation templatePresentation = new Presentation();
            
            // 设置幻灯片尺寸（与源PPT一致）
            templatePresentation.getSlideSize().setSize(
                (float)sourcePresentation.getSlideSize().getSize().getWidth(),
                (float)sourcePresentation.getSlideSize().getSize().getHeight(),
                sourcePresentation.getSlideSize().getType()
            );
            
            // 删除默认空白页
            if (templatePresentation.getSlides().size() > 0) {
                templatePresentation.getSlides().removeAt(0);
            }
            
            // 复制幻灯片（从 startPage-1 到 actualEndPage-1，因为索引从0开始）
            for (int i = startPage - 1; i < actualEndPage; i++) {
                ISlide slide = sourcePresentation.getSlides().get_Item(i);
                templatePresentation.getSlides().addClone(slide);
                System.out.println("  ✓ 已复制第" + (i + 1) + "页");
            }
            
            // 保存临时文件
            templatePresentation.save(outputFile, SaveFormat.Pptx);
            templatePresentation.dispose();
        } finally {
            if (sourcePresentation != null) {
                sourcePresentation.dispose();
            }
        }
    }
}
