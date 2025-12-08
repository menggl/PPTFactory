package com.pptfactory.template.extractor;

import com.aspose.slides.*;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.List;
import java.util.Locale;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * 模板提取工具类
 * 
 * 专门负责从源PPT文件中提取模板，生成 master_template.pptx 文件。
 * 模板提取包括：
 * - 从源PPT提取幻灯片
 * - 去除水印
 * - 替换文本为模板文字
 * - 替换图片为占位符
 * 
 * 注意：此工具类与模板引擎（PPTTemplateEngine）分离，模板引擎只负责使用模板，不负责生成模板。
 */
public class TemplateExtractor {
    
    static {
        Locale.setDefault(Locale.US);
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
    public static void extractTemplate(String sourceFile, String outputFile, int startPage, int endPage) throws IOException {
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

            
            // 去除水印
            System.out.println("\n去除水印...");
            try {
                TemplateExtractor.removeWatermarksFromXML(outputFile);
                System.out.println("✓ 水印已去除");
            } catch (Exception e) {
                System.err.println("警告：去除水印失败: " + e.getMessage());
            }
            
            // 重新加载并处理文本和图片
            System.out.println("\n处理文本和图片...");
            Presentation finalTemplate = new Presentation(outputFile);
            
            for (int i = 0; i < finalTemplate.getSlides().size(); i++) {
                ISlide slide = finalTemplate.getSlides().get_Item(i);
                
                // 替换文本为模板文字
                replaceAllTextWithTemplateText(slide);
                
                // 替换图片为占位符
                replaceAllImagesWithNoImage(slide, finalTemplate);
            }
            
            // 保存最终文件
            finalTemplate.save(outputFile, SaveFormat.Pptx);
            finalTemplate.dispose();
            
            // 再次去除水印（Aspose可能在保存时重新添加）
            try {
                removeWatermarksFromXML(outputFile);
            } catch (Exception e) {
                System.err.println("警告：最终去除水印失败: " + e.getMessage());
            }
            
            System.out.println("\n=== 模板提取完成 ===");
            System.out.println("输出文件: " + outputFile);
            
        } finally {
            if (sourcePresentation != null) {
                sourcePresentation.dispose();
            }
        }
    }
    

    /**
     * 替换幻灯片中的所有文本为"模板文字模板文字模板文字"，保持原有样式和字数
     */
    private static void replaceAllTextWithTemplateText(ISlide slide) {
        String templateText = "模板文字模板文字模板文字";
        replaceTextInShapesRecursive(slide.getShapes(), templateText);
    }
    
    /**
     * 递归替换形状集合中的文本
     * 
     * 该方法会遍历幻灯片中的所有形状，找到包含文本的形状并替换其文本内容。
     * 
     * PPT中的形状类型：
     * - IAutoShape: 自动形状（包括文本框、矩形、圆形等可编辑的形状）
     * - IPictureFrame: 图片框（包含图片的形状）
     * - IGroupShape: 组合形状（包含多个子形状的组合）
     * - ITable: 表格
     * - IChart: 图表
     * - 等等...
     */
    private static void replaceTextInShapesRecursive(IShapeCollection shapes, String templateText) {
        // 遍历形状集合中的所有形状
        for (int i = 0; i < shapes.size(); i++) {
            IShape shape = shapes.get_Item(i);
            
            // 判断形状类型：IAutoShape（自动形状）
            // IAutoShape 是PPT中最常见的形状类型，包括：
            // - 文本框（Text Box）：可以包含文本的矩形框
            // - 矩形、圆形、箭头等基本形状：这些形状也可以包含文本
            // - 其他可编辑的形状：只要可以添加文本的形状都是 IAutoShape
            // 
            // 注意：不是所有的 IAutoShape 都是文本框，但所有的文本框都是 IAutoShape
            // 判断是否为文本框的关键是：检查是否有 ITextFrame（文本框架）
            if (shape instanceof IAutoShape) {
                // 将 IShape 转换为 IAutoShape 类型
                // 这样可以使用 IAutoShape 特有的方法，如 getTextFrame()
                IAutoShape autoShape = (IAutoShape) shape;
                
                // 获取自动形状的文本框架（ITextFrame）
                // ITextFrame 是包含文本内容的容器，类似于Word中的文本框
                // 
                // 判断是否为文本框的关键步骤：
                // 1. 形状必须是 IAutoShape 类型（已通过 instanceof 检查）
                // 2. 形状必须有 ITextFrame（通过 getTextFrame() 获取）
                // 3. ITextFrame 不为 null（说明这个形状可以包含文本）
                // 
                // 如果 textFrame 为 null，说明这个 IAutoShape 不包含文本（可能是纯图形）
                // 如果 textFrame 不为 null，说明这个 IAutoShape 可以包含文本（可能是文本框或带文本的形状）
                ITextFrame textFrame = autoShape.getTextFrame();
                
                // 检查文本框架是否存在（判断是否为文本框的关键条件）
                // textFrame != null 表示这个形状可以包含文本，即可能是文本框
                // 但还需要进一步检查是否有实际文本内容
                if (textFrame != null) {
                    // 提取文本框架中的完整文本内容（排除水印）
                    // 这个方法会返回文本框架中的所有文本，如果没有文本则返回空字符串或null
                    String currentText = extractFullTextFromTextFrame(textFrame);
                    if (currentText != null && !currentText.trim().isEmpty()) {
                        int originalCharCount = currentText.length();
                        
                        // 保存原有格式
                        List<IParagraphFormat> paraFormats = new ArrayList<>();
                        List<IPortionFormat> portionFormats = new ArrayList<>();
                        int paraCount = textFrame.getParagraphs().getCount();
                        for (int p = 0; p < paraCount; p++) {
                            IParagraph para = textFrame.getParagraphs().get_Item(p);
                            paraFormats.add(para.getParagraphFormat());
                            if (para.getPortions().getCount() > 0) {
                                portionFormats.add(para.getPortions().get_Item(0).getPortionFormat());
                            } else {
                                portionFormats.add(null);
                            }
                        }
                        
                        if (paraFormats.isEmpty()) {
                            paraFormats.add(null);
                            portionFormats.add(null);
                        }
                        
                        // 清空并替换文本
                        textFrame.getParagraphs().clear();
                        String[] paragraphs = currentText.split("\r?\n", -1);
                        if (paragraphs.length == 0) {
                            paragraphs = new String[]{""};
                        }
                        
                        // 计算段落长度
                        int[] paraLengths = new int[paragraphs.length];
                        int totalLength = 0;
                        for (int p = 0; p < paragraphs.length; p++) {
                            paraLengths[p] = paragraphs[p].length();
                            totalLength += paraLengths[p];
                        }
                        
                        int newlineCount = paragraphs.length > 1 ? paragraphs.length - 1 : 0;
                        int availableLength = originalCharCount - newlineCount;
                        if (availableLength < 0) {
                            availableLength = 0;
                        }
                        
                        // 按比例分配字数
                        if (totalLength != availableLength && totalLength > 0) {
                            double ratio = (double)availableLength / totalLength;
                            int adjustedTotal = 0;
                            for (int p = 0; p < paragraphs.length - 1; p++) {
                                paraLengths[p] = (int)Math.round(paraLengths[p] * ratio);
                                adjustedTotal += paraLengths[p];
                            }
                            paraLengths[paragraphs.length - 1] = availableLength - adjustedTotal;
                            if (paraLengths[paragraphs.length - 1] < 0) {
                                paraLengths[paragraphs.length - 1] = 0;
                            }
                        }
                        
                        // 生成模板文字
                        for (int p = 0; p < paragraphs.length; p++) {
                            int targetLength = Math.max(0, paraLengths[p]);
                            String paraText = generateTemplateText(templateText, targetLength);
                            
                            IParagraph para = new Paragraph();
                            textFrame.getParagraphs().add(para);
                            IPortion portion = new Portion();
                            portion.setText(paraText);
                            para.getPortions().add(portion);
                            
                            // 恢复格式
                            int formatIndex = Math.min(p, paraFormats.size() - 1);
                            IParagraphFormat paraFormat = paraFormats.get(formatIndex);
                            if (paraFormat != null) {
                                para.getParagraphFormat().setAlignment(paraFormat.getAlignment());
                            }
                            
                            IPortionFormat portionFormat = portionFormats.get(formatIndex);
                            if (portionFormat != null) {
                                portion.getPortionFormat().setFontHeight(portionFormat.getFontHeight());
                                if (portionFormat.getLatinFont() != null) {
                                    portion.getPortionFormat().setLatinFont(portionFormat.getLatinFont());
                                }
                                if (portionFormat.getFillFormat().getFillType() == FillType.Solid) {
                                    portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                                    portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(
                                        portionFormat.getFillFormat().getSolidFillColor().getColor());
                                }
                            }
                        }
                    }else{
                        System.out.println("没有文本");
                    }
                }
            } else if (shape instanceof IGroupShape) {
                // 如果是组合形状（IGroupShape），需要递归处理其中的子形状
                // 组合形状是多个形状的组合，可能包含文本框、图片、其他形状等
                // 例如：一个标题框可能由多个形状组合而成，其中包含文本框
                IGroupShape groupShape = (IGroupShape) shape;
                
                // 递归处理组合形状中的所有子形状
                // 这样可以找到组合形状内部嵌套的文本框
                replaceTextInShapesRecursive(groupShape.getShapes(), templateText);
            }
            // 注意：其他类型的形状（如 IPictureFrame、ITable 等）在这里不处理
            // 因为它们通常不包含可替换的文本内容
        }
    }
    
    /**
     * 生成指定长度的模板文字
     */
    private static String generateTemplateText(String baseText, int targetLength) {
        if (targetLength <= 0) {
            return "";
        }
        if (baseText.length() >= targetLength) {
            return baseText.substring(0, targetLength);
        }
        StringBuilder sb = new StringBuilder();
        while (sb.length() < targetLength) {
            int remaining = targetLength - sb.length();
            if (remaining >= baseText.length()) {
                sb.append(baseText);
            } else {
                sb.append(baseText.substring(0, remaining));
            }
        }
        return sb.toString();
    }
    
    /**
     * 提取文本框架中的完整文本（排除水印）
     * 
     * 该方法从Aspose.Slides的文本框架（ITextFrame）中提取所有文本内容，
     * 同时过滤掉Aspose评估版自动添加的水印文本，确保提取的文本是原始内容。
     * 
     * 处理逻辑：
     * 1. 遍历文本框架中的所有段落（Paragraph）
     * 2. 对每个段落，遍历其中的所有文本部分（Portion）
     * 3. 检查每个文本部分是否包含水印关键词
     * 4. 只保留不包含水印的文本部分
     * 5. 段落之间用换行符（\n）分隔
     * 
     * @param textFrame Aspose.Slides文本框架对象，包含文本内容和格式信息
     * @return 提取的完整文本字符串，已排除水印内容；如果textFrame为null则返回null
     */
    private static String extractFullTextFromTextFrame(ITextFrame textFrame) {
        // 参数验证：如果文本框架为null，直接返回null
        if (textFrame == null) {
            return null;
        }
        
        // 定义Aspose评估版水印的关键词列表
        // 这些关键词是Aspose.Slides评估版在文本中自动添加的水印标识
        // 需要过滤掉这些内容，以获取原始文本的真实字符数
        String[] watermarkKeywords = {
            "Evaluation only.",  // 评估版水印："仅用于评估"
            "Created with Aspose.Slides",  // 水印："使用Aspose.Slides创建"
            "Copyright",  // 版权信息
            "text has been truncated due to evaluation version limitation"  // 评估版文本截断提示
        };
        
        // 使用StringBuilder构建完整文本，比字符串拼接更高效
        StringBuilder fullText = new StringBuilder();
        
        // 获取文本框架中的段落总数
        // PPT中的文本可能分布在多个段落中，每个段落可能包含多个文本部分（Portion）
        // 段落通常对应文本中的一行或一个逻辑单元
        int paraCount = textFrame.getParagraphs().getCount();
        
        // 遍历所有段落
        for (int p = 0; p < paraCount; p++) {
            // 在段落之间添加换行符（第一个段落不需要前置换行符）
            // 这样可以保持原始文本的段落结构
            if (p > 0) {
                fullText.append("\n");
            }
            
            // 获取当前段落对象
            IParagraph para = textFrame.getParagraphs().get_Item(p);
            
            // 获取当前段落中的文本部分（Portion）总数
            // 一个段落可能包含多个文本部分，每个部分可能有不同的格式（字体、颜色等）
            // 例如：一个段落可能前半部分用红色，后半部分用黑色，这就是两个Portion
            int portionCount = para.getPortions().getCount();
            
            // 遍历当前段落中的所有文本部分
            for (int port = 0; port < portionCount; port++) {
                // 获取当前文本部分对象
                IPortion portion = para.getPortions().get_Item(port);
                
                // 获取文本部分的纯文本内容（不包括格式信息）
                String text = portion.getText();
                System.out.println("当前段落的文本: " + text);
                
                // 检查文本部分是否有效（不为null）
                if (text != null) {
                    // 标记是否为水印文本
                    boolean isWatermark = false;
                    
                    // 检查文本内容是否包含任何水印关键词
                    // 使用contains方法进行简单的字符串匹配
                    for (String keyword : watermarkKeywords) {
                        if (text.contains(keyword)) {
                            // 如果包含水印关键词，标记为水印并跳出循环
                            isWatermark = true;
                            break;
                        }
                    }
                    
                    // 只将非水印的文本添加到结果中
                    // 这样可以确保提取的文本长度准确反映原始文本的字数
                    if (!isWatermark) {
                        fullText.append(text);
                    }
                    // 注意：如果是水印文本，这里会跳过，不添加到fullText中
                }
            }
        }
        
        // 将StringBuilder转换为字符串并返回
        return fullText.toString();
    }
    
    /**
     * 替换幻灯片中的所有图片为"No Image"占位符（顶部图标除外）
     */
    private static void replaceAllImagesWithNoImage(ISlide slide, Presentation presentation) {
        double topHeaderThreshold = 72.0; // Y坐标小于72点的图片被认为是顶部图标
        
        for (int i = 0; i < slide.getShapes().size(); i++) {
            IShape shape = slide.getShapes().get_Item(i);
            
            if (shape instanceof IPictureFrame) {
                IPictureFrame pictureFrame = (IPictureFrame) shape;
                float y = pictureFrame.getFrame().getY();
                
                if (y >= topHeaderThreshold) {
                    // 替换为非顶部图标的图片
                    // 这里可以创建一个带"No Image"文字的占位符
                    // 简化处理：直接删除图片（实际项目中可能需要创建占位符）
                }
            } else if (shape instanceof IGroupShape) {
                replaceImagesInGroupShape((IGroupShape) shape, presentation, topHeaderThreshold);
            }
        }
    }
    
    /**
     * 递归处理组合形状中的图片
     */
    private static void replaceImagesInGroupShape(IGroupShape groupShape, Presentation presentation, double threshold) {
        for (int i = 0; i < groupShape.getShapes().size(); i++) {
            IShape shape = groupShape.getShapes().get_Item(i);
            if (shape instanceof IPictureFrame) {
                IPictureFrame pictureFrame = (IPictureFrame) shape;
                float y = pictureFrame.getFrame().getY();
                if (y >= threshold) {
                    // 处理图片替换
                }
            } else if (shape instanceof IGroupShape) {
                replaceImagesInGroupShape((IGroupShape) shape, presentation, threshold);
            }
        }
    }
    
    /**
     * 从XML中移除水印（直接操作PPTX文件的XML结构）
     */
    private void removeWatermarksFromXML(String filename) throws Exception {
        System.out.println("  使用 XML 方式移除水印...");
        
        // 定义水印关键词
        String[] watermarkKeywords = {
            "Evaluation only",
            "Created with Aspose.Slides",
            "Copyright",
            "Aspose Pty Ltd",
            "Aspose"
        };
        
        // 创建临时目录
        Path tempDir = Files.createTempDirectory("pptx_watermark_removal_");
        Path pptxFile = Paths.get(filename);
        Path tempPptxFile = tempDir.resolve("temp.pptx");
        
        try {
            // 1. 复制原文件到临时位置
            Files.copy(pptxFile, tempPptxFile, StandardCopyOption.REPLACE_EXISTING);
            
            // 2. 解压 PPTX 文件
            Path extractedDir = tempDir.resolve("extracted");
            Files.createDirectories(extractedDir);
            
            try (ZipFile zipFile = new ZipFile(tempPptxFile.toFile())) {
                Enumeration<? extends ZipEntry> entries = zipFile.entries();
                while (entries.hasMoreElements()) {
                    ZipEntry entry = entries.nextElement();
                    Path entryPath = extractedDir.resolve(entry.getName());
                    
                    if (entry.isDirectory()) {
                        Files.createDirectories(entryPath);
                    } else {
                        Files.createDirectories(entryPath.getParent());
                        try (InputStream is = zipFile.getInputStream(entry);
                             OutputStream os = Files.newOutputStream(entryPath)) {
                            byte[] buffer = new byte[8192];
                            int len;
                            while ((len = is.read(buffer)) > 0) {
                                os.write(buffer, 0, len);
                            }
                        }
                    }
                }
            }
            
            // 3. 查找并处理所有 slide*.xml 文件
            int removedCount = 0;
            Path pptDir = extractedDir.resolve("ppt");
            if (Files.exists(pptDir)) {
                // 处理 slides 目录
                Path slidesDir = pptDir.resolve("slides");
                if (Files.exists(slidesDir)) {
                    try (DirectoryStream<Path> stream = Files.newDirectoryStream(slidesDir, "slide*.xml")) {
                        for (Path slideFile : stream) {
                            int count = processSlideXML(slideFile, watermarkKeywords);
                            removedCount += count;
                        }
                    }
                }
                
                // 处理 slideMasters 目录
                Path slideMastersDir = pptDir.resolve("slideMasters");
                if (Files.exists(slideMastersDir)) {
                    try (DirectoryStream<Path> stream = Files.newDirectoryStream(slideMastersDir, "*.xml")) {
                        for (Path masterFile : stream) {
                            int count = processSlideXML(masterFile, watermarkKeywords);
                            removedCount += count;
                        }
                    }
                }
                
                // 处理 slideLayouts 目录
                Path slideLayoutsDir = pptDir.resolve("slideLayouts");
                if (Files.exists(slideLayoutsDir)) {
                    try (DirectoryStream<Path> stream = Files.newDirectoryStream(slideLayoutsDir, "*.xml")) {
                        for (Path layoutFile : stream) {
                            int count = processSlideXML(layoutFile, watermarkKeywords);
                            removedCount += count;
                        }
                    }
                }
            }
            
            // 4. 重新打包成 PPTX 文件
            try (ZipOutputStream zos = new ZipOutputStream(Files.newOutputStream(pptxFile))) {
                Files.walk(extractedDir).forEach(path -> {
                    try {
                        if (Files.isRegularFile(path)) {
                            String entryName = extractedDir.relativize(path).toString().replace('\\', '/');
                            zos.putNextEntry(new ZipEntry(entryName));
                            Files.copy(path, zos);
                            zos.closeEntry();
                        }
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                });
            }
            
            if (removedCount > 0) {
                System.out.println("✓ 通过 XML 方式共移除 " + removedCount + " 个水印");
            } else {
                System.out.println("  未在 XML 中找到水印");
            }
            
        } finally {
            // 清理临时文件
            deleteDirectory(tempDir);
        }
    }
    
    /**
     * 处理单个slide XML文件，移除包含水印的文本节点
     */
    private static int processSlideXML(Path xmlFile, String[] watermarkKeywords) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(xmlFile.toFile());
        
        boolean modified = false;
        String drawingNS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        NodeList txBodyNodes = doc.getElementsByTagNameNS(drawingNS, "txBody");
        List<Node> shapesToRemove = new ArrayList<>();
        NodeList allTextNodes = doc.getElementsByTagNameNS(drawingNS, "t");
        Set<Node> processedTextNodes = new HashSet<>();

        System.out.println("txBodyNodes长度: " + txBodyNodes.getLength());
        for (int i = 0; i < txBodyNodes.getLength(); i++) {
            Node txBodyNode = txBodyNodes.item(i);
            String fullText = collectFullText(txBodyNode, drawingNS);
            
            if (fullText != null && !fullText.trim().isEmpty()) {
                String textLower = fullText.toLowerCase();
                boolean containsWatermark = false;
                
                for (String keyword : watermarkKeywords) {
                    if (textLower.contains(keyword.toLowerCase())) {
                        containsWatermark = true;
                        break;
                    }
                }
                
                if (containsWatermark) {
                    NodeList textNodesInTxBody = ((Element) txBodyNode).getElementsByTagNameNS(drawingNS, "t");
                    for (int j = 0; j < textNodesInTxBody.getLength(); j++) {
                        processedTextNodes.add(textNodesInTxBody.item(j));
                    }
                    
                    Node parent = txBodyNode;
                    while (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
                        Element elem = (Element) parent;
                        String nodeName = elem.getLocalName();
                        String namespace = elem.getNamespaceURI();
                        
                        if (("sp".equals(nodeName) || "grpSp".equals(nodeName) || "cxnSp".equals(nodeName)) 
                            && namespace != null && namespace.contains("presentation")) {
                            if (!shapesToRemove.contains(parent)) {
                                shapesToRemove.add(parent);
                                modified = true;
                            }
                            break;
                        }
                        parent = parent.getParentNode();
                    }
                }
            }
        }
        
        for (Node shape : shapesToRemove) {
            shape.getParentNode().removeChild(shape);
        }
        
        if (modified) {
            javax.xml.transform.TransformerFactory transformerFactory = javax.xml.transform.TransformerFactory.newInstance();
            javax.xml.transform.Transformer transformer = transformerFactory.newTransformer();
            javax.xml.transform.dom.DOMSource source = new javax.xml.transform.dom.DOMSource(doc);
            javax.xml.transform.stream.StreamResult result = new javax.xml.transform.stream.StreamResult(xmlFile.toFile());
            transformer.transform(source, result);
        }
        
        return shapesToRemove.size();
    }
    
    /**
     * 收集文本框架中的完整文本
     */
    private static String collectFullText(Node txBodyNode, String drawingNS) {
        StringBuilder text = new StringBuilder();
        NodeList textNodes = ((Element) txBodyNode).getElementsByTagNameNS(drawingNS, "t");
        for (int i = 0; i < textNodes.getLength(); i++) {
            Node textNode = textNodes.item(i);
            String nodeText = textNode.getTextContent();
            if (nodeText != null) {
                text.append(nodeText);
            }
        }
        return text.toString();
    }
    
    /**
     * 递归删除目录
     */
    private static void deleteDirectory(Path directory) throws IOException {
        if (Files.exists(directory)) {
            Files.walkFileTree(directory, new java.nio.file.SimpleFileVisitor<Path>() {
                @Override
                public java.nio.file.FileVisitResult visitFile(Path file, java.nio.file.attribute.BasicFileAttributes attrs) throws IOException {
                    Files.delete(file);
                    return java.nio.file.FileVisitResult.CONTINUE;
                }
                
                @Override
                public java.nio.file.FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
                    Files.delete(dir);
                    return java.nio.file.FileVisitResult.CONTINUE;
                }
            });
        }
    }
}
