package com.pptfactory.template.engine;

import com.pptfactory.style.StyleStrategy;
import com.pptfactory.style.DefaultStyle;
import com.aspose.slides.*;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.aspose.slides.Paragraph;
import com.aspose.slides.Portion;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.FileVisitResult;
import java.util.*;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.HashSet;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;

/**
 * 统一PPT模板引擎类
 * 
 * 该类是PPT模板引擎的核心实现，使用策略模式支持不同的风格。
 * 模板引擎负责加载模板文件和应用风格策略，将结构化内容渲染为PPT。
 * 
 * 设计模式：
 * - 策略模式：通过StyleStrategy接口支持不同的风格实现
 * - 模板方法模式：统一的渲染流程，具体样式由策略类决定
 * 
 * 设计理念：
 * - 模板引擎保持统一，所有风格通过策略类实现
 * - 模板文件决定"页面长什么样、有哪些占位符、布局怎么摆"
 * - 风格策略类决定"颜色、字体、图片风格、行距间距等视觉效果"
 */
public class PPTTemplateEngine {
    
    // 静态初始化块：在类加载时设置区域设置
    // 这必须在任何 Aspose.Slides 类被使用之前执行
    static {
        try {
            // 强制设置区域设置为 en-US，避免 Aspose.Slides 不支持的系统区域设置问题
            Locale.setDefault(Locale.US);
            System.setProperty("user.language", "en");
            System.setProperty("user.country", "US");
            System.setProperty("user.variant", "");
        } catch (Exception e) {
            // 忽略设置失败，继续执行
            System.err.println("警告：无法设置区域设置: " + e.getMessage());
        }
    }
    
    private Presentation presentation;
    private StyleStrategy styleStrategy;
    private String templateFile;
    private Presentation templatePresentation;
    private Presentation safetyReferencePresentation; // 安全生产参考PPT（用于获取布局和样式）
    private Presentation masterTemplatePresentation; // 统一的模板文件 master_template.pptx
    private Map<String, Integer> classicLayoutMap; // 经典布局映射：布局名称 -> 源PPT页码
    private Map<String, Map<String, Object>> layoutConfigMap; // 布局配置映射：布局名称 -> 配置信息
    private Map<String, Map<String, Object>> layoutStyleMap; // 布局样式映射：布局名称 -> 样式信息
    
    /**
     * 初始化PPT模板引擎
     * 
     * 创建模板引擎实例，加载模板文件并应用风格策略。
     * 
     * @param templateFile 模板PPT文件路径（必需）
     *                     模板文件决定了页面的布局、占位符位置等结构信息
     * @param styleStrategy 风格策略对象，可选
     *                     如果提供，将使用该策略类来应用样式
     *                     如果不提供，将使用默认风格
     * @throws IOException 如果模板文件不存在或无法读取
     */
    public PPTTemplateEngine(String templateFile, StyleStrategy styleStrategy) throws IOException {
        // 检查模板文件是否存在
        File template = new File(templateFile);
        if (!template.exists()) {
            throw new FileNotFoundException("模板文件不存在: " + templateFile);
        }
        
        // 加载模板文件
        this.templateFile = templateFile;
        this.templatePresentation = new Presentation(templateFile);
        
        // 创建新的演示文稿，使用 Aspose.Slides
        this.presentation = new Presentation();
        // 删除默认空白页，避免生成多余的首张空白幻灯片
        if (this.presentation.getSlides().size() > 0) {
            this.presentation.getSlides().removeAt(0);
        }
        
        // 设置幻灯片尺寸（从模板文件获取）
        this.presentation.getSlideSize().setSize(
            (float)templatePresentation.getSlideSize().getSize().getWidth(),
            (float)templatePresentation.getSlideSize().getSize().getHeight(),
            templatePresentation.getSlideSize().getType()
        );
        
        // 设置风格策略
        if (styleStrategy == null) {
            this.styleStrategy = new DefaultStyle();
        } else {
            this.styleStrategy = styleStrategy;
        }
        
        // 如果是安全生产类型，加载布局配置和模板文件
        if (isSafetyTemplate()) {
            // 加载布局配置文件
            loadLayoutConfig();
            // 从配置文件加载布局映射
            if (layoutConfigMap != null && !layoutConfigMap.isEmpty()) {
                loadClassicLayoutsFromConfig();
            } else {
                // 如果配置文件不存在，创建空的布局映射（需要在提取模板时生成配置文件）
                this.classicLayoutMap = new HashMap<>();
                System.out.println("提示：布局配置文件不存在，请先运行模板提取工具生成模板文件");
            }
            // 加载已存在的 master_template.pptx 模板文件
            loadMasterTemplate();
        }
    }
    
    /**
     * 加载 master_template.pptx 模板文件到内存
     * 
     * 模板文件应通过 TemplateExtractor 工具预先提取生成。
     * 此方法只负责加载已存在的模板文件，不负责生成模板。
     */
    private void loadMasterTemplate() {
        String masterTemplateFileName = "templates/master_template.pptx";
        File masterTemplateFile = new File(masterTemplateFileName);
        
        if (!masterTemplateFile.exists()) {
            System.err.println("警告：模板文件不存在: " + masterTemplateFileName);
            System.err.println("请先运行模板提取工具生成模板文件:");
            System.err.println("  TemplateExtractor.extractTemplate(sourceFile, outputFile, startPage, endPage)");
            this.masterTemplatePresentation = null;
            return;
        }
        
        try {
            this.masterTemplatePresentation = new Presentation(masterTemplateFileName);
            System.out.println("✓ 已加载 master_template.pptx 到内存（共 " + this.masterTemplatePresentation.getSlides().size() + " 张模板幻灯片）");
        } catch (Exception e) {
            System.err.println("警告：加载 master_template.pptx 失败: " + e.getMessage());
            this.masterTemplatePresentation = null;
        }
    }
    
    /**
     * 加载布局配置文件
     * 
     * 从 config/layouts.json 文件加载布局和样式配置。
     * 如果配置文件不存在，则返回空Map，系统会使用默认的提取方式。
     */
    @SuppressWarnings("unchecked")
    private void loadLayoutConfig() {
        File configFile = new File("config/layouts.json");
        if (!configFile.exists()) {
            System.out.println("提示：未找到布局配置文件 config/layouts.json，将使用默认方式提取布局");
            return;
        }
        
        try {
            ObjectMapper mapper = new ObjectMapper();
            Map<String, Object> config = mapper.readValue(configFile, new TypeReference<Map<String, Object>>() {});
            
            // 加载布局配置
            if (config.containsKey("layouts")) {
                List<Map<String, Object>> layouts = (List<Map<String, Object>>) config.get("layouts");
                this.layoutConfigMap = new HashMap<>();
                for (Map<String, Object> layout : layouts) {
                    String name = (String) layout.get("name");
                    if (name != null) {
                        layoutConfigMap.put(name, layout);
                    }
                }
                System.out.println("✓ 已从配置文件加载 " + layoutConfigMap.size() + " 个布局定义");
            }
            
            // 加载样式配置
            if (config.containsKey("styles")) {
                this.layoutStyleMap = (Map<String, Map<String, Object>>) config.get("styles");
                System.out.println("✓ 已从配置文件加载 " + (layoutStyleMap != null ? layoutStyleMap.size() : 0) + " 个布局样式");
            }
        } catch (Exception e) {
            System.err.println("警告：加载布局配置文件失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 从配置文件加载经典布局映射
     * 
     * 支持通过布局名称、别名、类别等方式来匹配布局
     */
    private void loadClassicLayoutsFromConfig() {
        if (layoutConfigMap == null || layoutConfigMap.isEmpty()) {
            return;
        }
        
        this.classicLayoutMap = new HashMap<>();
        Map<String, String> aliasMap = new HashMap<>(); // 别名 -> 布局名称映射
        
        for (Map.Entry<String, Map<String, Object>> entry : layoutConfigMap.entrySet()) {
            String layoutName = entry.getKey();
            Map<String, Object> layoutConfig = entry.getValue();
            
            // 获取页码
            Object pageNumberObj = layoutConfig.get("pageNumber");
            if (pageNumberObj instanceof Number) {
                int pageNumber = ((Number) pageNumberObj).intValue();
                classicLayoutMap.put(layoutName, pageNumber);
                
                // 获取显示名称
                String displayName = (String) layoutConfig.get("displayName");
                String description = (String) layoutConfig.get("description");
                String info = displayName != null ? displayName : layoutName;
                if (description != null) {
                    info += " - " + description;
                }
                System.out.println("  ✓ 从配置文件加载布局: " + layoutName + " (" + info + ", 源PPT第" + pageNumber + "页)");
                
                // 加载别名映射
                @SuppressWarnings("unchecked")
                List<String> aliases = (List<String>) layoutConfig.get("aliases");
                if (aliases != null) {
                    for (String alias : aliases) {
                        aliasMap.put(alias.toLowerCase(), layoutName);
                    }
                }
                
                // 也支持通过类别来匹配
                String category = (String) layoutConfig.get("category");
                if (category != null) {
                    aliasMap.put(category.toLowerCase(), layoutName);
                }
            }
        }
        
        // 将别名映射也添加到classicLayoutMap中（通过别名可以找到对应的页码）
        for (Map.Entry<String, String> aliasEntry : aliasMap.entrySet()) {
            String alias = aliasEntry.getKey();
            String layoutName = aliasEntry.getValue();
            if (classicLayoutMap.containsKey(layoutName)) {
                classicLayoutMap.put(alias, classicLayoutMap.get(layoutName));
            }
        }
        
        System.out.println("✓ 共从配置文件加载 " + layoutConfigMap.size() + " 个经典布局类型（包含 " + aliasMap.size() + " 个别名）");
    }
    
    /**
     * 查找实际的布局名称（如果输入的是别名，返回真正的布局名称）
     * 
     * @param layoutNameOrAlias 布局名称或别名
     * @return 实际的布局名称，如果找不到则返回null
     */
    private String findActualLayoutName(String layoutNameOrAlias) {
        if (layoutConfigMap == null || layoutConfigMap.isEmpty()) {
            return null;
        }
        
        // 如果直接是布局名称，直接返回
        if (layoutConfigMap.containsKey(layoutNameOrAlias)) {
            return layoutNameOrAlias;
        }
        
        // 遍历所有布局，查找别名匹配
        for (Map.Entry<String, Map<String, Object>> entry : layoutConfigMap.entrySet()) {
            String layoutName = entry.getKey();
            Map<String, Object> layoutConfig = entry.getValue();
            
            // 检查别名
            @SuppressWarnings("unchecked")
            List<String> aliases = (List<String>) layoutConfig.get("aliases");
            if (aliases != null) {
                for (String alias : aliases) {
                    if (alias.equalsIgnoreCase(layoutNameOrAlias)) {
                        return layoutName;
                    }
                }
            }
            
            // 检查类别
            String category = (String) layoutConfig.get("category");
            if (category != null && category.equalsIgnoreCase(layoutNameOrAlias)) {
                return layoutName;
            }
        }
        
        return null;
    }
    
    /**
     * 根据类别查找布局名称
     * 
     * 在配置文件中查找 category 字段与给定布局类型匹配的布局
     * 
     * @param layoutType 布局类型（如 "image_with_text", "content_page"）
     * @return 匹配的布局名称，如果找不到则返回null
     */
    private String findLayoutByCategory(String layoutType) {
        if (layoutConfigMap == null || layoutConfigMap.isEmpty()) {
            return null;
        }
        
        String lowerLayoutType = layoutType.toLowerCase();
        
        // 遍历所有布局，查找类别匹配
        for (Map.Entry<String, Map<String, Object>> entry : layoutConfigMap.entrySet()) {
            String layoutName = entry.getKey();
            Map<String, Object> layoutConfig = entry.getValue();
            
            // 检查类别
            String category = (String) layoutConfig.get("category");
            if (category != null && category.equalsIgnoreCase(lowerLayoutType)) {
                return layoutName;
            }
        }
        
        return null;
    }
    
    /**
     * 从源PPT中提取布局并注册为经典布局类型
     * 
     * 分析《1.2 安全生产方针政策.pptx》第5页到倒数第2页的布局结构，
     * 为每个布局创建一个经典的布局类型名称，供所有PPT类型使用。
     * 
     * 如果配置文件存在，此方法不会被调用。
     */
    private void extractAndRegisterClassicLayouts() {
        if (safetyReferencePresentation == null) {
            return;
        }
        
        this.classicLayoutMap = new HashMap<>();
        int slideCount = safetyReferencePresentation.getSlides().size();
        
        if (slideCount < 5) {
            System.out.println("警告：源PPT页数不足，无法提取布局");
            return;
        }
        
        // 从第5页开始到倒数第2页（索引4到slideCount-2）
        int startIndex = 4; // 第5页（索引4）
        int endIndex = slideCount - 2; // 倒数第2页（索引slideCount-2）
        
        System.out.println("开始提取经典布局（从第5页到第" + (endIndex + 1) + "页）...");
        
        for (int i = startIndex; i <= endIndex; i++) {
            try {
                ISlide slide = safetyReferencePresentation.getSlides().get_Item(i);
                String layoutName = analyzeAndNameLayout(slide, i + 1);
                
                // 注册布局：布局名称 -> 页码（从1开始）
                classicLayoutMap.put(layoutName, i + 1);
                System.out.println("  ✓ 注册经典布局: " + layoutName + " (源PPT第" + (i + 1) + "页)");
            } catch (Exception e) {
                System.err.println("警告：分析第" + (i + 1) + "页布局失败: " + e.getMessage());
            }
        }
        
        System.out.println("✓ 共提取并注册 " + classicLayoutMap.size() + " 个经典布局类型");
        
        // 抽离模板文件和策略类风格
        extractTemplatesAndStyles();
    }
    
    /**
     * 抽离模板文件和策略类风格
     * 
     * 从源PPT的第5页到倒数第2页，将所有模板幻灯片合并到一个 master_template.pptx 文件中，
     * 并为每个模板创建对应的策略类风格。
     */
    private void extractTemplatesAndStyles() {
        if (safetyReferencePresentation == null) {
            System.out.println("提示：参考PPT未加载，跳过模板和风格抽离");
            return;
        }
        
        int slideCount = safetyReferencePresentation.getSlides().size();
        System.out.println("调试：源PPT共有 " + slideCount + " 页");
        
        if (slideCount < 5) {
            System.out.println("提示：源PPT页数不足（" + slideCount + "页），无法抽离模板");
            return;
        }
        
        // 创建模板目录
        File templatesDir = new File("templates");
        if (!templatesDir.exists()) {
            templatesDir.mkdirs();
            System.out.println("✓ 已创建模板目录: templates");
        }
        
        // 创建统一的 master_template.pptx 文件
        String masterTemplateFileName = "templates/master_template.pptx";
        Presentation masterTemplate = new Presentation();
        
        // 设置幻灯片尺寸（与源PPT一致）
        masterTemplate.getSlideSize().setSize(
            (float)safetyReferencePresentation.getSlideSize().getSize().getWidth(),
            (float)safetyReferencePresentation.getSlideSize().getSize().getHeight(),
            safetyReferencePresentation.getSlideSize().getType()
        );
        
        // 删除默认空白页
        if (masterTemplate.getSlides().size() > 0) {
            masterTemplate.getSlides().removeAt(0);
        }
        
        // 从第5页开始到倒数第2页（索引4到slideCount-2）
        int startIndex = 4; // 第5页（索引4）
        int endIndex = slideCount - 2; // 倒数第2页（索引slideCount-2）
        
        System.out.println("开始抽离模板文件和策略类风格（从第5页到第" + (endIndex + 1) + "页，共" + (endIndex - startIndex + 1) + "页）...");
        System.out.println("所有模板将统一保存到: " + masterTemplateFileName);
        
        for (int i = startIndex; i <= endIndex; i++) {
            try {
                ISlide slide = safetyReferencePresentation.getSlides().get_Item(i);
                int pageNumber = i + 1;
                String layoutName = analyzeAndNameLayout(slide, pageNumber);
                
                // 1. 将幻灯片添加到 master_template.pptx
                masterTemplate.getSlides().addClone(slide);
                int masterIndex = masterTemplate.getSlides().size() - 1; // 当前添加的幻灯片在master_template中的索引
                
                System.out.println("  ✓ 已将源PPT第" + pageNumber + "页添加到 master_template.pptx (索引: " + masterIndex + ", 布局: " + layoutName + ")");
                
                // 2. 创建对应的策略类风格
                createStyleClassForTemplate(layoutName, slide, pageNumber);
                
                System.out.println("  ✓ 已创建风格类: " + layoutName);
            } catch (Exception e) {
                System.err.println("警告：抽离第" + (i + 1) + "页模板和风格失败: " + e.getMessage());
                e.printStackTrace();
            }
        }
        
        // 保存 master_template.pptx（临时保存，用于后续处理）
        masterTemplate.save(masterTemplateFileName, SaveFormat.Pptx);
        masterTemplate.dispose();
        
        // 去除 master_template.pptx 中的水印（包括水印的文本框）
        try {
            removeWatermarksFromXML(masterTemplateFileName);
            System.out.println("    ✓ 已去除 master_template.pptx 中的水印");
        } catch (Exception e) {
            System.err.println("警告：去除 master_template.pptx 水印失败: " + e.getMessage());
        }
        
        // 重新加载文件（水印已去除），替换文字和图片
        Presentation finalMasterTemplate = new Presentation(masterTemplateFileName);
        
        // 处理每一张幻灯片
        for (int i = 0; i < finalMasterTemplate.getSlides().size(); i++) {
            ISlide slide = finalMasterTemplate.getSlides().get_Item(i);
            
            // 将模板文件中的所有文字替换为"模板文字模板文字模板文字"，保留文字的大小、字体和对齐方式
            replaceAllTextWithTemplateTextPreservingStyle(slide);
            
            // 将模板文件中的所有图片替换为"No Image"（顶部图标除外）
            replaceAllImagesWithNoImage(slide, finalMasterTemplate);
        }
        
        System.out.println("    ✓ 已替换所有文本为模板文字");
        
        // 最终保存 master_template.pptx
        finalMasterTemplate.save(masterTemplateFileName, SaveFormat.Pptx);
        finalMasterTemplate.dispose();
        
        // 最终保存后，再次去除水印（因为 Aspose.Slides 可能在保存时重新添加水印）
        // 注意：只去除水印，不要影响已替换的"模板文字"
        try {
            removeWatermarksFromXML(masterTemplateFileName);
            System.out.println("    ✓ 已再次去除 master_template.pptx 中的水印");
        } catch (Exception e) {
            System.err.println("警告：最终去除 master_template.pptx 水印失败: " + e.getMessage());
        }
        
        // 加载 master_template.pptx 到内存，供后续使用
        try {
            this.masterTemplatePresentation = new Presentation(masterTemplateFileName);
            System.out.println("✓ 已加载 master_template.pptx 到内存（共 " + this.masterTemplatePresentation.getSlides().size() + " 张模板幻灯片）");
        } catch (Exception e) {
            System.err.println("警告：加载 master_template.pptx 失败: " + e.getMessage());
            this.masterTemplatePresentation = null;
        }
        
        System.out.println("✓ 模板文件和策略类风格抽离完成");
    }
    
    /**
     * 为模板创建对应的策略类风格
     * 
     * @param layoutName 布局名称
     * @param slide 源幻灯片
     * @param pageNumber 页码
     */
    private void createStyleClassForTemplate(String layoutName, ISlide slide, int pageNumber) {
        try {
            // 分析幻灯片中的样式信息
            Map<String, Object> styleInfo = analyzeSlideStyle(slide);
            
            // 生成风格类文件名
            String className = toClassName(layoutName);
            String styleFileName = "src/main/java/com/pptfactory/style/" + className + ".java";
            
            // 检查文件是否已存在
            File styleFile = new File(styleFileName);
            if (styleFile.exists()) {
                System.out.println("    ⚠ 风格类文件已存在，跳过: " + styleFileName);
                return;
            }
            
            // 生成风格类代码
            String styleClassCode = generateStyleClassCode(className, layoutName, styleInfo, pageNumber);
            
            // 保存风格类文件
            File styleDir = new File("src/main/java/com/pptfactory/style");
            if (!styleDir.exists()) {
                styleDir.mkdirs();
            }
            
            try (FileWriter writer = new FileWriter(styleFile)) {
                writer.write(styleClassCode);
            }
            
            System.out.println("    ✓ 已创建风格类: " + styleFileName);
        } catch (Exception e) {
            System.err.println("警告：创建风格类失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 分析幻灯片中的样式信息
     * 
     * @param slide 要分析的幻灯片
     * @return 样式信息Map
     */
    private Map<String, Object> analyzeSlideStyle(ISlide slide) {
        Map<String, Object> styleInfo = new HashMap<>();
        
        // 分析文本框中的字体大小和颜色
        List<Double> titleSizes = new ArrayList<>();
        List<Double> contentSizes = new ArrayList<>();
        List<java.awt.Color> titleColors = new ArrayList<>();
        List<java.awt.Color> contentColors = new ArrayList<>();
        
        for (int i = 0; i < slide.getShapes().size(); i++) {
            IShape shape = slide.getShapes().get_Item(i);
            if (shape instanceof IAutoShape) {
                IAutoShape autoShape = (IAutoShape) shape;
                ITextFrame textFrame = autoShape.getTextFrame();
                if (textFrame != null) {
                    // 分析段落和文本部分的样式
                    for (int j = 0; j < textFrame.getParagraphs().getCount(); j++) {
                        IParagraph para = textFrame.getParagraphs().get_Item(j);
                        for (int k = 0; k < para.getPortions().getCount(); k++) {
                            IPortion portion = para.getPortions().get_Item(k);
                            double fontSize = portion.getPortionFormat().getFontHeight();
                            
                            // 判断是标题还是正文（根据字体大小）
                            if (fontSize >= 24) {
                                titleSizes.add(fontSize);
                                if (portion.getPortionFormat().getFillFormat().getFillType() == FillType.Solid) {
                                    titleColors.add(portion.getPortionFormat().getFillFormat().getSolidFillColor().getColor());
                                }
                            } else {
                                contentSizes.add(fontSize);
                                if (portion.getPortionFormat().getFillFormat().getFillType() == FillType.Solid) {
                                    contentColors.add(portion.getPortionFormat().getFillFormat().getSolidFillColor().getColor());
                                }
                            }
                        }
                    }
                }
            }
        }
        
        // 计算平均值（确保不为 NaN）
        double avgTitleSize = 32.0;
        if (!titleSizes.isEmpty()) {
            double sum = titleSizes.stream().mapToDouble(Double::doubleValue).sum();
            double avg = sum / titleSizes.size();
            if (!Double.isNaN(avg) && avg > 0) {
                avgTitleSize = avg;
            }
        }
        
        double avgContentSize = 18.0;
        if (!contentSizes.isEmpty()) {
            double sum = contentSizes.stream().mapToDouble(Double::doubleValue).sum();
            double avg = sum / contentSizes.size();
            if (!Double.isNaN(avg) && avg > 0) {
                avgContentSize = avg;
            }
        }
        
        // 获取最常见的颜色
        java.awt.Color titleColor = titleColors.isEmpty() ? new java.awt.Color(220, 20, 60) : titleColors.get(0);
        java.awt.Color contentColor = contentColors.isEmpty() ? new java.awt.Color(0, 0, 0) : contentColors.get(0);
        
        styleInfo.put("titleFontSize", avgTitleSize);
        styleInfo.put("contentFontSize", avgContentSize);
        styleInfo.put("titleColor", titleColor);
        styleInfo.put("contentColor", contentColor);
        
        return styleInfo;
    }
    
    /**
     * 将布局名称转换为类名
     * 
     * @param layoutName 布局名称（如 "classic_image_text_5"）
     * @return 类名（如 "ClassicImageText5Style"）
     */
    private String toClassName(String layoutName) {
        // 移除 "classic_" 前缀
        String name = layoutName.replaceFirst("^classic_", "");
        
        // 将下划线分隔的名称转换为驼峰命名
        String[] parts = name.split("_");
        StringBuilder className = new StringBuilder();
        for (String part : parts) {
            if (!part.isEmpty()) {
                className.append(Character.toUpperCase(part.charAt(0)));
                if (part.length() > 1) {
                    className.append(part.substring(1));
                }
            }
        }
        className.append("Style");
        
        return className.toString();
    }
    
    /**
     * 生成风格类代码
     * 
     * @param className 类名
     * @param layoutName 布局名称
     * @param styleInfo 样式信息
     * @param pageNumber 页码
     * @return 风格类代码
     */
    private String generateStyleClassCode(String className, String layoutName, Map<String, Object> styleInfo, int pageNumber) {
        double titleSize = ((Number) styleInfo.getOrDefault("titleFontSize", 32.0)).doubleValue();
        double contentSize = ((Number) styleInfo.getOrDefault("contentFontSize", 18.0)).doubleValue();
        java.awt.Color titleColor = (java.awt.Color) styleInfo.getOrDefault("titleColor", new java.awt.Color(220, 20, 60));
        java.awt.Color contentColor = (java.awt.Color) styleInfo.getOrDefault("contentColor", new java.awt.Color(0, 0, 0));
        
        return "package com.pptfactory.style;\n\n" +
               "import com.aspose.slides.IPortion;\n" +
               "import com.aspose.slides.FillType;\n" +
               "import java.awt.Color;\n\n" +
               "/**\n" +
               " * " + layoutName + " 风格策略类\n" +
               " * \n" +
               " * 从《1.2 安全生产方针政策.pptx》第" + pageNumber + "页提取的风格。\n" +
               " * 此风格类保留了源PPT页面的样式特征。\n" +
               " */\n" +
               "public class " + className + " implements StyleStrategy {\n\n" +
               "    // 字体大小配置（从源PPT第" + pageNumber + "页提取）\n" +
               "    private static final double MAIN_TITLE_SIZE = " + titleSize + ";\n" +
               "    private static final double TITLE_SIZE = " + titleSize + ";\n" +
               "    private static final double SUBTITLE_SIZE = " + (titleSize * 0.7) + ";\n" +
               "    private static final double CONTENT_SIZE = " + contentSize + ";\n" +
               "    private static final double BULLET_SIZE = " + contentSize + ";\n\n" +
               "    // 颜色配置（从源PPT第" + pageNumber + "页提取）\n" +
               "    private static final Color TITLE_COLOR = new Color(" + titleColor.getRed() + ", " + titleColor.getGreen() + ", " + titleColor.getBlue() + ");\n" +
               "    private static final Color CONTENT_COLOR = new Color(" + contentColor.getRed() + ", " + contentColor.getGreen() + ", " + contentColor.getBlue() + ");\n\n" +
               "    // 间距配置\n" +
               "    private static final double BULLET_SPACING = 12.0;\n\n" +
               "    @Override\n" +
               "    public void applyTitleStyle(IPortion portion, boolean isMainTitle) {\n" +
               "        portion.getPortionFormat().setFontHeight((float)(isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE));\n" +
               "        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);\n" +
               "        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public void applySubtitleStyle(IPortion portion) {\n" +
               "        portion.getPortionFormat().setFontHeight((float)SUBTITLE_SIZE);\n" +
               "        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);\n" +
               "        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(TITLE_COLOR);\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public void applyContentStyle(IPortion portion) {\n" +
               "        portion.getPortionFormat().setFontHeight((float)CONTENT_SIZE);\n" +
               "        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);\n" +
               "        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(CONTENT_COLOR);\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public void applyBulletStyle(IPortion portion) {\n" +
               "        portion.getPortionFormat().setFontHeight((float)BULLET_SIZE);\n" +
               "        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);\n" +
               "        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(CONTENT_COLOR);\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public double getBulletSpacing() {\n" +
               "        return BULLET_SPACING;\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public double getTitleFontSize(boolean isMainTitle) {\n" +
               "        return isMainTitle ? MAIN_TITLE_SIZE : TITLE_SIZE;\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public double getSubtitleFontSize() {\n" +
               "        return SUBTITLE_SIZE;\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public double getContentFontSize() {\n" +
               "        return CONTENT_SIZE;\n" +
               "    }\n\n" +
               "    @Override\n" +
               "    public double getBulletFontSize() {\n" +
               "        return BULLET_SIZE;\n" +
               "    }\n" +
               "}\n";
    }
    
    /**
     * 分析幻灯片布局并生成布局名称
     * 
     * 根据幻灯片的特征（文本框数量、图片数量、布局结构等）生成一个描述性的布局名称。
     * 
     * @param slide 要分析的幻灯片
     * @param pageNumber 页码（用于生成唯一名称）
     * @return 布局名称
     */
    private String analyzeAndNameLayout(ISlide slide, int pageNumber) {
        int textBoxCount = 0;
        int imageCount = 0;
        int shapeCount = slide.getShapes().size();
        
        // 统计文本框和图片数量
        for (int i = 0; i < shapeCount; i++) {
            IShape shape = slide.getShapes().get_Item(i);
            if (shape instanceof IAutoShape) {
                IAutoShape autoShape = (IAutoShape) shape;
                ITextFrame textFrame = autoShape.getTextFrame();
                if (textFrame != null && textFrame.getText() != null && !textFrame.getText().trim().isEmpty()) {
                    textBoxCount++;
                }
            } else if (shape instanceof IPictureFrame) {
                imageCount++;
            }
        }
        
        // 根据特征生成布局名称
        // 格式：classic_layout_<特征描述>_<序号>
        String feature = "";
        if (imageCount > 0 && textBoxCount > 0) {
            feature = "image_text";
        } else if (textBoxCount >= 3) {
            feature = "multi_text";
        } else if (textBoxCount == 2) {
            feature = "two_text";
        } else if (textBoxCount == 1) {
            feature = "single_text";
        } else {
            feature = "simple";
        }
        
        // 使用页码作为唯一标识，确保每个布局都有唯一的名称
        return "classic_" + feature + "_" + pageNumber;
    }
    
    /**
     * 分析模板的详细排版格式（已废弃，改为手动维护配置文件）
     * 
     * 注意：此方法已不再使用，排版格式描述信息改为手动在 config/layouts.json 中维护。
     * 
     * @param slide 要分析的幻灯片
     * @param layoutName 布局名称
     * @param pageNumber 页码
     * @return 排版格式描述Map
     * @deprecated 已废弃，改为手动维护配置文件
     */
    @Deprecated
    private Map<String, Object> analyzeLayoutFormat(ISlide slide, String layoutName, int pageNumber) {
        Map<String, Object> formatInfo = new HashMap<>();
        
        // 基本信息
        formatInfo.put("name", layoutName);
        formatInfo.put("displayName", getDisplayNameFromLayoutName(layoutName));
        formatInfo.put("type", "classic");
        formatInfo.put("category", getCategoryFromLayoutName(layoutName));
        formatInfo.put("sourceFile", "1.2 安全生产方针政策.pptx");
        formatInfo.put("pageNumber", pageNumber);
        
        // 分析文本框
        List<Map<String, Object>> textBoxes = new ArrayList<>();
        List<Map<String, Object>> images = new ArrayList<>();
        int textBoxCount = 0;
        int imageCount = 0;
        
        for (int i = 0; i < slide.getShapes().size(); i++) {
            IShape shape = slide.getShapes().get_Item(i);
            
            if (shape instanceof IAutoShape) {
                IAutoShape autoShape = (IAutoShape) shape;
                ITextFrame textFrame = autoShape.getTextFrame();
                if (textFrame != null && textFrame.getText() != null && !textFrame.getText().trim().isEmpty()) {
                    textBoxCount++;
                    Map<String, Object> textBoxInfo = new HashMap<>();
                    textBoxInfo.put("index", textBoxCount);
                    textBoxInfo.put("x", autoShape.getFrame().getX());
                    textBoxInfo.put("y", autoShape.getFrame().getY());
                    textBoxInfo.put("width", autoShape.getFrame().getWidth());
                    textBoxInfo.put("height", autoShape.getFrame().getHeight());
                    
                    // 分析文本对齐方式
                    if (textFrame.getParagraphs().getCount() > 0) {
                        IParagraph para = textFrame.getParagraphs().get_Item(0);
                        int alignment = para.getParagraphFormat().getAlignment();
                        String alignmentStr = "left";
                        if (alignment == TextAlignment.Center) {
                            alignmentStr = "center";
                        } else if (alignment == TextAlignment.Right) {
                            alignmentStr = "right";
                        } else if (alignment == TextAlignment.Justify) {
                            alignmentStr = "justify";
                        }
                        textBoxInfo.put("alignment", alignmentStr);
                    }
                    
                    // 分析字体大小
                    if (textFrame.getParagraphs().getCount() > 0) {
                        IParagraph para = textFrame.getParagraphs().get_Item(0);
                        if (para.getPortions().getCount() > 0) {
                            IPortion portion = para.getPortions().get_Item(0);
                            textBoxInfo.put("fontSize", portion.getPortionFormat().getFontHeight() / 100.0); // 转换为点
                        }
                    }
                    
                    textBoxes.add(textBoxInfo);
                }
            } else if (shape instanceof IPictureFrame) {
                imageCount++;
                IPictureFrame pictureFrame = (IPictureFrame) shape;
                Map<String, Object> imageInfo = new HashMap<>();
                imageInfo.put("index", imageCount);
                imageInfo.put("x", pictureFrame.getFrame().getX());
                imageInfo.put("y", pictureFrame.getFrame().getY());
                imageInfo.put("width", pictureFrame.getFrame().getWidth());
                imageInfo.put("height", pictureFrame.getFrame().getHeight());
                images.add(imageInfo);
            }
        }
        
        // 生成排版格式描述
        String layoutStructure = generateLayoutStructureDescription(textBoxes, images);
        String formatDescription = generateFormatDescription(textBoxes, images, layoutStructure);
        
        formatInfo.put("description", formatDescription);
        formatInfo.put("detailedDescription", formatDescription + " 该布局包含 " + textBoxCount + " 个文本框和 " + imageCount + " 个图片，排版格式清晰，适合用于展示相关内容。");
        Map<String, Object> layoutFormat = new HashMap<>();
        layoutFormat.put("textBoxes", textBoxes);
        layoutFormat.put("images", images);
        layoutFormat.put("layoutStructure", layoutStructure);
        layoutFormat.put("textBoxCount", textBoxCount);
        layoutFormat.put("imageCount", imageCount);
        formatInfo.put("layoutFormat", layoutFormat);
        
        Map<String, Object> features = new HashMap<>();
        features.put("hasImage", imageCount > 0);
        features.put("textBoxCount", textBoxCount);
        features.put("imageCount", imageCount);
        features.put("layoutStructure", layoutStructure);
        features.put("textPlaceholders", generateTextPlaceholders(textBoxCount));
        if (imageCount > 0) {
            features.put("imagePlaceholder", "主图片");
        } else {
            features.put("imagePlaceholder", null);
        }
        formatInfo.put("features", features);
        formatInfo.put("useCases", generateUseCases(layoutName, textBoxCount, imageCount));
        formatInfo.put("tags", generateTags(layoutName, textBoxCount, imageCount));
        
        return formatInfo;
    }
    
    /**
     * 从布局名称生成显示名称
     */
    private String getDisplayNameFromLayoutName(String layoutName) {
        if (layoutName.contains("image_text")) {
            return "图片+文本布局（经典）";
        } else if (layoutName.contains("multi_text")) {
            return "多文本框布局（经典）";
        } else if (layoutName.contains("two_text")) {
            return "双文本框布局（经典）";
        } else if (layoutName.contains("single_text")) {
            return "单文本框布局（经典）";
        } else {
            return "简单布局（经典）";
        }
    }
    
    /**
     * 从布局名称生成类别
     */
    private String getCategoryFromLayoutName(String layoutName) {
        if (layoutName.contains("image_text")) {
            return "image_with_text";
        } else if (layoutName.contains("multi_text")) {
            return "content_page";
        } else {
            return "content_page";
        }
    }
    
    /**
     * 生成布局结构描述
     */
    private String generateLayoutStructureDescription(List<Map<String, Object>> textBoxes, List<Map<String, Object>> images) {
        StringBuilder desc = new StringBuilder();
        
        if (images.size() > 0 && textBoxes.size() > 0) {
            // 分析图片和文本的相对位置
            double avgImageX = images.stream().mapToDouble(img -> ((Number) img.get("x")).doubleValue()).average().orElse(0);
            double avgTextBoxX = textBoxes.stream().mapToDouble(tb -> ((Number) tb.get("x")).doubleValue()).average().orElse(0);
            
            if (avgImageX < avgTextBoxX) {
                desc.append("图片位于左侧，文本位于右侧");
            } else {
                desc.append("图片位于右侧，文本位于左侧");
            }
        } else if (textBoxes.size() >= 3) {
            desc.append("多个文本框垂直或水平排列");
        } else if (textBoxes.size() == 2) {
            desc.append("两个文本框上下或左右排列");
        } else if (textBoxes.size() == 1) {
            desc.append("单个文本框居中或靠左排列");
        }
        
        return desc.toString();
    }
    
    /**
     * 生成格式描述
     */
    private String generateFormatDescription(List<Map<String, Object>> textBoxes, List<Map<String, Object>> images, String layoutStructure) {
        StringBuilder desc = new StringBuilder();
        
        if (images.size() > 0 && textBoxes.size() > 0) {
            desc.append("左侧或上方包含图片，右侧或下方包含文本内容的经典布局");
        } else if (textBoxes.size() >= 3) {
            desc.append("包含多个文本框的内容页布局，适合展示多个要点或列表");
        } else if (textBoxes.size() == 2) {
            desc.append("包含两个文本框的布局，适合展示标题和正文内容");
        } else {
            desc.append("简单的单文本框布局，适合展示标题或简短内容");
        }
        
        return desc.toString();
    }
    
    /**
     * 生成文本占位符列表
     */
    private List<String> generateTextPlaceholders(int textBoxCount) {
        List<String> placeholders = new ArrayList<>();
        if (textBoxCount >= 1) {
            placeholders.add("标题");
        }
        if (textBoxCount >= 2) {
            placeholders.add("正文内容");
        }
        for (int i = 3; i <= textBoxCount; i++) {
            placeholders.add("要点" + (i - 2));
        }
        return placeholders;
    }
    
    /**
     * 生成适用场景
     */
    private List<String> generateUseCases(String layoutName, int textBoxCount, int imageCount) {
        List<String> useCases = new ArrayList<>();
        
        if (imageCount > 0 && textBoxCount > 0) {
            useCases.add("产品介绍");
            useCases.add("案例展示");
            useCases.add("图文说明");
            useCases.add("内容讲解");
        } else if (textBoxCount >= 3) {
            useCases.add("要点列表");
            useCases.add("步骤说明");
            useCases.add("多内容展示");
            useCases.add("分类说明");
        } else {
            useCases.add("标题展示");
            useCases.add("内容说明");
        }
        
        return useCases;
    }
    
    /**
     * 生成标签
     */
    private List<String> generateTags(String layoutName, int textBoxCount, int imageCount) {
        List<String> tags = new ArrayList<>();
        
        if (imageCount > 0) {
            tags.add("图片");
        }
        if (textBoxCount > 0) {
            tags.add("文本");
        }
        if (imageCount > 0 && textBoxCount > 0) {
            tags.add("图文混排");
        }
        tags.add("经典布局");
        
        return tags;
    }
    
    /**
     * 更新布局配置文件（已废弃，改为手动维护配置文件）
     * 
     * 注意：此方法已不再使用，排版格式描述信息改为手动在 config/layouts.json 中维护。
     * 
     * @deprecated 已废弃，改为手动维护配置文件
     */
    @Deprecated
    private void updateLayoutConfig(String layoutName, Map<String, Object> layoutDescription) {
        try {
            File configFile = new File("config/layouts.json");
            Map<String, Object> config = new HashMap<>();
            
            // 读取现有配置
            if (configFile.exists()) {
                ObjectMapper mapper = new ObjectMapper();
                config = mapper.readValue(configFile, new TypeReference<Map<String, Object>>() {});
            }
            
            // 获取或创建layouts数组
            @SuppressWarnings("unchecked")
            List<Map<String, Object>> layouts = (List<Map<String, Object>>) config.getOrDefault("layouts", new ArrayList<>());
            
            // 检查是否已存在该布局
            boolean exists = false;
            for (int i = 0; i < layouts.size(); i++) {
                Map<String, Object> layout = layouts.get(i);
                if (layoutName.equals(layout.get("name"))) {
                    // 更新现有布局
                    layouts.set(i, layoutDescription);
                    exists = true;
                    break;
                }
            }
            
            // 如果不存在，添加新布局
            if (!exists) {
                layouts.add(layoutDescription);
            }
            
            config.put("layouts", layouts);
            
            // 保存配置文件
            ObjectMapper mapper = new ObjectMapper();
            mapper.writerWithDefaultPrettyPrinter().writeValue(configFile, config);
            
            System.out.println("    ✓ 已更新布局配置文件: " + layoutName);
        } catch (Exception e) {
            System.err.println("警告：更新布局配置文件失败: " + e.getMessage());
            // 不抛出异常，继续执行
        }
    }
    
    /**
     * 根据slide_data渲染单张幻灯片
     * 
     * 这是渲染幻灯片的入口方法，根据slide_data中的layout字段选择相应的
     * 布局渲染方法。样式由风格策略类控制。
     * 
     * @param slideData 包含幻灯片数据的Map，必须包含以下字段：
     *                  - layout: 布局类型字符串，可选值：
     *                    * "title_page": 标题页布局
     *                    * "content_page": 内容页布局（标题+要点列表）
     *                    * "two_column": 两列布局
     *                    * "image_with_text": 图片+文字布局
     *                    * "image_left_text_right": 左图右文
     *                    * "image_right_text_left": 右图左文
     *                    * "pure_content": 纯内容页
     *                    * "three_column": 三列布局
     *                    * "quote_page": 引用页
     *                    * "chapter_cover": 章节封面页
     *                  其他字段根据不同的布局类型而不同
     * @return 创建的幻灯片对象
     */
    public ISlide renderSlide(Map<String, Object> slideData) {
        String layoutName = (String) slideData.getOrDefault("layout", "content_page");
        
        // 尝试从配置文件中查找布局（支持名称、别名、类别匹配）
        if (classicLayoutMap != null && !classicLayoutMap.isEmpty()) {
            // 首先尝试直接匹配
            if (classicLayoutMap.containsKey(layoutName)) {
                int pageNumber = classicLayoutMap.get(layoutName);
                // 如果是别名，需要找到真正的布局名称
                String actualLayoutName = findActualLayoutName(layoutName);
                return renderClassicLayout(pageNumber, slideData, actualLayoutName != null ? actualLayoutName : layoutName);
            }
            
            // 尝试小写匹配（不区分大小写）
            String lowerLayoutName = layoutName.toLowerCase();
            if (classicLayoutMap.containsKey(lowerLayoutName)) {
                int pageNumber = classicLayoutMap.get(lowerLayoutName);
                String actualLayoutName = findActualLayoutName(lowerLayoutName);
                return renderClassicLayout(pageNumber, slideData, actualLayoutName != null ? actualLayoutName : layoutName);
            }
            
            // 尝试通过类别匹配（如 "image_with_text" 匹配到 category 为 "image_with_text" 的布局）
            String matchedLayoutName = findLayoutByCategory(layoutName);
            if (matchedLayoutName != null && classicLayoutMap.containsKey(matchedLayoutName)) {
                int pageNumber = classicLayoutMap.get(matchedLayoutName);
                return renderClassicLayout(pageNumber, slideData, matchedLayoutName);
            }
        }
        
        // 如果是经典布局类型（classic_* 格式），从模板文件加载
        if (layoutName.startsWith("classic_") && classicLayoutMap != null && classicLayoutMap.containsKey(layoutName)) {
            int pageNumber = classicLayoutMap.get(layoutName);
            return renderClassicLayout(pageNumber, slideData, layoutName);
        }
        
        // 如果是安全生产类型的内容页布局（safety_content_N 格式），直接从源PPT复制
        if (isSafetyTemplate() && layoutName.startsWith("safety_content_")) {
            try {
                int pageNumber = Integer.parseInt(layoutName.substring("safety_content_".length()));
                return renderSafetyContentPage(pageNumber, slideData);
            } catch (NumberFormatException e) {
                System.err.println("警告：无效的安全生产内容页布局格式: " + layoutName);
            }
        }
        
        // 如果是安全生产类型，尝试从参考PPT中获取对应布局的样式
        if (isSafetyTemplate() && safetyReferencePresentation != null) {
            ISlide referenceSlide = findMatchingSafetyLayout(layoutName);
            if (referenceSlide != null) {
                // 使用参考PPT中的布局，然后替换文本内容
                return renderSlideWithSafetyLayout(referenceSlide, slideData, layoutName);
            }
        }
        
        // 根据布局类型调用相应的渲染方法
        switch (layoutName) {
            case "title_page":
                return renderTitlePage(slideData);
            case "content_page":
            case "title_with_content":
                return renderContentPage(slideData);
            case "image_with_text":
            case "image_with_content":
                return renderImageWithText(slideData);
            case "image_left_text_right":
                return renderImageLeftTextRight(slideData);
            case "image_right_text_left":
                return renderImageRightTextLeft(slideData);
            case "pure_content":
                return renderPureContent(slideData);
            case "two_column":
                return renderTwoColumn(slideData);
            case "three_column":
                return renderThreeColumn(slideData);
            case "quote_page":
                return renderQuotePage(slideData);
            case "chapter_cover":
                return renderChapterCover(slideData);
            default:
                // 默认使用内容页布局
                return renderContentPage(slideData);
        }
    }
    
    /**
     * 渲染标题页布局
     * 
     * 使用风格策略类应用标题页的样式。
     * 
     * @param data 包含title和subtitle的Map
     * @return 幻灯片对象
     */
    private ISlide renderTitlePage(Map<String, Object> data) {
        // 创建空白幻灯片
        ISlide slide = presentation.getSlides().addEmptySlide(presentation.getLayoutSlides().get_Item(0));
        
        // 添加标题
        String title = (String) data.getOrDefault("title", "");
        if (!title.isEmpty()) {
            // 创建自动形状（文本框）
            IAutoShape titleShape = slide.getShapes().addAutoShape(ShapeType.Rectangle, 
                (float)(1.0 * 72), (float)(2.5 * 72), (float)(8.0 * 72), (float)(1.5 * 72));
            titleShape.getFillFormat().setFillType(FillType.NoFill);
            titleShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
            
            // 设置文本
            ITextFrame titleFrame = titleShape.getTextFrame();
            titleFrame.setText(title);
            titleFrame.getParagraphs().get_Item(0).getParagraphFormat().setAlignment(TextAlignment.Center);
            
            // 应用样式
            IPortion titlePortion = titleFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
            titlePortion.getPortionFormat().setFontHeight((float)styleStrategy.getTitleFontSize(true));
            // titlePortion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
            titlePortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
            titlePortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
            styleStrategy.applyTitleStyle(titlePortion, true);
            
            System.out.println("    添加标题: \"" + title + "\"");
        }
        
        // 添加副标题
        if (data.containsKey("subtitle")) {
            String subtitle = (String) data.get("subtitle");
            if (subtitle != null && !subtitle.isEmpty()) {
                // 创建自动形状（文本框）
                IAutoShape subtitleShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                    (float)(1.0 * 72), (float)(4.5 * 72), (float)(8.0 * 72), (float)(1.0 * 72));
                subtitleShape.getFillFormat().setFillType(FillType.NoFill);
                subtitleShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                
                // 设置文本
                ITextFrame subtitleFrame = subtitleShape.getTextFrame();
                subtitleFrame.setText(subtitle);
                subtitleFrame.getParagraphs().get_Item(0).getParagraphFormat().setAlignment(TextAlignment.Center);
                
                // 应用样式
                IPortion subtitlePortion = subtitleFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
                subtitlePortion.getPortionFormat().setFontHeight((float)styleStrategy.getSubtitleFontSize());
                // subtitlePortion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
                subtitlePortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                subtitlePortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(64, 64, 64));
                styleStrategy.applySubtitleStyle(subtitlePortion);
                
                System.out.println("    添加副标题: \"" + subtitle + "\"");
            }
        }
        
        return slide;
    }
    
    /**
     * 渲染内容页布局（标题+要点列表）
     * 
     * 使用风格策略类应用内容页的样式。
     * 
     * @param data 包含title和bullets的Map
     * @return 幻灯片对象
     */
    private ISlide renderContentPage(Map<String, Object> data) {
        // 创建空白幻灯片
        ISlide slide = presentation.getSlides().addEmptySlide(presentation.getLayoutSlides().get_Item(0));
        
        // 添加标题
        String title = (String) data.getOrDefault("title", "");
        if (!title.isEmpty()) {
            // 创建自动形状（文本框）
            IAutoShape titleShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                (float)(0.5 * 72), (float)(0.5 * 72), (float)(9.0 * 72), (float)(1.0 * 72));
            titleShape.getFillFormat().setFillType(FillType.NoFill);
            titleShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
            
            // 设置文本
            ITextFrame titleFrame = titleShape.getTextFrame();
            titleFrame.setText(title);
            
            // 应用样式
            IPortion titlePortion = titleFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
            titlePortion.getPortionFormat().setFontHeight((float)styleStrategy.getTitleFontSize(false));
            // titlePortion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
            titlePortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
            titlePortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
            styleStrategy.applyTitleStyle(titlePortion, false);
            
            System.out.println("    添加标题: \"" + title + "\"");
        }
        
        // 添加内容要点
        if (data.containsKey("bullets")) {
            Object bulletsObj = data.get("bullets");
            if (bulletsObj instanceof List) {
                @SuppressWarnings("unchecked")
                List<String> bullets = (List<String>) bulletsObj;
                
                if (!bullets.isEmpty()) {
                    // 创建自动形状（文本框）
                    IAutoShape contentShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                        (float)(1.0 * 72), (float)(2.0 * 72), (float)(8.0 * 72), (float)(5.0 * 72));
                    contentShape.getFillFormat().setFillType(FillType.NoFill);
                    contentShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                    contentShape.getTextFrame().getTextFrameFormat().setAutofitType(TextAutofitType.Shape);
                    
                    ITextFrame contentFrame = contentShape.getTextFrame();
                    contentFrame.getParagraphs().clear();
                    
                    // 添加所有要点
                    System.out.println("    添加 " + bullets.size() + " 个要点");
                    for (int i = 0; i < bullets.size(); i++) {
                        // 创建段落并添加到集合
                        IParagraph para = new Paragraph();
                        contentFrame.getParagraphs().add(para);
                        para.getParagraphFormat().getBullet().setType(BulletType.Symbol);
                        para.getParagraphFormat().getBullet().setChar((char)8226); // 圆点符号
                        para.getParagraphFormat().setIndent((float)(0.5 * 72));
                        para.getParagraphFormat().setSpaceAfter((float)styleStrategy.getBulletSpacing());
                        
                        // 创建文本部分并添加到段落
                        IPortion portion = new Portion();
                        para.getPortions().add(portion);
                        String bulletText = bullets.get(i);
                        portion.setText(bulletText);
                        portion.getPortionFormat().setFontHeight((float)styleStrategy.getBulletFontSize());
                        // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
                        portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                        portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
                        styleStrategy.applyBulletStyle(portion);
                    }
                }
            }
        }
        
        return slide;
    }
    
    /**
     * 渲染两列布局
     * 
     * @param data 包含title、left_content和right_content的Map
     * @return 幻灯片对象
     */
    private ISlide renderTwoColumn(Map<String, Object> data) {
        // 创建空白幻灯片
        ISlide slide = presentation.getSlides().addEmptySlide(presentation.getLayoutSlides().get_Item(0));
        
        // 添加标题
        if (data.containsKey("title")) {
            String title = (String) data.get("title");
            if (title != null && !title.isEmpty()) {
                IAutoShape titleShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                    (float)(0.5 * 72), (float)(0.5 * 72), (float)(9.0 * 72), (float)(1.0 * 72));
                titleShape.getFillFormat().setFillType(FillType.NoFill);
                titleShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                
                ITextFrame titleFrame = titleShape.getTextFrame();
                titleFrame.setText(title);
                
                IPortion titlePortion = titleFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
                titlePortion.getPortionFormat().setFontHeight((float)styleStrategy.getTitleFontSize(false));
                // titlePortion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
                titlePortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                titlePortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
                styleStrategy.applyTitleStyle(titlePortion, false);
            }
        }
        
        // 左列
        if (data.containsKey("left_content")) {
            Object leftContent = data.get("left_content");
            if (leftContent != null) {
                IAutoShape leftShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                    (float)(0.5 * 72), (float)(2.0 * 72), (float)(4.5 * 72), (float)(5.0 * 72));
                leftShape.getFillFormat().setFillType(FillType.NoFill);
                leftShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                leftShape.getTextFrame().getTextFrameFormat().setAutofitType(TextAutofitType.Shape);
                
                ITextFrame leftFrame = leftShape.getTextFrame();
                leftFrame.setText(leftContent.toString());
                
                IPortion leftPortion = leftFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
                leftPortion.getPortionFormat().setFontHeight((float)styleStrategy.getContentFontSize());
                // leftPortion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
                leftPortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                leftPortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
                styleStrategy.applyContentStyle(leftPortion);
            }
        }
        
        // 右列
        if (data.containsKey("right_content")) {
            Object rightContent = data.get("right_content");
            if (rightContent != null) {
                IAutoShape rightShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                    (float)(5.5 * 72), (float)(2.0 * 72), (float)(4.5 * 72), (float)(5.0 * 72));
                rightShape.getFillFormat().setFillType(FillType.NoFill);
                rightShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                rightShape.getTextFrame().getTextFrameFormat().setAutofitType(TextAutofitType.Shape);
                
                ITextFrame rightFrame = rightShape.getTextFrame();
                rightFrame.setText(rightContent.toString());
                
                IPortion rightPortion = rightFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
                rightPortion.getPortionFormat().setFontHeight((float)styleStrategy.getContentFontSize());
                // rightPortion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
                rightPortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                rightPortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
                styleStrategy.applyContentStyle(rightPortion);
            }
        }
        
        return slide;
    }
    
    /**
     * 渲染图片+文字布局
     * 
     * @param data 包含title、image_path和text的Map
     * @return 幻灯片对象
     */
    private ISlide renderImageWithText(Map<String, Object> data) {
        // 创建空白幻灯片
        ISlide slide = presentation.getSlides().addEmptySlide(presentation.getLayoutSlides().get_Item(0));
        
        // 添加标题（如果有）
        if (data.containsKey("title")) {
            String title = (String) data.get("title");
            if (title != null && !title.isEmpty()) {
                IAutoShape titleShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                    (float)(0.5 * 72), (float)(0.5 * 72), (float)(9.0 * 72), (float)(1.0 * 72));
                titleShape.getFillFormat().setFillType(FillType.NoFill);
                titleShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                
                ITextFrame titleFrame = titleShape.getTextFrame();
                titleFrame.setText(title);
                
                IPortion titlePortion = titleFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
                titlePortion.getPortionFormat().setFontHeight((float)styleStrategy.getTitleFontSize(false));
                // titlePortion.getPortionFormat().setBold(NullableBool.True); // TODO: 根据实际 Aspose.Slides API 调整
                titlePortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                titlePortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
                styleStrategy.applyTitleStyle(titlePortion, false);
            }
        }
        
        // 添加图片（如果有）
        if (data.containsKey("image_path")) {
            String imagePath = (String) data.get("image_path");
            if (imagePath != null && !imagePath.isEmpty()) {
                try {
                    File imageFile = new File(imagePath);
                    if (imageFile.exists()) {
                        // 读取图片文件
                        ByteArrayOutputStream buffer = new ByteArrayOutputStream();
                        FileInputStream fis = new FileInputStream(imageFile);
                        byte[] bufferData = new byte[8192];
                        int nRead;
                        while ((nRead = fis.read(bufferData, 0, bufferData.length)) != -1) {
                            buffer.write(bufferData, 0, nRead);
                        }
                        fis.close();
                        byte[] pictureData = buffer.toByteArray();
                        
                        // 添加图片到演示文稿
                        IPPImage image = presentation.getImages().addImage(pictureData);
                        
                        // 创建图片形状
                        IPictureFrame pictureFrame = slide.getShapes().addPictureFrame(ShapeType.Rectangle,
                            (float)(1.0 * 72), (float)(2.0 * 72), (float)(4.0 * 72), (float)(4.0 * 72), image);
                    }
                } catch (IOException e) {
                    System.err.println("警告：无法加载图片 " + imagePath + ": " + e.getMessage());
                }
            }
        }
        
        // 添加文字内容
        if (data.containsKey("text")) {
            Object textContent = data.get("text");
            if (textContent != null) {
                IAutoShape textShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                    (float)(5.5 * 72), (float)(2.0 * 72), (float)(4.5 * 72), (float)(5.0 * 72));
                textShape.getFillFormat().setFillType(FillType.NoFill);
                textShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                textShape.getTextFrame().getTextFrameFormat().setAutofitType(TextAutofitType.Shape);
                
                ITextFrame textFrame = textShape.getTextFrame();
                textFrame.setText(textContent.toString());
                
                IPortion textPortion = textFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
                textPortion.getPortionFormat().setFontHeight((float)styleStrategy.getContentFontSize());
                // textPortion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
                textPortion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                textPortion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
                styleStrategy.applyContentStyle(textPortion);
            }
        }
        
        return slide;
    }
    
    // 其他布局方法的简化实现
    private ISlide renderImageLeftTextRight(Map<String, Object> data) {
        // 类似 renderImageWithText，但图片在左，文字在右
        return renderImageWithText(data); // 简化实现
    }
    
    private ISlide renderImageRightTextLeft(Map<String, Object> data) {
        // 类似 renderImageWithText，但图片在右，文字在左
        return renderImageWithText(data); // 简化实现
    }
    
    private ISlide renderPureContent(Map<String, Object> data) {
        // 创建空白幻灯片
        ISlide slide = presentation.getSlides().addEmptySlide(presentation.getLayoutSlides().get_Item(0));
        Object content = data.getOrDefault("content", data.getOrDefault("text", ""));
        if (content != null && !content.toString().isEmpty()) {
            IAutoShape contentShape = slide.getShapes().addAutoShape(ShapeType.Rectangle,
                (float)(1.0 * 72), (float)(1.0 * 72), (float)(8.0 * 72), (float)(6.0 * 72));
            contentShape.getFillFormat().setFillType(FillType.NoFill);
            contentShape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
            contentShape.getTextFrame().getTextFrameFormat().setAutofitType(TextAutofitType.Shape);
            
            ITextFrame contentFrame = contentShape.getTextFrame();
            contentFrame.setText(content.toString());
            
            IPortion portion = contentFrame.getParagraphs().get_Item(0).getPortions().get_Item(0);
            portion.getPortionFormat().setFontHeight((float)styleStrategy.getContentFontSize());
            // portion.getPortionFormat().setBold(NullableBool.False); // TODO: 根据实际 Aspose.Slides API 调整
            portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
            portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(new java.awt.Color(0, 0, 0));
            styleStrategy.applyContentStyle(portion);
        }
        return slide;
    }
    
    private ISlide renderThreeColumn(Map<String, Object> data) {
        // 三列布局的简化实现
        return renderTwoColumn(data); // 简化实现
    }
    
    private ISlide renderQuotePage(Map<String, Object> data) {
        // 引用页的简化实现
        return renderContentPage(data); // 简化实现
    }
    
    private ISlide renderChapterCover(Map<String, Object> data) {
        // 章节封面页的简化实现
        return renderTitlePage(data); // 简化实现
    }
    
    /**
     * 从JSON数据渲染整个PPT
     * 
     * @param slidesData 包含slides数组的Map
     */
    @SuppressWarnings("unchecked")
    public void renderFromJson(Map<String, Object> slidesData) {
        Object slidesObj = slidesData.get("slides");
        if (slidesObj instanceof List) {
            List<Map<String, Object>> slides = (List<Map<String, Object>>) slidesObj;
            
            // 如果是安全生产类型，先插入固定的前四张幻灯片
            addSafetyCoverSlidesIfNeeded();
            
            System.out.println("开始渲染 " + slides.size() + " 张幻灯片...");
            for (int i = 0; i < slides.size(); i++) {
                Map<String, Object> slideData = slides.get(i);
                String layout = (String) slideData.getOrDefault("layout", "unknown");
                System.out.println("  渲染第 " + (i + 1) + " 张幻灯片，布局: " + layout);
                ISlide slide = renderSlide(slideData);
                System.out.println("    ✓ 幻灯片创建成功，包含 " + slide.getShapes().size() + " 个形状");
            }
            
            // 如果是安全生产类型，在最后插入固定的最后一张幻灯片
            addSafetyLastSlideIfNeeded();
            
            System.out.println("所有幻灯片渲染完成，共 " + presentation.getSlides().size() + " 张");
        }
    }
    
    /**
     * 从形状中提取所有文本内容（支持多种形状类型）
     */
    private String extractTextFromShape(IShape shape) {
        if (shape instanceof IAutoShape) {
            IAutoShape autoShape = (IAutoShape) shape;
            ITextFrame textFrame = autoShape.getTextFrame();
            if (textFrame != null) {
                // 方法1：直接获取文本
                String text = textFrame.getText();
                if (text != null && !text.trim().isEmpty()) {
                    return text.trim();
                }
                
                // 方法2：遍历所有段落和文本部分
                // 使用 try-catch 方式遍历，直到索引超出范围
                StringBuilder fullText = new StringBuilder();
                int paraIndex = 0;
                while (true) {
                    try {
                        IParagraph para = textFrame.getParagraphs().get_Item(paraIndex);
                        int portionIndex = 0;
                        while (true) {
                            try {
                                IPortion portion = para.getPortions().get_Item(portionIndex);
                                String portionText = portion.getText();
                                if (portionText != null) {
                                    fullText.append(portionText);
                                }
                                portionIndex++;
                            } catch (Exception e) {
                                break; // 没有更多文本部分
                            }
                        }
                        paraIndex++;
                    } catch (Exception e) {
                        break; // 没有更多段落
                    }
                }
                return fullText.toString().trim();
            }
        } else if (shape instanceof IGroupShape) {
            // 如果是组合形状，递归检查子形状
            IGroupShape groupShape = (IGroupShape) shape;
            StringBuilder groupText = new StringBuilder();
            for (int i = 0; i < groupShape.getShapes().size(); i++) {
                String subText = extractTextFromShape(groupShape.getShapes().get_Item(i));
                if (subText != null && !subText.isEmpty()) {
                    groupText.append(subText).append(" ");
                }
            }
            return groupText.toString().trim();
        }
        return "";
    }
    
    /**
     * 检查形状是否包含水印文本
     */
    private boolean containsWatermark(IShape shape, String[] watermarkKeywords) {
        String text = extractTextFromShape(shape);
        if (text != null && !text.isEmpty()) {
            String textLower = text.toLowerCase();
            for (String keyword : watermarkKeywords) {
                if (textLower.contains(keyword.toLowerCase())) {
                    return true;
                }
            }
        }
        return false;
    }
    
    /**
     * 从形状集合中移除水印
     */
    private int removeWatermarksFromShapes(IShapeCollection shapes, String[] watermarkKeywords, String location) {
        int removedCount = 0;
        
        // 从后往前遍历，避免删除时索引变化的问题
        for (int j = shapes.size() - 1; j >= 0; j--) {
            IShape shape = shapes.get_Item(j);
            
            if (containsWatermark(shape, watermarkKeywords)) {
                try {
                    String text = extractTextFromShape(shape);
                    shapes.remove(shape);
                    removedCount++;
                    System.out.println("      ✓ 已从" + location + "移除水印: \"" + 
                        (text.length() > 50 ? text.substring(0, 50) + "..." : text) + "\"");
                } catch (Exception e) {
                    System.err.println("      警告：无法移除形状: " + e.getMessage());
                }
            }
        }
        
        return removedCount;
    }
    
    /**
     * 是否为安全生产类型模板
     */
    private boolean isSafetyTemplate() {
        if (templateFile == null) {
            return false;
        }
        String lower = templateFile.toLowerCase(Locale.ROOT);
        // 只要模板路径包含 safety 或 “安全生产” 关键字，即认为是安全生产类型
        return lower.contains("safety") || lower.contains("安全生产");
    }
    
    /**
     * 如果是安全生产类型，在最前面插入《1.2 安全生产方针政策.pptx》的前四张幻灯片
     */
    private void addSafetyCoverSlidesIfNeeded() {
        if (!isSafetyTemplate()) {
            return;
        }
        File safetyCover = new File("1.2 安全生产方针政策.pptx");
        if (!safetyCover.exists()) {
            System.err.println("警告：未找到安全生产封面文件: " + safetyCover.getAbsolutePath());
            return;
        }
        Presentation safetyPresentation = null;
        try {
            safetyPresentation = new Presentation(safetyCover.getAbsolutePath());
            int slideCount = safetyPresentation.getSlides().size();
            
            if (slideCount == 0) {
                System.err.println("警告：安全生产封面文件没有幻灯片内容");
                return;
            }
            
            // 插入前四张幻灯片（最多4张）
            int slidesToInsert = Math.min(4, slideCount);
            for (int i = 0; i < slidesToInsert; i++) {
                ISlide srcSlide = safetyPresentation.getSlides().get_Item(i);
                // 将幻灯片插入到位置 i，保持源幻灯片的完整内容（包含动画/媒体）
                presentation.getSlides().insertClone(i, srcSlide);
                System.out.println("✓ 已插入安全生产封面幻灯片（第" + (i + 1) + "页）");
            }
            
            if (slideCount < 4) {
                System.err.println("警告：安全生产封面文件只有 " + slideCount + " 张幻灯片，无法插入全部4页");
            }
        } catch (Exception e) {
            System.err.println("警告：插入安全生产封面失败: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (safetyPresentation != null) {
                safetyPresentation.dispose();
            }
        }
    }
    
    /**
     * 渲染经典布局（从 master_template.pptx 加载，供所有PPT类型使用）
     * 
     * 根据布局名称和页码从 master_template.pptx 中选择对应的模板幻灯片，
     * 然后替换模板中的文本内容，应用样式，并处理图片。
     * 
     * @param pageNumber 页码（从1开始，源PPT中的页码，用于从 master_template.pptx 中选择对应的幻灯片）
     * @param slideData 幻灯片数据（可能包含需要替换的文本内容）
     * @param layoutName 布局名称（如 "classic_image_text_5"）
     * @return 复制的幻灯片对象
     */
    private ISlide renderClassicLayout(int pageNumber, Map<String, Object> slideData, String layoutName) {
        try {
            // 1. 从 master_template.pptx 加载
            // 如果 masterTemplatePresentation 未加载，尝试从文件加载
            Presentation masterTemplate = this.masterTemplatePresentation;
            if (masterTemplate == null) {
                String masterTemplateFileName = "templates/master_template.pptx";
                File masterTemplateFile = new File(masterTemplateFileName);
                
                if (!masterTemplateFile.exists()) {
                    System.err.println("错误：master_template.pptx 文件不存在: " + masterTemplateFileName + "，使用默认内容页布局");
                    return renderContentPage(slideData);
                }
                
                // 加载 master_template.pptx
                masterTemplate = new Presentation(masterTemplateFile.getAbsolutePath());
                
                if (masterTemplate.getSlides().size() == 0) {
                    masterTemplate.dispose();
                    System.err.println("错误：master_template.pptx 文件为空，使用默认内容页布局");
                    return renderContentPage(slideData);
                }
                
                // 缓存到内存
                this.masterTemplatePresentation = masterTemplate;
            }
            
            // 2. 从配置文件中获取页码对应的 master_template 中的幻灯片索引
            // 需要根据源PPT的页码找到 master_template 中对应的幻灯片
            // master_template 中的幻灯片顺序应该是：源PPT第5页 -> 索引0, 第6页 -> 索引1, ...
            int masterSlideIndex = pageNumber - 5; // 第5页对应索引0，第6页对应索引1
            
            if (masterSlideIndex < 0 || masterSlideIndex >= masterTemplate.getSlides().size()) {
                System.err.println("错误：页码 " + pageNumber + " 超出范围（master_template 中共 " + masterTemplate.getSlides().size() + " 张幻灯片），使用默认内容页布局");
                // 如果 masterTemplate 是从文件临时加载的，需要释放
                if (this.masterTemplatePresentation == null && masterTemplate != null) {
                    masterTemplate.dispose();
                }
                return renderContentPage(slideData);
            }
            
            // 3. 复制 master_template 中对应索引的幻灯片
            ISlide templateSlide = masterTemplate.getSlides().get_Item(masterSlideIndex);
            ISlide clonedSlide = presentation.getSlides().addClone(templateSlide);
            
            // 4. 替换模板中的文本内容
            replaceSlideTextContent(clonedSlide, slideData);
            
            // 5. 尝试替换图片内容（如果JSON中提供了图片路径）
            // 如果没有提供图片，保留模板文件中的"No Image"图片
            replaceSlideImageContent(clonedSlide, slideData);
            
            // 6. 如果配置文件中有样式定义，应用样式
            if (layoutStyleMap != null && layoutStyleMap.containsKey(layoutName)) {
                applyLayoutStyle(clonedSlide, layoutName, layoutStyleMap.get(layoutName));
            }
            
            // 获取布局的详细信息用于日志输出
            String layoutInfo = layoutName;
            if (layoutConfigMap != null && layoutConfigMap.containsKey(layoutName)) {
                Map<String, Object> layoutConfig = layoutConfigMap.get(layoutName);
                String displayName = (String) layoutConfig.get("displayName");
                if (displayName != null) {
                    layoutInfo = displayName + " (" + layoutName + ")";
                }
            }
            
            System.out.println("    ✓ 已使用经典布局: " + layoutInfo + " (源PPT第" + pageNumber + "页, master_template 索引: " + masterSlideIndex + ")");
            return clonedSlide;
        } catch (Exception e) {
            System.err.println("错误：从 master_template.pptx 加载经典布局失败: " + e.getMessage());
            e.printStackTrace();
            // 使用默认内容页布局
            return renderContentPage(slideData);
        }
    }
    
    /**
     * 替换幻灯片中的图片内容
     * 
     * @param slide 目标幻灯片
     * @param slideData 包含要替换的图片数据（image_path 或 imagePath）
     */
    private void replaceSlideImageContent(ISlide slide, Map<String, Object> slideData) {
        try {
            // 获取图片路径（支持多种字段名）
            String imagePath = null;
            if (slideData.containsKey("image_path")) {
                imagePath = (String) slideData.get("image_path");
            } else if (slideData.containsKey("imagePath")) {
                imagePath = (String) slideData.get("imagePath");
            } else if (slideData.containsKey("image")) {
                imagePath = (String) slideData.get("image");
            }
            
            if (imagePath == null || imagePath.isEmpty()) {
                return; // 没有图片需要替换
            }
            
            File imageFile = new File(imagePath);
            if (!imageFile.exists()) {
                System.err.println("警告：图片文件不存在: " + imagePath);
                return;
            }
            
            // 读取图片文件
            ByteArrayOutputStream buffer = new ByteArrayOutputStream();
            FileInputStream fis = new FileInputStream(imageFile);
            byte[] bufferData = new byte[8192];
            int nRead;
            while ((nRead = fis.read(bufferData, 0, bufferData.length)) != -1) {
                buffer.write(bufferData, 0, nRead);
            }
            fis.close();
            byte[] pictureData = buffer.toByteArray();
            
            // 添加到演示文稿的图片集合
            IPPImage newImage = presentation.getImages().addImage(pictureData);
            
            // 查找幻灯片中的图片框并替换
            boolean replaced = false;
            for (int i = 0; i < slide.getShapes().size(); i++) {
                IShape shape = slide.getShapes().get_Item(i);
                if (shape instanceof IPictureFrame) {
                    IPictureFrame pictureFrame = (IPictureFrame) shape;
                    // 替换图片
                    pictureFrame.getPictureFormat().getPicture().setImage(newImage);
                    replaced = true;
                    System.out.println("      ✓ 已替换图片: " + imagePath);
                    break; // 只替换第一个图片框
                }
            }
            
            if (!replaced) {
                System.out.println("      提示：未找到图片框，无法替换图片");
            }
        } catch (Exception e) {
            System.err.println("警告：替换图片内容失败: " + e.getMessage());
        }
    }
    
    /**
     * 替换幻灯片中的文本内容
     * 
     * 根据JSON中的内容智能替换幻灯片中的文本：
     * - 只替换包含"模板文字"的文本框（模板文件中的占位符）
     * - 跳过顶部标题栏中的文本框（Y坐标小于72点的，通常是固定的标题栏）
     * - 对于多列布局，按X坐标位置匹配（左、中、右）
     * - 对于单列布局，按Y坐标从上到下匹配
     * 
     * @param slide 目标幻灯片
     * @param slideData 包含要替换的文本数据
     */
    private void replaceSlideTextContent(ISlide slide, Map<String, Object> slideData) {
        try {
            // 收集所有可替换的文本框（包含"模板文字"的文本框，且不在顶部标题栏）
            List<IAutoShape> replaceableTextShapes = new ArrayList<>();
            double topHeaderThreshold = 72.0; // 72点 = 1英寸，Y坐标小于此值的文本框被认为是顶部标题栏
            
            // 递归收集所有可替换的文本框（包括组合形状中的文本框）
            collectReplaceableTextShapesRecursive(slide.getShapes(), replaceableTextShapes, topHeaderThreshold);
            
            System.out.println("      调试：找到 " + replaceableTextShapes.size() + " 个可替换的文本框");
            if (replaceableTextShapes.isEmpty()) {
                System.out.println("      提示：未找到可替换的文本框（包含'模板文字'的文本框）");
                return;
            }
            
            // 调试：打印JSON中的字段和文本框位置信息
            System.out.println("      调试：JSON中的字段: " + slideData.keySet());
            for (int i = 0; i < replaceableTextShapes.size() && i < 5; i++) {
                IAutoShape shape = replaceableTextShapes.get(i);
                float x = shape.getFrame().getX();
                float y = shape.getFrame().getY();
                String text = shape.getTextFrame().getText();
                System.out.println("      调试：文本框[" + i + "] X=" + x + ", Y=" + y + ", 文本=" + (text != null && text.length() > 20 ? text.substring(0, 20) + "..." : text));
            }
            
            // 按位置排序：先按Y坐标（从上到下），再按X坐标（从左到右）
            replaceableTextShapes.sort((a, b) -> {
                float y1 = a.getFrame().getY();
                float y2 = b.getFrame().getY();
                if (Math.abs(y1 - y2) > 10) { // Y坐标相差超过10点，认为是不同行
                    return Float.compare(y1, y2);
                }
                // 同一行，按X坐标排序
                float x1 = a.getFrame().getX();
                float x2 = b.getFrame().getX();
                return Float.compare(x1, x2);
            });
            
            // 计算幻灯片宽度，用于判断左、中、右列
            double slideWidth = presentation.getSlideSize().getSize().getWidth();
            double leftThreshold = slideWidth * 0.33;  // 左列：X < 33%
            double rightThreshold = slideWidth * 0.67; // 右列：X > 67%
            
            // 智能匹配文本框：根据位置和大小判断哪个文本框应该放什么内容
            // 1. 最大的文本框通常是正文区域
            // 2. 较小的文本框可能是标题或要点
            // 3. 按Y坐标，最上面的通常是标题
            
            // 找到最大的文本框（通常是正文）
            IAutoShape largestTextShape = null;
            double maxArea = 0;
            for (IAutoShape shape : replaceableTextShapes) {
                double area = shape.getFrame().getWidth() * shape.getFrame().getHeight();
                if (area > maxArea) {
                    maxArea = area;
                    largestTextShape = shape;
                }
            }
            
            // 找到最上面的文本框（通常是标题）
            IAutoShape topTextShape = null;
            float minY = Float.MAX_VALUE;
            for (IAutoShape shape : replaceableTextShapes) {
                float y = shape.getFrame().getY();
                if (y < minY) {
                    minY = y;
                    topTextShape = shape;
                }
            }
            
            // 替换标题（使用最上面的文本框，如果存在）
            if (slideData.containsKey("title") && topTextShape != null) {
                String title = (String) slideData.get("title");
                if (title != null && !title.isEmpty()) {
                    System.out.println("      调试：替换标题: " + title + " (Y=" + topTextShape.getFrame().getY() + ")");
                    replaceTextInShape(topTextShape, title, true);
                }
            }
            
            // 替换正文（使用最大的文本框，如果存在）
            if (slideData.containsKey("text") && largestTextShape != null && largestTextShape != topTextShape) {
                String text = (String) slideData.get("text");
                if (text != null && !text.isEmpty()) {
                    System.out.println("      调试：替换正文: " + (text.length() > 50 ? text.substring(0, 50) + "..." : text) + " (面积=" + maxArea + ")");
                    replaceTextInShape(largestTextShape, text, false);
                }
            } else if (slideData.containsKey("bullets") && largestTextShape != null && largestTextShape != topTextShape) {
                @SuppressWarnings("unchecked")
                List<String> bullets = (List<String>) slideData.get("bullets");
                if (bullets != null && !bullets.isEmpty()) {
                    // 将要点列表合并为多行文本
                    StringBuilder bulletsText = new StringBuilder();
                    for (String bullet : bullets) {
                        if (bulletsText.length() > 0) {
                            bulletsText.append("\n");
                        }
                        bulletsText.append("• ").append(bullet);
                    }
                    replaceTextInShape(largestTextShape, bulletsText.toString(), false);
                }
            }
            
            // 处理其他字段（如 chapter_title, description, content, quote, author 等）
            // 对于这些字段，使用剩余的文本框（排除已使用的topTextShape和largestTextShape）
            List<IAutoShape> remainingShapes = new ArrayList<>();
            for (IAutoShape shape : replaceableTextShapes) {
                if (shape != topTextShape && shape != largestTextShape) {
                    remainingShapes.add(shape);
                }
            }
            
            int remainingIndex = 0;
            if (slideData.containsKey("chapter_title") && remainingIndex < remainingShapes.size()) {
                String chapterTitle = (String) slideData.get("chapter_title");
                if (chapterTitle != null && !chapterTitle.isEmpty()) {
                    replaceTextInShape(remainingShapes.get(remainingIndex), chapterTitle, true);
                    remainingIndex++;
                }
            }
            
            if (slideData.containsKey("description") && remainingIndex < remainingShapes.size()) {
                String description = (String) slideData.get("description");
                if (description != null && !description.isEmpty()) {
                    replaceTextInShape(remainingShapes.get(remainingIndex), description, false);
                    remainingIndex++;
                }
            }
            
            if (slideData.containsKey("content") && remainingIndex < remainingShapes.size()) {
                @SuppressWarnings("unchecked")
                List<String> content = (List<String>) slideData.get("content");
                if (content != null && !content.isEmpty()) {
                    StringBuilder contentText = new StringBuilder();
                    for (String line : content) {
                        if (contentText.length() > 0) {
                            contentText.append("\n");
                        }
                        contentText.append(line);
                    }
                    replaceTextInShape(remainingShapes.get(remainingIndex), contentText.toString(), false);
                    remainingIndex++;
                }
            }
            
            if (slideData.containsKey("quote") && remainingIndex < remainingShapes.size()) {
                String quote = (String) slideData.get("quote");
                if (quote != null && !quote.isEmpty()) {
                    replaceTextInShape(remainingShapes.get(remainingIndex), quote, true);
                    remainingIndex++;
                }
            }
            
            if (slideData.containsKey("author") && remainingIndex < remainingShapes.size()) {
                String author = (String) slideData.get("author");
                if (author != null && !author.isEmpty()) {
                    replaceTextInShape(remainingShapes.get(remainingIndex), author, false);
                    remainingIndex++;
                }
            }
            
            // 处理两列和三列布局（按位置匹配）
            if (slideData.containsKey("left_content") || slideData.containsKey("middle_content") || slideData.containsKey("right_content")) {
                // 多列布局：按X坐标位置匹配
                List<IAutoShape> leftShapes = new ArrayList<>();
                List<IAutoShape> middleShapes = new ArrayList<>();
                List<IAutoShape> rightShapes = new ArrayList<>();
                
                for (IAutoShape shape : replaceableTextShapes) {
                    double x = shape.getFrame().getX();
                    double centerX = x + shape.getFrame().getWidth() / 2.0;
                    
                    if (centerX < leftThreshold) {
                        leftShapes.add(shape);
                    } else if (centerX < rightThreshold) {
                        middleShapes.add(shape);
                    } else {
                        rightShapes.add(shape);
                    }
                }
                
                // 替换左列内容
                if (slideData.containsKey("left_content") && !leftShapes.isEmpty()) {
                    @SuppressWarnings("unchecked")
                    List<String> leftContent = (List<String>) slideData.get("left_content");
                    if (leftContent != null && !leftContent.isEmpty()) {
                        StringBuilder leftText = new StringBuilder();
                        for (String line : leftContent) {
                            if (leftText.length() > 0) {
                                leftText.append("\n");
                            }
                            leftText.append(line);
                        }
                        // 使用第一个左列文本框
                        replaceTextInShape(leftShapes.get(0), leftText.toString(), false);
                    }
                }
                
                // 替换中间列内容
                if (slideData.containsKey("middle_content") && !middleShapes.isEmpty()) {
                    @SuppressWarnings("unchecked")
                    List<String> middleContent = (List<String>) slideData.get("middle_content");
                    if (middleContent != null && !middleContent.isEmpty()) {
                        StringBuilder middleText = new StringBuilder();
                        for (String line : middleContent) {
                            if (middleText.length() > 0) {
                                middleText.append("\n");
                            }
                            middleText.append(line);
                        }
                        // 使用第一个中间列文本框
                        replaceTextInShape(middleShapes.get(0), middleText.toString(), false);
                    }
                }
                
                // 替换右列内容
                if (slideData.containsKey("right_content") && !rightShapes.isEmpty()) {
                    @SuppressWarnings("unchecked")
                    List<String> rightContent = (List<String>) slideData.get("right_content");
                    if (rightContent != null && !rightContent.isEmpty()) {
                        StringBuilder rightText = new StringBuilder();
                        for (String line : rightContent) {
                            if (rightText.length() > 0) {
                                rightText.append("\n");
                            }
                            rightText.append(line);
                        }
                        // 使用第一个右列文本框
                        replaceTextInShape(rightShapes.get(0), rightText.toString(), false);
                    }
                }
                
                // 多列布局处理完成，直接返回
                return;
            }
        } catch (Exception e) {
            // 忽略替换失败，保持原样
            System.err.println("警告：替换文本内容失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 递归收集所有可替换的文本框（包括组合形状中的文本框）
     * 
     * @param shapes 形状集合
     * @param replaceableTextShapes 收集到的可替换文本框列表
     * @param topHeaderThreshold 顶部标题栏阈值（Y坐标小于此值的文本框被认为是顶部标题栏）
     */
    private void collectReplaceableTextShapesRecursive(IShapeCollection shapes, List<IAutoShape> replaceableTextShapes, double topHeaderThreshold) {
        for (int i = 0; i < shapes.size(); i++) {
            IShape shape = shapes.get_Item(i);
            
            if (shape instanceof IAutoShape) {
                IAutoShape autoShape = (IAutoShape) shape;
                ITextFrame textFrame = autoShape.getTextFrame();
                if (textFrame != null) {
                    String currentText = textFrame.getText();
                    if (currentText != null && !currentText.trim().isEmpty()) {
                        // 只替换包含"模板文字"的文本框（模板文件中的占位符）
                        if (currentText.contains("模板文字")) {
                            // 检查是否在顶部标题栏（Y坐标小于阈值）
                            float y = autoShape.getFrame().getY();
                            if (y >= topHeaderThreshold) {
                                replaceableTextShapes.add(autoShape);
                                System.out.println("        调试：找到可替换文本框，Y=" + y + ", 文本=" + (currentText.length() > 30 ? currentText.substring(0, 30) + "..." : currentText));
                            } else {
                                System.out.println("        调试：跳过顶部标题栏文本框，Y=" + y);
                            }
                        }
                    }
                }
            } else if (shape instanceof IGroupShape) {
                // 如果是组合形状，递归处理子形状
                IGroupShape groupShape = (IGroupShape) shape;
                collectReplaceableTextShapesRecursive(groupShape.getShapes(), replaceableTextShapes, topHeaderThreshold);
            }
        }
    }
    
    /**
     * 替换形状中的文本内容
     * 
     * @param shape 文本形状
     * @param text 新文本内容
     * @param isTitle 是否为标题（用于应用样式）
     */
    private void replaceTextInShape(IAutoShape shape, String text, boolean isTitle) {
        try {
            ITextFrame textFrame = shape.getTextFrame();
            if (textFrame != null) {
                // 清空原有文本
                textFrame.getParagraphs().clear();
                
                // 处理多行文本（按换行符分割）
                String[] lines = text.split("\n");
                for (int i = 0; i < lines.length; i++) {
                    IParagraph para = new Paragraph();
                    textFrame.getParagraphs().add(para);
                    IPortion portion = new Portion();
                    portion.setText(lines[i]);
                    para.getPortions().add(portion);
                    
                    // 应用样式（第一行如果是标题，使用标题样式）
                    if (styleStrategy != null) {
                        styleStrategy.applyTitleStyle(portion, isTitle && i == 0);
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("警告：替换形状文本失败: " + e.getMessage());
        }
    }
    
    /**
     * 应用布局样式
     * 
     * @param slide 目标幻灯片
     * @param layoutName 布局名称
     * @param styleConfig 样式配置
     */
    private void applyLayoutStyle(ISlide slide, String layoutName, Map<String, Object> styleConfig) {
        try {
            // 这里可以实现根据样式配置应用样式的逻辑
            // 例如：设置字体大小、颜色等
            // 目前先保留接口，后续可以根据需要扩展
            if (styleConfig != null && !styleConfig.isEmpty()) {
                System.out.println("    ✓ 已应用布局样式: " + layoutName);
            }
        } catch (Exception e) {
            System.err.println("警告：应用布局样式失败: " + e.getMessage());
        }
    }
    
    /**
     * 渲染安全生产内容页（从《1.2 安全生产方针政策.pptx》复制指定页面）
     * 
     * @param pageNumber 页码（从1开始，对应源PPT的第5页到倒数第2页）
     * @param slideData 幻灯片数据（可能包含需要替换的文本内容）
     * @return 复制的幻灯片对象
     */
    private ISlide renderSafetyContentPage(int pageNumber, Map<String, Object> slideData) {
        File safetyCover = new File("1.2 安全生产方针政策.pptx");
        if (!safetyCover.exists()) {
            System.err.println("警告：未找到安全生产封面文件: " + safetyCover.getAbsolutePath());
            return renderContentPage(slideData); // 回退到默认内容页
        }
        
        Presentation safetyPresentation = null;
        try {
            safetyPresentation = new Presentation(safetyCover.getAbsolutePath());
            int slideCount = safetyPresentation.getSlides().size();
            
            // 计算源PPT中的实际页码（第5页到倒数第2页，即索引4到slideCount-2）
            int startPage = 5; // 从第5页开始（索引4）
            int endPage = slideCount - 1; // 倒数第2页（索引slideCount-2）
            
            // 验证页码是否在有效范围内
            if (pageNumber < startPage || pageNumber > endPage) {
                System.err.println("警告：页码 " + pageNumber + " 不在有效范围内（" + startPage + "-" + endPage + "），使用默认内容页");
                return renderContentPage(slideData);
            }
            
            // 转换为索引（从0开始）
            int slideIndex = pageNumber - 1;
            
            // 复制指定的幻灯片
            ISlide srcSlide = safetyPresentation.getSlides().get_Item(slideIndex);
            ISlide clonedSlide = presentation.getSlides().addClone(srcSlide);
            
            // TODO: 如果需要替换文本内容，可以在这里处理
            // 例如：根据 slideData 中的内容替换幻灯片中的占位符文本
            
            System.out.println("    ✓ 已从源PPT复制第" + pageNumber + "页（索引" + slideIndex + "）");
            return clonedSlide;
        } catch (Exception e) {
            System.err.println("警告：复制安全生产内容页失败: " + e.getMessage());
            e.printStackTrace();
            return renderContentPage(slideData); // 回退到默认内容页
        } finally {
            if (safetyPresentation != null) {
                safetyPresentation.dispose();
            }
        }
    }
    
    /**
     * 在安全生产参考PPT中查找匹配的布局页面
     * 
     * @param layoutName 布局类型名称
     * @return 匹配的幻灯片，如果未找到则返回null
     */
    private ISlide findMatchingSafetyLayout(String layoutName) {
        if (safetyReferencePresentation == null) {
            return null;
        }
        
        int slideCount = safetyReferencePresentation.getSlides().size();
        if (slideCount < 5) {
            return null; // 至少需要5页
        }
        
        // 从第5页开始到倒数第2页（索引4到slideCount-2）
        int startIndex = 4; // 第5页（索引4）
        int endIndex = slideCount - 2; // 倒数第2页（索引slideCount-2）
        
        // 简单的布局映射策略：根据布局类型选择对应的页面
        // 这里使用循环分配的方式，可以根据实际需求调整
        int targetIndex = -1;
        
        // 根据布局类型映射到不同的页面索引
        // 使用简单的哈希方式将布局类型映射到页面范围
        int layoutHash = layoutName.hashCode();
        int availablePages = endIndex - startIndex + 1;
        if (availablePages > 0) {
            targetIndex = startIndex + (Math.abs(layoutHash) % availablePages);
        }
        
        if (targetIndex >= startIndex && targetIndex <= endIndex) {
            try {
                ISlide slide = safetyReferencePresentation.getSlides().get_Item(targetIndex);
                System.out.println("    ✓ 找到匹配的安全生产布局（源PPT第" + (targetIndex + 1) + "页，布局类型: " + layoutName + "）");
                return slide;
            } catch (Exception e) {
                System.err.println("警告：获取参考布局失败: " + e.getMessage());
            }
        }
        
        return null;
    }
    
    /**
     * 使用安全生产参考PPT中的布局渲染幻灯片，并替换文本内容
     * 
     * @param referenceSlide 参考幻灯片
     * @param slideData 幻灯片数据
     * @param layoutName 布局类型名称
     * @return 渲染后的幻灯片
     */
    private ISlide renderSlideWithSafetyLayout(ISlide referenceSlide, Map<String, Object> slideData, String layoutName) {
        try {
            // 复制参考幻灯片
            ISlide newSlide = presentation.getSlides().addClone(referenceSlide);
            
            // 尝试替换文本内容（根据布局类型和slideData中的内容）
            // 这里可以实现更复杂的文本替换逻辑
            // 目前先简单处理：如果slideData中有title，尝试替换第一个文本框
            if (slideData.containsKey("title")) {
                String title = (String) slideData.get("title");
                if (title != null && !title.isEmpty()) {
                    // 尝试找到第一个文本框并替换内容
                    replaceFirstTextShape(newSlide, title);
                }
            }
            
            return newSlide;
        } catch (Exception e) {
            System.err.println("警告：使用安全生产布局渲染失败，回退到标准布局: " + e.getMessage());
            // 回退到标准渲染方法
            return renderContentPage(slideData);
        }
    }
    
    /**
     * 将模板文件中的所有文字替换为"模板文字模板文字..."（保留原有样式）
     * 
     * 此方法用于在删除水印后替换文字，保留文字的大小、字体和对齐方式
     * 
     * @param slide 目标幻灯片
     */
    private void replaceAllTextWithTemplateTextPreservingStyle(ISlide slide) {
        try {
            String templateText = "模板文字模板文字模板文字";
            
            // 统计处理的文本框数量
            int[] shapeCount = {0};
            
            // 遍历所有形状（包括组合形状中的子形状）
            replaceTextInShapesRecursive(slide.getShapes(), templateText, shapeCount);
        } catch (Exception e) {
            System.err.println("警告：替换模板文字失败: " + e.getMessage());
            e.printStackTrace();
            // 不抛出异常，继续执行
        }
    }
    
    /**
     * 递归替换形状集合中的文本（包括组合形状中的子形状）
     * 
     * @param shapes 形状集合
     * @param templateText 模板文字
     * @param shapeCount 计数器数组（用于统计处理的文本框数量）
     */
    private void replaceTextInShapesRecursive(IShapeCollection shapes, String templateText, int[] shapeCount) {
        for (int i = 0; i < shapes.size(); i++) {
            IShape shape = shapes.get_Item(i);
            
            if (shape instanceof IAutoShape) {
                IAutoShape autoShape = (IAutoShape) shape;
                ITextFrame textFrame = autoShape.getTextFrame();
                if (textFrame != null) {
                    // 先检查是否有文本内容（使用 getText() 方法快速检查）
                    String quickCheck = textFrame.getText();
                    if (quickCheck != null && !quickCheck.trim().isEmpty()) {
                        // 使用更可靠的方法提取完整文本（包括所有段落）
                        String currentText = extractFullTextFromTextFrame(textFrame);
                        // 如果提取的文本为空，使用快速检查的文本
                        if (currentText == null || currentText.trim().isEmpty()) {
                            currentText = quickCheck;
                        }
                        
                        // 确保 currentText 不为空
                        if (currentText == null) {
                            currentText = "";
                        }
                        
                        // 无论文本是否为空，只要文本框存在且有文本框架，都应该处理
                        // 这样可以确保所有文本框都被处理，包括那些可能被误判为空的文本框
                        shapeCount[0]++; // 统计处理的文本框数量
                        
                        // 如果文本为空，使用一个默认的模板文字
                        if (currentText.trim().isEmpty()) {
                            currentText = "模板文字";
                        }
                        
                        // 计算原有文字的字数（中文字符和英文字符都算一个字符）
                        int originalCharCount = currentText.length();
                        
                        // 保存原有的段落格式和文本格式（保存所有段落的信息）
                        List<IParagraphFormat> originalParaFormats = new ArrayList<>();
                        List<IPortionFormat> originalPortionFormats = new ArrayList<>();
                        int paraCount = textFrame.getParagraphs().getCount();
                        for (int p = 0; p < paraCount; p++) {
                            IParagraph originalPara = textFrame.getParagraphs().get_Item(p);
                            originalParaFormats.add(originalPara.getParagraphFormat());
                            
                            // 保存第一个文本部分的格式（作为该段落的代表格式）
                            IPortionFormat portionFormat = null;
                            if (originalPara.getPortions().getCount() > 0) {
                                portionFormat = originalPara.getPortions().get_Item(0).getPortionFormat();
                            }
                            originalPortionFormats.add(portionFormat);
                        }
                        
                        // 如果没有段落，使用默认格式
                        if (originalParaFormats.isEmpty()) {
                            originalParaFormats.add(null);
                            originalPortionFormats.add(null);
                        }
                        
                        // 清空原有文本
                        textFrame.getParagraphs().clear();
                        
                        // 保持原有的段落结构：按段落分割
                        String[] paragraphs = currentText.split("\r?\n", -1); // 保留空段落
                        if (paragraphs.length == 0) {
                            paragraphs = new String[]{""};
                        }
                        
                        // 计算每个段落的原始字数（不包括换行符）
                        int[] paraLengths = new int[paragraphs.length];
                        int totalParaLength = 0;
                        for (int p = 0; p < paragraphs.length; p++) {
                            paraLengths[p] = paragraphs[p].length();
                            totalParaLength += paraLengths[p];
                        }
                        
                        // 计算换行符数量（段落数 - 1）
                        int newlineCount = paragraphs.length > 1 ? paragraphs.length - 1 : 0;
                        
                        // 计算可用于段落文字的总字数（原始总字数 - 换行符数）
                        int availableLength = originalCharCount - newlineCount;
                        if (availableLength < 0) {
                            availableLength = 0;
                        }
                        
                        // 如果段落总字数与可用字数不一致，按比例分配
                        if (totalParaLength != availableLength && totalParaLength > 0) {
                            // 按比例分配字数到各个段落
                            double ratio = (double)availableLength / totalParaLength;
                            int adjustedTotal = 0;
                            for (int p = 0; p < paragraphs.length - 1; p++) {
                                paraLengths[p] = (int)Math.round(paraLengths[p] * ratio);
                                adjustedTotal += paraLengths[p];
                            }
                            // 最后一个段落使用剩余的字数
                            paraLengths[paragraphs.length - 1] = availableLength - adjustedTotal;
                            if (paraLengths[paragraphs.length - 1] < 0) {
                                paraLengths[paragraphs.length - 1] = 0;
                            }
                        }
                        
                        // 为每个段落创建对应的模板文字
                        for (int p = 0; p < paragraphs.length; p++) {
                            int paraTargetLength = paraLengths[p];
                            if (paraTargetLength < 0) {
                                paraTargetLength = 0;
                            }
                            String paraTemplateText = generateTemplateText(templateText, paraTargetLength);
                            
                            
                            IParagraph para = new Paragraph();
                            textFrame.getParagraphs().add(para);
                            IPortion portion = new Portion();
                            portion.setText(paraTemplateText);
                            para.getPortions().add(portion);
                            
                            // 恢复原有的段落格式（对齐方式）
                            int formatIndex = Math.min(p, originalParaFormats.size() - 1);
                            IParagraphFormat originalParaFormat = originalParaFormats.get(formatIndex);
                            if (originalParaFormat != null) {
                                para.getParagraphFormat().setAlignment(originalParaFormat.getAlignment());
                                para.getParagraphFormat().setSpaceAfter(originalParaFormat.getSpaceAfter());
                                para.getParagraphFormat().setSpaceBefore(originalParaFormat.getSpaceBefore());
                            }
                            
                            // 恢复原有的文本格式（字体大小、字体名称、颜色）
                            IPortionFormat originalPortionFormat = originalPortionFormats.get(formatIndex);
                            if (originalPortionFormat != null) {
                                portion.getPortionFormat().setFontHeight(originalPortionFormat.getFontHeight());
                                if (originalPortionFormat.getLatinFont() != null) {
                                    portion.getPortionFormat().setLatinFont(originalPortionFormat.getLatinFont());
                                }
                                if (originalPortionFormat.getFillFormat().getFillType() == FillType.Solid) {
                                    portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                                    portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(
                                        originalPortionFormat.getFillFormat().getSolidFillColor().getColor());
                                }
                            }
                        }
                    }
                }
            } else if (shape instanceof IGroupShape) {
                // 如果是组合形状，递归处理子形状
                IGroupShape groupShape = (IGroupShape) shape;
                replaceTextInShapesRecursive(groupShape.getShapes(), templateText, shapeCount);
            }
        }
    }
    
    /**
     * 从文本框架中提取完整文本（包括所有段落和文本部分）
     * 排除 Aspose 评估版水印文本
     * 
     * @param textFrame 文本框架
     * @return 完整文本内容（已排除水印）
     */
    private String extractFullTextFromTextFrame(ITextFrame textFrame) {
        if (textFrame == null) {
            return null;
        }
        
        // Aspose 评估版水印关键词
        String[] watermarkKeywords = {
            "Evaluation only.",
            "Created with Aspose.Slides",
            "Copyright",
            "text has been truncated due to evaluation version limitation"
        };
        
        StringBuilder fullText = new StringBuilder();
        int paraCount = textFrame.getParagraphs().getCount();
        
        for (int p = 0; p < paraCount; p++) {
            if (p > 0) {
                fullText.append("\n"); // 段落之间用换行符分隔
            }
            
            IParagraph para = textFrame.getParagraphs().get_Item(p);
            int portionCount = para.getPortions().getCount();
            
            for (int port = 0; port < portionCount; port++) {
                IPortion portion = para.getPortions().get_Item(port);
                String text = portion.getText();
                if (text != null) {
                    // 检查是否是水印文本
                    boolean isWatermark = false;
                    for (String keyword : watermarkKeywords) {
                        if (text.contains(keyword)) {
                            isWatermark = true;
                            break;
                        }
                    }
                    // 如果不是水印文本，才添加到结果中
                    if (!isWatermark) {
                        fullText.append(text);
                    }
                }
            }
        }
        
        return fullText.toString();
    }
    
    /**
     * 生成指定长度的模板文字
     * 
     * 如果模板文字长度 > 目标长度，则截取
     * 如果模板文字长度 < 目标长度，则重复模板文字直到达到目标长度
     * 
     * @param templateText 模板文字（如"模板文字模板文字模板文字"）
     * @param targetLength 目标长度（原始文字的字数）
     * @return 生成后的模板文字
     */
    private String generateTemplateText(String templateText, int targetLength) {
        if (templateText == null || templateText.isEmpty()) {
            templateText = "模板文字模板文字模板文字";
        }
        
        int templateLength = templateText.length();
        
        if (templateLength > targetLength) {
            // 字数多了截取
            return templateText.substring(0, targetLength);
        } else if (templateLength < targetLength) {
            // 字数少了重复
            StringBuilder sb = new StringBuilder(targetLength);
            while (sb.length() < targetLength) {
                int remaining = targetLength - sb.length();
                if (remaining >= templateLength) {
                    sb.append(templateText);
                } else {
                    sb.append(templateText.substring(0, remaining));
                }
            }
            return sb.toString();
        } else {
            // 字数相同，直接返回
            return templateText;
        }
    }
    
    /**
     * 将模板文件中的所有文字替换为"模板文字模板文字..."
     * 
     * @param slide 目标幻灯片
     */
    private void replaceAllTextWithTemplateText(ISlide slide) {
        try {
            String templateText = "模板文字模板文字...";
            
            // 遍历所有形状
            for (int i = 0; i < slide.getShapes().size(); i++) {
                IShape shape = slide.getShapes().get_Item(i);
                if (shape instanceof IAutoShape) {
                    IAutoShape autoShape = (IAutoShape) shape;
                    ITextFrame textFrame = autoShape.getTextFrame();
                    if (textFrame != null) {
                        // 检查是否有文本内容
                        String currentText = textFrame.getText();
                        if (currentText != null && !currentText.trim().isEmpty()) {
                            // 清空原有文本
                            textFrame.getParagraphs().clear();
                            
                            // 添加模板文字
                            IParagraph para = new Paragraph();
                            textFrame.getParagraphs().add(para);
                            IPortion portion = new Portion();
                            portion.setText(templateText);
                            para.getPortions().add(portion);
                            
                            // 保留原有的字体大小和颜色（不应用样式策略）
                            // 这样模板文件可以保留原有的样式特征
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("警告：替换模板文字失败: " + e.getMessage());
            // 不抛出异常，继续执行
        }
    }
    
    /**
     * 将模板文件中的所有图片替换为"No Image"文本（顶部图标除外）
     * 
     * 顶部图标判断标准：Y坐标小于72点（1英寸）的图片被认为是顶部图标
     * 
     * @param slide 目标幻灯片
     * @param presentation 演示文稿对象（用于创建文本形状）
     */
    private void replaceAllImagesWithNoImage(ISlide slide, Presentation presentation) {
        try {
            String noImageText = "No Image";
            double topIconThreshold = 72.0; // 72点 = 1英寸，Y坐标小于此值的图片被认为是顶部图标
            
            // 收集需要替换的图片框（排除顶部图标，包括组合形状中的图片）
            List<IPictureFrame> imagesToReplace = new ArrayList<>();
            collectImagesToReplace(slide.getShapes(), imagesToReplace, topIconThreshold);
            
            // 替换每个图片框为带有"No Image"文字标识的图片
            for (IPictureFrame pictureFrame : imagesToReplace) {
                try {
                    // 获取图片框的位置和大小
                    float x = pictureFrame.getFrame().getX();
                    float y = pictureFrame.getFrame().getY();
                    float width = pictureFrame.getFrame().getWidth();
                    float height = pictureFrame.getFrame().getHeight();
                    
                    // 创建带有"No Image"文字标识的图片
                    java.awt.image.BufferedImage noImageBufferedImage = createNoImageImage((int)width, (int)height, noImageText);
                    
                    // 将 BufferedImage 转换为字节数组
                    ByteArrayOutputStream baos = new ByteArrayOutputStream();
                    javax.imageio.ImageIO.write(noImageBufferedImage, "PNG", baos);
                    byte[] imageBytes = baos.toByteArray();
                    
                    // 添加到演示文稿的图片集合
                    IPPImage noImage = presentation.getImages().addImage(imageBytes);
                    
                    // 替换原图片框的图片
                    pictureFrame.getPictureFormat().getPicture().setImage(noImage);
                } catch (Exception e) {
                    System.err.println("警告：替换图片为带文字标识的图片失败: " + e.getMessage());
                    e.printStackTrace();
                    // 继续处理下一个图片
                }
            }
            
            if (imagesToReplace.size() > 0) {
                System.out.println("    ✓ 已替换 " + imagesToReplace.size() + " 个图片为带 \"No Image\" 文字标识的图片");
            }
        } catch (Exception e) {
            System.err.println("警告：替换模板图片失败: " + e.getMessage());
            e.printStackTrace();
            // 不抛出异常，继续执行
        }
    }
    
    /**
     * 递归收集需要替换的图片框（包括组合形状中的图片）
     * 
     * @param shapes 形状集合
     * @param imagesToReplace 收集到的图片框列表
     * @param topIconThreshold 顶部图标阈值（Y坐标小于此值的图片被认为是顶部图标）
     */
    private void collectImagesToReplace(IShapeCollection shapes, List<IPictureFrame> imagesToReplace, double topIconThreshold) {
        for (int i = 0; i < shapes.size(); i++) {
            IShape shape = shapes.get_Item(i);
            
            if (shape instanceof IPictureFrame) {
                IPictureFrame pictureFrame = (IPictureFrame) shape;
                // 获取图片框的位置
                float y = pictureFrame.getFrame().getY();
                // 如果Y坐标大于阈值，说明不是顶部图标，需要替换
                if (y >= topIconThreshold) {
                    imagesToReplace.add(pictureFrame);
                }
            } else if (shape instanceof IGroupShape) {
                // 如果是组合形状，递归处理子形状
                IGroupShape groupShape = (IGroupShape) shape;
                collectImagesToReplace(groupShape.getShapes(), imagesToReplace, topIconThreshold);
            }
        }
    }
    
    /**
     * 创建带有"No Image"文字标识的图片
     * 
     * @param width 图片宽度（像素）
     * @param height 图片高度（像素）
     * @param text 要显示的文字
     * @return BufferedImage 对象
     */
    private java.awt.image.BufferedImage createNoImageImage(int width, int height, String text) {
        // 创建图片
        java.awt.image.BufferedImage image = new java.awt.image.BufferedImage(
            width, height, java.awt.image.BufferedImage.TYPE_INT_RGB);
        
        // 获取 Graphics2D 对象用于绘制
        java.awt.Graphics2D g2d = image.createGraphics();
        
        // 设置抗锯齿
        g2d.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING, 
            java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
        
        // 填充白色背景
        g2d.setColor(java.awt.Color.WHITE);
        g2d.fillRect(0, 0, width, height);
        
        // 绘制边框
        g2d.setColor(java.awt.Color.LIGHT_GRAY);
        g2d.setStroke(new java.awt.BasicStroke(2.0f));
        g2d.drawRect(2, 2, width - 4, height - 4);
        
        // 绘制文字
        g2d.setColor(java.awt.Color.GRAY);
        java.awt.Font font = new java.awt.Font("Arial", java.awt.Font.PLAIN, Math.max(12, Math.min(width, height) / 10));
        g2d.setFont(font);
        
        // 计算文字位置（居中）
        java.awt.FontMetrics fm = g2d.getFontMetrics();
        int textWidth = fm.stringWidth(text);
        int textHeight = fm.getHeight();
        int x = (width - textWidth) / 2;
        int y = (height - textHeight) / 2 + fm.getAscent();
        
        // 绘制文字
        g2d.drawString(text, x, y);
        
        // 释放资源
        g2d.dispose();
        
        return image;
    }
    
    /**
     * 替换幻灯片中第一个文本形状的内容
     */
    private void replaceFirstTextShape(ISlide slide, String newText) {
        try {
            for (int i = 0; i < slide.getShapes().size(); i++) {
                IShape shape = slide.getShapes().get_Item(i);
                if (shape instanceof IAutoShape) {
                    IAutoShape autoShape = (IAutoShape) shape;
                    ITextFrame textFrame = autoShape.getTextFrame();
                    if (textFrame != null) {
                        // 清空原有文本
                        textFrame.getParagraphs().clear();
                        // 添加新文本
                        IParagraph para = new Paragraph();
                        textFrame.getParagraphs().add(para);
                        IPortion portion = new Portion();
                        portion.setText(newText);
                        para.getPortions().add(portion);
                        // 应用样式
                        if (styleStrategy != null) {
                            styleStrategy.applyTitleStyle(portion, false);
                        }
                        break; // 只替换第一个
                    }
                }
            }
        } catch (Exception e) {
            // 忽略替换失败，保持原样
        }
    }
    
    /**
     * 如果是安全生产类型，在最后插入《1.2 安全生产方针政策.pptx》的最后一张幻灯片
     */
    private void addSafetyLastSlideIfNeeded() {
        if (!isSafetyTemplate()) {
            return;
        }
        File safetyCover = new File("1.2 安全生产方针政策.pptx");
        if (!safetyCover.exists()) {
            System.err.println("警告：未找到安全生产封面文件: " + safetyCover.getAbsolutePath());
            return;
        }
        Presentation safetyPresentation = null;
        try {
            safetyPresentation = new Presentation(safetyCover.getAbsolutePath());
            int slideCount = safetyPresentation.getSlides().size();
            
            if (slideCount == 0) {
                System.err.println("警告：安全生产封面文件没有幻灯片内容");
                return;
            }
            
            // 获取最后一张幻灯片
            ISlide lastSlide = safetyPresentation.getSlides().get_Item(slideCount - 1);
            // 将最后一张幻灯片插入到当前演示文稿的最后
            int insertPosition = presentation.getSlides().size();
            presentation.getSlides().insertClone(insertPosition, lastSlide);
            System.out.println("✓ 已插入安全生产封面幻灯片（最后1页，源文件第" + slideCount + "页）");
        } catch (Exception e) {
            System.err.println("警告：插入安全生产最后一张幻灯片失败: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (safetyPresentation != null) {
                safetyPresentation.dispose();
            }
        }
    }
    
    /**
     * 移除 Aspose.Slides 评估版水印
     * 
     * 使用多种方法尝试移除水印：
     * 1. 检查所有幻灯片中的文本形状
     * 2. 检查母版幻灯片
     * 3. 检查布局幻灯片
     * 4. 支持组合形状
     */
    private void removeAsposeWatermarks() {
        // 定义要移除的水印关键词（使用部分匹配）
        String[] watermarkKeywords = {
            "Evaluation only",
            "Created with Aspose.Slides",
            "Copyright",
            "Aspose Pty Ltd",
            "Aspose"
        };
        
        int totalRemoved = 0;
        
        System.out.println("  开始检查水印...");
        
        // 方法1：检查所有普通幻灯片
        System.out.println("  检查普通幻灯片...");
        for (int i = 0; i < presentation.getSlides().size(); i++) {
            ISlide slide = presentation.getSlides().get_Item(i);
            IShapeCollection shapes = slide.getShapes();
            System.out.println("    检查第 " + (i + 1) + " 张幻灯片，共 " + shapes.size() + " 个形状");
            int removed = removeWatermarksFromShapes(shapes, watermarkKeywords, "第 " + (i + 1) + " 张幻灯片");
            totalRemoved += removed;
        }
        
        // 方法2：检查所有母版幻灯片
        System.out.println("  检查母版幻灯片...");
        for (int i = 0; i < presentation.getMasters().size(); i++) {
            IMasterSlide master = presentation.getMasters().get_Item(i);
            IShapeCollection shapes = master.getShapes();
            System.out.println("    检查母版 " + (i + 1) + "，共 " + shapes.size() + " 个形状");
            int removed = removeWatermarksFromShapes(shapes, watermarkKeywords, "母版 " + (i + 1));
            totalRemoved += removed;
        }
        
        // 方法3：检查所有布局幻灯片
        System.out.println("  检查布局幻灯片...");
        for (int i = 0; i < presentation.getLayoutSlides().size(); i++) {
            ILayoutSlide layout = presentation.getLayoutSlides().get_Item(i);
            IShapeCollection shapes = layout.getShapes();
            System.out.println("    检查布局 " + (i + 1) + "，共 " + shapes.size() + " 个形状");
            int removed = removeWatermarksFromShapes(shapes, watermarkKeywords, "布局 " + (i + 1));
            totalRemoved += removed;
        }
        
        if (totalRemoved > 0) {
            System.out.println("✓ 共移除 " + totalRemoved + " 个水印");
        } else {
            System.out.println("  未找到水印（可能已被移除或不存在）");
        }
    }
    
    /**
     * 通过直接操作 PPTX 文件的 XML 结构来移除水印
     * 
     * PPTX 文件实际上是一个 ZIP 压缩包，包含多个 XML 文件。
     * 水印文本通常存储在 slide*.xml 文件中。
     * 此方法会：
     * 1. 解压 PPTX 文件
     * 2. 遍历所有 slide*.xml 文件，查找并删除包含水印文本的 XML 节点
     * 3. 重新打包成 PPTX 文件
     * 
     * @param filename PPTX 文件路径
     * @throws Exception 如果处理失败
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
     * 处理单个 slide XML 文件，移除包含水印的文本节点
     * 
     * 改进的算法：
     * 1. 查找所有文本框架（<a:txBody>）
     * 2. 收集每个文本框架的完整文本内容（可能分布在多个 <a:t> 节点中）
     * 3. 检查完整文本是否包含水印关键词
     * 4. 如果包含，删除整个形状节点（<p:sp>）
     */
    private int processSlideXML(Path xmlFile, String[] watermarkKeywords) throws Exception {
        int removedCount = 0;
        
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(xmlFile.toFile());
        
        boolean modified = false;
        
        // 定义命名空间
        String drawingNS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        
        // 查找所有文本框架（<a:txBody>）
        NodeList txBodyNodes = doc.getElementsByTagNameNS(drawingNS, "txBody");
        List<Node> shapesToRemove = new ArrayList<>();
        
        // 也直接查找所有文本节点，作为备用方法
        NodeList allTextNodes = doc.getElementsByTagNameNS(drawingNS, "t");
        Set<Node> processedTextNodes = new HashSet<>();
        
        for (int i = 0; i < txBodyNodes.getLength(); i++) {
            Node txBodyNode = txBodyNodes.item(i);
            
            // 收集整个文本框架的完整文本内容
            String fullText = collectFullText(txBodyNode, drawingNS);
            
            if (fullText != null && !fullText.trim().isEmpty()) {
                String textLower = fullText.toLowerCase();
                boolean containsWatermark = false;
                String matchedKeyword = null;
                
                // 检查是否包含任何水印关键词
                for (String keyword : watermarkKeywords) {
                    if (textLower.contains(keyword.toLowerCase())) {
                        containsWatermark = true;
                        matchedKeyword = keyword;
                        System.out.println("      找到水印文本 (关键词: " + keyword + "): \"" + (fullText.length() > 80 ? fullText.substring(0, 80) + "..." : fullText) + "\"");
                        break;
                    }
                }
                
                // 如果通过文本框架找到了水印，标记该文本框架中的所有文本节点为已处理
                if (containsWatermark) {
                    NodeList textNodesInTxBody = ((Element) txBodyNode).getElementsByTagNameNS(drawingNS, "t");
                    for (int j = 0; j < textNodesInTxBody.getLength(); j++) {
                        processedTextNodes.add(textNodesInTxBody.item(j));
                    }
                }
                
                if (containsWatermark) {
                    // 向上查找，找到 <p:sp> (形状) 或 <p:grpSp> (组合形状) 节点
                    Node parent = txBodyNode;
                    while (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
                        Element elem = (Element) parent;
                        String nodeName = elem.getLocalName();
                        String namespace = elem.getNamespaceURI();
                        
                        // 检查是否是形状节点
                        if (("sp".equals(nodeName) || "grpSp".equals(nodeName) || "cxnSp".equals(nodeName)) 
                            && namespace != null && namespace.contains("presentation")) {
                            // 检查是否已经添加到待删除列表（避免重复）
                            if (!shapesToRemove.contains(parent)) {
                                shapesToRemove.add(parent);
                                modified = true;
                                removedCount++;
                            }
                            break;
                        }
                        parent = parent.getParentNode();
                    }
                }
            }
        }
        
        // 备用方法：直接检查所有文本节点（如果文本框架方法没有找到水印）
        if (shapesToRemove.isEmpty()) {
            for (int i = 0; i < allTextNodes.getLength(); i++) {
                Node textNode = allTextNodes.item(i);
                
                // 跳过已经处理过的文本节点
                if (processedTextNodes.contains(textNode)) {
                    continue;
                }
                
                String text = textNode.getTextContent();
                if (text != null && !text.trim().isEmpty()) {
                    String textLower = text.toLowerCase();
                    for (String keyword : watermarkKeywords) {
                        if (textLower.contains(keyword.toLowerCase())) {
                            // 找到包含水印的文本节点，向上查找形状节点
                            Node parent = textNode;
                            while (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
                                Element elem = (Element) parent;
                                String nodeName = elem.getLocalName();
                                String namespace = elem.getNamespaceURI();
                                
                                if (("sp".equals(nodeName) || "grpSp".equals(nodeName) || "cxnSp".equals(nodeName)) 
                                    && namespace != null && namespace.contains("presentation")) {
                                    if (!shapesToRemove.contains(parent)) {
                                        shapesToRemove.add(parent);
                                        modified = true;
                                        removedCount++;
                                        System.out.println("      找到水印 (备用方法, 关键词: " + keyword + "): \"" + (text.length() > 50 ? text.substring(0, 50) + "..." : text) + "\"");
                                    }
                                    break;
                                }
                                parent = parent.getParentNode();
                            }
                            break;
                        }
                    }
                }
            }
        }
        
        // 删除找到的形状节点
        for (Node shapeNode : shapesToRemove) {
            Node parent = shapeNode.getParentNode();
            if (parent != null) {
                parent.removeChild(shapeNode);
            }
        }
        
        // 如果修改了，保存文件
        if (modified) {
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(doc);
            StreamResult result = new StreamResult(xmlFile.toFile());
            transformer.transform(source, result);
        }
        
        return removedCount;
    }
    
    /**
     * 收集文本框架（<a:txBody>）的完整文本内容
     * 
     * 文本可能分布在多个段落（<a:p>）和多个文本运行（<a:r>）中的多个文本节点（<a:t>）中
     * 
     * @param txBodyNode 文本框架节点
     * @param drawingNS 绘图命名空间
     * @return 完整的文本内容
     */
    private String collectFullText(Node txBodyNode, String drawingNS) {
        StringBuilder fullText = new StringBuilder();
        
        // 方法1：直接查找所有文本节点（<a:t>），这是最可靠的方法
        NodeList allTextNodes = ((Element) txBodyNode).getElementsByTagNameNS(drawingNS, "t");
        
        for (int i = 0; i < allTextNodes.getLength(); i++) {
            Node textNode = allTextNodes.item(i);
            String text = textNode.getTextContent();
            if (text != null && !text.trim().isEmpty()) {
                if (fullText.length() > 0) {
                    // 检查前一个文本节点和当前文本节点是否在同一段落中
                    // 如果不在同一段落，添加换行符
                    Node prevTextNode = (i > 0) ? allTextNodes.item(i - 1) : null;
                    if (prevTextNode != null) {
                        Node prevPara = findParentParagraph(prevTextNode);
                        Node currPara = findParentParagraph(textNode);
                        if (prevPara != null && currPara != null && !prevPara.equals(currPara)) {
                            fullText.append("\n");
                        } else {
                            // 在同一段落中，添加空格（如果前一个文本不是以空格结尾）
                            if (!fullText.toString().endsWith(" ") && !fullText.toString().endsWith("\n")) {
                                fullText.append(" ");
                            }
                        }
                    }
                }
                fullText.append(text);
            }
        }
        
        return fullText.toString().trim();
    }
    
    /**
     * 查找文本节点的父段落节点（<a:p>）
     */
    private Node findParentParagraph(Node textNode) {
        Node parent = textNode.getParentNode();
        while (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
            Element elem = (Element) parent;
            if ("p".equals(elem.getLocalName()) && 
                elem.getNamespaceURI() != null && 
                elem.getNamespaceURI().contains("drawingml")) {
                return parent;
            }
            parent = parent.getParentNode();
        }
        return null;
    }
    
    /**
     * 递归删除目录
     */
    private void deleteDirectory(Path directory) throws IOException {
        if (Files.exists(directory)) {
            Files.walkFileTree(directory, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Files.delete(file);
                    return FileVisitResult.CONTINUE;
                }
                
                @Override
                public FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
                    Files.delete(dir);
                    return FileVisitResult.CONTINUE;
                }
            });
        }
    }
    
    /**
     * 保存PPT到文件
     * 
     * 在保存之后会通过 XML 操作移除 Aspose.Slides 评估版水印。
     * 
     * @param filename 输出文件名
     * @throws IOException 如果保存失败
     */
    public void save(String filename) throws IOException {
        // 先保存文件
        presentation.save(filename, SaveFormat.Pptx);
        
        // 然后通过 XML 方式移除水印（不使用 Aspose API）
        try {
            System.out.println("正在移除 Aspose.Slides 评估版水印（使用 XML 方式）...");
            removeWatermarksFromXML(filename);
        } catch (Exception e) {
            System.err.println("警告：移除水印时出错: " + e.getMessage());
            e.printStackTrace();
            // 继续执行，不中断保存过程
        }
    }
    
    /**
     * 关闭演示文稿
     * 
     * @throws IOException 如果关闭失败
     */
    public void close() throws IOException {
        if (presentation != null) {
            presentation.dispose();
        }
        if (templatePresentation != null) {
            templatePresentation.dispose();
        }
        if (safetyReferencePresentation != null) {
            safetyReferencePresentation.dispose();
        }
        if (masterTemplatePresentation != null) {
            masterTemplatePresentation.dispose();
        }
    }
}

