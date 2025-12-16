package com.pptfactory.util;

import com.aspose.slides.ISlide;
import com.aspose.slides.IShape;
import com.aspose.slides.IPictureFrame;
import com.aspose.slides.IShapeCollection;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.FileVisitResult;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathFactory;
import javax.xml.xpath.XPathConstants;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import java.util.Optional;

/**
 * PPT生产工具类
 * 
 * 功能：
 * 1. 解析ppt内容映射.txt文件
 * 2. 根据模板页编号找到对应的元信息文件，获取page_index
 * 3. 从master_template.pptx拷贝对应页面到新文件
 * 4. 清除水印
 * 5. 清除备注信息
 */
public class ProduceUtil {
    
    // 设置Locale为US，避免Aspose.Slides不支持某些Locale格式的问题
    static {
        Locale.setDefault(Locale.US);
    }
    
    private static final String PROJECT_ROOT = System.getProperty("user.dir");
    private static final String MAPPING_FILE = PROJECT_ROOT + "/produce/ppt内容映射.txt";
    private static final String METADATA_DIR = PROJECT_ROOT + "/templates/metadata";
    private static final String TEMPLATE_FILE = PROJECT_ROOT + "/templates/master_template.pptx";
    private static final String OUTPUT_DIR = PROJECT_ROOT + "/produce";
    
    /**
     * 主方法
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== PPT生产工具 ===");
            String outputFile = producePPT();
            System.out.println("\n✓ 完成！输出文件: " + outputFile);
            
            // 生成图片映射
            System.out.println("\n=== 生成图片映射 ===");
            generateImageMappings(outputFile);
            System.out.println("\n✓ 图片映射生成完成！");
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 生产PPT文件
     * 
     * @return 生成的PPT文件路径
     * @throws Exception 如果处理失败
     */
    public static String producePPT() throws Exception {
        // 1. 解析映射文件
        System.out.println("1. 解析映射文件: " + MAPPING_FILE);
        List<Map<String, Object>> mappings = parseMappingFile();
        System.out.println("   ✓ 解析到 " + mappings.size() + " 个页面映射");
        
        // 2. 收集需要拷贝的页面索引
        System.out.println("\n2. 收集需要拷贝的页面索引");
        List<Integer> pageIndices = new ArrayList<>();
        for (Map<String, Object> mapping : mappings) {
            String templateId = (String) mapping.get("模板页编号");
            int pageIndex = getPageIndexFromMetadata(templateId);
            pageIndices.add(pageIndex);
            System.out.println("   ✓ 模板 " + templateId + " -> 页面索引 " + pageIndex);
        }
        
        // 3. 生成输出文件名
        String outputFileName = generateOutputFileName();
        String outputFile = OUTPUT_DIR + "/" + outputFileName;
        System.out.println("\n3. 输出文件: " + outputFile);
        
        // 4. 拷贝幻灯片
        System.out.println("\n4. 拷贝幻灯片");
        copySlidesFromTemplate(pageIndices, outputFile);
        System.out.println("   ✓ 已拷贝 " + pageIndices.size() + " 个页面");
        
        // 5. 清除水印
        System.out.println("\n5. 清除水印");
        CleanWatermarksUtil.removeWatermarksFromXML(outputFile);
        
        // 6. 清除备注
        System.out.println("\n6. 清除备注信息");
        CleanAllNoteTextUtil.cleanAllNoteText(outputFile);
        
        // 7. 替换文本
        System.out.println("\n7. 替换文本内容");
        replaceTextsInPPT(outputFile, mappings);
        
        return outputFile;
    }
    
    /**
     * 解析映射文件
     */
    @SuppressWarnings("unchecked")
    private static List<Map<String, Object>> parseMappingFile() throws Exception {
        File mappingFile = new File(MAPPING_FILE);
        if (!mappingFile.exists()) {
            throw new RuntimeException("映射文件不存在: " + MAPPING_FILE);
        }
        
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readValue(mappingFile, new TypeReference<List<Map<String, Object>>>() {});
    }
    
    /**
     * 从元信息文件中获取page_index
     */
    @SuppressWarnings("unchecked")
    private static int getPageIndexFromMetadata(String templateId) throws Exception {
        String metadataFile = METADATA_DIR + "/" + templateId + ".json";
        File file = new File(metadataFile);
        
        if (!file.exists()) {
            throw new RuntimeException("元信息文件不存在: " + metadataFile);
        }
        
        ObjectMapper mapper = new ObjectMapper();
        Map<String, Object> metadata = mapper.readValue(file, new TypeReference<Map<String, Object>>() {});
        
        Object pageIndexObj = metadata.get("page_index");
        if (pageIndexObj == null) {
            throw new RuntimeException("元信息文件中没有找到 page_index: " + metadataFile);
        }
        
        // 处理可能是Integer或Long的情况
        int pageIndex;
        if (pageIndexObj instanceof Integer) {
            pageIndex = (Integer) pageIndexObj;
        } else if (pageIndexObj instanceof Long) {
            pageIndex = ((Long) pageIndexObj).intValue();
        } else {
            pageIndex = Integer.parseInt(pageIndexObj.toString());
        }
        
        return pageIndex;
    }
    
    /**
     * 生成输出文件名
     * 格式：new_ppt_[年月日时分秒].pptx
     */
    private static String generateOutputFileName() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        String timestamp = sdf.format(new Date());
        return "new_ppt_" + timestamp + ".pptx";
    }
    
    /**
     * 从模板文件拷贝指定的幻灯片到新文件
     * 
     * @param pageIndices 需要拷贝的页面索引列表（从1开始）
     * @param outputFile 输出文件路径
     */
    private static void copySlidesFromTemplate(List<Integer> pageIndices, String outputFile) throws Exception {
        // 检查模板文件是否存在
        File templateFile = new File(TEMPLATE_FILE);
        if (!templateFile.exists()) {
            throw new RuntimeException("模板文件不存在: " + TEMPLATE_FILE);
        }
        
        // 确保输出目录存在
        File outputDir = new File(OUTPUT_DIR);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }
        
        Presentation templatePresentation = null;
        Presentation newPresentation = null;
        
        try {
            // 加载模板PPT
            templatePresentation = new Presentation(TEMPLATE_FILE);
            int totalSlides = templatePresentation.getSlides().size();
            System.out.println("   模板PPT共有 " + totalSlides + " 页");
            
            // 验证页面索引
            for (int pageIndex : pageIndices) {
                if (pageIndex < 1 || pageIndex > totalSlides) {
                    throw new IllegalArgumentException("页面索引超出范围: " + pageIndex + " (模板共有 " + totalSlides + " 页)");
                }
            }
            
            // 创建新PPT
            newPresentation = new Presentation();
            
            // 设置幻灯片尺寸（与模板PPT一致）
            newPresentation.getSlideSize().setSize(
                (float) templatePresentation.getSlideSize().getSize().getWidth(),
                (float) templatePresentation.getSlideSize().getSize().getHeight(),
                templatePresentation.getSlideSize().getType()
            );
            
            // 删除默认空白页
            if (newPresentation.getSlides().size() > 0) {
                newPresentation.getSlides().removeAt(0);
            }
            
            // 按照顺序拷贝幻灯片（pageIndex从1开始，数组索引从0开始）
            for (int i = 0; i < pageIndices.size(); i++) {
                int pageIndex = pageIndices.get(i);
                ISlide sourceSlide = templatePresentation.getSlides().get_Item(pageIndex - 1);
                newPresentation.getSlides().addClone(sourceSlide);
                System.out.println("   ✓ 已拷贝第 " + (i + 1) + " 个页面 (模板第 " + pageIndex + " 页)");
            }
            
            // 保存新PPT
            newPresentation.save(outputFile, SaveFormat.Pptx);
            
        } finally {
            if (templatePresentation != null) {
                templatePresentation.dispose();
            }
            if (newPresentation != null) {
                newPresentation.dispose();
            }
        }
    }
    
    /**
     * 根据映射关系替换PPT中的文本
     * 
     * @param pptxFile PPTX文件路径
     * @param mappings 映射关系列表（第1条对应第1页，第2条对应第2页）
     */
    @SuppressWarnings("unchecked")
    private static void replaceTextsInPPT(String pptxFile, List<Map<String, Object>> mappings) throws Exception {
        // 创建临时目录
        Path tempDir = Files.createTempDirectory("pptx_replace_");
        
        try {
            // 1. 解压PPTX文件
            System.out.println("   解压PPTX文件...");
            unzipPPTX(pptxFile, tempDir.toString());
            
            // 2. 获取所有幻灯片文件
            List<String> slideFiles = getAvailableSlides(tempDir);
            System.out.println("   找到 " + slideFiles.size() + " 个幻灯片文件");
            
            if (slideFiles.size() < mappings.size()) {
                throw new RuntimeException("幻灯片数量(" + slideFiles.size() + ")少于映射数量(" + mappings.size() + ")");
            }
            
            // 3. 按照映射顺序替换文本（第1条映射对应第1页，第2条映射对应第2页）
            for (int i = 0; i < mappings.size(); i++) {
                int slideIndex = i + 1; // 幻灯片页码从1开始
                Map<String, Object> mapping = mappings.get(i);
                
                // 获取文本映射
                Map<String, Object> textMapping = (Map<String, Object>) mapping.get("文本映射");
                if (textMapping == null || textMapping.isEmpty()) {
                    System.out.println("   跳过第 " + slideIndex + " 页（无文本映射）");
                    continue;
                }
                
                // 获取对应的slide文件
                String slideFileName = slideFiles.get(slideIndex - 1);
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
                
                System.out.println("   处理第 " + slideIndex + " 页: " + slideFileName);
                
                // 替换该页面的所有文本
                boolean replaced = false;
                for (Map.Entry<String, Object> entry : textMapping.entrySet()) {
                    String oldText = entry.getKey();
                    String newText = entry.getValue().toString();
                    
                    if (processSlideXML(slidePath, oldText, newText)) {
                        System.out.println("     ✓ 替换: '" + oldText + "' -> '" + 
                            (newText.length() > 30 ? newText.substring(0, 30) + "..." : newText) + "'");
                        replaced = true;
                    } else {
                        System.out.println("     ⚠ 未找到: '" + oldText + "'");
                    }
                }
                
                if (!replaced) {
                    System.out.println("     ⚠ 第 " + slideIndex + " 页未进行任何替换");
                }
            }
            
            // 4. 重新打包为PPTX
            System.out.println("   重新打包为PPTX...");
            Path tempOutput = Files.createTempFile("pptx_output_", ".pptx");
            try {
                zipDirectory(tempDir.toString(), tempOutput.toString());
                // 原子性地替换原文件
                Files.move(tempOutput, Paths.get(pptxFile), StandardCopyOption.REPLACE_EXISTING);
            } catch (Exception e) {
                try {
                    Files.deleteIfExists(tempOutput);
                } catch (IOException ignored) {}
                throw e;
            }
            
        } finally {
            // 5. 清理临时目录
            cleanTempDirectory(tempDir);
        }
    }
    
    /**
     * 解压PPTX文件
     */
    private static void unzipPPTX(String zipFilePath, String destDirectory) throws IOException {
        try (ZipFile zipFile = new ZipFile(zipFilePath)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                File entryDestination = new File(destDirectory, entry.getName());
                
                if (entry.isDirectory()) {
                    entryDestination.mkdirs();
                } else {
                    entryDestination.getParentFile().mkdirs();
                    
                    try (InputStream in = zipFile.getInputStream(entry);
                         OutputStream out = new FileOutputStream(entryDestination)) {
                        byte[] buffer = new byte[8192];
                        int len;
                        while ((len = in.read(buffer)) > 0) {
                            out.write(buffer, 0, len);
                        }
                    }
                }
            }
        }
    }
    
    /**
     * 获取可用的幻灯片文件列表
     */
    private static List<String> getAvailableSlides(Path tempDir) throws IOException {
        Path slidesDir = tempDir.resolve("ppt/slides");
        if (!Files.exists(slidesDir)) {
            return Collections.emptyList();
        }
        
        List<String> slideFiles = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(slidesDir, "slide*.xml")) {
            for (Path path : stream) {
                slideFiles.add(path.getFileName().toString());
            }
        }
        
        // 按数字排序
        slideFiles.sort((a, b) -> {
            Pattern pattern = Pattern.compile("slide(\\d+)\\.xml");
            Matcher matcherA = pattern.matcher(a);
            Matcher matcherB = pattern.matcher(b);
            
            if (matcherA.find() && matcherB.find()) {
                return Integer.compare(
                    Integer.parseInt(matcherA.group(1)), 
                    Integer.parseInt(matcherB.group(1))
                );
            }
            return a.compareTo(b);
        });
        
        return slideFiles;
    }
    
    /**
     * 处理slide.xml文件，替换文本
     */
    private static boolean processSlideXML(Path slidePath, String oldText, String newText) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        
        Document doc = builder.parse(slidePath.toFile());
        
        // 先尝试使用XPath方法（更精确，针对PPTX的a:t元素）
        boolean replaced = replaceTextUsingXPath(doc, slidePath, oldText, newText);
        
        // 如果XPath方法没找到，再尝试递归方法
        if (!replaced) {
            replaced = replaceTextInElement(doc.getDocumentElement(), oldText, newText);
            if (replaced) {
                saveXMLDocument(doc, slidePath);
            }
        }
        
        return replaced;
    }
    
    /**
     * 使用XPath更精确地定位文本节点（针对PPTX的a:t元素）
     */
    private static boolean replaceTextUsingXPath(Document doc, Path slidePath, String oldText, String newText) throws Exception {
        XPathFactory xpathFactory = XPathFactory.newInstance();
        XPath xpath = xpathFactory.newXPath();
        
        // 注册命名空间
        xpath.setNamespaceContext(new javax.xml.namespace.NamespaceContext() {
            @Override
            public String getNamespaceURI(String prefix) {
                switch (prefix) {
                    case "a":
                        return "http://schemas.openxmlformats.org/drawingml/2006/main";
                    case "p": 
                        return "http://schemas.openxmlformats.org/presentationml/2006/main";
                    case "r": 
                        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    default: 
                        return null;
                }
            }
            
            @Override
            public String getPrefix(String namespaceURI) {
                return null;
            }
            
            @Override
            public Iterator<String> getPrefixes(String namespaceURI) {
                return null;
            }
        });
        
        // 查找所有a:t元素
        String expression = "//a:t";
        NodeList textNodes = (NodeList) xpath.evaluate(expression, doc, XPathConstants.NODESET);
        
        boolean replaced = false;
        for (int i = 0; i < textNodes.getLength(); i++) {
            Node textNode = textNodes.item(i);
            String text = textNode.getTextContent();
            if (text != null && text.contains(oldText)) {
                String replacedText = text.replace(oldText, newText);
                textNode.setTextContent(replacedText);
                replaced = true;
            }
        }
        
        if (replaced) {
            saveXMLDocument(doc, slidePath);
        }
        
        return replaced;
    }
    
    /**
     * 递归搜索并替换XML元素中的文本
     */
    private static boolean replaceTextInElement(Element element, String oldText, String newText) {
        boolean replaced = false;
        
        // 处理当前元素的文本节点
        NodeList childNodes = element.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node child = childNodes.item(i);
            
            if (child.getNodeType() == Node.TEXT_NODE) {
                String text = child.getNodeValue();
                if (text != null && text.contains(oldText)) {
                    String replacedText = text.replace(oldText, newText);
                    child.setNodeValue(replacedText);
                    replaced = true;
                }
            }
        }
        
        // 递归处理子元素
        NodeList children = element.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child.getNodeType() == Node.ELEMENT_NODE) {
                if (replaceTextInElement((Element) child, oldText, newText)) {
                    replaced = true;
                }
            }
        }
        
        return replaced;
    }
    
    /**
     * 保存XML文档
     */
    private static void saveXMLDocument(Document doc, Path filePath) throws Exception {
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        
        // 设置输出属性
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
        
        // 保存文档
        DOMSource source = new DOMSource(doc);
        StreamResult result = new StreamResult(filePath.toFile());
        transformer.transform(source, result);
    }
    
    /**
     * 重新打包目录为PPTX
     */
    private static void zipDirectory(String sourceDir, String zipFilePath) throws IOException {
        Path sourcePath = Paths.get(sourceDir);
        
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFilePath))) {
            Files.walkFileTree(sourcePath, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    // 计算相对路径
                    Path relativePath = sourcePath.relativize(file);
                    
                    // 添加到ZIP
                    ZipEntry zipEntry = new ZipEntry(relativePath.toString().replace("\\", "/"));
                    zos.putNextEntry(zipEntry);
                    Files.copy(file, zos);
                    zos.closeEntry();
                    
                    return FileVisitResult.CONTINUE;
                }
                
                @Override
                public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                    // 对于非根目录，添加目录条目
                    if (!dir.equals(sourcePath)) {
                        Path relativePath = sourcePath.relativize(dir);
                        ZipEntry zipEntry = new ZipEntry(relativePath.toString().replace("\\", "/") + "/");
                        zos.putNextEntry(zipEntry);
                        zos.closeEntry();
                    }
                    return FileVisitResult.CONTINUE;
                }
            });
        }
    }
    
    /**
     * 清理临时目录
     */
    private static void cleanTempDirectory(Path tempDir) {
        try {
            if (Files.exists(tempDir)) {
                Files.walkFileTree(tempDir, new SimpleFileVisitor<Path>() {
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
        } catch (IOException e) {
            System.err.println("清理临时目录时出错: " + e.getMessage());
        }
    }
    
    /**
     * 生成图片映射
     * 遍历新生成的pptx文件，查找图片的标题标注，根据文本映射生成图片提示词，并更新映射文件
     * 
     * @param pptxFile 新生成的PPTX文件路径
     * @throws Exception 如果处理失败
     */
    @SuppressWarnings("unchecked")
    public static void generateImageMappings(String pptxFile) throws Exception {
        // 1. 解析映射文件
        System.out.println("1. 解析映射文件: " + MAPPING_FILE);
        List<Map<String, Object>> mappings = parseMappingFile();
        System.out.println("   ✓ 解析到 " + mappings.size() + " 个页面映射");
        
        // 2. 创建临时目录并解压PPTX
        Path tempDir = Files.createTempDirectory("pptx_image_mapping_");
        try {
            System.out.println("\n2. 解压PPTX文件: " + pptxFile);
            unzipPPTX(pptxFile, tempDir.toString());
            
            // 3. 使用 ScanPPTImageInfoUtil 扫描 PPTX，获取图片注释与显示尺寸
            System.out.println("\n3. 使用 ScanPPTImageInfoUtil 扫描 PPTX，获取图片注释与显示尺寸");
            // ScanPPTImageInfoUtil 会在 produce 目录下生成 ppt_image_info_*.json
            ScanPPTImageInfoUtil.scanPptx(pptxFile);

            // 找到刚生成的最新文件
            File produceDir = new File(PROJECT_ROOT, "produce");
            File[] jsonFiles = produceDir.listFiles((d, name) -> name.startsWith("ppt_image_info_") && name.endsWith(".json"));
            if (jsonFiles == null || jsonFiles.length == 0) {
                System.out.println("   未找到 ScanPPTImageInfoUtil 生成的结果文件，跳过图片映射生成");
                return;
            }
            File latest = Arrays.stream(jsonFiles).max(Comparator.comparingLong(File::lastModified)).get();

            ObjectMapper mapper = new ObjectMapper();
            @SuppressWarnings("unchecked")
            Map<String, Object> scanResult = mapper.readValue(latest, Map.class);
            @SuppressWarnings("unchecked")
            List<Map<String, Object>> items = (List<Map<String, Object>>) scanResult.getOrDefault("items", Collections.emptyList());

            // 按 slide_index 分组
            Map<Integer, List<Map<String, Object>>> bySlide = new HashMap<>();
            for (Map<String, Object> it : items) {
                Object si = it.get("slide_index");
                if (si == null) continue;
                int idx = (si instanceof Number) ? ((Number) si).intValue() : Integer.parseInt(si.toString());
                bySlide.computeIfAbsent(idx, k -> new ArrayList<>()).add(it);
            }

            boolean hasNewMappings = false;

            // 遍历映射并匹配扫描结果
            for (int i = 0; i < mappings.size(); i++) {
                int slideIndex = i + 1;
                Map<String, Object> mapping = mappings.get(i);
                Map<String, Object> textMapping = (Map<String, Object>) mapping.get("文本映射");
                if (textMapping == null || textMapping.isEmpty()) {
                    System.out.println("   跳过第 " + slideIndex + " 页（无文本映射）");
                    continue;
                }

                Map<String, String> imageAnnotationMapping = (Map<String, String>) mapping.get("图片提示词准备");
                if (imageAnnotationMapping == null) {
                    Map<String, String> legacy = (Map<String, String>) mapping.get("图片标注映射");
                    if (legacy != null) {
                        imageAnnotationMapping = new LinkedHashMap<>(legacy);
                        hasNewMappings = true;
                    } else {
                        imageAnnotationMapping = new LinkedHashMap<>();
                    }
                    mapping.put("图片提示词准备", imageAnnotationMapping);
                }

                List<Map<String, Object>> slideItems = bySlide.getOrDefault(slideIndex, Collections.emptyList());
                if (slideItems.isEmpty()) {
                    System.out.println("   第 " + slideIndex + " 页未发现图片注释");
                    continue;
                }

                System.out.println("   处理第 " + slideIndex + " 页，发现 " + slideItems.size() + " 个注释图片");

                for (Map<String, Object> it : slideItems) {
                    String title = Optional.ofNullable(it.get("annotation")).map(Object::toString).orElse("");
                    if (title == null || title.trim().isEmpty()) continue;
                    if ("警告".equals(title.trim())) continue;
                    if (!(title.contains("我是文本") || title.contains("我是长文本"))) continue;

                    if (imageAnnotationMapping.containsKey(title)) continue;

                    // 拆分标注并匹配文本映射
                    String[] parts = title.split("\\|");
                    List<String> replacementTexts = new ArrayList<>();
                    List<String> otherInfo = new ArrayList<>();
                    for (String part : parts) {
                        part = part.trim();
                        if (part.isEmpty()) continue;
                        String replacement = null;
                        if (textMapping.containsKey(part)) replacement = textMapping.get(part).toString();
                        else {
                            String longKey = part.replace("我是文本", "我是长文本");
                            if (textMapping.containsKey(longKey)) replacement = textMapping.get(longKey).toString();
                            else {
                                String shortKey = part.replace("我是长文本", "我是文本");
                                if (textMapping.containsKey(shortKey)) replacement = textMapping.get(shortKey).toString();
                            }
                        }
                        if (replacement != null && !replacement.isEmpty()) replacementTexts.add(replacement);
                        else { otherInfo.add(part); }
                    }

                    // 图片大小信息：优先使用扫描结果中的 width_px_120dpi/height_px_120dpi
                    String imageSizeInfo = "";
                    int widthPx = -1, heightPx = -1;
                    Object wObj = it.get("width_px_120dpi");
                    Object hObj = it.get("height_px_120dpi");
                    if (wObj instanceof Number) widthPx = ((Number) wObj).intValue();
                    if (hObj instanceof Number) heightPx = ((Number) hObj).intValue();
                    if (widthPx <= 0 || heightPx <= 0) {
                        Object wc = it.get("width_cm");
                        Object hc = it.get("height_cm");
                        if (wc instanceof Number && hc instanceof Number) {
                            widthPx = (int) Math.round(((Number) wc).doubleValue() * 120);
                            heightPx = (int) Math.round(((Number) hc).doubleValue() * 120);
                        }
                    }
                    if (widthPx > 0 && heightPx > 0) {
                        imageSizeInfo = String.format("图片尺寸为像素%d宽 × %d像素高", widthPx, heightPx);
                    }

                    if (!replacementTexts.isEmpty() || !otherInfo.isEmpty()) {
                        List<String> promptParts = new ArrayList<>();
                        promptParts.addAll(replacementTexts);
                        promptParts.addAll(otherInfo);
                        if (!imageSizeInfo.isEmpty()) promptParts.add(imageSizeInfo);
                        String imagePrompt = String.join("|", promptParts);
                        imageAnnotationMapping.put(title, imagePrompt);
                        hasNewMappings = true;
                        System.out.println("       生成图片提示词: " + title + " => " + (imagePrompt.length() > 80 ? imagePrompt.substring(0,80)+"..." : imagePrompt));
                    }
                }
            }

            // 保存映射
            if (hasNewMappings) {
                System.out.println("\n4. 更新映射文件");
                mapper.writerWithDefaultPrettyPrinter().writeValue(new File(MAPPING_FILE), mappings);
                System.out.println("   ✓ 已更新映射文件: " + MAPPING_FILE);
            } else {
                System.out.println("\n4. 无需更新映射文件（没有新的图片映射）");
            }
        } finally {
            // 清理临时目录
            cleanTempDirectory(tempDir);
        }
    }
    
    /**
     * 解析幻灯片关系文件
     */
    private static Map<String, String> parseSlideRelations(Path relPath) throws Exception {
        Map<String, String> relations = new HashMap<>();
        
        if (!Files.exists(relPath)) {
            return relations;
        }
        
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(relPath.toFile());
        
        NodeList relationshipNodes = doc.getElementsByTagName("Relationship");
        for (int i = 0; i < relationshipNodes.getLength(); i++) {
            Node node = relationshipNodes.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element elem = (Element) node;
                String id = elem.getAttribute("Id");
                String target = elem.getAttribute("Target");
                if (id != null && target != null) {
                    relations.put(id, target);
                }
            }
        }
        
        return relations;
    }
}
