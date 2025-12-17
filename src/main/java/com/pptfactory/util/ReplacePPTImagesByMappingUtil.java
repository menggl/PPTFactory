package com.pptfactory.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * 根据ppt内容映射.txt文件中的图片路径映射，替换PPT中的图片
 * 参考ReplacePPTImageUtil.java的实现逻辑
 */
/**
 * 批量替换PPTX文件中的图片，依据“ppt内容映射.txt”中的“图片路径映射”字段。
 * 支持每页独立图片映射，自动处理PPTX解压、图片关系、XML更新、重新打包等。
 * 主要流程：
 *   1. 读取内容映射文件，获取每页的图片路径映射。
 *   2. 解压最新生成的PPTX文件到临时目录。
 *   3. 遍历每一页幻灯片，按标注（alt/title/descr）查找图片并替换。
 *   4. 处理PPTX图片关系（embedId），确保每张图片独立。
 *   5. 重新打包为PPTX，清理临时目录。
 *
 * 兼容POI无法直接替换图片的情况，采用直接操作PPTX的XML和media文件。
 *
 * 适用场景：
 *   - 批量自动化替换PPT图片，图片与内容页一一对应。
 *   - 支持图片标注模糊匹配，防止标注微小差异导致替换失败。
 *   - 适合与AI生成图片、下载图片等流程配合使用。
 */
public class ReplacePPTImagesByMappingUtil {
    
    private static final String PROJECT_ROOT = System.getProperty("user.dir");
    private static final String MAPPING_FILE = PROJECT_ROOT + "/produce/ppt内容映射.txt";
    private static final String PPT_DIR = PROJECT_ROOT + "/produce";
    private static final ObjectMapper MAPPER = new ObjectMapper();

    // XML命名空间
    private static final String PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    
    /**
     * 主入口：读取映射文件，查找最新PPT，批量替换图片。
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== 根据图片路径映射替换PPT图片 ===");
            
        // 1. 读取映射文件
            System.out.println("1. 读取映射文件: " + MAPPING_FILE);
        List<Map<String, Object>> mappings = readMappings();
            if (mappings == null || mappings.isEmpty()) {
            System.err.println("未能读取映射文件或格式错误");
            return;
        }
            System.out.println("   ✓ 解析到 " + mappings.size() + " 个页面映射");
            
        // 2. 获取最新PPT文件
            System.out.println("\n2. 查找最新生成的PPT文件");
        String pptFileName = getLatestPptFileName();
        if (pptFileName == null) {
            System.err.println("未找到新生成的PPT文件");
            return;
        }
        String pptPath = PPT_DIR + "/" + pptFileName;
            System.out.println("   ✓ 找到PPT文件: " + pptPath);
            
        // 3. 替换图片
            System.out.println("\n3. 开始替换图片...");
        boolean changed = replaceImages(pptPath, mappings);
            
        if (changed) {
                System.out.println("\n✓ 图片批量替换完成: " + pptPath);
        } else {
                System.out.println("\n未检测到可替换的图片");
            }
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 读取内容映射文件，返回每一页的映射对象。
     */
    private static List<Map<String, Object>> readMappings() throws IOException {
        File mappingFile = new File(MAPPING_FILE);
        if (!mappingFile.exists()) {
            System.err.println("映射文件不存在: " + MAPPING_FILE);
            return null;
        }
        try (InputStream is = new FileInputStream(mappingFile)) {
            return MAPPER.readValue(is, new TypeReference<List<Map<String, Object>>>() {});
        }
    }

    /**
     * 查找produce目录下最新生成的PPTX文件。
     */
    private static String getLatestPptFileName() {
        File dir = new File(PPT_DIR);
        File[] files = dir.listFiles((d, name) -> name.startsWith("new_ppt_") && name.endsWith(".pptx"));
        if (files == null || files.length == 0) {
            return null;
        }
        Arrays.sort(files, Comparator.comparing(File::getName).reversed());
        return files[0].getName();
    }

    /**
     * 替换PPTX文件中的图片，按每页的图片路径映射进行替换。
     * @param pptxPath PPTX文件路径
     * @param mappings 每页的内容映射
     * @return 是否有图片被替换
     */
    private static boolean replaceImages(String pptxPath, List<Map<String, Object>> mappings) throws Exception {
        Path tempDir = null;
        try {
            // 创建临时目录
            tempDir = createTempDirInProject("pptx_replace_images_");
            System.out.println("   临时目录: " + tempDir.toString());
            // 1. 解压PPTX文件
            System.out.println("   解压PPTX文件...");
            unzipPPTX(pptxPath, tempDir.toString());
            // 2. 获取所有幻灯片文件
            List<String> slideFiles = getAvailableSlides(tempDir);
            System.out.println("   找到 " + slideFiles.size() + " 个幻灯片");
            // 3. 遍历每一页，替换图片
            // 注意：每页独立替换，每页的图片使用独立的media文件，避免相互影响
            int replacedCount = 0;
            for (int pageIndex = 0; pageIndex < slideFiles.size() && pageIndex < mappings.size(); pageIndex++) {
                String slideFile = slideFiles.get(pageIndex);
                Map<String, Object> mapping = mappings.get(pageIndex);
                // 获取图片路径映射（只使用当前页的映射）
                Map<String, String> imagePathMap = getStringMap(mapping.get("图片路径映射"));
                if (imagePathMap == null || imagePathMap.isEmpty()) {
                    System.out.println("   第" + (pageIndex + 1) + "页: 无图片路径映射，跳过");
                    continue;
                }
                System.out.println("   第" + (pageIndex + 1) + "页: 开始替换图片（使用该页的独立映射）");
                System.out.println("   该页的图片路径映射: " + imagePathMap);
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFile);
                Path slideRelPath = tempDir.resolve("ppt/slides/_rels").resolve(slideFile + ".rels");
                int pageReplacedCount = replaceImagesInSlide(slidePath, slideRelPath, imagePathMap, tempDir, pageIndex + 1);
                replacedCount += pageReplacedCount;
                System.out.println("   第" + (pageIndex + 1) + "页: 完成，替换了 " + pageReplacedCount + " 张图片");
            }
            if (replacedCount > 0) {
                // 4. 重新打包为PPTX
                System.out.println("   重新打包为PPTX...");
                zipDirectory(tempDir.toString(), pptxPath);
                System.out.println("   共替换了 " + replacedCount + " 个图片");
                return true;
            }
            return false;
        } finally {
            // 清理临时目录
            if (tempDir != null) {
                cleanTempDirectory(tempDir);
            }
        }
    }
    
    /**
     * 替换单个幻灯片(slide)中的图片。
     * 通过标注（title/descr）与图片路径映射匹配，支持精确和模糊匹配。
     * 自动处理图片关系ID，确保每张图片独立。
     * @param slidePath 幻灯片XML路径
     * @param slideRelPath 幻灯片关系XML路径
     * @param imagePathMap 当前页的图片路径映射
     * @param tempDir 临时目录
     * @param pageNum 页码（仅用于日志）
     * @return 替换的图片数量
     */
    private static int replaceImagesInSlide(Path slidePath, Path slideRelPath, 
                                            Map<String, String> imagePathMap, 
                                            Path tempDir, int pageNum) throws Exception {
        if (!Files.exists(slidePath)) {
            return 0;
        }
        
        // 读取幻灯片关系文件
        Map<String, String> relations = new HashMap<>();
            if (Files.exists(slideRelPath)) {
                relations = parseSlideRelations(slideRelPath);
            }
        
        // 解析幻灯片XML
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(slidePath.toFile());
        
        // 查找所有图片元素
        NodeList picNodes = doc.getElementsByTagNameNS(PML_NS, "pic");
        if (picNodes == null || picNodes.getLength() == 0) {
            return 0;
        }
        
        int replacedCount = 0;
        
        // 遍历每个图片元素
        System.out.println("   第" + pageNum + "页: 找到 " + picNodes.getLength() + " 张图片，开始逐一匹配替换");
        System.out.println("   该页的映射key列表: " + imagePathMap.keySet());
        System.out.println("   该页的完整映射: " + imagePathMap);
        
        // 用于跟踪已使用的映射key，避免重复使用
        Set<String> usedKeys = new HashSet<>();
        
        // 用于跟踪embedId的使用情况，如果多张图片共享同一个embedId，需要为它们创建新的关系ID
        Map<String, Integer> embedIdUsageCount = new HashMap<>();
        Map<String, String> embedIdToNewRelId = new HashMap<>(); // 记录embedId到新关系ID的映射
        
        for (int i = 0; i < picNodes.getLength(); i++) {
            Element pic = (Element) picNodes.item(i);
            
            System.out.println("\n   --- 处理第" + pageNum + "页第" + (i+1) + "张图片 ---");
            
            // 获取图片标注（title或descr）
            String annotation = getImageAnnotation(pic);
                if (annotation == null || annotation.trim().isEmpty()) {
                    System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 无标注，跳过");
                    continue; // 没有标注的图片跳过
                }
            
            System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 读取到标注=\"" + annotation + "\"");
            
            // 在映射中查找对应的图片路径（精确匹配）
            String imagePath = imagePathMap.get(annotation);
            String matchedKey = annotation;
            
            // 如果精确匹配失败，尝试模糊匹配（忽略空格、顺序等）
            if (imagePath == null) {
                System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 精确匹配失败，尝试模糊匹配...");
                Map.Entry<String, String> fuzzyMatch = findImagePathByFuzzyMatchWithKey(annotation, imagePathMap);
                if (fuzzyMatch != null) {
                    matchedKey = fuzzyMatch.getKey();
                    imagePath = fuzzyMatch.getValue();
                }
            }
            
            if (imagePath == null) {
                System.err.println("   ❌ 第" + pageNum + "页 图片" + (i+1) + ": 标注 \"" + annotation + "\" 在映射中未找到，跳过此图片");
                System.out.println("   可用的映射key: " + imagePathMap.keySet());
                continue;
            }
            
            // 检查该key是否已被使用（如果一页中有多张图片使用相同的标注，这是正常的）
            if (usedKeys.contains(matchedKey)) {
                System.out.println("   ⚠️  第" + pageNum + "页 图片" + (i+1) + ": 映射key \"" + matchedKey + "\" 已被使用，但继续替换（可能有多张图片使用相同标注）");
            } else {
                usedKeys.add(matchedKey);
            }
            
            System.out.println("   ✓ 第" + pageNum + "页 图片" + (i+1) + ": 匹配成功！");
            System.out.println("      匹配的key: \"" + matchedKey + "\"");
            System.out.println("      将替换为: " + imagePath);
            
            // 检查图片文件是否存在
            File imageFile = new File(PROJECT_ROOT, imagePath);
            if (!imageFile.exists()) {
                System.err.println("   第" + pageNum + "页: 图片文件不存在: " + imagePath);
                continue;
            }
            
            // 获取图片的embed关系ID
            String embedId = getImageEmbedId(pic);
            if (embedId == null || !relations.containsKey(embedId)) {
                System.err.println("   第" + pageNum + "页 图片" + (i+1) + ": 无法找到图片的embed关系");
                continue;
            }
            
            // 跟踪embedId的使用次数
            embedIdUsageCount.put(embedId, embedIdUsageCount.getOrDefault(embedId, 0) + 1);
            int usageCount = embedIdUsageCount.get(embedId);
            
            System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": embedId=" + embedId + " (使用次数: " + usageCount + ")");
            
            // 如果该embedId已被其他图片使用（usageCount > 1），需要为当前图片创建新的关系ID
            // 这样可以确保每张图片都有独立的media文件，不会相互影响
            String actualRelId = embedId;
            if (usageCount > 1) {
                // 检查是否已经为该embedId创建了新的关系ID
                if (!embedIdToNewRelId.containsKey(embedId)) {
                    // 为第一个使用该embedId的图片创建新的关系ID（实际上第一个图片应该保持原关系ID）
                    // 从第二个开始才需要新关系ID
                    System.out.println("   ⚠️  第" + pageNum + "页 图片" + (i+1) + ": 检测到embedId \"" + embedId + "\" 被多张图片共享");
                }
                // 为当前图片创建新的关系ID
                String newRelId = "rId" + (1000 + i); // 使用一个较大的数字避免冲突
                embedIdToNewRelId.put(embedId + "_" + i, newRelId);
                actualRelId = newRelId;
                System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 创建新的关系ID: " + actualRelId + " (原embedId: " + embedId + ")");
            }
            
            String oldImageTarget = relations.get(embedId);
            if (oldImageTarget == null || !oldImageTarget.startsWith("../media/")) {
                System.err.println("   第" + pageNum + "页 图片" + (i+1) + ": embedId对应的target无效: " + oldImageTarget);
                continue;
            }
            
            System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 原图片target=" + oldImageTarget);
            
            Path mediaDir = tempDir.resolve("ppt/media");
            
            // 确定新图片的文件名：使用页面编号、图片索引和embedId，确保每张图片都有独立的文件
            // 格式：image_pageNum_index_embedId.ext，例如：image_3_4_rId5.png
            // 这样可以避免如果两张图片共享同一个embedId时，第二张图片覆盖第一张的问题
            String newImageExtension = getFileExtension(imageFile.getName());
            String embedIdSuffix = embedId.replaceAll("[^a-zA-Z0-9]", "_"); // 清理embedId中的特殊字符
            String newImageFileName = String.format("image_%d_%d_%s.%s", pageNum, i, embedIdSuffix, newImageExtension);
            Path newImagePath = mediaDir.resolve(newImageFileName);
            
            // 检查是否已存在同名文件（避免重复）
            int counter = 0;
            while (Files.exists(newImagePath)) {
                newImageFileName = String.format("image_%d_%d_%s_%d.%s", pageNum, i, embedIdSuffix, counter, newImageExtension);
                newImagePath = mediaDir.resolve(newImageFileName);
                counter++;
            }
            
            System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 新图片文件名=" + newImageFileName);
            
            // 复制新图片到media目录
            Files.copy(imageFile.toPath(), newImagePath, StandardCopyOption.REPLACE_EXISTING);
            
            // 更新关系文件，指向新的图片文件
            // 如果该图片需要新的关系ID（因为共享embedId），需要先在关系文件中添加新关系，然后更新幻灯片XML中的引用
            if (!actualRelId.equals(embedId)) {
                // 需要创建新的关系并更新幻灯片XML中的引用
                System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 需要创建新关系ID并更新幻灯片XML");
                addNewRelationship(slideRelPath, actualRelId, "../media/" + newImageFileName);
                // 更新幻灯片XML中的embed引用
                updateSlideXMLEmbedReference(doc, pic, embedId, actualRelId);
                // 保存更新后的幻灯片XML
                saveXMLDocument(doc, slidePath);
            } else {
                // 直接更新现有关系
                updateSlideRelations(slideRelPath, embedId, "../media/" + newImageFileName);
            }
            
            // 重新读取关系文件，确保后续图片能获取到最新的关系
            relations = parseSlideRelations(slideRelPath);
            
            System.out.println("   第" + pageNum + "页 图片" + (i+1) + ": 关系文件已更新，关系ID=" + actualRelId + " => " + newImageFileName);
            
            replacedCount++;
            System.out.println("   ✓✓ 第" + pageNum + "页 图片" + (i+1) + ": 替换完成！(标注=\"" + annotation + "\" => " + imagePath + ", media文件=" + newImageFileName + ", embedId=" + embedId + ")");
        }
        
        System.out.println("\n   第" + pageNum + "页: 所有图片处理完成，共替换 " + replacedCount + " 张图片");
        return replacedCount;
    }
    
    /**
     * 获取图片标注（title或descr）
     */
    private static String getImageAnnotation(Element pic) {
        NodeList cNvPrList = pic.getElementsByTagNameNS(PML_NS, "cNvPr");
        if (cNvPrList == null || cNvPrList.getLength() == 0) {
            return null;
        }
        
        Element cNvPr = (Element) cNvPrList.item(0);
        String title = cNvPr.getAttribute("title");
        if (title != null && !title.trim().isEmpty()) {
            return title.trim();
        }
        
        String descr = cNvPr.getAttribute("descr");
        if (descr != null && !descr.trim().isEmpty()) {
            return descr.trim();
        }
        
        return null;
    }
    
    /**
     * 通过模糊匹配查找图片路径
     * 支持忽略空格、顺序差异等
     */
    private static String findImagePathByFuzzyMatch(String annotation, Map<String, String> imagePathMap) {
        Map.Entry<String, String> result = findImagePathByFuzzyMatchWithKey(annotation, imagePathMap);
        return result != null ? result.getValue() : null;
    }
    
    /**
     * 通过模糊匹配查找图片路径，并返回匹配的key
     */
    private static Map.Entry<String, String> findImagePathByFuzzyMatchWithKey(String annotation, Map<String, String> imagePathMap) {
            /**
             * 通过模糊匹配查找图片路径，返回匹配的key和value。
             * 支持忽略空格、顺序等。
             */
            // ...existing code...
        // 标准化标注：去除空格，按|分割后排序
        String normalizedAnnotation = normalizeAnnotation(annotation);
        
        // 遍历所有映射key，进行模糊匹配
        for (Map.Entry<String, String> entry : imagePathMap.entrySet()) {
            String key = entry.getKey();
            String normalizedKey = normalizeAnnotation(key);
            
            // 如果标准化后的字符串相同，则认为匹配
            if (normalizedAnnotation.equals(normalizedKey)) {
                System.out.println("   模糊匹配成功: \"" + annotation + "\" => \"" + key + "\"");
                return entry;
            }
        }
        
        return null;
    }
    
    /**
     * 标准化标注字符串：去除空格，按|分割后排序并重新拼接
     */
    private static String normalizeAnnotation(String annotation) {
            /**
             * 标准化标注字符串：去除空格，按|分割后排序。
             * 便于模糊匹配。
             */
            // ...existing code...
        if (annotation == null) {
            return "";
        }
        
        // 按|分割
        String[] parts = annotation.split("\\|");
        List<String> normalizedParts = new ArrayList<>();
        
        for (String part : parts) {
            String trimmed = part.trim();
            if (!trimmed.isEmpty()) {
                normalizedParts.add(trimmed);
            }
        }
        
        // 排序后重新拼接
        Collections.sort(normalizedParts);
        return String.join("|", normalizedParts);
    }
    
    /**
     * 获取图片的embed关系ID
     */
    private static String getImageEmbedId(Element pic) {
            /**
             * 获取图片的embed关系ID。
             * @param pic <pic>元素
             * @return embedId
             */
            // ...existing code...
        NodeList blipList = pic.getElementsByTagNameNS(DML_NS, "blip");
        if (blipList == null || blipList.getLength() == 0) {
            return null;
        }
        
        Element blip = (Element) blipList.item(0);
        return blip.getAttributeNS(REL_NS, "embed");
    }
    
    /**
     * 解析幻灯片关系文件
     */
    private static Map<String, String> parseSlideRelations(Path relPath) throws Exception {
            /**
             * 解析幻灯片关系文件，返回图片关系ID到图片文件的映射。
             * @param relPath 关系XML路径
             * @return 关系映射
             */
            // ...existing code...
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
                String type = elem.getAttribute("Type");
                String target = elem.getAttribute("Target");
                
                // 查找图片类型的关系
                if (type != null && (type.contains("image") || type.contains("picture"))) {
                    relations.put(id, target);
                }
            }
        }
        
        return relations;
    }
    
    /**
     * 更新幻灯片关系文件
     */
    private static void updateSlideRelations(Path relPath, String relId, String newTarget) throws Exception {
            /**
             * 更新幻灯片关系文件，将指定关系ID指向新的图片文件。
             * @param relPath 关系XML路径
             * @param relId 关系ID
             * @param newTarget 新图片文件路径
             */
            // ...existing code...
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
                if (relId.equals(id)) {
                    elem.setAttribute("Target", newTarget);
                    break;
                }
            }
        }
        
        // 保存修改后的XML
        saveXMLDocument(doc, relPath);
    }
    
    /**
     * 在关系文件中添加新的关系
     */
    private static void addNewRelationship(Path relPath, String newRelId, String target) throws Exception {
            /**
             * 在关系文件中添加新的图片关系。
             * @param relPath 关系XML路径
             * @param newRelId 新关系ID
             * @param target 新图片文件路径
             */
            // ...existing code...
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(relPath.toFile());
        
        Element root = doc.getDocumentElement();
        if (root == null) {
            throw new Exception("关系文件格式错误：找不到根元素");
        }
        
        // 创建新的Relationship元素
        Element newRel = doc.createElement("Relationship");
        newRel.setAttribute("Id", newRelId);
        newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        newRel.setAttribute("Target", target);
        
        root.appendChild(newRel);
        
        // 保存修改后的XML
        saveXMLDocument(doc, relPath);
    }
    
    /**
     * 更新幻灯片XML中的embed引用
     */
    private static void updateSlideXMLEmbedReference(Document doc, Element pic, String oldEmbedId, String newEmbedId) {
            /**
             * 更新幻灯片XML中的embed引用。
             * @param doc 幻灯片XML文档
             * @param pic <pic>元素
             * @param oldEmbedId 原embedId
             * @param newEmbedId 新embedId
             */
            // ...existing code...
        // 查找pic元素中的blip元素
        NodeList blipList = pic.getElementsByTagNameNS(DML_NS, "blip");
        if (blipList != null && blipList.getLength() > 0) {
            Element blip = (Element) blipList.item(0);
            String currentEmbedId = blip.getAttributeNS(REL_NS, "embed");
            if (oldEmbedId.equals(currentEmbedId)) {
                blip.setAttributeNS(REL_NS, "embed", newEmbedId);
            }
        }
    }
    
    /**
     * 获取文件扩展名
     */
    private static String getFileExtension(String fileName) {
            /**
             * 获取文件扩展名（不含.）。
             * @param fileName 文件名
             * @return 扩展名
             */
            // ...existing code...
        int lastDot = fileName.lastIndexOf('.');
        if (lastDot > 0 && lastDot < fileName.length() - 1) {
            return fileName.substring(lastDot + 1).toLowerCase();
        }
        return "png"; // 默认扩展名
    }
    
    /**
     * 获取项目目录下的temp目录
     */
    private static Path getProjectTempDir() throws IOException {
            /**
             * 获取项目根目录下的temp目录。
             */
            // ...existing code...
        Path projectRoot = Paths.get(PROJECT_ROOT);
        Path tempDir = projectRoot.resolve("temp");
        if (!Files.exists(tempDir)) {
            Files.createDirectories(tempDir);
        }
        return tempDir;
    }
    
    /**
     * 在项目temp目录下创建临时目录
     */
    private static Path createTempDirInProject(String prefix) throws IOException {
            /**
             * 在项目temp目录下创建唯一临时目录。
             * @param prefix 目录前缀
             * @return 临时目录路径
             */
            // ...existing code...
        Path projectTempDir = getProjectTempDir();
        String dirName = prefix + System.currentTimeMillis() + "_" + Thread.currentThread().getId();
        Path tempDir = projectTempDir.resolve(dirName);
        Files.createDirectories(tempDir);
        return tempDir;
    }
    
    /**
     * 解压PPTX文件
     */
    private static void unzipPPTX(String zipFilePath, String destDirectory) throws IOException {
            /**
             * 解压PPTX文件到指定目录。
             * @param zipFilePath PPTX文件路径
             * @param destDirectory 目标目录
             */
            // ...existing code...
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
            /**
             * 获取临时目录下所有幻灯片文件名（slide*.xml），按页码排序。
             * @param tempDir 临时目录
             * @return 幻灯片文件名列表
             */
            // ...existing code...
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
     * 重新打包目录为PPTX
     */
    private static void zipDirectory(String sourceDir, String zipFilePath) throws IOException {
            /**
             * 将目录重新打包为PPTX（zip格式）。
             * @param sourceDir 源目录
             * @param zipFilePath 目标PPTX路径
             */
            // ...existing code...
        Path sourcePath = Paths.get(sourceDir);
        
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFilePath))) {
            Files.walkFileTree(sourcePath, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Path relativePath = sourcePath.relativize(file);
                    ZipEntry zipEntry = new ZipEntry(relativePath.toString().replace("\\", "/"));
                    zos.putNextEntry(zipEntry);
                    Files.copy(file, zos);
                    zos.closeEntry();
                    return FileVisitResult.CONTINUE;
                }
                
                @Override
                public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
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
     * 保存XML文档
     */
    private static void saveXMLDocument(Document doc, Path filePath) throws Exception {
            /**
             * 保存XML文档到文件。
             * @param doc XML文档
             * @param filePath 目标文件路径
             */
            // ...existing code...
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
        
        DOMSource source = new DOMSource(doc);
        StreamResult result = new StreamResult(filePath.toFile());
        transformer.transform(source, result);
    }
    
    /**
     * 清理临时目录
     */
    private static void cleanTempDirectory(Path tempDir) {
        try {
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
            System.out.println("   临时目录已清理");
        } catch (IOException e) {
            System.err.println("清理临时目录时出错: " + e.getMessage());
        }
    }
    
    /**
     * 将Object转换为Map<String, String>
     */
    private static Map<String, String> getStringMap(Object obj) {
        if (obj instanceof Map) {
            Map<?, ?> map = (Map<?, ?>) obj;
            Map<String, String> result = new LinkedHashMap<>();
            for (Map.Entry<?, ?> e : map.entrySet()) {
                if (e.getKey() != null && e.getValue() != null) {
                    result.put(e.getKey().toString(), e.getValue().toString());
                }
            }
            return result;
        }
        return null;
    }
}
