package com.pptfactory.util;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.FileVisitResult;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * 提取PPTX文件中每页的备注信息（metadata）并保存为JSON文件
 * 
 * 功能：
 * 1. 解析PPTX文件中每页的备注信息
 * 2. 备注信息为JSON格式的metadata数据
 * 3. 将备注信息提取并保存为JSON文件（T001.json, T002.json等）
 * 4. JSON文件中的template_id字段值为文件名，page_index字段值为页码
 * 5. 将JSON文件保存到与模板PPTX目录相同的metadata文件夹中
 */
public class ExtractPPTNotesMetadataUtil {
    
    private static final ObjectMapper objectMapper = new ObjectMapper();
    
    /**
     * 主方法
     */
    public static void main(String[] args) {
        String pptxPath = "/Users/menggl/workspace/PPTFactory/templates/type_purchase/master_template.pptx";
        extractNotesMetadata(pptxPath);
    }
    
    private static final String DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    
    /**
     * 提取PPTX文件中备注的metadata信息并保存为JSON文件
     * 
     * @param pptxPath PPTX文件路径
     */
    public static void extractNotesMetadata(String pptxPath) {
        Path tempDir = null;
        try {
            System.out.println("=== 开始提取备注metadata信息 ===");
            System.out.println("PPTX文件路径: " + pptxPath);
            
            // 获取PPTX文件所在目录
            File pptxFile = new File(pptxPath);
            if (!pptxFile.exists()) {
                System.err.println("错误: PPTX文件不存在: " + pptxPath);
                return;
            }
            
            String pptxDir = pptxFile.getParent();
            String metadataDir = pptxDir + File.separator + "metadata";
            
            // 创建metadata文件夹（如果不存在）
            Path metadataPath = Paths.get(metadataDir);
            if (!Files.exists(metadataPath)) {
                Files.createDirectories(metadataPath);
                System.out.println("创建metadata文件夹: " + metadataDir);
            } else {
                System.out.println("metadata文件夹已存在: " + metadataDir);
            }
            
            // 同时使用POI和XML两种方法提取备注
            int successCount = 0;
            int failCount = 0;
            
            // 使用POI API读取备注（优先使用，因为更可靠）
            try (FileInputStream fis = new FileInputStream(pptxPath);
                 XMLSlideShow ppt = new XMLSlideShow(fis)) {
                
                int slideIndex = 0;
                
                // 遍历每一页幻灯片
                for (XSLFSlide slide : ppt.getSlides()) {
                    slideIndex++; // 页码从1开始
                    String templateId = String.format("T%03d", slideIndex);
                    
                    System.out.println("\n处理第 " + slideIndex + " 页 (模板ID: " + templateId + ")");
                    
                    try {
                        // 使用POI API提取备注文本（参考CleanAllNoteTextUtil的方法）
                        String noteText = extractNoteTextUsingPOI(slide);
                        
                        if (noteText == null || noteText.trim().isEmpty()) {
                            System.out.println("  ⚠ 该页没有备注信息，跳过");
                            failCount++;
                            continue;
                        }
                        
                        // 打印备注内容的前200个字符用于调试
                        String preview = noteText.length() > 200 ? noteText.substring(0, 200) + "..." : noteText;
                        System.out.println("  备注内容预览: " + preview.replace("\n", "\\n").replace("\r", "\\r"));
                        
                        // 解析JSON并设置template_id和page_index
                        JsonNode metadataJson = parseAndUpdateMetadata(noteText, templateId, slideIndex);
                        
                        // 保存为JSON文件
                        String jsonFileName = templateId + ".json";
                        File jsonFile = new File(metadataDir, jsonFileName);
                        
                        try (FileWriter writer = new FileWriter(jsonFile, java.nio.charset.StandardCharsets.UTF_8)) {
                            objectMapper.writerWithDefaultPrettyPrinter().writeValue(writer, metadataJson);
                            System.out.println("  ✓ 成功保存: " + jsonFile.getAbsolutePath());
                            successCount++;
                        }
                        
                    } catch (Exception e) {
                        System.err.println("  ✗ 处理第 " + slideIndex + " 页时出错: " + e.getMessage());
                        e.printStackTrace();
                        failCount++;
                    }
                }
            }
            
            System.out.println("\n=== 提取完成 ===");
            System.out.println("成功: " + successCount + " 个文件");
            System.out.println("失败: " + failCount + " 个文件");
            
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // 清理临时目录（如果有的话）
            if (tempDir != null && Files.exists(tempDir)) {
                try {
                    deleteDirectory(tempDir);
                    System.out.println("\n已清理临时目录: " + tempDir);
                } catch (Exception e) {
                    System.err.println("清理临时目录失败: " + e.getMessage());
                }
            }
        }
    }
    
    /**
     * 使用POI API提取备注文本（参考CleanAllNoteTextUtil的方法）
     * 
     * @param slide 幻灯片对象
     * @return 备注文本，如果没有备注则返回null
     */
    private static String extractNoteTextUsingPOI(XSLFSlide slide) {
        try {
            XSLFNotes notes = slide.getNotes();
            if (notes == null) {
                return null;
            }
            
            StringBuilder noteTextBuilder = new StringBuilder();
            
            // 遍历备注页中的所有形状，参考CleanAllNoteTextUtil的实现
            for (XSLFShape noteShape : notes.getShapes()) {
                if (noteShape instanceof XSLFTextShape) {
                    XSLFTextShape noteTextShape = (XSLFTextShape) noteShape;
                    String noteText = noteTextShape.getText();
                    if (noteText != null && !noteText.trim().isEmpty()) {
                        if (noteTextBuilder.length() > 0) {
                            noteTextBuilder.append("\n");
                        }
                        noteTextBuilder.append(noteText.trim());
                        System.out.println("    发现备注文本: " + (noteText.length() > 100 ? noteText.substring(0, 100) + "..." : noteText));
                    }
                }
            }
            
            String result = noteTextBuilder.toString().trim();
            if (!result.isEmpty()) {
                System.out.println("    提取到备注文本，总长度: " + result.length());
            }
            
            return result.isEmpty() ? null : result;
            
        } catch (Exception e) {
            System.err.println("    提取备注文本时出错: " + e.getMessage());
            e.printStackTrace();
            return null;
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
     * 获取所有备注文件列表（按序号排序）
     */
    private static List<String> getNotesSlideFiles(Path notesSlidesDir) throws IOException {
        if (!Files.exists(notesSlidesDir)) {
            return Collections.emptyList();
        }
        List<String> notesFiles = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(notesSlidesDir, "notesSlide*.xml")) {
            for (Path path : stream) {
                notesFiles.add(path.getFileName().toString());
            }
        }
        // 按文件名的数字序号排序
        notesFiles.sort((a, b) -> {
            java.util.regex.Pattern pattern = java.util.regex.Pattern.compile("notesSlide(\\d+)\\.xml");
            java.util.regex.Matcher ma = pattern.matcher(a);
            java.util.regex.Matcher mb = pattern.matcher(b);
            if (ma.find() && mb.find()) {
                return Integer.compare(Integer.parseInt(ma.group(1)), Integer.parseInt(mb.group(1)));
            }
            return a.compareTo(b);
        });
        return notesFiles;
    }
    
    /**
     * 创建临时目录
     */
    private static Path createTempDir() throws IOException {
        Path projectRoot = Paths.get(System.getProperty("user.dir"));
        Path tempDir = projectRoot.resolve("temp");
        if (!Files.exists(tempDir)) {
            Files.createDirectories(tempDir);
        }
        String dirName = "extract_notes_" + System.currentTimeMillis() + "_" + Thread.currentThread().getId();
        Path extractTempDir = tempDir.resolve(dirName);
        Files.createDirectories(extractTempDir);
        return extractTempDir;
    }
    
    /**
     * 递归删除目录
     */
    private static void deleteDirectory(Path directory) throws IOException {
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
     * 从备注XML文件中提取文本内容（使用XML直接解析）
     * 
     * @param notesPath 备注XML文件路径
     * @return 备注文本，如果没有备注则返回null
     */
    /**
     * 从备注XML文件中提取文本内容
     * 备注存储在 ppt/notesSlides/notesSlide*.xml 文件中
     * 
     * @param notesPath 备注XML文件路径
     * @return 备注文本，如果没有备注则返回null
     */
    private static String extractNoteTextFromXML(Path notesPath) throws Exception {
        if (!Files.exists(notesPath)) {
            return null;
        }
        
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(notesPath.toFile());
        
        // 使用 getElementsByTagNameNS 方法查找所有文本节点 (a:t 元素)
        // 文本节点位于: <p:txBody><a:p><a:r><a:t>文本内容</a:t></a:r></a:p></p:txBody>
        NodeList textNodes = doc.getElementsByTagNameNS(DML_NS, "t");
        
        System.out.println("    找到 " + textNodes.getLength() + " 个文本节点");
        
        StringBuilder noteTextBuilder = new StringBuilder();
        for (int i = 0; i < textNodes.getLength(); i++) {
            Node textNode = textNodes.item(i);
            String text = textNode.getTextContent();
            if (text != null && !text.trim().isEmpty()) {
                // 检查前一个文本节点和当前文本节点是否在同一段落中
                // 如果不在同一段落，添加换行
                if (noteTextBuilder.length() > 0 && i > 0) {
                    Node prevTextNode = textNodes.item(i - 1);
                    Node prevPara = findParentParagraph(prevTextNode);
                    Node currPara = findParentParagraph(textNode);
                    if (prevPara != null && currPara != null && !prevPara.equals(currPara)) {
                        noteTextBuilder.append("\n");
                    }
                }
                noteTextBuilder.append(text);
            }
        }
        
        String noteText = noteTextBuilder.toString().trim();
        if (!noteText.isEmpty()) {
            System.out.println("    提取到文本，总长度: " + noteText.length());
            // 打印前200个字符用于调试
            String preview = noteText.length() > 200 ? noteText.substring(0, 200) + "..." : noteText;
            System.out.println("    文本预览: " + preview.replace("\n", "\\n"));
        } else {
            // 如果没找到文本，读取完整文件内容用于调试
            String fileContent = new String(Files.readAllBytes(notesPath), java.nio.charset.StandardCharsets.UTF_8);
            System.out.println("    XML文件完整内容:");
            System.out.println(fileContent);
            
            // 尝试查找是否有转义的JSON内容
            if (fileContent.contains("&lt;") || fileContent.contains("&gt;") || fileContent.contains("&quot;")) {
                System.out.println("    注意：XML中可能包含转义字符");
            }
            
            // 尝试直接查找JSON（可能在注释或其他位置）
            int jsonStart = fileContent.indexOf('{');
            int jsonEnd = fileContent.lastIndexOf('}');
            if (jsonStart >= 0 && jsonEnd > jsonStart) {
                String potentialJson = fileContent.substring(jsonStart, jsonEnd + 1);
                System.out.println("    发现可能的JSON内容（位置 " + jsonStart + "-" + jsonEnd + "）:");
                System.out.println(potentialJson);
            }
        }
        
        return noteText.isEmpty() ? null : noteText;
    }
    
    /**
     * 查找文本节点的父段落节点（<a:p>）
     */
    private static Node findParentParagraph(Node textNode) {
        Node parent = textNode.getParentNode();
        while (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
            Element elem = (Element) parent;
            if (DML_NS.equals(elem.getNamespaceURI()) && "p".equals(elem.getLocalName())) {
                return parent;
            }
            parent = parent.getParentNode();
        }
        return null;
    }
    
    /**
     * 解析备注文本中的JSON并更新template_id和page_index字段
     * 
     * @param noteText 备注文本（应该是JSON格式）
     * @param templateId 模板ID（如T001）
     * @param pageIndex 页码（从1开始）
     * @return 解析并更新后的JsonNode对象
     * @throws IOException 如果JSON解析失败
     */
    private static JsonNode parseAndUpdateMetadata(String noteText, String templateId, int pageIndex) 
            throws IOException {
        // 先尝试直接解析JSON
        JsonNode jsonNode = null;
        String jsonText = noteText.trim();
        
        try {
            jsonNode = objectMapper.readTree(jsonText);
        } catch (Exception e) {
            // 如果解析失败，尝试提取JSON部分
            System.out.println("  直接解析失败，尝试提取JSON部分...");
            jsonText = extractJsonFromText(noteText);
            if (jsonText != null && !jsonText.isEmpty()) {
                try {
                    jsonNode = objectMapper.readTree(jsonText);
                } catch (Exception e2) {
                    System.err.println("  提取JSON后解析仍然失败: " + e2.getMessage());
                    throw new IOException("无法从备注中提取有效的JSON: " + e2.getMessage(), e2);
                }
            } else {
                throw new IOException("无法从备注中提取有效的JSON: " + e.getMessage(), e);
            }
        }
        
        // 确保是ObjectNode类型以便修改
        ObjectNode metadataObject;
        if (jsonNode != null && jsonNode.isObject()) {
            // 深拷贝以避免修改原始节点
            metadataObject = (ObjectNode) jsonNode.deepCopy();
        } else {
            // 如果不是对象，创建一个新的对象并将原内容作为内部字段
            metadataObject = objectMapper.createObjectNode();
            if (jsonNode != null) {
                metadataObject.set("data", jsonNode);
            }
        }
        
        // 更新template_id和page_index字段（覆盖原有值）
        metadataObject.put("template_id", templateId);
        metadataObject.put("page_index", pageIndex);
        
        return metadataObject;
    }
    
    /**
     * 从文本中提取JSON部分
     * 尝试找到第一个{到最后一个}之间的内容
     * 
     * @param text 包含JSON的文本
     * @return 提取的JSON文本，如果找不到则返回null
     */
    private static String extractJsonFromText(String text) {
        if (text == null || text.trim().isEmpty()) {
            return null;
        }
        
        String trimmed = text.trim();
        
        // 如果整个文本看起来就是JSON，直接返回
        if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
            return trimmed;
        }
        
        // 尝试找到第一个{和最后一个}
        int firstBrace = trimmed.indexOf('{');
        int lastBrace = trimmed.lastIndexOf('}');
        
        if (firstBrace >= 0 && lastBrace > firstBrace) {
            return trimmed.substring(firstBrace, lastBrace + 1);
        }
        
        // 如果找不到大括号，返回null
        return null;
    }
}
