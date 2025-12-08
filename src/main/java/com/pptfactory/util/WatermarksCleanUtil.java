package com.pptfactory.util;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.FileVisitResult;
import java.util.*;
import java.util.List;
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
public class WatermarksCleanUtil {
    public static void main(String[] args) throws Exception {
        removeWatermarksFromXML("src/test/resources/test.pptx");
    }

    private static void removeWatermarksFromXML(String filename) throws Exception {
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
    private static int processSlideXML(Path xmlFile, String[] watermarkKeywords) throws Exception {
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
     * 收集文本框架（<a:txBody>）的完整文本内容
     * 
     * 文本可能分布在多个段落（<a:p>）和多个文本运行（<a:r>）中的多个文本节点（<a:t>）中
     * 
     * @param txBodyNode 文本框架节点
     * @param drawingNS 绘图命名空间
     * @return 完整的文本内容
     */
    private static String collectFullText(Node txBodyNode, String drawingNS) {
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
    private static Node findParentParagraph(Node textNode) {
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
}
