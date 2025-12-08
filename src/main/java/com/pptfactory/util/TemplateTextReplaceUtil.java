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
public class TemplateTextReplaceUtil {
    public static void main(String[] args) throws Exception {
        replaceTextFromXML("/Users/menggl/workspace/PPTFactory/master_template.pptx");
    }
    private static void replaceTextFromXML(String filename) throws Exception {
        System.out.println("  使用 XML 方式替换文本...");
        
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
            Path pptDir = extractedDir.resolve("ppt");
            if (Files.exists(pptDir)) {
                // 处理 slides 目录
                Path slidesDir = pptDir.resolve("slides");
                if (Files.exists(slidesDir)) {
                    try (DirectoryStream<Path> stream = Files.newDirectoryStream(slidesDir, "slide*.xml")) {
                        for (Path slideFile : stream) {
                            processSlideXML(slideFile);
                        }
                    }
                }
                
                // 处理 slideMasters 目录
                Path slideMastersDir = pptDir.resolve("slideMasters");
                if (Files.exists(slideMastersDir)) {
                    try (DirectoryStream<Path> stream = Files.newDirectoryStream(slideMastersDir, "*.xml")) {
                        for (Path masterFile : stream) {
                            processSlideXML(masterFile);
                        }
                    }
                }
                
                // 处理 slideLayouts 目录
                Path slideLayoutsDir = pptDir.resolve("slideLayouts");
                if (Files.exists(slideLayoutsDir)) {
                    try (DirectoryStream<Path> stream = Files.newDirectoryStream(slideLayoutsDir, "*.xml")) {
                        for (Path layoutFile : stream) {
                            processSlideXML(layoutFile);
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
    private static void processSlideXML(Path xmlFile) throws Exception {
        
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(xmlFile.toFile());

        // 定义命名空间
        String drawingNS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        
        // 查找所有文本框架（<a:txBody>）
        NodeList txBodyNodes = doc.getElementsByTagNameNS(drawingNS, "txBody");
        
        for (int i = 0; i < txBodyNodes.getLength(); i++) {
            Node txBodyNode = txBodyNodes.item(i);
            
            // 替换文本框架中的所有文本内容为模板文字
            replaceAllTextWithTemplateText(txBodyNode, drawingNS);
        }
        
        // // 备用方法：直接检查所有文本节点（如果文本框架方法没有找到水印）
        // if (shapesToRemove.isEmpty()) {
        //     for (int i = 0; i < allTextNodes.getLength(); i++) {
        //         Node textNode = allTextNodes.item(i);
                
        //         // 跳过已经处理过的文本节点
        //         if (processedTextNodes.contains(textNode)) {
        //             continue;
        //         }
                
        //         String text = textNode.getTextContent();
        //         if (text != null && !text.trim().isEmpty()) {
        //             String textLower = text.toLowerCase();
        //             for (String keyword : watermarkKeywords) {
        //                 if (textLower.contains(keyword.toLowerCase())) {
        //                     // 找到包含水印的文本节点，向上查找形状节点
        //                     Node parent = textNode;
        //                     while (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
        //                         Element elem = (Element) parent;
        //                         String nodeName = elem.getLocalName();
        //                         String namespace = elem.getNamespaceURI();
                                
        //                         if (("sp".equals(nodeName) || "grpSp".equals(nodeName) || "cxnSp".equals(nodeName)) 
        //                             && namespace != null && namespace.contains("presentation")) {
        //                             if (!shapesToRemove.contains(parent)) {
        //                                 shapesToRemove.add(parent);
        //                                 modified = true;
        //                                 removedCount++;
        //                                 System.out.println("      找到水印 (备用方法, 关键词: " + keyword + "): \"" + (text.length() > 50 ? text.substring(0, 50) + "..." : text) + "\"");
        //                             }
        //                             break;
        //                         }
        //                         parent = parent.getParentNode();
        //                     }
        //                     break;
        //                 }
        //             }
        //         }
        //     }
        // }
        
        
        // 如果修改了，保存文件
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        DOMSource source = new DOMSource(doc);
        StreamResult result = new StreamResult(xmlFile.toFile());
        transformer.transform(source, result);
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
    private static void replaceAllTextWithTemplateText(Node txBodyNode, String drawingNS) {
        
        // 方法1：直接查找所有文本节点（<a:t>），这是最可靠的方法
        NodeList allTextNodes = ((Element) txBodyNode).getElementsByTagNameNS(drawingNS, "t");
        
        for (int i = 0; i < allTextNodes.getLength(); i++) {
            Node textNode = allTextNodes.item(i);
            String text = textNode.getTextContent();
            if (text != null && !text.trim().isEmpty()) {
                String templateText = TemplateTextGenerate.generateTemplateText(text.length());
                textNode.setTextContent(templateText);
                System.out.println("替换文本: " + text + " -> " + templateText);
            } else {
                NodeList allChildNodes = ((Element) textNode).getElementsByTagNameNS(drawingNS, "t");
                if(allChildNodes != null && allChildNodes.getLength() > 0) {
                    replaceAllTextWithTemplateText(textNode,drawingNS);
                }
            }
        }
    }
}
