package com.pptfactory.util;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipFile;
import java.util.zip.ZipEntry;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

/**
 * 检查PPT文件中指定页面的图片
 */
public class CheckSlideImages {
    
    public static void main(String[] args) throws Exception {
        // 找到最新的PPT文件
        String produceDir = "produce";
        File dir = new File(produceDir);
        File[] files = dir.listFiles((d, name) -> name.startsWith("new_ppt_") && name.endsWith(".pptx"));
        
        if (files == null || files.length == 0) {
            System.out.println("未找到PPT文件");
            return;
        }
        
        // 按文件名排序，获取最新的
        File latestFile = null;
        for (File f : files) {
            if (latestFile == null || f.getName().compareTo(latestFile.getName()) > 0) {
                latestFile = f;
            }
        }
        
        System.out.println("检查文件: " + latestFile.getAbsolutePath());
        
        // 解压PPTX
        Path tempDir = Files.createTempDirectory("check_slide_");
        try {
            unzipPPTX(latestFile.getAbsolutePath(), tempDir.toString());
            
            // 获取所有幻灯片文件
            List<String> slideFiles = getAvailableSlides(tempDir);
            System.out.println("找到 " + slideFiles.size() + " 个幻灯片");
            
            // 查找T008对应的页面（在映射文件中，第一个T008是第9个条目，索引8，对应第9页）
            // 第二个T008是第15个条目，索引14，对应第15页
            int[] t008Pages = {8, 14}; // 页面索引（从0开始，对应第9页和第15页）
            
            final String PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
            
            for (int pageIdx : t008Pages) {
                if (pageIdx >= slideFiles.size()) {
                    System.out.println("\n页面索引 " + pageIdx + " 超出范围（总页数: " + slideFiles.size() + "）");
                    continue;
                }
                
                int slideIndex = pageIdx + 1; // 页码从1开始
                String slideFileName = slideFiles.get(pageIdx);
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
                
                System.out.println("\n=== 检查第 " + slideIndex + " 页 (T008) ===");
                System.out.println("文件: " + slideFileName);
                
                if (!Files.exists(slidePath)) {
                    System.out.println("文件不存在: " + slidePath);
                    continue;
                }
                
                // 解析幻灯片XML
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(slidePath.toFile());
                
                // 查找所有图片
                NodeList picNodes = doc.getElementsByTagNameNS(PML_NS, "pic");
                System.out.println("找到 " + picNodes.getLength() + " 个图片元素");
                
                if (picNodes.getLength() == 0) {
                    System.out.println("  → 该页面没有图片");
                    continue;
                }
                
                // 检查每个图片的标注
                for (int i = 0; i < picNodes.getLength(); i++) {
                    Element pic = (Element) picNodes.item(i);
                    
                    // 读取标题标注
                    String title = "";
                    NodeList cNvPrList = pic.getElementsByTagNameNS(PML_NS, "cNvPr");
                    if (cNvPrList != null && cNvPrList.getLength() > 0) {
                        Element cNvPr = (Element) cNvPrList.item(0);
                        title = cNvPr.getAttribute("title");
                    }
                    
                    // 读取描述
                    String descr = "";
                    NodeList cNvPrList2 = pic.getElementsByTagNameNS(PML_NS, "cNvPr");
                    if (cNvPrList2 != null && cNvPrList2.getLength() > 0) {
                        Element cNvPr = (Element) cNvPrList2.item(0);
                        descr = cNvPr.getAttribute("descr");
                    }
                    
                    System.out.println("\n  图片 " + (i + 1) + ":");
                    System.out.println("    标题(title): " + (title.isEmpty() ? "(空)" : title));
                    System.out.println("    描述(descr): " + (descr.isEmpty() ? "(空)" : descr));
                    
                    if (title.isEmpty() && descr.isEmpty()) {
                        System.out.println("    → 该图片没有标注！");
                    } else if (!title.contains("我是文本") && !title.contains("我是长文本") 
                            && !descr.contains("我是文本") && !descr.contains("我是长文本")) {
                        System.out.println("    → 标注不包含'我是文本'或'我是长文本'，会被跳过");
                    }
                }
            }
            
        } finally {
            // 清理临时目录
            deleteDirectory(tempDir.toFile());
        }
    }
    
    private static void unzipPPTX(String pptxFile, String destDir) throws Exception {
        try (ZipFile zipFile = new ZipFile(pptxFile)) {
            java.util.Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                File file = new File(destDir, entry.getName());
                if (entry.isDirectory()) {
                    file.mkdirs();
                } else {
                    file.getParentFile().mkdirs();
                    try (java.io.InputStream is = zipFile.getInputStream(entry);
                         java.io.FileOutputStream fos = new java.io.FileOutputStream(file)) {
                        byte[] buffer = new byte[8192];
                        int len;
                        while ((len = is.read(buffer)) > 0) {
                            fos.write(buffer, 0, len);
                        }
                    }
                }
            }
        }
    }
    
    private static List<String> getAvailableSlides(Path tempDir) throws Exception {
        List<String> slides = new ArrayList<>();
        Path slidesDir = tempDir.resolve("ppt/slides");
        if (!Files.exists(slidesDir)) {
            return slides;
        }
        
        Files.list(slidesDir)
            .filter(p -> p.getFileName().toString().startsWith("slide") 
                    && p.getFileName().toString().endsWith(".xml"))
            .sorted()
            .forEach(p -> slides.add(p.getFileName().toString()));
        
        return slides;
    }
    
    private static void deleteDirectory(File dir) {
        if (dir.exists()) {
            File[] files = dir.listFiles();
            if (files != null) {
                for (File f : files) {
                    if (f.isDirectory()) {
                        deleteDirectory(f);
                    } else {
                        f.delete();
                    }
                }
            }
            dir.delete();
        }
    }
}
