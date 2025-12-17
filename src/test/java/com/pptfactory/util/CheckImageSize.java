package com.pptfactory.util;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
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
 * 检查PPT文件中图片的实际尺寸信息
 */
public class CheckImageSize {
    
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
        Path tempDir = Files.createTempDirectory("check_image_size_");
        try {
            unzipPPTX(latestFile.getAbsolutePath(), tempDir.toString());
            
            // 获取所有幻灯片文件
            List<String> slideFiles = getAvailableSlides(tempDir);
            System.out.println("找到 " + slideFiles.size() + " 个幻灯片\n");
            
            final String PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
            final String DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
            
            // 检查第15页（T012模板页）
            int pageIdx = 14; // 第15页，索引14
            if (pageIdx >= slideFiles.size()) {
                System.out.println("页面索引超出范围");
                return;
            }
            
            int slideIndex = pageIdx + 1;
            String slideFileName = slideFiles.get(pageIdx);
            Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
            
            System.out.println("=== 检查第 " + slideIndex + " 页 (T012) ===");
            System.out.println("文件: " + slideFileName + "\n");
            
            // 解析幻灯片XML
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(slidePath.toFile());
            
            // 查找所有图片
            NodeList picNodes = doc.getElementsByTagNameNS(PML_NS, "pic");
            System.out.println("找到 " + picNodes.getLength() + " 个图片元素\n");
            
            for (int i = 0; i < picNodes.getLength(); i++) {
                Element pic = (Element) picNodes.item(i);
                
                // 读取标题标注
                String title = "";
                NodeList cNvPrList = pic.getElementsByTagNameNS(PML_NS, "cNvPr");
                if (cNvPrList != null && cNvPrList.getLength() > 0) {
                    Element cNvPr = (Element) cNvPrList.item(0);
                    title = cNvPr.getAttribute("title");
                }
                
                System.out.println("图片 " + (i + 1) + ":");
                System.out.println("  标注: " + (title.isEmpty() ? "(空)" : title));
                
                // 查找 <a:xfrm> 元素
                NodeList xfrmList = pic.getElementsByTagNameNS(DML_NS, "xfrm");
                if (xfrmList != null && xfrmList.getLength() > 0) {
                    Element xfrm = (Element) xfrmList.item(0);
                    
                    // 读取 ext (尺寸)
                    NodeList extList = xfrm.getElementsByTagNameNS(DML_NS, "ext");
                    if (extList != null && extList.getLength() > 0) {
                        Element ext = (Element) extList.item(0);
                        String cxStr = ext.getAttribute("cx"); // 宽度（EMU）
                        String cyStr = ext.getAttribute("cy"); // 高度（EMU）
                        
                        System.out.println("  <a:ext> 尺寸:");
                        System.out.println("    cx (宽度, EMU): " + cxStr);
                        System.out.println("    cy (高度, EMU): " + cyStr);
                        
                        if (!cxStr.isEmpty() && !cyStr.isEmpty()) {
                            try {
                                long cxEmu = Long.parseLong(cxStr);
                                long cyEmu = Long.parseLong(cyStr);
                                
                                // 转换为厘米
                                double widthCm = cxEmu / 360000.0;
                                double heightCm = cyEmu / 360000.0;
                                
                                System.out.println("    转换为厘米: " + String.format("%.2f cm × %.2f cm", widthCm, heightCm));
                                
                                // 转换为像素（120 DPI）
                                int widthPx = (int) Math.round(widthCm * 120);
                                int heightPx = (int) Math.round(heightCm * 120);
                                
                                System.out.println("    转换为像素(120 DPI): " + widthPx + " × " + heightPx);
                            } catch (NumberFormatException e) {
                                System.out.println("    无法解析数值");
                            }
                        }
                    }
                    
                    // 读取 off (位置)
                    NodeList offList = xfrm.getElementsByTagNameNS(DML_NS, "off");
                    if (offList != null && offList.getLength() > 0) {
                        Element off = (Element) offList.item(0);
                        String xStr = off.getAttribute("x");
                        String yStr = off.getAttribute("y");
                        System.out.println("  <a:off> 位置:");
                        System.out.println("    x: " + xStr + " EMU");
                        System.out.println("    y: " + yStr + " EMU");
                    }
                }
                
                // 检查是否有缩放信息
                NodeList blipList = pic.getElementsByTagNameNS(DML_NS, "blip");
                if (blipList != null && blipList.getLength() > 0) {
                    System.out.println("  找到 <a:blip> 元素（图片数据）");
                }
                
                System.out.println();
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
