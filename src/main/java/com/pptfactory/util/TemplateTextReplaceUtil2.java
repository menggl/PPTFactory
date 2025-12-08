package com.pptfactory.util;
import org.xml.sax.*;
import org.xml.sax.helpers.*;
import javax.xml.parsers.*;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.*;
import java.nio.file.*;
public class TemplateTextReplaceUtil2 {

    public static void main(String[] args) throws Exception {
        replaceTextFromXML("/Users/menggl/workspace/PPTFactory/master_template_clean.pptx");
    }

    public static void replaceTextFromXML(String pptxPath) throws Exception {
        // PPTX 实际上是 ZIP 文件，需要先解压
        File tempDir = PPTXSAXProcessor.unzipPPTX(pptxPath);
        
        // 处理所有幻灯片
        File slidesDir = new File(tempDir, "ppt/slides");
        for (File slideFile : slidesDir.listFiles((dir, name) -> name.endsWith(".xml"))) {
            System.out.println("\n=== 解析幻灯片: " + slideFile.getName() + " ===");
            parseSlideWithSAX(slideFile);
        }
    }
    
    // 使用 SAX 解析单个幻灯片
    public static void parseSlideWithSAX(File slideFile) throws Exception {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        factory.setNamespaceAware(true);
        SAXParser parser = factory.newSAXParser();
        
        SlideHandler slideHandler = new SlideHandler();
        parser.parse(slideFile, slideHandler);
        
        // 输出结果
        slideHandler.printResults();
    }

    public static class SlideHandler extends DefaultHandler {
        // 状态跟踪
        private boolean inTextBox = false;
        private boolean inParagraph = false;
        private boolean inTextRun = false;
        private boolean inText = false;
        
        // 当前收集的文本
        private StringBuilder currentText = new StringBuilder();
        private StringBuilder paragraphText = new StringBuilder();
        private List<String> paragraphs = new ArrayList<>();
        
        // 当前格式信息
        private TextFormat currentFormat = new TextFormat();
        private List<TextFormat> paragraphFormats = new ArrayList<>();
        
        // 结果存储
        private Map<String, List<String>> shapeTexts = new LinkedHashMap<>();
        private String currentShapeId = null;
        
        // 命名空间常量
        private static final String PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static final String DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        
        @Override
        public void startElement(String uri, String localName, 
                                 String qName, Attributes attributes) throws SAXException {
            
            // 检测形状开始
            if (uri.equals(PRESENTATION_NS) && localName.equals("sp")) {
                currentShapeId = getShapeId(attributes);
                inTextBox = false;
            }
            
            // 检测文本框开始
            else if (uri.equals(PRESENTATION_NS) && localName.equals("txBody")) {
                inTextBox = true;
                paragraphs.clear();
                paragraphFormats.clear();
            }
            
            // 检测段落开始
            else if (uri.equals(DRAWING_NS) && localName.equals("p")) {
                inParagraph = true;
                paragraphText.setLength(0);
                currentFormat = new TextFormat(); // 重置格式
            }
            
            // 检测文本运行开始
            else if (uri.equals(DRAWING_NS) && localName.equals("r")) {
                inTextRun = true;
            }
            
            // 检测文本内容开始
            else if (uri.equals(DRAWING_NS) && localName.equals("t")) {
                inText = true;
                currentText.setLength(0);
            }
            
            // 解析格式属性（rPr元素）
            else if (inTextRun && uri.equals(DRAWING_NS) && localName.equals("rPr")) {
                parseFormatAttributes(attributes);
            }
            
            // 解析颜色
            else if (inTextRun && uri.equals(DRAWING_NS) && localName.equals("solidFill")) {
                // 颜色信息可能在子元素中，这里只是标记
            }
            
            // 解析具体颜色值
            else if (inTextRun && uri.equals(DRAWING_NS) && 
                     (localName.equals("srgbClr") || localName.equals("prstClr"))) {
                if (localName.equals("srgbClr")) {
                    String colorVal = attributes.getValue("val");
                    if (colorVal != null) {
                        currentFormat.color = "#" + colorVal;
                    }
                }
            }
            
            // 解析字体
            else if (inTextRun && uri.equals(DRAWING_NS) && localName.equals("latin")) {
                String typeface = attributes.getValue("typeface");
                if (typeface != null) {
                    currentFormat.fontFamily = typeface;
                }
            }
        }
        
        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            if (inText) {
                currentText.append(ch, start, length);
            }
        }
        
        @Override
        public void endElement(String uri, String localName, String qName) throws SAXException {
            
            // 文本内容结束
            if (uri.equals(DRAWING_NS) && localName.equals("t")) {
                inText = false;
                if (currentText.length() > 0) {
                    paragraphText.append(currentText.toString());
                }
            }
            
            // 文本运行结束
            else if (uri.equals(DRAWING_NS) && localName.equals("r")) {
                inTextRun = false;
            }
            
            // 段落结束
            else if (uri.equals(DRAWING_NS) && localName.equals("p")) {
                inParagraph = false;
                if (paragraphText.length() > 0) {
                    paragraphs.add(paragraphText.toString());
                    paragraphFormats.add(new TextFormat(currentFormat)); // 保存副本
                }
            }
            
            // 文本框结束
            else if (uri.equals(PRESENTATION_NS) && localName.equals("txBody")) {
                inTextBox = false;
                
                // 保存当前形状的文本
                if (!paragraphs.isEmpty() && currentShapeId != null) {
                    shapeTexts.put(currentShapeId, new ArrayList<>(paragraphs));
                }
            }
            
            // 形状结束
            else if (uri.equals(PRESENTATION_NS) && localName.equals("sp")) {
                currentShapeId = null;
            }
        }
        
        // 解析格式属性
        private void parseFormatAttributes(Attributes attributes) {
            // 字体大小
            String sz = attributes.getValue("sz");
            if (sz != null) {
                try {
                    currentFormat.fontSize = Integer.parseInt(sz) / 100.0;
                } catch (NumberFormatException e) {
                    currentFormat.fontSize = 18.0;
                }
            }
            
            // 粗体
            String bold = attributes.getValue("b");
            currentFormat.bold = "1".equals(bold);
            
            // 斜体
            String italic = attributes.getValue("i");
            currentFormat.italic = "1".equals(italic);
            
            // 下划线
            String underline = attributes.getValue("u");
            currentFormat.underline = underline != null && !"none".equals(underline);
            
            // 字间距
            String kern = attributes.getValue("kern");
            if (kern != null) {
                try {
                    currentFormat.letterSpacing = Integer.parseInt(kern) / 100.0;
                } catch (NumberFormatException e) {
                    // 忽略
                }
            }
        }
        
        // 获取形状ID
        private String getShapeId(Attributes attributes) {
            // 实际实现需要解析嵌套结构
            return "shape_" + System.currentTimeMillis(); // 简化示例
        }
        
        // 打印结果
        public void printResults() {
            System.out.println("找到 " + shapeTexts.size() + " 个文本框:");
            
            for (Map.Entry<String, List<String>> entry : shapeTexts.entrySet()) {
                System.out.println("\n形状ID: " + entry.getKey());
                for (int i = 0; i < entry.getValue().size(); i++) {
                    String para = entry.getValue().get(i);
                    TextFormat format = i < paragraphFormats.size() ? 
                        paragraphFormats.get(i) : null;
                    
                    System.out.println("  段落 " + (i+1) + ": " + para);
                    if (format != null) {
                        System.out.println("    格式: " + format);
                    }
                }
            }
        }
        
        // 获取所有文本
        public Map<String, String> getAllText() {
            Map<String, String> result = new LinkedHashMap<>();
            
            for (Map.Entry<String, List<String>> entry : shapeTexts.entrySet()) {
                StringBuilder fullText = new StringBuilder();
                for (String para : entry.getValue()) {
                    if (fullText.length() > 0) {
                        fullText.append("\n");
                    }
                    fullText.append(para);
                }
                result.put(entry.getKey(), fullText.toString());
            }
            
            return result;
        }
        
        // 文本格式类
        static class TextFormat {
            double fontSize = 18.0;
            boolean bold = false;
            boolean italic = false;
            boolean underline = false;
            double letterSpacing = 0;
            String color = "#000000";
            String fontFamily = "Calibri";
            
            TextFormat() {}
            
            TextFormat(TextFormat other) {
                this.fontSize = other.fontSize;
                this.bold = other.bold;
                this.italic = other.italic;
                this.underline = other.underline;
                this.letterSpacing = other.letterSpacing;
                this.color = other.color;
                this.fontFamily = other.fontFamily;
            }
            
            @Override
            public String toString() {
                return String.format("字体: %s, 大小: %.1fpt, 颜色: %s%s%s", 
                    fontFamily, fontSize, color,
                    bold ? " 粗体" : "",
                    italic ? " 斜体" : "");
            }
        }
    }

    public static class PPTXSAXProcessor {
    
        // 解压 PPTX 文件
        public static File unzipPPTX(String pptxPath) throws IOException {
            File tempDir = Files.createTempDirectory("pptx_").toFile();
            
            try (ZipFile zipFile = new ZipFile(pptxPath)) {
                Enumeration<? extends ZipEntry> entries = zipFile.entries();
                
                while (entries.hasMoreElements()) {
                    ZipEntry entry = entries.nextElement();
                    File entryFile = new File(tempDir, entry.getName());
                    
                    // 创建目录
                    if (entry.isDirectory()) {
                        entryFile.mkdirs();
                    } else {
                        // 创建父目录
                        entryFile.getParentFile().mkdirs();
                        
                        // 复制文件
                        try (InputStream is = zipFile.getInputStream(entry);
                             OutputStream os = new FileOutputStream(entryFile)) {
                            byte[] buffer = new byte[1024];
                            int length;
                            while ((length = is.read(buffer)) > 0) {
                                os.write(buffer, 0, length);
                            }
                        }
                    }
                }
            }
            
            return tempDir;
        }
        
        // 处理幻灯片目录
        public static Map<String, Map<String, List<String>>> processAllSlides(File pptDir) throws Exception {
            Map<String, Map<String, List<String>>> allSlidesText = new LinkedHashMap<>();
            
            File slidesDir = new File(pptDir, "ppt/slides");
            if (!slidesDir.exists()) {
                return allSlidesText;
            }
            
            File[] slideFiles = slidesDir.listFiles((dir, name) -> 
                name.endsWith(".xml") && name.startsWith("slide"));
            
            if (slideFiles != null) {
                Arrays.sort(slideFiles); // 按文件名排序
                
                for (File slideFile : slideFiles) {
                    String slideName = slideFile.getName().replace(".xml", "");
                    
                    // 使用SAX解析
                    Map<String, List<String>> slideText = parseSingleSlideSAX(slideFile);
                    allSlidesText.put(slideName, slideText);
                    
                    System.out.println("已处理: " + slideName + " - " + 
                                     slideText.size() + "个文本框");
                }
            }
            
            return allSlidesText;
        }
        
        // 使用SAX解析单个幻灯片
        public static Map<String, List<String>> parseSingleSlideSAX(File slideFile) throws Exception {
            SAXParserFactory factory = SAXParserFactory.newInstance();
            factory.setNamespaceAware(true);
            SAXParser parser = factory.newSAXParser();
            
            SlideContentHandler handler = new SlideContentHandler();
            parser.parse(slideFile, handler);
            
            return handler.getShapeParagraphs();
        }
        
        // 专门用于提取内容的处理器
        static class SlideContentHandler extends DefaultHandler {
            private Map<String, List<String>> shapeParagraphs = new LinkedHashMap<>();
            private List<String> currentParagraphs = new ArrayList<>();
            private StringBuilder currentText = new StringBuilder();
            private StringBuilder currentParagraph = new StringBuilder();
            
            private boolean inTxBody = false;
            private boolean inParagraph = false;
            private boolean inText = false;
            private String currentShapeId = null;
            
            private static final String PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
            private static final String DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
            
            @Override
            public void startElement(String uri, String localName, 
                                     String qName, Attributes attributes) {
                
                if (uri.equals(PRESENTATION_NS)) {
                    if (localName.equals("sp")) {
                        // 开始新形状
                        currentShapeId = extractShapeId(attributes);
                        currentParagraphs.clear();
                    } else if (localName.equals("txBody")) {
                        inTxBody = true;
                    }
                } else if (uri.equals(DRAWING_NS)) {
                    if (localName.equals("p")) {
                        inParagraph = true;
                        currentParagraph.setLength(0);
                    } else if (localName.equals("t")) {
                        inText = true;
                        currentText.setLength(0);
                    }
                }
                // 检查是否隐藏
                String hidden = attributes.getValue("hidden");
                if ("1".equals(hidden) || "true".equals(hidden)) {
                    System.out.println("警告：这是隐藏形状！");
                }
            }
            
            @Override
            public void characters(char[] ch, int start, int length) {
                if (inText) {
                    currentText.append(ch, start, length);
                }
            }
            
            @Override
            public void endElement(String uri, String localName, String qName) {
                if (uri.equals(DRAWING_NS)) {
                    if (localName.equals("t")) {
                        inText = false;
                        if (currentText.length() > 0) {
                            currentParagraph.append(currentText.toString());
                        }
                    } else if (localName.equals("p")) {
                        inParagraph = false;
                        if (currentParagraph.length() > 0) {
                            currentParagraphs.add(currentParagraph.toString());
                        }
                    }
                } else if (uri.equals(PRESENTATION_NS)) {
                    if (localName.equals("txBody")) {
                        inTxBody = false;
                        // 保存当前形状的段落
                        if (!currentParagraphs.isEmpty() && currentShapeId != null) {
                            shapeParagraphs.put(currentShapeId, 
                                new ArrayList<>(currentParagraphs));
                        }
                    } else if (localName.equals("sp")) {
                        currentShapeId = null;
                    }
                }
            }
            
            private String extractShapeId(Attributes attributes) {
                // 简化实现，实际需要解析嵌套结构
                return "shape_" + System.currentTimeMillis() + "_" + 
                       shapeParagraphs.size();
            }
            
            public Map<String, List<String>> getShapeParagraphs() {
                return new LinkedHashMap<>(shapeParagraphs);
            }
        }
    }
    
}
