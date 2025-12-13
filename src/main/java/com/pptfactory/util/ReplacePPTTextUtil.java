package com.pptfactory.util;
import org.w3c.dom.*;

import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.*;
import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class ReplacePPTTextUtil {
    private static final Map<Integer, Map<String, String>> replacements = new HashMap<>();
    static {
        replacements.put(1, new HashMap<String, String>(){{
            put("安全第一", "二我是文本");
            put("预防为主", "三我是文本");
            put("综合治理", "四我是文本");
        }});
        replacements.put(2, new HashMap<String, String>(){{
            put("煤矿安全生产核心方针", "一我是主标题");
            put("优先保障生命安全", "二我是副标题");
            put("在煤矿生产全流程中，必须将人员安全置于产量、效率之上", "三我是文本");
            put("在煤矿生产的整个过程里，人员安全的重要性要远超产量和效率。对采煤机司机而言，当操作中发现顶底板冒顶预兆、瓦斯超限等安全隐患时，需立即停机处理，严禁“冒险作业”“带病开机”。", "四我是长文本");
            put("对应岗位操作中的“紧急停机情形”", "五我是文本");
            put("在岗位操作方面，当遇到顶底板有冒顶预兆等情况时，采煤机司机必须立即停机。这体现了“安全第一”的方针，将人员安全放在首位，避免因继续作业而导致安全事故。", "六我是长文本");
        }});
        replacements.put(3, new HashMap<String, String>(){{
            put("煤矿安全生产核心方针", "一我是主标题");
            put("提前排查规避风险", "二我是副标题");
            put("通过提前检查、规范操作，从源头减少事故发生", "三我是文本");
            put("要从源头上减少事故的发生，就需要提前进行检查和规范操作。对于采煤机司机来说，开机前必须完成“设备检查”，如电缆、截齿、喷雾装置等，以及“环境检查”，如瓦斯浓度、煤壁稳定性等。", "四我是长文本");
            put("《煤矿安全规程》对“事前预防”的明确要求", "五我是文本");
            put("《煤矿安全规程》明确要求采煤机司机做好“事前预防”。司机开机前要进行设备和环境检查，避免因检查不到位引发设备损坏或人身伤害，这是保障安全生产的重要环节。", "六我是长文本");
        }});
        replacements.put(4, new HashMap<String, String>(){{
            put("煤矿安全生产核心方针", "一我是主标题");
            put("多维度协同保障", "二我是副标题");
            put("核心内涵：结合技术、管理、人员协作等多方面措施保障安全", "三我是文本");
            put("详细描述：保障煤矿安全需要结合技术、管理、人员协作等多方面措施。采煤机司机一方面要熟练掌握设备操作技术，另一方面要与支架工、输送机司机等岗位协同，同时遵守矿内安全管理制度。", "四我是长文本");
            put("采煤机司机需掌握的设备操作技术", "五我是文本");
            put("详细描述：采煤机司机需要熟练掌握设备操作技术，例如正确进刀、调速等。这是保障采煤机正常运行和安全生产的基础，只有掌握好这些技术，才能更好地完成工作任务。", "六我是长文本");
        }});
        replacements.put(5, new HashMap<String, String>(){{
            put("煤矿安全生产核心方针", "一我是主标题");
            put("多维度协同保障", "二我是副标题");
            put("采煤机司机与其他岗位的协同要求", "三我是文本");
            put("采煤机司机要与支架工、输送机司机等岗位协同作业，按顺序开机停机。这种协同作业能够提高工作效率，同时也能保障整个采煤过程的安全，形成良好的工作秩序。", "四我是长文本");
            put("采煤机司机需遵守的矿内安全管理制度", "五我是文本");
            put("采煤机司机要遵守矿内的安全管理制度，如交接班、日志填写等。这些制度有助于规范司机的行为，保证工作的连续性和可追溯性，从而为安全生产提供有力保障。", "六我是长文本");
        }});
        replacements.put(7, new HashMap<String, String>(){{
            put("与采煤机司机岗位密切相关的法律法规", "一我是主标题");
            put("中华人民共和国安全生产法", "二我是副标题");
            put("从业人员须经专门安全培训并取得相应资格上岗", "三我是文本");
            put("从业人员必须经专门安全培训，取得相应资格方可上岗。对应采煤机司机 “上岗条件” 中 “培训考试合格、持证上岗”，司机需主动参加矿内安全培训，不无证上岗。", "四我是长文本");
            put("从业人员有权拒绝违章指挥、强令冒险作业", "五我是文本");
            put("如采煤机司机可拒绝 “强行截割硬岩”“带载启动” 等违规指令，发现管理人员强令违规操作时，有权拒绝并上报。", "六我是长文本");
            put("生产经营单位需为从业人员提供符合标准的劳动防护用品", "七我是文本");
            put("像防尘口罩、安全帽等，以保障采煤机司机在喷雾降尘不足时的健康，确保其在工作环境中的安全。", "八我是长文本");
        }});
        replacements.put(8, new HashMap<String, String>(){{
            put("与采煤机司机岗位密切相关的法律法规", "一我是主标题");
            put("煤矿安全规程", "二我是副标题");
            put("电动机、开关附近", "三我是文本");
            put("米内风流中瓦斯浓度达到", "四我是长文本");
            put("采煤机截煤时喷雾装置的使用要求", "五我是文本");
            put("规程要求采煤机截煤时必须开启喷雾装置", "六我是长文本");
            put("采煤机停止工作或检修时的电源操作要求", "七我是文本");
            put("当采煤机停止工作或检修时", "八我是长文本");
        }});
        replacements.put(9, new HashMap<String, String>(){{
            put("与采煤机司机岗位密切相关的法律法规", "一我是主标题");
            put("煤矿安全培训规定", "二我是副标题");
            put("煤矿特种作业人员须经专门培训并取得《特种作业操作证》上岗", "三我是文本");
            put("煤矿特种作业人员（含采煤机司机）必须接受专门培训", "四我是长文本");
            put("证书有效期内需定期复审", "五我是文本");
            put("复审不合格不得继续上岗", "六我是长文本");
        }});
        replacements.put(10, new HashMap<String, String>(){{
            put("法律法规的岗位落实要求", "一我是主标题");
            put("日常操作对标法规", "二我是副标题");
            put("每次开机前、作业中、停机后，对照等条款自查操作行为", "三我是文本");
            put("每次开机前，需严格对照《煤矿安全规程》等条款，检查周围是否无人员、瓦斯浓度是否正常等；", "四我是文本");
            put("作业中，确认喷雾正常、设备运行无异常；", "五我是文本");
            put("停机后，切断电源闭锁，避免 “习惯性违规”。", "六我是文本");
        }});
        replacements.put(11, new HashMap<String, String>(){{
            put("法律法规的岗位落实要求", "一我是主标题");
            put("隐患处置依规报告", "二我是副标题");
            put("发现超出自身处理能力的隐患，需立即停机并按流程上报跟班队长或分管领导", "三我是文本");
            put("当发现如瓦斯持续超限、设备重大故障等超出自身处理能力的隐患时，采煤机司机必须立即停机，然后严格按照流程上报跟班队长或分管领导。", "四我是长文本");
        }});
        replacements.put(12, new HashMap<String, String>(){{
            put("法律法规的岗位落实要求", "一我是主标题");
            put("责任认知知法懂责", "二我是副标题");
            put("明确违规操作的法律后果", "三我是文本");
            put("采煤机司机要清楚若因无证上岗、违反规程导致事故，需承担相应责任，如行政处分、经济处罚，情节严重者还会被追究法律责任，从而树立 “违法必担责” 的敬畏意识。", "四我是长文本");
        }});
    }
    private static final String pptxPath = "/Users/menggl/workspace/PPTFactory/templates/master_template.pptx";
    private static final String outputPath = "/Users/menggl/workspace/PPTFactory/templates/master_template.pptx";
    
    /**
     * 获取项目目录下的temp目录
     * 如果目录不存在则创建
     */
    private static Path getProjectTempDir() throws IOException {
        // 基于pptxPath推断项目根目录
        Path projectRoot = Paths.get(pptxPath).getParent().getParent(); // 从 templates/master_template.pptx 向上两级
        Path tempDir = projectRoot.resolve("temp");
        
        // 确保目录存在
        if (!Files.exists(tempDir)) {
            Files.createDirectories(tempDir);
        }
        
        return tempDir;
    }
    
    /**
     * 在项目temp目录下创建临时目录
     */
    private static Path createTempDirInProject(String prefix) throws IOException {
        Path projectTempDir = getProjectTempDir();
        // 创建带时间戳的唯一目录名
        String dirName = prefix + System.currentTimeMillis() + "_" + Thread.currentThread().getId();
        Path tempDir = projectTempDir.resolve(dirName);
        Files.createDirectories(tempDir);
        return tempDir;
    }
    
    /**
     * 在项目temp目录下创建临时文件
     */
    private static Path createTempFileInProject(String prefix, String suffix) throws IOException {
        Path projectTempDir = getProjectTempDir();
        // 创建带时间戳的唯一文件名
        String fileName = prefix + System.currentTimeMillis() + "_" + Thread.currentThread().getId() + suffix;
        Path tempFile = projectTempDir.resolve(fileName);
        return tempFile;
    }
    
    public static void main(String[] args) throws Exception {
        // 使用批量替换方法，只解压一次，处理所有替换，然后打包一次
        // 这样可以避免重复解压和打包导致的文件损坏问题
        Map<Integer, Boolean> results = batchReplaceInSlides(pptxPath, outputPath, replacements);
        
        // 打印结果
        System.out.println("\n替换结果:");
        results.forEach((slideNum, success) -> 
            System.out.println("幻灯片 " + slideNum + ": " + (success ? "成功" : "失败"))
        );
    }
    
    /**
     * 主方法：替换指定页面的指定文本
     * @param pptxPath PPTX文件路径
     * @param outputPath 输出文件路径
     * @param slideNumber 要替换的幻灯片页码（从1开始）
     * @param oldText 要替换的旧文本
     * @param newText 要替换的新文本
     * @return 是否成功替换
     */
    public static boolean replaceTextInSpecificSlide(String pptxPath, String outputPath, 
                                                     int slideNumber, String oldText, String newText) throws Exception {
        
        // 参数验证
        if (slideNumber < 1) {
            throw new IllegalArgumentException("幻灯片页码必须从1开始");
        }
        
        if (oldText == null || oldText.trim().isEmpty()) {
            throw new IllegalArgumentException("要替换的文本不能为空");
        }
        
        // 创建临时目录（在项目temp目录下）
        Path tempDir = createTempDirInProject("pptx_edit_");
        System.out.println("临时目录: " + tempDir.toString());
        
        try {
            // 1. 解压PPTX文件到临时目录
            System.out.println("解压PPTX文件...");
            unzipPPTX(pptxPath, tempDir.toString());
            
            // 2. 确定要处理的slide文件
            String slideFileName = "slide" + slideNumber + ".xml";
            Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
            
            if (!Files.exists(slidePath)) {
                // 检查实际存在的slide文件
                System.out.println("正在检查可用的幻灯片文件...");
                List<String> availableSlides = getAvailableSlides(tempDir);
                if (availableSlides.isEmpty()) {
                    throw new FileNotFoundException("在PPTX文件中未找到任何幻灯片");
                }
                
                System.out.println("可用的幻灯片: " + availableSlides);
                
                if (slideNumber > availableSlides.size()) {
                    throw new IllegalArgumentException("幻灯片页码 " + slideNumber + " 超出范围，最大页码为 " + availableSlides.size());
                }
                
                // 使用实际的文件名
                slideFileName = availableSlides.get(slideNumber - 1);
                slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
                System.out.println("使用实际文件名: " + slideFileName);
            }
            
            // 3. 解析并修改slide.xml
            System.out.println("处理文件: " + slidePath);
            boolean textReplaced = processSlideXML(slidePath, oldText, newText);
            
            if (!textReplaced) {
                System.out.println("警告: 在第 " + slideNumber + " 页中未找到文本 '" + oldText + "'");
            }
            
            // 4. 重新打包为PPTX
            System.out.println("重新打包为PPTX...");
            // 如果输入和输出路径相同，先写入临时文件，然后原子性地替换
            if (pptxPath.equals(outputPath)) {
                Path tempOutput = createTempFileInProject("pptx_output_", ".pptx");
                try {
                    zipDirectory(tempDir.toString(), tempOutput.toString());
                    // 原子性地替换原文件
                    Files.move(tempOutput, Paths.get(outputPath), StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception e) {
                    // 如果失败，尝试删除临时文件
                    try {
                        Files.deleteIfExists(tempOutput);
                    } catch (IOException ignored) {}
                    throw e;
                }
            } else {
                zipDirectory(tempDir.toString(), outputPath);
            }
            
            System.out.println("处理完成！输出文件: " + outputPath);
            return textReplaced;
            
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
        
        // 解析XML文档
        Document doc = builder.parse(slidePath.toFile());
        
        // 先尝试使用XPath方法（更精确，针对PPTX的a:t元素）
        boolean replaced = replaceTextUsingXPath(slidePath, oldText, newText);
        
        // 如果XPath方法没找到，再尝试递归方法
        if (!replaced) {
            replaced = replaceTextInElement(doc.getDocumentElement(), oldText, newText, 0);
            if (replaced) {
                // 保存修改后的XML
                saveXMLDocument(doc, slidePath);
            }
        }
        
        return replaced;
    }
    
    /**
     * 递归搜索并替换XML元素中的文本
     */
    private static boolean replaceTextInElement(Element element, String oldText, String newText, int depth) {
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
                    System.out.println("在深度 " + depth + " 找到并替换文本: " + 
                                     text.substring(0, Math.min(text.length(), 50)) + "...");
                    replaced = true;
                }
            }
        }
        
        // 递归处理子元素
        NodeList children = element.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child.getNodeType() == Node.ELEMENT_NODE) {
                if (replaceTextInElement((Element) child, oldText, newText, depth + 1)) {
                    replaced = true;
                }
            }
        }
        
        return replaced;
    }
    
    /**
     * 使用XPath更精确地定位文本节点（针对PPTX的a:t元素）
     */
    private static boolean replaceTextUsingXPath(Path slidePath, String oldText, String newText) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        
        Document doc = builder.parse(slidePath.toFile());
        
        // 创建XPath
        XPathFactory xpathFactory = XPathFactory.newInstance();
        XPath xpath = xpathFactory.newXPath();
        // 注册命名空间
        xpath.setNamespaceContext(new javax.xml.namespace.NamespaceContext() {
            @Override
            public String getNamespaceURI(String prefix) {
                switch (prefix) {
                    case "a":
                        return "http://schemas.openxmlformats.org/drawingml/2006/main";
                    case "p": return "http://schemas.openxmlformats.org/presentationml/2006/main";
                    case "r": return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    default: return null;
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
                System.out.println("使用XPath找到并替换文本: '" + text + "' -> '" + replacedText + "'");
                replaced = true;
            }
        }
        
        if (replaced) {
            saveXMLDocument(doc, slidePath);
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
            System.out.println("临时目录已清理");
        } catch (IOException e) {
            System.err.println("清理临时目录时出错: " + e.getMessage());
        }
    }
    
    /**
     * 高级功能：获取所有幻灯片中的文本内容
     */
    public static Map<Integer, List<String>> extractAllSlideTexts(String pptxPath) throws Exception {
        Path tempDir = createTempDirInProject("pptx_extract_");
        Map<Integer, List<String>> slideTexts = new TreeMap<>();
        
        try {
            unzipPPTX(pptxPath, tempDir.toString());
            
            List<String> slideFiles = getAvailableSlides(tempDir);
            for (int i = 0; i < slideFiles.size(); i++) {
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFiles.get(i));
                List<String> texts = extractTextFromSlide(slidePath);
                slideTexts.put(i + 1, texts);
            }
            
        } finally {
            cleanTempDirectory(tempDir);
        }
        
        return slideTexts;
    }
    
    /**
     * 从单个slide.xml提取所有文本
     */
    private static List<String> extractTextFromSlide(Path slidePath) throws Exception {
        List<String> texts = new ArrayList<>();
        
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(slidePath.toFile());
        
        // 使用XPath查找所有文本节点
        XPathFactory xpathFactory = XPathFactory.newInstance();
        XPath xpath = xpathFactory.newXPath();

        xpath.setNamespaceContext(new javax.xml.namespace.NamespaceContext() {
            @Override
            public String getNamespaceURI(String prefix) {
                if ("a".equals(prefix)) {
                    return "http://schemas.openxmlformats.org/drawingml/2006/main";
                }
                return javax.xml.XMLConstants.NULL_NS_URI;
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
        
        NodeList textNodes = (NodeList) xpath.evaluate("//a:t", doc, XPathConstants.NODESET);
        
        for (int i = 0; i < textNodes.getLength(); i++) {
            Node textNode = textNodes.item(i);
            String text = textNode.getTextContent();
            if (text != null && !text.trim().isEmpty()) {
                texts.add(text.trim());
            }
        }
        
        return texts;
    }
    
    /**
     * 批量替换多个幻灯片中的文本
     */
    public static Map<Integer, Boolean> batchReplaceInSlides(String pptxPath, String outputPath,
                                                            Map<Integer, Map<String, String>> replacements) 
                                                            throws Exception {
        Path tempDir = createTempDirInProject("pptx_batch_");
        Map<Integer, Boolean> results = new TreeMap<>();
        
        try {
            // 解压
            unzipPPTX(pptxPath, tempDir.toString());
            
            // 获取所有幻灯片文件
            List<String> slideFiles = getAvailableSlides(tempDir);
            System.out.println("找到 " + slideFiles.size() + " 个幻灯片文件: " + slideFiles);
            
            // 处理每个指定的幻灯片
            for (Map.Entry<Integer, Map<String, String>> entry : replacements.entrySet()) {
                int slideNum = entry.getKey();
                
                if (slideNum < 1 || slideNum > slideFiles.size()) {
                    System.err.println("警告: 幻灯片 " + slideNum + " 不存在，跳过（共有 " + slideFiles.size() + " 个幻灯片）");
                    results.put(slideNum, false);
                    continue;
                }
                
                String slideFileName = slideFiles.get(slideNum - 1);
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
                System.out.println("处理幻灯片 " + slideNum + ": " + slideFileName);
                
                // 先提取并显示当前幻灯片的所有文本，用于调试
                List<String> currentTexts = extractTextFromSlide(slidePath);
                System.out.println("幻灯片 " + slideNum + " 中的文本内容: " + currentTexts);
                
                boolean slideReplaced = false;
                Map<String, String> slideReplacements = entry.getValue();
                
                // 对当前幻灯片应用所有替换
                for (Map.Entry<String, String> textReplacement : slideReplacements.entrySet()) {
                    String oldText = textReplacement.getKey();
                    String newText = textReplacement.getValue();
                    System.out.println("尝试替换: '" + oldText + "' -> '" + newText + "'");
                    if (processSlideXML(slidePath, oldText, newText)) {
                        slideReplaced = true;
                        System.out.println("成功替换: '" + oldText + "' -> '" + newText + "'");
                    } else {
                        System.out.println("未找到文本: '" + oldText + "'");
                    }
                }
                
                results.put(slideNum, slideReplaced);
            }
            
            // 重新打包
            // 如果输入和输出路径相同，先写入临时文件，然后原子性地替换
            if (pptxPath.equals(outputPath)) {
                Path tempOutput = createTempFileInProject("pptx_output_", ".pptx");
                try {
                    zipDirectory(tempDir.toString(), tempOutput.toString());
                    // 原子性地替换原文件
                    Files.move(tempOutput, Paths.get(outputPath), StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception e) {
                    // 如果失败，尝试删除临时文件
                    try {
                        Files.deleteIfExists(tempOutput);
                    } catch (IOException ignored) {}
                    throw e;
                }
            } else {
                zipDirectory(tempDir.toString(), outputPath);
            }
            
        } finally {
            cleanTempDirectory(tempDir);
        }
        
        return results;
    }
}
