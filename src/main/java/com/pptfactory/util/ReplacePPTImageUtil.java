package com.pptfactory.util;

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
 * 替换指定pptx文件中指定页面的图片
 */
public class ReplacePPTImageUtil {
    
    private static final String PPTX_PATH = "/Users/menggl/workspace/PPTFactory/templates/master_template.pptx";
    
    public static void main(String[] args) {
        // 将PPTX中所有图片替换为带有"No Image"文字的图像
        replaceAllImagesWithNoImage(PPTX_PATH, PPTX_PATH);
    }

    /**
     * 替换PPTX文件中指定页面的图片
     * @param pptxPath PPTX文件路径
     * @param page 页码（从1开始）
     * @param imagePath 新图片文件路径
     * @param outputPath 输出文件路径（如果与输入路径相同，会先写入临时文件再替换）
     */
    public static void replacePPTImage(String pptxPath, int page, String imagePath, String outputPath) {
        Path tempDir = null;
        try {
            // 参数验证
            if (page < 1) {
                throw new IllegalArgumentException("页码必须从1开始");
            }
            
            File imageFile = new File(imagePath);
            if (!imageFile.exists()) {
                throw new FileNotFoundException("图片文件不存在: " + imagePath);
            }
            
            // 创建临时目录（在项目temp目录下）
            tempDir = createTempDirInProject("pptx_image_replace_");
            System.out.println("临时目录: " + tempDir.toString());
            
            // 1. 解压PPTX文件
            System.out.println("解压PPTX文件...");
            unzipPPTX(pptxPath, tempDir.toString());
            
            // 2. 找到指定幻灯片的XML文件
            String slideFileName = "slide" + page + ".xml";
            Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
            
            if (!Files.exists(slidePath)) {
                // 检查实际存在的slide文件
                List<String> availableSlides = getAvailableSlides(tempDir);
                if (availableSlides.isEmpty()) {
                    throw new FileNotFoundException("在PPTX文件中未找到任何幻灯片");
                }
                
                System.out.println("可用的幻灯片: " + availableSlides);
                
                if (page > availableSlides.size()) {
                    throw new IllegalArgumentException("页码 " + page + " 超出范围，最大页码为 " + availableSlides.size());
                }
                
                slideFileName = availableSlides.get(page - 1);
                slidePath = tempDir.resolve("ppt/slides").resolve(slideFileName);
                System.out.println("使用实际文件名: " + slideFileName);
            }
            
            // 3. 找到幻灯片的关系文件
            String slideRelFileName = "slide" + page + ".xml.rels";
            Path slideRelPath = tempDir.resolve("ppt/slides/_rels").resolve(slideRelFileName);
            
            // 4. 替换图片
            System.out.println("处理幻灯片: " + slideFileName);
            boolean replaced = replaceImageInSlide(slidePath, slideRelPath, imageFile, tempDir);
            
            if (!replaced) {
                System.out.println("警告: 在第 " + page + " 页中未找到图片");
            } else {
                System.out.println("成功替换图片");
            }
            
            // 5. 重新打包为PPTX
            System.out.println("重新打包为PPTX...");
            if (pptxPath.equals(outputPath)) {
                // 如果输入输出路径相同，先写入临时文件
                Path tempOutput = createTempFileInProject("pptx_output_", ".pptx");
                try {
                    zipDirectory(tempDir.toString(), tempOutput.toString());
                    Files.move(tempOutput, Paths.get(outputPath), StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception e) {
                    try {
                        Files.deleteIfExists(tempOutput);
                    } catch (IOException ignored) {}
                    throw e;
                }
            } else {
                zipDirectory(tempDir.toString(), outputPath);
            }
            
            System.out.println("处理完成！输出文件: " + outputPath);
            
        } catch (Exception e) {
            System.err.println("替换图片失败: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // 清理临时目录
            if (tempDir != null) {
                cleanTempDirectory(tempDir);
            }
        }
    }
    
    /**
     * 替换幻灯片中的图片
     */
    private static boolean replaceImageInSlide(Path slidePath, Path slideRelPath, File newImageFile, Path tempDir) 
            throws Exception {
        
        // 读取幻灯片关系文件，找到图片关系
        Map<String, String> imageRelations = new HashMap<>();
        if (Files.exists(slideRelPath)) {
            imageRelations = parseSlideRelations(slideRelPath);
        }
        
        if (imageRelations.isEmpty()) {
            System.out.println("未找到图片关系，尝试直接解析幻灯片XML...");
            // 如果没有关系文件，尝试直接从XML中查找图片引用
            return replaceImageDirectly(slidePath, newImageFile, tempDir);
        }
        
        // 找到第一个图片关系
        String imageRelId = null;
        String imageTarget = null;
        for (Map.Entry<String, String> entry : imageRelations.entrySet()) {
            String target = entry.getValue();
            if (target != null && target.startsWith("../media/")) {
                imageRelId = entry.getKey();
                imageTarget = target;
                break;
            }
        }
        
        if (imageRelId == null) {
            System.out.println("未找到图片关系");
            return false;
        }
        
        // 获取图片文件名
        String imageFileName = imageTarget.substring("../media/".length());
        Path mediaDir = tempDir.resolve("ppt/media");
        Path oldImagePath = mediaDir.resolve(imageFileName);
        
        // 确定新图片的文件名和扩展名
        String newImageExtension = getFileExtension(newImageFile.getName());
        String newImageFileName;
        
        // 如果原图片存在，保持相同的文件名（只替换内容）
        if (Files.exists(oldImagePath)) {
            newImageFileName = imageFileName;
        } else {
            // 否则使用新的文件名
            newImageFileName = "image" + System.currentTimeMillis() + "." + newImageExtension;
        }
        
        Path newImagePath = mediaDir.resolve(newImageFileName);
        
        // 复制新图片到media目录
        Files.copy(newImageFile.toPath(), newImagePath, StandardCopyOption.REPLACE_EXISTING);
        System.out.println("已复制图片到: " + newImagePath);
        
        // 如果文件名改变了，需要更新关系文件
        if (!newImageFileName.equals(imageFileName)) {
            updateSlideRelations(slideRelPath, imageRelId, "../media/" + newImageFileName);
        }
        
        return true;
    }
    
    /**
     * 直接从幻灯片XML中查找并替换图片（当没有关系文件时）
     */
    private static boolean replaceImageDirectly(Path slidePath, File newImageFile, Path tempDir) throws Exception {
        // 读取幻灯片XML
        String xmlContent = new String(Files.readAllBytes(slidePath), "UTF-8");
        
        // 查找图片引用模式：r:embed="rIdX" 或 r:link="rIdX"
        Pattern embedPattern = Pattern.compile("r:embed=\"(rId\\d+)\"");
        Pattern linkPattern = Pattern.compile("r:link=\"(rId\\d+)\"");
        
        Matcher embedMatcher = embedPattern.matcher(xmlContent);
        Matcher linkMatcher = linkPattern.matcher(xmlContent);
        
        // 查找关系文件
        Path slideRelPath = slidePath.getParent().getParent().resolve("_rels")
                .resolve(slidePath.getFileName().toString() + ".rels");
        
        if (!Files.exists(slideRelPath)) {
            System.out.println("未找到关系文件，无法替换图片");
            return false;
        }
        
        // 解析关系文件
        Map<String, String> relations = parseSlideRelations(slideRelPath);
        
        // 查找图片关系
        String imageRelId = null;
        if (embedMatcher.find()) {
            imageRelId = embedMatcher.group(1);
        } else if (linkMatcher.find()) {
            imageRelId = linkMatcher.group(1);
        }
        
        if (imageRelId == null || !relations.containsKey(imageRelId)) {
            return false;
        }
        
        String imageTarget = relations.get(imageRelId);
        if (imageTarget == null || !imageTarget.startsWith("../media/")) {
            return false;
        }
        
        // 替换图片文件
        String imageFileName = imageTarget.substring("../media/".length());
        Path mediaDir = tempDir.resolve("ppt/media");
        Path imagePath = mediaDir.resolve(imageFileName);
        
        if (Files.exists(imagePath)) {
            Files.copy(newImageFile.toPath(), imagePath, StandardCopyOption.REPLACE_EXISTING);
            System.out.println("已替换图片: " + imagePath);
            return true;
        }
        
        return false;
    }
    
    /**
     * 解析幻灯片关系文件
     */
    private static Map<String, String> parseSlideRelations(Path relPath) throws Exception {
        Map<String, String> relations = new HashMap<>();
        
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
     * 获取文件扩展名
     */
    private static String getFileExtension(String fileName) {
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
        Path projectRoot = Paths.get(PPTX_PATH).getParent().getParent();
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
        Path projectTempDir = getProjectTempDir();
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
        String fileName = prefix + System.currentTimeMillis() + "_" + Thread.currentThread().getId() + suffix;
        Path tempFile = projectTempDir.resolve(fileName);
        return tempFile;
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
     * 重新打包目录为PPTX
     */
    private static void zipDirectory(String sourceDir, String zipFilePath) throws IOException {
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
     * 将PPTX文件中所有图片替换为带有"No Image"文字的图像
     * @param pptxPath PPTX文件路径
     * @param outputPath 输出文件路径（如果与输入路径相同，会先写入临时文件再替换）
     */
    public static void replaceAllImagesWithNoImage(String pptxPath, String outputPath) {
        Path tempDir = null;
        try {
            // 创建临时目录（在项目temp目录下）
            tempDir = createTempDirInProject("pptx_replace_all_images_");
            System.out.println("临时目录: " + tempDir.toString());
            
            // 1. 解压PPTX文件
            System.out.println("解压PPTX文件...");
            unzipPPTX(pptxPath, tempDir.toString());
            
            // 2. 获取所有幻灯片文件
            List<String> slideFiles = getAvailableSlides(tempDir);
            System.out.println("找到 " + slideFiles.size() + " 个幻灯片");
            
            // 3. 收集背景使用的图片，后续跳过
            Set<String> backgroundImages = collectBackgroundImages(tempDir);
            if (!backgroundImages.isEmpty()) {
                System.out.println("检测到背景图片文件（将跳过替换）: " + backgroundImages);
            }
            
            // 4. 收集标题不为空的图片（只替换这些）
            Set<String> titleMarkedImages = collectImagesWithTitle(tempDir);
            System.out.println("检测到需替换的图片(标题不为空)文件: " + titleMarkedImages);
            
            // 5. 获取所有图片文件
            Path mediaDir = tempDir.resolve("ppt/media");
            List<Path> imageFiles = new ArrayList<>();
            if (Files.exists(mediaDir)) {
                try (DirectoryStream<Path> stream = Files.newDirectoryStream(mediaDir)) {
                    for (Path path : stream) {
                        if (Files.isRegularFile(path)) {
                            String fileName = path.getFileName().toString().toLowerCase();
                            if (fileName.endsWith(".png") || fileName.endsWith(".jpg") || 
                                fileName.endsWith(".jpeg") || fileName.endsWith(".gif") ||
                                fileName.endsWith(".bmp") || fileName.endsWith(".tiff")) {
                                imageFiles.add(path);
                            }
                        }
                    }
                }
            }
            
            System.out.println("找到 " + imageFiles.size() + " 个图片文件");
            
            // 6. 替换所有图片（跳过背景图；仅替换标题不为空的）
            int replacedCount = 0;
            for (Path imagePath : imageFiles) {
                if (backgroundImages.contains(imagePath.getFileName().toString())) {
                    System.out.println("跳过背景图片: " + imagePath.getFileName());
                    continue;
                }
                if (!titleMarkedImages.isEmpty() &&
                    !titleMarkedImages.contains(imagePath.getFileName().toString())) {
                    // 标题为空的不替换
                    continue;
                }
                try {
                    // 读取原图片尺寸
                    java.awt.image.BufferedImage originalImage = javax.imageio.ImageIO.read(imagePath.toFile());
                    int width = originalImage != null ? originalImage.getWidth() : 400;
                    int height = originalImage != null ? originalImage.getHeight() : 300;
                    
                    // 创建带有"No Image"文字的图片
                    java.awt.image.BufferedImage noImageImage = createNoImageImage(width, height, "No Image");
                    
                    // 保存替换后的图片
                    String extension = getFileExtension(imagePath.getFileName().toString());
                    javax.imageio.ImageIO.write(noImageImage, extension.equals("jpg") || extension.equals("jpeg") ? "jpg" : "png", imagePath.toFile());
                    
                    replacedCount++;
                    System.out.println("已替换图片: " + imagePath.getFileName());
                } catch (Exception e) {
                    System.err.println("替换图片失败 " + imagePath.getFileName() + ": " + e.getMessage());
                }
            }
            
            System.out.println("共替换了 " + replacedCount + " 个图片");
            
            // 7. 重新打包为PPTX
            System.out.println("重新打包为PPTX...");
            if (pptxPath.equals(outputPath)) {
                // 如果输入输出路径相同，先写入临时文件
                Path tempOutput = createTempFileInProject("pptx_output_", ".pptx");
                try {
                    zipDirectory(tempDir.toString(), tempOutput.toString());
                    Files.move(tempOutput, Paths.get(outputPath), StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception e) {
                    try {
                        Files.deleteIfExists(tempOutput);
                    } catch (IOException ignored) {}
                    throw e;
                }
            } else {
                zipDirectory(tempDir.toString(), outputPath);
            }
            
            System.out.println("处理完成！输出文件: " + outputPath);
            
        } catch (Exception e) {
            System.err.println("替换图片失败: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // 清理临时目录
            if (tempDir != null) {
                cleanTempDirectory(tempDir);
            }
        }
    }
    
    /**
     * 创建带有指定文字的图片
     * @param width 图片宽度
     * @param height 图片高度
     * @param text 要显示的文字
     * @return BufferedImage对象
     */
    private static java.awt.image.BufferedImage createNoImageImage(int width, int height, String text) {
        // 创建图片
        java.awt.image.BufferedImage image = new java.awt.image.BufferedImage(
            width, height, java.awt.image.BufferedImage.TYPE_INT_RGB);
        
        // 获取 Graphics2D 对象用于绘制
        java.awt.Graphics2D g2d = image.createGraphics();
        
        // 设置抗锯齿
        g2d.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING, 
            java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setRenderingHint(java.awt.RenderingHints.KEY_TEXT_ANTIALIASING,
            java.awt.RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        
        // 填充白色背景
        g2d.setColor(java.awt.Color.WHITE);
        g2d.fillRect(0, 0, width, height);
        
        // 绘制边框
        g2d.setColor(java.awt.Color.LIGHT_GRAY);
        g2d.setStroke(new java.awt.BasicStroke(2.0f));
        g2d.drawRect(2, 2, width - 4, height - 4);
        
        // 绘制文字
        g2d.setColor(java.awt.Color.GRAY);
        // 根据图片大小调整字体大小
        int fontSize = Math.max(12, Math.min(width, height) / 10);
        java.awt.Font font = new java.awt.Font("Arial", java.awt.Font.PLAIN, fontSize);
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
     * 收集所有被幻灯片背景引用的图片文件名（仅文件名）
     */
    private static Set<String> collectBackgroundImages(Path tempDir) {
        Set<String> backgroundImages = new HashSet<>();
        try {
            List<String> slideFiles = getAvailableSlides(tempDir);
            for (String slideFile : slideFiles) {
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFile);
                Path relPath = tempDir.resolve("ppt/slides/_rels").resolve(slideFile + ".rels");
                if (!Files.exists(slidePath) || !Files.exists(relPath)) {
                    continue;
                }

                // 查找 slide.xml 中的背景引用 r:embed
                String xml = new String(Files.readAllBytes(slidePath), "UTF-8");
                Matcher matcher = Pattern.compile("p:bgRef[^>]*r:embed=\\\"(rId\\d+)\\\"").matcher(xml);
                if (!matcher.find()) {
                    continue;
                }
                String bgRelId = matcher.group(1);

                // 从关系文件解析目标
                Map<String, String> relations = parseSlideRelations(relPath);
                if (relations.containsKey(bgRelId)) {
                    String target = relations.get(bgRelId);
                    if (target != null && target.startsWith("../media/")) {
                        String fileName = target.substring("../media/".length());
                        backgroundImages.add(fileName);
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("收集背景图片时出错: " + e.getMessage());
        }
        return backgroundImages;
    }

    /**
     * 收集所有“标题/描述”不为空的图片对应的文件名（仅文件名）
     */
    private static Set<String> collectImagesWithTitle(Path tempDir) {
        Set<String> images = new HashSet<>();
        final String PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
        final String DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        final String REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        try {
            List<String> slideFiles = getAvailableSlides(tempDir);
            for (String slideFile : slideFiles) {
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFile);
                Path relPath = tempDir.resolve("ppt/slides/_rels").resolve(slideFile + ".rels");
                if (!Files.exists(slidePath) || !Files.exists(relPath)) {
                    continue;
                }

                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(slidePath.toFile());

                NodeList picNodes = doc.getElementsByTagNameNS(PML_NS, "pic");
                if (picNodes == null) continue;

                // 关系映射
                Map<String, String> relations = parseSlideRelations(relPath);

                for (int i = 0; i < picNodes.getLength(); i++) {
                    Element pic = (Element) picNodes.item(i);

                    // 取 cNvPr 的 title 或 descr
                    String title = "";
                    NodeList cNvPrList = pic.getElementsByTagNameNS(PML_NS, "cNvPr");
                    if (cNvPrList != null && cNvPrList.getLength() > 0) {
                        Element cNvPr = (Element) cNvPrList.item(0);
                        title = Optional.ofNullable(cNvPr.getAttribute("title")).orElse("");
                        if (title.isEmpty()) {
                            title = Optional.ofNullable(cNvPr.getAttribute("descr")).orElse("");
                        }
                    }
                    if (title == null || title.trim().isEmpty()) {
                        continue; // 标题/描述为空，不处理
                    }

                    // 找 blip 的 r:embed
                    NodeList blipList = pic.getElementsByTagNameNS(DML_NS, "blip");
                    if (blipList == null || blipList.getLength() == 0) continue;
                    Element blip = (Element) blipList.item(0);
                    String embedId = blip.getAttributeNS(REL_NS, "embed");
                    if (embedId == null || embedId.isEmpty()) continue;

                    // 通过关系找到目标文件
                    if (relations.containsKey(embedId)) {
                        String target = relations.get(embedId);
                        if (target != null && target.startsWith("../media/")) {
                            String fileName = target.substring("../media/".length());
                            images.add(fileName);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("收集标题不为空的图片时出错: " + e.getMessage());
        }
        return images;
    }
}
