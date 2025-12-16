package com.pptfactory.util;

import org.w3c.dom.*;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.imageio.ImageIO;

/**
 * 扫描 PPTX 中每页的图片，查找图片的可选文字（cNvPr title/descr）不为空的图片，
 * 并且读取图片在幻灯片中展示时的宽/高（从 slide XML 中的 a:ext cx/cy 获取，单位为 EMU）。
 * 输出结果到控制台并保存到 produce/ppt_image_info_<timestamp>.json
 */
public class ScanPPTImageInfoUtil {

    private static final String PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static void main(String[] args) throws Exception {
        String pptxPath = args != null && args.length > 0 ? args[0] : "templates/master_template.pptx";
        scanPptx(pptxPath);
    }

    public static void scanPptx(String pptxPath) throws Exception {
        Path pptx = Paths.get(pptxPath);
        if (!Files.exists(pptx)) {
            System.err.println("PPTX 文件不存在: " + pptxPath);
            return;
        }

        Path tempDir = createTempDirInProject("scan_pptx_");
        try {
            unzipPPTX(pptx.toString(), tempDir.toString());

            List<String> slideFiles = getAvailableSlides(tempDir);
            Map<String, Object> report = new LinkedHashMap<>();
            List<Map<String, Object>> items = new ArrayList<>();

            for (int si = 0; si < slideFiles.size(); si++) {
                String slideFile = slideFiles.get(si);
                Path slidePath = tempDir.resolve("ppt/slides").resolve(slideFile);
                Path relPath = tempDir.resolve("ppt/slides/_rels").resolve(slideFile + ".rels");
                Map<String, String> relations = new HashMap<>();
                if (Files.exists(relPath)) {
                    relations = parseSlideRelations(relPath);
                }

                // parse slide xml
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(slidePath.toFile());

                final String PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
                final String DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
                final String REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                NodeList picNodes = doc.getElementsByTagNameNS(PML_NS, "pic");
                if (picNodes == null) continue;

                for (int i = 0; i < picNodes.getLength(); i++) {
                    Element pic = (Element) picNodes.item(i);

                    // 获取 cNvPr 的 title 或 descr
                    String title = "";
                    NodeList cNvPrList = pic.getElementsByTagNameNS(PML_NS, "cNvPr");
                    if (cNvPrList != null && cNvPrList.getLength() > 0) {
                        Element cNvPr = (Element) cNvPrList.item(0);
                        title = Optional.ofNullable(cNvPr.getAttribute("title")).orElse("");
                        if (title == null || title.trim().isEmpty()) {
                            title = Optional.ofNullable(cNvPr.getAttribute("descr")).orElse("");
                        }
                    }

                    if (title == null) title = "";
                    if (title.trim().isEmpty()) {
                        continue; // 只关心有标注的图片
                    }

                    // 找 blip 的 r:embed
                    NodeList blipList = pic.getElementsByTagNameNS(DML_NS, "blip");
                    String embedId = null;
                    if (blipList != null && blipList.getLength() > 0) {
                        Element blip = (Element) blipList.item(0);
                        embedId = blip.getAttributeNS(REL_NS, "embed");
                    }

                    String mediaFileName = null;
                    if (embedId != null && relations.containsKey(embedId)) {
                        String target = relations.get(embedId);
                        if (target != null && target.startsWith("../media/")) {
                            mediaFileName = target.substring("../media/".length());
                        }
                    }

                    // 计算图片在幻灯片上的最终显示尺寸（考虑祖先 group 的 chOff/chExt -> ext 映射）
                    long cx = -1L, cy = -1L; // EMU
                    long[] finalExt = computeFinalExt(pic);
                    if (finalExt != null && finalExt.length >= 2) {
                        cx = finalExt[0];
                        cy = finalExt[1];
                    }

                    Map<String, Object> record = new LinkedHashMap<>();
                    record.put("slide_file", slideFile);
                    record.put("slide_index", si + 1);
                    record.put("picture_index_on_slide", i + 1);
                    record.put("annotation", title);
                    record.put("media_file", mediaFileName == null ? "" : mediaFileName);

                    // convert EMU to cm and pixels (120 dpi)
                    if (cx > 0 && cy > 0) {
                        double widthCm = emuToCm(cx);
                        double heightCm = emuToCm(cy);
                        double widthPx120 = emuToPixels(cx, 120);
                        double heightPx120 = emuToPixels(cy, 120);
                        record.put("width_emu", cx);
                        record.put("height_emu", cy);
                        record.put("width_cm", round(widthCm, 2));
                        record.put("height_cm", round(heightCm, 2));
                        record.put("width_px_120dpi", (int)Math.round(widthPx120));
                        record.put("height_px_120dpi", (int)Math.round(heightPx120));
                    }

                    // 如果有 media 文件，读取实际像素尺寸
                    if (mediaFileName != null) {
                        Path mediaPath = tempDir.resolve("ppt/media").resolve(mediaFileName);
                        if (Files.exists(mediaPath)) {
                            try {
                                BufferedImage img = ImageIO.read(mediaPath.toFile());
                                if (img != null) {
                                    record.put("image_pixel_width", img.getWidth());
                                    record.put("image_pixel_height", img.getHeight());
                                }
                            } catch (Exception e) {
                                // ignore
                            }
                        }
                    }

                    items.add(record);
                }
            }

            report.put("scanned_at", System.currentTimeMillis());
            report.put("pptx", pptxPath);
            report.put("items", items);

            // 输出到控制台（简洁格式）
            for (Map<String, Object> it : items) {
                System.out.println(it);
            }

            // 写入 produce 目录的 json 文件
            // 如果 pptx 是相对路径且没有上两级父路径，fallback 到当前工作目录
            Path projectRoot;
            if (pptx.getParent() != null && pptx.getParent().getParent() != null) {
                projectRoot = pptx.getParent().getParent();
            } else {
                projectRoot = Paths.get(System.getProperty("user.dir"));
            }
            Path produceDir = projectRoot.resolve("produce");
            if (!Files.exists(produceDir)) Files.createDirectories(produceDir);
            String outName = "ppt_image_info_" + System.currentTimeMillis() + ".json";
            Path outPath = produceDir.resolve(outName);

            // 过滤掉 annotation 为 "警告" 的条目
            if (report.get("items") instanceof List) {
                @SuppressWarnings("unchecked")
                List<Map<String, Object>> repItems = (List<Map<String, Object>>) report.get("items");
                repItems.removeIf(it -> {
                    Object a = it.get("annotation");
                    return a != null && "警告".equals(a.toString());
                });
            }

            // 使用内部实现将 JSON 写为 UTF-8、漂亮格式（不转义非 ASCII）
            writeJsonPretty(outPath, report);
            System.out.println("保存结果到: " + outPath.toString());

        } finally {
            // 清理临时目录
            // 留下临时目录以便调试；如果想删除可解除下面注释
            // cleanTempDirectory(tempDir);
        }
    }

    private static String toJson(Object obj) {
        // 简单实现，适用于本处生成的 Map/List 结构
        if (obj instanceof Map) {
            StringBuilder sb = new StringBuilder();
            sb.append("{");
            Map<?, ?> m = (Map<?, ?>) obj;
            boolean first = true;
            for (Map.Entry<?, ?> e : m.entrySet()) {
                if (!first) sb.append(","); first = false;
                sb.append(jsonEscape(e.getKey().toString())).append(":").append(toJson(e.getValue()));
            }
            sb.append("}");
            return sb.toString();
        } else if (obj instanceof List) {
            StringBuilder sb = new StringBuilder();
            sb.append("[");
            List<?> l = (List<?>) obj;
            boolean first = true;
            for (Object o : l) {
                if (!first) sb.append(","); first = false;
                sb.append(toJson(o));
            }
            sb.append("]");
            return sb.toString();
        } else if (obj instanceof String) {
            return jsonEscape((String) obj);
        } else if (obj instanceof Number || obj instanceof Boolean) {
            return obj.toString();
        } else {
            return jsonEscape(String.valueOf(obj));
        }
    }

    private static String jsonEscape(String s) {
        if (s == null) return "\"\"";
        return "\"" + s.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
    }

    private static void writeJsonPretty(Path outPath, Object obj) throws IOException {
        try (BufferedWriter w = Files.newBufferedWriter(outPath, StandardCharsets.UTF_8)) {
            writeValue(w, obj, 0);
            w.newLine();
        }
    }

    private static void writeValue(BufferedWriter w, Object obj, int indent) throws IOException {
        String ind = "";
        for (int i = 0; i < indent; i++) ind += "    ";

        if (obj == null) {
            w.write("null");
        } else if (obj instanceof Map) {
            @SuppressWarnings("unchecked")
            Map<String, Object> m = (Map<String, Object>) obj;
            w.write("{");
            if (!m.isEmpty()) {
                w.newLine();
                Iterator<Map.Entry<String, Object>> it = m.entrySet().iterator();
                while (it.hasNext()) {
                    Map.Entry<String, Object> e = it.next();
                    w.write(ind + "    " + quote(e.getKey()) + ": ");
                    writeValue(w, e.getValue(), indent + 1);
                    if (it.hasNext()) w.write(",");
                    w.newLine();
                }
                w.write(ind);
            }
            w.write("}");
        } else if (obj instanceof List) {
            @SuppressWarnings("unchecked")
            List<Object> l = (List<Object>) obj;
            w.write("[");
            if (!l.isEmpty()) {
                w.newLine();
                Iterator<Object> it = l.iterator();
                while (it.hasNext()) {
                    Object o = it.next();
                    w.write(ind + "    ");
                    writeValue(w, o, indent + 1);
                    if (it.hasNext()) w.write(",");
                    w.newLine();
                }
                w.write(ind);
            }
            w.write("]");
        } else if (obj instanceof String) {
            w.write(quote((String) obj));
        } else if (obj instanceof Number || obj instanceof Boolean) {
            w.write(String.valueOf(obj));
        } else {
            w.write(quote(String.valueOf(obj)));
        }
    }

    private static String quote(String s) {
        if (s == null) return "\"\"";
        StringBuilder sb = new StringBuilder();
        sb.append('"');
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '\\': sb.append("\\\\"); break;
                case '"': sb.append("\\\""); break;
                case '\b': sb.append("\\b"); break;
                case '\f': sb.append("\\f"); break;
                case '\n': sb.append("\\n"); break;
                case '\r': sb.append("\\r"); break;
                case '\t': sb.append("\\t"); break;
                default:
                    sb.append(c);
            }
        }
        sb.append('"');
        return sb.toString();
    }

    private static double emuToCm(long emu) {
        return emu * 2.54 / 914400.0;
    }

    private static double emuToPixels(long emu, int dpi) {
        return (emu / 914400.0) * dpi;
    }

    private static double round(double v, int digits) {
        double p = Math.pow(10, digits);
        return Math.round(v * p) / p;
    }

    /**
     * 计算给定图片元素在幻灯片上的最终显示尺寸和位置。
     *
     * 算法：
     * - 首先读取图片自身的 `a:xfrm` 的 `off`/`ext`（如果存在），作为局部坐标。
     * - 向上遍历祖先元素，若祖先存在 `a:xfrm`，并包含 `chOff`/`chExt`（表示子坐标系）和 `off`/`ext`（表示在父坐标系中的映射），
     *   则将子坐标系中的偏移与尺寸按比例映射到父坐标系：
     *   newOff = parentOff + round((childOff - parentChOff) * scale)
     *   newExt = round(childExt * scale)
     * - cascade 到根（slide 坐标系），得到最终的 off/ext（单位：EMU）。
     *
     * 返回值数组格式：{extCx, extCy, offX, offY}，其中 extCx/extCy 为最终宽高（EMU），
     * 如果无法确定最终宽高则返回 extCx/extCy 为 -1。
     */
    private static long[] computeFinalExt(Element pic) {
        long offX = 0, offY = 0, extCx = -1, extCy = -1;

        // 初始层：pic 自身的 xfrm
        NodeList xfrmList = pic.getElementsByTagNameNS(DML_NS, "xfrm");
        if (xfrmList != null && xfrmList.getLength() > 0) {
            Element xfrm = (Element) xfrmList.item(0);
            Element offEl = getFirstChildElementByTagNameNS(xfrm, DML_NS, "off");
            Element extEl = getFirstChildElementByTagNameNS(xfrm, DML_NS, "ext");
            offX = parseLongAttr(offEl, "x", 0);
            offY = parseLongAttr(offEl, "y", 0);
            extCx = parseLongAttr(extEl, "cx", -1);
            extCy = parseLongAttr(extEl, "cy", -1);
        }

        // 向上遍历祖先，应用 group's (off,ext,chOff,chExt) 映射
        Node anc = pic.getParentNode();
        while (anc != null && anc.getNodeType() == Node.ELEMENT_NODE) {
            Element aElem = (Element) anc;
            NodeList aXfrms = aElem.getElementsByTagNameNS(DML_NS, "xfrm");
            if (aXfrms != null && aXfrms.getLength() > 0) {
                Element axfrm = (Element) aXfrms.item(0);
                Element aOff = getFirstChildElementByTagNameNS(axfrm, DML_NS, "off");
                Element aExt = getFirstChildElementByTagNameNS(axfrm, DML_NS, "ext");
                Element aChOff = getFirstChildElementByTagNameNS(axfrm, DML_NS, "chOff");
                Element aChExt = getFirstChildElementByTagNameNS(axfrm, DML_NS, "chExt");

                long aOffX = parseLongAttr(aOff, "x", 0);
                long aOffY = parseLongAttr(aOff, "y", 0);
                long aExtCx = parseLongAttr(aExt, "cx", -1);
                long aExtCy = parseLongAttr(aExt, "cy", -1);
                long aChOffX = parseLongAttr(aChOff, "x", 0);
                long aChOffY = parseLongAttr(aChOff, "y", 0);
                long aChExtCx = parseLongAttr(aChExt, "cx", 0);
                long aChExtCy = parseLongAttr(aChExt, "cy", 0);

                double scaleX = 1.0, scaleY = 1.0;
                if (aChExtCx > 0) {
                    if (aExtCx > 0) {
                        scaleX = (double) aExtCx / (double) aChExtCx;
                    } else if (extCx > 0) {
                        scaleX = (double) aExtCx / (double) aChExtCx;
                    }
                } else if (aExtCx > 0 && aChExtCx == 0 && extCx > 0) {
                    scaleX = (double) aExtCx / (double) extCx;
                }
                if (aChExtCy > 0) {
                    if (aExtCy > 0) {
                        scaleY = (double) aExtCy / (double) aChExtCy;
                    } else if (extCy > 0) {
                        scaleY = (double) aExtCy / (double) aChExtCy;
                    }
                } else if (aExtCy > 0 && aChExtCy == 0 && extCy > 0) {
                    scaleY = (double) aExtCy / (double) extCy;
                }

                long newOffX = aOffX + Math.round((offX - aChOffX) * scaleX);
                long newOffY = aOffY + Math.round((offY - aChOffY) * scaleY);
                long newExtCx = extCx > 0 ? Math.round(extCx * scaleX) : aExtCx;
                long newExtCy = extCy > 0 ? Math.round(extCy * scaleY) : aExtCy;

                offX = newOffX;
                offY = newOffY;
                extCx = newExtCx;
                extCy = newExtCy;
            }
            anc = anc.getParentNode();
        }

        if (extCx <= 0 || extCy <= 0) {
            return new long[]{-1L, -1L, offX, offY};
        }
        return new long[]{extCx, extCy, offX, offY};
    }

    /**
     * 从元素属性中安全解析 long 值，解析失败或元素/属性缺失时返回默认值。
     */
    private static long parseLongAttr(Element el, String attr, long def) {
        if (el == null) return def;
        String v = el.getAttribute(attr);
        if (v == null || v.isEmpty()) return def;
        try {
            return Long.parseLong(v);
        } catch (NumberFormatException e) {
            return def;
        }
    }

    /**
     * 返回父元素中第一个匹配给定命名空间和本地名的子元素，找不到时返回 null。
     */
    private static Element getFirstChildElementByTagNameNS(Element parent, String ns, String localName) {
        if (parent == null) return null;
        NodeList nl = parent.getElementsByTagNameNS(ns, localName);
        if (nl != null && nl.getLength() > 0) return (Element) nl.item(0);
        return null;
    }

    // --- 下面是 PPTX 解压与关系解析的辅助方法（与项目中其他 util 方法类似） ---

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

    private static List<String> getAvailableSlides(Path tempDir) throws IOException {
        Path slidesDir = tempDir.resolve("ppt/slides");
        if (!Files.exists(slidesDir)) return Collections.emptyList();
        List<String> slideFiles = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(slidesDir, "slide*.xml")) {
            for (Path path : stream) slideFiles.add(path.getFileName().toString());
        }
        slideFiles.sort((a, b) -> {
            Pattern pattern = Pattern.compile("slide(\\d+)\\.xml");
            Matcher ma = pattern.matcher(a); Matcher mb = pattern.matcher(b);
            if (ma.find() && mb.find()) {
                return Integer.compare(Integer.parseInt(ma.group(1)), Integer.parseInt(mb.group(1)));
            }
            return a.compareTo(b);
        });
        return slideFiles;
    }

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
                if (type != null && (type.contains("image") || type.contains("picture"))) {
                    relations.put(id, target);
                }
            }
        }
        return relations;
    }

    private static Path getProjectTempDir() throws IOException {
        Path projectRoot = Paths.get(System.getProperty("user.dir"));
        Path tempDir = projectRoot.resolve("temp");
        if (!Files.exists(tempDir)) Files.createDirectories(tempDir);
        return tempDir;
    }

    private static Path createTempDirInProject(String prefix) throws IOException {
        Path projectTempDir = getProjectTempDir();
        String dirName = prefix + System.currentTimeMillis() + "_" + Thread.currentThread().getId();
        Path tempDir = projectTempDir.resolve(dirName);
        Files.createDirectories(tempDir);
        return tempDir;
    }

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
        } catch (IOException e) {
            System.err.println("清理临时目录时出错: " + e.getMessage());
        }
    }
}
