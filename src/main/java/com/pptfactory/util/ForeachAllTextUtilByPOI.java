package com.pptfactory.util;

import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlObject;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.regex.*;

public class ForeachAllTextUtilByPOI {
    private static Map<Integer,List<String>> textMap = new HashMap<>();
    static {
        textMap.put(1, Arrays.asList("一、煤矿安全生产核心方针", "安全第一", "预防为主", "综合治理"));
        textMap.put(2, Arrays.asList(
            "一、煤矿安全生产核心方针",
            "1.“安全第一”：优先保障生命安全",
            "在煤矿生产全流程中，必须将人员安全置于产量、效率之上",
            "在煤矿生产的整个过程里，人员安全的重要性要远超产量和效率。对采煤机司机而言，当操作中发现顶底板冒顶预兆、瓦斯超限等安全隐患时，需立即停机处理，严禁“冒险作业”“带病开机”。",
            "对应岗位操作中的“紧急停机情形”",
            "在岗位操作方面，当遇到顶底板有冒顶预兆等情况时，采煤机司机必须立即停机。这体现了“安全第一”的方针，将人员安全放在首位，避免因继续作业而导致安全事故。"
        ));
    }
    // 用于去重：内容 + 位置完全一致
    private static class TextKey {
        int slide;
        String shapeName;
        public TextKey(int slide, String shapeName) {
            this.slide = slide;
            this.shapeName = shapeName;
        }

        @Override
        public int hashCode() {
            return Objects.hash(slide, shapeName);
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            TextKey t = (TextKey) o;
            return slide == t.slide && shapeName.equals(t.shapeName);
        }
    }
    // 待删除的 shape 队列
    private static class ShapeDeleteTask {
        int slide;
        XSLFSlide slideRef;
        XSLFShapeContainer parent; // 如果是组内子形状，需要用父容器删除
        XSLFShape shape;
        String shapeName;
        String text;

        public ShapeDeleteTask(int slide, XSLFSlide slideRef, XSLFShapeContainer parent, XSLFShape shape, String shapeName, String text) {
            this.slide = slide;
            this.slideRef = slideRef;
            this.parent = parent;
            this.shape = shape;
            this.shapeName = shapeName;
            this.text = text;
        }
    }
    private static final Deque<ShapeDeleteTask> deleteQueue = new ArrayDeque<>();
    // 记录每个唯一 key 对应的最新 shape，用于覆盖时删除旧 shape
    private static final Map<String, ShapeDeleteTask> latestShapeMap = new HashMap<>();

    public static void main(String[] args) {
        try {
            Map<Integer,Map<String,TextItem>> results = extractPPTX(
                "/Users/menggl/workspace/PPTFactory/templates/master_template2.pptx",
          "/Users/menggl/workspace/PPTFactory/templates/master_template3.pptx"
            );
            for (Integer page : results.keySet()) {
                System.out.println("******************************第" + page + "页");
                for (TextItem item : results.get(page).values()) {
                    System.out.println(item);
                }
                System.out.println("--------------------------------第" + page + "页结束");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static class TextItem {
        public int slide;
        public String shapeName;
        public String shapeId;
        public String text;
        public String source; // "slide", "layout", "master"
        // 坐标位置
        public Rectangle2D anchor;
        public TextItem(int s, String n, String id, String t, String src, Rectangle2D anchor) {
            slide = s; shapeName = n; shapeId = id; text = t; source = src; this.anchor = anchor;
        }

        public String toString() {
            String combinedShapeId = slide + "_" + (shapeId != null ? shapeId : "N/A");
            return "shapeId: 【" + combinedShapeId + "】" 
                   + ", shapeName: 【" + shapeName + "】" 
                   + ", source: 【" + source + "】"
                   + ", text: 【" + text + "】"
                   + ", anchor: 【" + anchor + "】";
        }
    }

    // 后面的覆盖前面的，只保留最后一个
    private static void addUnique(Map<Integer,Map<String,TextItem>> results, int slide, String shapeName, String shapeId, String text, String source, Rectangle2D anchor, XSLFSlide slideRef, XSLFShapeContainer parent, XSLFShape shape) {
        Map<String,TextItem> shapeMap = results.get(slide);
        if(shapeMap == null) {
            shapeMap = new HashMap<>();
            results.put(slide, shapeMap);
        }
        TextItem textItem = new TextItem(slide, shapeName, shapeId, text, source, anchor);
        // 如果shapeName相同，文本内容相同，则后面的覆盖前面的
        String key = shapeName + "_" + text;
        String uniqueKey = slide + ":" + key;
        if(shapeMap.containsKey(key)){
            System.out.println("覆盖旧数据，旧数据："+shapeMap.get(key)+"，新数据："+textItem);
            // 如果之前的shape被覆盖了，需要删除掉之前的shape，放入删除任务队列中，调用enqueueDelete方法
            ShapeDeleteTask oldTask = latestShapeMap.get(uniqueKey);
            if (oldTask != null) {
                enqueueDelete(oldTask.slide, oldTask.slideRef, oldTask.parent, oldTask.shape, oldTask.shapeName, oldTask.text);
            }
        }
        shapeMap.put(key, textItem);
        // 更新最新 shape 记录，便于后续覆盖时删除旧 shape
        latestShapeMap.put(uniqueKey, new ShapeDeleteTask(slide, slideRef, parent, shape, shapeName, text));
    }

    private static void enqueueDelete(int slide, XSLFSlide slideRef, XSLFShapeContainer parent, XSLFShape shape, String shapeName, String text) {
        deleteQueue.addLast(new ShapeDeleteTask(slide, slideRef, parent, shape, shapeName, text));
        System.out.println("需要删除的shape: " + shapeName + ", page: " + slide + ", txt: " + text);
    }

    private static void processDeleteQueue() {
        while (!deleteQueue.isEmpty()) {
            ShapeDeleteTask task = deleteQueue.pollFirst();
            try {
                // 如果有父容器，则优先用父容器删除（解决组内子形状无法从 slide 直接删除的问题）
                if (task.parent != null) {
                    task.parent.removeShape(task.shape);
                } else {
                    task.slideRef.removeShape(task.shape);
                }
                System.out.println("已删除shape: " + task.shapeName + ", page: " + task.slide + ", txt: " + task.text);
            } catch (Exception e) {
                System.out.println("删除shape失败: " + task.shapeName + ", page: " + task.slide + ", 原因: " + e.getMessage());
            }
        }
    }

    public static Map<Integer,Map<String,TextItem>> extractPPTX(String path) throws Exception {
        return extractPPTX(path, path);
    }

    public static Map<Integer,Map<String,TextItem>> extractPPTX(String path, String outputPath) throws Exception {
        try (FileInputStream fis = new FileInputStream(path);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {
            Map<Integer,Map<String,TextItem>> results = new HashMap<>();
            deleteQueue.clear(); // 清理上一轮的残留任务
            latestShapeMap.clear(); // 清理覆盖记录

            int page = 1;
            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    extractShape(slide, shape, results, page);
                }
                page++;
            }
            processDeleteQueue(); // 扫描完成后统一删除

            if (outputPath != null && !outputPath.isEmpty()) {
                // 删除outputPath文件
                File file = new File(outputPath);
                if(file.exists()) {
                    file.delete();
                }
                // 创建outputPath文件
                file.createNewFile();
                try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                    ppt.write(fos);
                    System.out.println("已保存清理后的PPT到: " + outputPath);
                }
            }
            Thread.sleep(1000);
            return results;
        }
    }

    private static void extractShape(XSLFSlide slide, XSLFShape shape, Map<Integer,Map<String,TextItem>> results, int page) {
        // 如果 shape 是 GroupShape，则递归处理其中的子形状
        if (shape instanceof XSLFGroupShape) {
            XSLFGroupShape group = (XSLFGroupShape) shape;
            for (XSLFShape child : group.getShapes()) {
                extractShapeInContainer(slide, child, results, page, group);
            }
            return;
        }
        if (isFromMasterOrLayout(shape)) {
            System.out.println("从母版或布局中提取文本. shapeName: " + shape.getShapeName());
            return;
        }
        
        // 检查 shape 是否被隐藏
        if (isShapeHidden(shape)) {
            System.out.println("shape被隐藏. shapeName: " + shape.getShapeName());
            return;
        }

        // 如果 shape 的名称包含 SmartArt 或 组合，则忽略
        String name = shape.getShapeName();
        if (name.contains("SmartArt") || name.contains("组合")) return;

        String shapeName = shape.getShapeName();
        String shapeId = extractXmlId(shape);
        String source = detectShapeSource(slide, shape);

        // ① TextShape（优先）
        if (shape instanceof XSLFTextShape ts) {
            if (ts.getPlaceholder() != null) return; //过滤母版/布局占位符

            Rectangle2D anchor = ts.getAnchor();

            // 如果一个文本框里包含多个不同的小文本框内容，且 shape 名称像自动生成的
            List<XSLFTextParagraph> paragraphs = ts.getTextParagraphs();
            if (paragraphs.size() > 1 && shape.getShapeName().matches("文本框 \\d+")) {
                // 合并paragraphs的文本内容
                StringBuilder txt = new StringBuilder();
                for (XSLFTextParagraph paragraph : paragraphs) {
                    txt.append(paragraph.getText());
                }
                // 如果文本不在textMap中，则跳过
                if(textMap.get(page) != null && !textMap.get(page).contains(txt.toString().trim())) {
                    return;
                }
                // 如果在textMap中，则添加到results中
                addUnique(results, page, shapeName, shapeId, txt.toString().trim(), source, anchor, slide, null, shape);
                // 这是 GroupShape 的汇总文本，忽略
                return;
            }
            // 如果shapeName是矩形，忽略
            if(shapeName != null && shapeName.matches("矩形 \\d+")) return;


            // 如果shapeName不是文本框，忽略
            // if (!shapeName.matches("文本框 \\d+")) return;

            String txt = ts.getText();
            if(txt == null || txt.trim().isEmpty()) return;
            
            if (txt != null && !txt.trim().isEmpty()) {
                // 如果文本不在textMap中，则添加到删除shape列表，否则添加到results中
                List<String> textList = textMap.get(page);
                if(textList != null && !textList.isEmpty() && !textList.contains(txt.trim())) {
                    // 文本不在textMap中，则添加到删除shape队列，稍后统一删除
                    enqueueDelete(page, slide, null, shape, shapeName, txt.trim());
                    return;
                }
                addUnique(results, page, shapeName, shapeId, txt.trim(), source, anchor, slide, null, shape);
            }
        }

        // // ② Table （表格）
        // if (shape instanceof XSLFTable table) {
        //     for (XSLFTableRow r : table) {
        //         for (XSLFTableCell c : r) {
        //             String t = c.getText();
        //             if (t != null && !t.trim().isEmpty()) {
        //                 addUnique(list, page, shapeName + ":cell", shapeId, t.trim(), source);
        //             }
        //         }
        //     }
        // }

        // // ④ SmartArt or Chart（只有在 TextShape 没有文本时才读取 XML）
        // if (!hasTextByPOI) {
        //     try {
        //         String xml = shape.getXmlObject().xmlText();
        //         Matcher m = Pattern.compile("<a:t>(.*?)</a:t>").matcher(xml);

        //         while (m.find()) {
        //             String t = m.group(1).trim();
        //             if (!t.isEmpty()) {
        //                 addUnique(list, page, shapeName + ":SmartArt", shapeId, t, source);
        //             }
        //         }
        //     } catch (Exception ignore) {}
        // }
    }

    /**
     * 在指定容器中递归处理 shape，便于后续删除时使用父容器
     */
    private static void extractShapeInContainer(XSLFSlide slide, XSLFShape shape, Map<Integer,Map<String,TextItem>> results, int page, XSLFShapeContainer parent) {
        if (shape instanceof XSLFGroupShape) {
            XSLFGroupShape group = (XSLFGroupShape) shape;
            for (XSLFShape child : group.getShapes()) {
                extractShapeInContainer(slide, child, results, page, group);
            }
            return;
        }

        // 重用现有逻辑：临时将 parent 信息传入删除队列
        if (isFromMasterOrLayout(shape)) {
            System.out.println("从母版或布局中提取文本. shapeName: " + shape.getShapeName());
            return;
        }

        if (isShapeHidden(shape)) {
            System.out.println("shape被隐藏. shapeName: " + shape.getShapeName());
            return;
        }

        String shapeName = shape.getShapeName();
        if (shapeName != null) {
            String lowerName = shapeName;
            if (lowerName.contains("SmartArt") || lowerName.contains("组合")) return;
        }

        String shapeId = extractXmlId(shape);
        String source = detectShapeSource(slide, shape);

        if (shape instanceof XSLFTextShape ts) {
            if (ts.getPlaceholder() != null) return;

            List<XSLFTextParagraph> paragraphs = ts.getTextParagraphs();
            if (paragraphs.size() > 1 && shape.getShapeName().matches("文本框 \\d+")) {
                return;
            }
            if(shapeName != null && shapeName.matches("矩形 \\d+")) return;

            String txt = ts.getText();
            if(txt == null || txt.trim().isEmpty()) return;

            Rectangle2D anchor = ts.getAnchor();

            if (!txt.trim().isEmpty()) {
                List<String> textList = textMap.get(page);
                if(textList != null && !textList.isEmpty() && !textList.contains(txt.trim())) {
                    enqueueDelete(page, slide, parent, shape, shapeName, txt.trim());
                    return;
                }
                addUnique(results, page, shapeName, shapeId, txt.trim(), source, anchor, slide, parent, shape);
            }
        }
    }

    private static boolean isFromMasterOrLayout(XSLFShape shape) {
        // 1. 占位符
        if (shape.isPlaceholder()) {
            System.out.println("占位符. shapeName: " + shape.getShapeName());
            return true;
        }
        
        // 2. 检查形状名称
        String name = shape.getShapeName();
        if (name == null){
            System.out.println("name is null. shapeName: " + shape.getShapeName());
            return false;
        }
        
        String lower = name.toLowerCase();
        String[] keywords = {
            "master", "layout", "placeholder", 
            "标题", "页脚", "页眉", "日期", "编号",
            "footer", "header", "slide number", "date"
        };
        
        for (String kw : keywords) {
            if (lower.contains(kw)) {
                System.out.println("name contains " + kw + ". shapeName: " + shape.getShapeName());
                return true;
            }
        }
        
        return false;
    }

    /**
     * 从XML中提取形状ID
     */
    private static String extractXmlId(XSLFShape shape) {
        try {
            XmlObject xmlObj = shape.getXmlObject();
            String xml = xmlObj.xmlText();
            
            // 查找 id="xxx" 或 id='xxx'
            java.util.regex.Pattern pattern = 
                java.util.regex.Pattern.compile("id\\s*=\\s*[\"']([^\"']+)[\"']");
            java.util.regex.Matcher matcher = pattern.matcher(xml);
            
            if (matcher.find()) {
                return matcher.group(1);
            }
        } catch (Exception e) {
            // 忽略异常
        }
        return null;
    }

    /**
     * 检测 shape 的来源：slide、layout 或 master
     * 通过检查 XML 底层来判断
     */
    private static String detectShapeSource(XSLFSlide slide, XSLFShape shape) {
        try {
            // 方法1: 通过检查 shape 的 XML 对象所在的文档部分来判断
            XmlObject xmlObj = shape.getXmlObject();
            if (xmlObj == null) {
                return "unknown";
            }
            
            // 尝试获取 shape 所在的文档部分
            try {
                // 通过反射获取文档部分
                java.lang.reflect.Method getParentMethod = shape.getClass().getMethod("getParent");
                Object parent = getParentMethod.invoke(shape);
                
                if (parent != null) {
                    // 检查 parent 的类型
                    String parentClassName = parent.getClass().getName();
                    
                    if (parentClassName.contains("XSLFSlide")) {
                        // 检查是否在 slide 的直接 shapes 中
                        if (isShapeInList(slide.getShapes(), shape)) {
                            return "slide";
                        }
                    } else if (parentClassName.contains("XSLFSlideLayout")) {
                        return "layout";
                    } else if (parentClassName.contains("XSLFSlideMaster")) {
                        return "master";
                    }
                }
            } catch (Exception e) {
                // 反射失败，使用备用方法
            }
            
            // 方法2: 通过比较 shape ID 来判断
            // 检查 shape 是否在 slide、layout 或 master 的 shapes 列表中
            String shapeId = extractXmlId(shape);
            if (shapeId != null) {
                // 先检查 slide（最可能的情况）
                if (containsShapeById(slide.getShapes(), shapeId)) {
                    // 进一步确认：检查是否在 slide 的直接 shapes 中（不是继承的）
                    if (isShapeDirectlyInSlide(slide, shapeId)) {
                        return "slide";
                    }
                }
                
                // 检查 layout
                XSLFSlideLayout layout = slide.getSlideLayout();
                if (layout != null && containsShapeById(layout.getShapes(), shapeId)) {
                    return "layout";
                }
                
                // 检查 master
                if (layout != null) {
                    XSLFSlideMaster master = layout.getSlideMaster();
                    if (master != null && containsShapeById(master.getShapes(), shapeId)) {
                        return "master";
                    }
                }
            }
            
            // 方法3: 通过检查 XML 对象的文档 URI 来判断
            try {
                org.apache.xmlbeans.XmlDocumentProperties props = xmlObj.documentProperties();
                if (props != null) {
                    String sourceUri = props.getSourceName();
                    if (sourceUri != null) {
                        if (sourceUri.contains("/slides/slide")) {
                            return "slide";
                        } else if (sourceUri.contains("/slideLayouts/slideLayout")) {
                            return "layout";
                        } else if (sourceUri.contains("/slideMasters/slideMaster")) {
                            return "master";
                        }
                    }
                }
            } catch (Exception e) {
                // 忽略异常
            }
            
            // 默认返回 slide（大多数情况下 shape 是在 slide 中）
            return "slide";
            
        } catch (Exception e) {
            e.printStackTrace();
            return "unknown";
        }
    }
    
    /**
     * 检查 shape 是否在 shapes 列表中（通过对象引用）
     */
    private static boolean isShapeInList(List<XSLFShape> shapes, XSLFShape targetShape) {
        for (XSLFShape s : shapes) {
            if (s == targetShape) {
                return true;
            }
            // 如果是 GroupShape，递归检查子形状
            if (s instanceof XSLFGroupShape) {
                if (isShapeInList(((XSLFGroupShape) s).getShapes(), targetShape)) {
                    return true;
                }
            }
        }
        return false;
    }
    
    /**
     * 检查 shape 是否直接在 slide 中（不是从 layout/master 继承的）
     */
    private static boolean isShapeDirectlyInSlide(XSLFSlide slide, String shapeId) {
        // 这个方法需要更深入的实现
        // 暂时返回 true，表示在 slide 中
        return true;
    }
    
    /**
     * 检查 shapes 列表中是否包含指定 ID 的 shape
     */
    private static boolean containsShapeById(List<XSLFShape> shapes, String targetId) {
        for (XSLFShape s : shapes) {
            String id = extractXmlId(s);
            if (targetId.equals(id)) {
                return true;
            }
            // 如果是 GroupShape，递归检查子形状
            if (s instanceof XSLFGroupShape) {
                if (containsShapeById(((XSLFGroupShape) s).getShapes(), targetId)) {
                    return true;
                }
            }
        }
        return false;
    }
    
    /**
     * 通过多种方法判断 shape 是否被隐藏
     * 检查方式：
     * 1. XML 中的 hidden 属性
     * 2. 样式中的 visibility 属性
     * 3. 透明度（alpha）设置为 0 或接近 0
     * 4. 位置在可见区域外（可选）
     * 5. 填充和线条都设置为无
     */
    private static boolean isShapeHidden(XSLFShape shape) {
        try {
            // 方法1: 检查 XML 中的 hidden 属性
            if (isHiddenByXmlAttribute(shape)) {
                System.out.println("隐藏匹配: hidden attr, shapeName=" + shape.getShapeName());
                return true;
            }
            
            // 方法2: 检查样式中的 visibility 属性
            if (isHiddenByVisibility(shape)) {
                System.out.println("隐藏匹配: visibility attr, shapeName=" + shape.getShapeName());
                return true;
            }
            
            // 方法3: 检查透明度
            if (isHiddenByTransparency(shape)) {
                System.out.println("隐藏匹配: transparency alpha≈0, shapeName=" + shape.getShapeName());
                return true;
            }
            
            // 方法4: 检查填充和线条是否都设置为无（可能是隐藏的）
            if (isHiddenByNoFillAndNoLine(shape)) {
                System.out.println("隐藏匹配: noFill & noLine, shapeName=" + shape.getShapeName());
                return true;
            }
            
            // 方法5: 检查是否在可见区域外（可选，根据需求决定是否启用）
            // if (isHiddenByPosition(shape)) {
            //     return true;
            // }
            
        } catch (Exception e) {
            // 如果检查过程中出现异常，默认认为不隐藏
            return false;
        }
        
        return false;
    }
    
    /**
     * 方法1: 检查 XML 中的 hidden 属性
     */
    private static boolean isHiddenByXmlAttribute(XSLFShape shape) {
        try {
            XmlObject xmlObj = shape.getXmlObject();
            if (xmlObj == null) {
                return false;
            }
            
            String xml = xmlObj.xmlText();
            if (xml == null) {
                return false;
            }
            
            // 检查 hidden="1" 或 hidden="true"
            Pattern hiddenPattern = Pattern.compile("hidden\\s*=\\s*[\"'](1|true)[\"']", Pattern.CASE_INSENSITIVE);
            Matcher matcher = hiddenPattern.matcher(xml);
            if (matcher.find()) {
                return true;
            }
            
            // 检查 <p:spPr><a:xfrm hidden="1">
            Pattern xfrmHiddenPattern = Pattern.compile("<a:xfrm[^>]*hidden\\s*=\\s*[\"'](1|true)[\"']", Pattern.CASE_INSENSITIVE);
            matcher = xfrmHiddenPattern.matcher(xml);
            if (matcher.find()) {
                return true;
            }
            
        } catch (Exception e) {
            // 忽略异常
        }
        return false;
    }
    
    /**
     * 方法2: 检查样式中的 visibility 属性
     */
    private static boolean isHiddenByVisibility(XSLFShape shape) {
        try {
            XmlObject xmlObj = shape.getXmlObject();
            if (xmlObj == null) {
                return false;
            }
            
            String xml = xmlObj.xmlText();
            if (xml == null) {
                return false;
            }
            
            // 检查 visibility="hidden" 或 visibility="none"
            Pattern visibilityPattern = Pattern.compile("visibility\\s*=\\s*[\"'](hidden|none)[\"']", Pattern.CASE_INSENSITIVE);
            Matcher matcher = visibilityPattern.matcher(xml);
            if (matcher.find()) {
                return true;
            }
            
            // 检查 display="none"
            Pattern displayPattern = Pattern.compile("display\\s*=\\s*[\"']none[\"']", Pattern.CASE_INSENSITIVE);
            matcher = displayPattern.matcher(xml);
            if (matcher.find()) {
                return true;
            }
            
        } catch (Exception e) {
            // 忽略异常
        }
        return false;
    }
    
    /**
     * 方法3: 检查透明度（alpha 值）
     */
    private static boolean isHiddenByTransparency(XSLFShape shape) {
        try {
            XmlObject xmlObj = shape.getXmlObject();
            if (xmlObj == null) {
                return false;
            }
            
            String xml = xmlObj.xmlText();
            if (xml == null) {
                return false;
            }
            
            // 检查 alpha 属性，如果 alpha="0" 或接近 0，则认为隐藏
            // alpha 值范围通常是 0-100000（100000 = 100% = 完全不透明）
            Pattern alphaPattern = Pattern.compile("alpha\\s*=\\s*[\"'](\\d+)[\"']", Pattern.CASE_INSENSITIVE);
            Matcher matcher = alphaPattern.matcher(xml);
            
            while (matcher.find()) {
                try {
                    int alpha = Integer.parseInt(matcher.group(1));
                    // 如果 alpha 值小于 100（即小于 0.1%），认为完全透明（隐藏）
                    if (alpha < 100) {
                        return true;
                    }
                } catch (NumberFormatException e) {
                    // 忽略数字解析错误
                }
            }
            
            // 检查 val 属性中的 alpha 值（在某些格式中）
            Pattern valAlphaPattern = Pattern.compile("<a:alpha[^>]*val\\s*=\\s*[\"'](\\d+)[\"']", Pattern.CASE_INSENSITIVE);
            matcher = valAlphaPattern.matcher(xml);
            
            while (matcher.find()) {
                try {
                    int alpha = Integer.parseInt(matcher.group(1));
                    if (alpha < 100) {
                        return true;
                    }
                } catch (NumberFormatException e) {
                    // 忽略数字解析错误
                }
            }
            
        } catch (Exception e) {
            // 忽略异常
        }
        return false;
    }
    
    /**
     * 方法4: 检查填充和线条是否都设置为无
     * 注意：这个方法可能不够准确，因为有些 shape 可能本身就是无填充无线条的
     * 但结合其他方法使用，可以作为辅助判断
     */
    private static boolean isHiddenByNoFillAndNoLine(XSLFShape shape) {
        try {
            XmlObject xmlObj = shape.getXmlObject();
            if (xmlObj == null) {
                return false;
            }
            
            String xml = xmlObj.xmlText();
            if (xml == null) {
                return false;
            }
            
            // 检查是否有 noFill 和 noLn（无线条）
            boolean hasNoFill = xml.contains("noFill") || xml.contains("NoFill");
            boolean hasNoLine = xml.contains("noLn") || xml.contains("NoLn") || 
                               (xml.contains("<a:ln") && xml.contains("w=\"0\""));
            
            // 如果既无填充又无线条，且不是文本形状，可能是隐藏的
            // 但对于文本形状，即使无填充无线条，文本仍然可见，所以需要特殊处理
            if (hasNoFill && hasNoLine && !(shape instanceof XSLFTextShape)) {
                // 进一步检查：如果形状大小也很小（接近 0），可能是隐藏的
                try {
                    Rectangle2D anchor = shape.getAnchor();
                    if (anchor != null && (anchor.getWidth() < 1 || anchor.getHeight() < 1)) {
                        return true;
                    }
                } catch (Exception e) {
                    // 忽略异常
                }
            }
            
        } catch (Exception e) {
            // 忽略异常
        }
        return false;
    }
    
    /**
     * 方法5: 检查是否在可见区域外（可选方法）
     * 注意：这个方法可能不够准确，因为 shape 可能被故意放在可见区域外
     */
    private static boolean isHiddenByPosition(XSLFShape shape) {
        try {
            Rectangle2D anchor = shape.getAnchor();
            if (anchor == null) {
                return false;
            }
            
            // 如果形状完全在负坐标区域或超出常见幻灯片尺寸（如 1920x1080），可能是隐藏的
            // 但这只是辅助判断，不能作为唯一依据
            double x = anchor.getX();
            double y = anchor.getY();
            double width = anchor.getWidth();
            double height = anchor.getHeight();
            
            // 如果形状完全在负坐标区域
            if (x + width < 0 || y + height < 0) {
                return true;
            }
            
            // 如果形状超出常见幻灯片尺寸（假设最大为 20000x20000）
            if (x > 20000 || y > 20000) {
                return true;
            }
            
        } catch (Exception e) {
            // 忽略异常
        }
        return false;
    }
    
}
