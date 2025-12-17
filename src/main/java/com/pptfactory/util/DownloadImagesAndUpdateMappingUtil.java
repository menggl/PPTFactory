package com.pptfactory.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 根据 ppt内容映射.txt 的图片链接映射，下载图片到本地，并生成图片路径映射。
 * 图片保存路径 produce/images/new_ppt_[年月日时分秒]/[第几页]_[第几张图片].png
 * 图片路径映射写回 produce/ppt内容映射.txt
 */
public class DownloadImagesAndUpdateMappingUtil {
    private static final String MAPPING_FILE = "produce/ppt内容映射.txt";
    private static final String IMAGE_BASE_DIR = "produce/images";
    private static final ObjectMapper MAPPER = new ObjectMapper();

    public static void main(String[] args) throws Exception {
        // 1. 读取映射文件
        List<Map<String, Object>> mappings = readMappings();
        if (mappings == null) {
            System.err.println("未能读取映射文件或格式错误");
            return;
        }
        // 2. 获取新PPT文件名（假设和图片目录名一致）
        String pptFileName = getLatestPptFileName();
        if (pptFileName == null) {
            System.err.println("未找到新生成的PPT文件");
            return;
        }
        String imageDir = IMAGE_BASE_DIR + "/" + pptFileName.substring(0, pptFileName.lastIndexOf('.'));
        Files.createDirectories(Paths.get(imageDir));

        // 3. 遍历每一页，下载图片，生成图片路径映射
        for (int i = 0; i < mappings.size(); i++) {
            Map<String, Object> mapping = mappings.get(i);
            Map<String, String> imageUrlMap = getStringMap(mapping.get("图片链接映射"));
            if (imageUrlMap == null || imageUrlMap.isEmpty()) continue;
            Map<String, String> imagePathMap = new LinkedHashMap<>();
            int imgIdx = 1;
            for (Map.Entry<String, String> entry : imageUrlMap.entrySet()) {
                String label = entry.getKey();
                String url = entry.getValue();
                if (url == null || url.isEmpty()) continue;
                String imgFileName = (i + 1) + "_" + imgIdx + ".png";
                String imgPath = imageDir + "/" + imgFileName;
                boolean ok = downloadImage(url, imgPath);
                if (ok) {
                    imagePathMap.put(label, imgPath);
                    System.out.println("已下载: " + label + " => " + imgPath);
                } else {
                    System.err.println("下载失败: " + label + " => " + url);
                }
                imgIdx++;
            }
            if (!imagePathMap.isEmpty()) {
                mapping.put("图片路径映射", imagePathMap);
            }
        }
        // 4. 写回映射文件
        writeMappings(mappings);
        System.out.println("图片下载及路径映射已完成，结果已写回: " + MAPPING_FILE);
    }

    private static List<Map<String, Object>> readMappings() throws IOException {
        try (InputStream is = new FileInputStream(MAPPING_FILE)) {
            return MAPPER.readValue(is, new TypeReference<List<Map<String, Object>>>() {});
        }
    }

    private static void writeMappings(List<Map<String, Object>> mappings) throws IOException {
        try (Writer writer = new OutputStreamWriter(new FileOutputStream(MAPPING_FILE), StandardCharsets.UTF_8)) {
            MAPPER.writerWithDefaultPrettyPrinter().writeValue(writer, mappings);
        }
    }

    private static Map<String, String> getStringMap(Object obj) {
        if (obj instanceof Map) {
            Map<?, ?> map = (Map<?, ?>) obj;
            Map<String, String> result = new LinkedHashMap<>();
            for (Map.Entry<?, ?> e : map.entrySet()) {
                if (e.getKey() != null && e.getValue() != null) {
                    result.put(e.getKey().toString(), e.getValue().toString());
                }
            }
            return result;
        }
        return null;
    }

    private static boolean downloadImage(String urlStr, String filePath) {
        try (InputStream in = new URL(urlStr).openStream()) {
            Files.copy(in, Paths.get(filePath));
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private static String getLatestPptFileName() {
        File dir = new File("produce");
        File[] files = dir.listFiles((d, name) -> name.startsWith("new_ppt_") && name.endsWith(".pptx"));
        if (files == null || files.length == 0) return null;
        Arrays.sort(files, Comparator.comparing(File::getName).reversed());
        return files[0].getName();
    }
}
