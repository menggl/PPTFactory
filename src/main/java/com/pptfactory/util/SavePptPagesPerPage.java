package com.pptfactory.util;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.util.Iterator;
import java.util.Optional;

/**
 * 将 produce/ppt内容页.txt 的每一页保存为单独记录到 ppt_page 表，主键为 (ppt_name, page_number)。
 * 用法：
 * java -cp <classpath-with-mysql-connector> com.pptfactory.util.SavePptPagesPerPage <ppt_name> [path/to/produce/ppt内容页.txt]
 *
 * DB 配置通过环境变量：DB_URL, DB_USER, DB_PASS
 */
public class SavePptPagesPerPage {
    private static final ObjectMapper MAPPER = new ObjectMapper();

    public static void main(String[] args) {
        if (args == null || args.length == 0) {
            System.err.println("必须提供 ppt_name（示例 new_ppt_20251217123045）作为第一个参数");
            System.exit(1);
        }
        String pptName = args[0];
        String inputPath = args.length > 1 ? args[1] : "produce/ppt内容页.txt";

        String dbUrl = Optional.ofNullable(System.getenv("DB_URL")).orElse("jdbc:mysql://127.0.0.1:3306/pptfactory?useUnicode=true&characterEncoding=utf8mb4&serverTimezone=UTC");
        String dbUser = Optional.ofNullable(System.getenv("DB_USER")).orElse("root");
        String dbPass = Optional.ofNullable(System.getenv("DB_PASS")).orElse("");

        try {
            Path p = Paths.get(inputPath);
            if (!Files.exists(p)) {
                System.err.println("输入文件不存在: " + inputPath);
                System.exit(2);
            }

            String raw = Files.readString(p, StandardCharsets.UTF_8);

            JsonNode rootNode;
            try {
                rootNode = MAPPER.readTree(raw);
            } catch (IOException e) {
                // 如果不是 JSON，则把整个文本视为单页文本
                ArrayNode arr = MAPPER.createArrayNode();
                arr.add(raw);
                rootNode = arr;
            }

            if (!rootNode.isArray()) {
                ArrayNode arr = MAPPER.createArrayNode();
                arr.add(rootNode);
                rootNode = arr;
            }

            upsertPages(dbUrl, dbUser, dbPass, pptName, (ArrayNode) rootNode, inputPath);
            System.out.println("按页保存完成: " + pptName);
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
            System.exit(10);
        }
    }

    private static void upsertPages(String url, String user, String pass, String pptName, ArrayNode pages, String sourceFile) throws SQLException {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
        } catch (ClassNotFoundException ignored) {
        }

        String insertSql = "INSERT INTO ppt_page (ppt_name, page_number, page_json, page_text) VALUES (?,?,?,?)";
        String updateSql = "UPDATE ppt_page SET page_json = ?, page_text = ?, updated_at = CURRENT_TIMESTAMP WHERE ppt_name = ? AND page_number = ?";

        try (Connection conn = DriverManager.getConnection(url, user, pass)) {
            conn.setAutoCommit(false);

            for (int i = 0; i < pages.size(); i++) {
                JsonNode pageNode = pages.get(i);
                int pageNumber = extractPageNumber(pageNode, i + 1);

                String pageJsonStr = pageNode.isTextual() ? null : pageNode.toString();
                String pageText = extractPageText(pageNode);

                try (PreparedStatement ins = conn.prepareStatement(insertSql)) {
                    ins.setString(1, pptName);
                    ins.setInt(2, pageNumber);
                    if (pageJsonStr != null) ins.setString(3, pageJsonStr); else ins.setNull(3, Types.VARCHAR);
                    ins.setString(4, pageText);
                    ins.executeUpdate();
                    conn.commit();
                } catch (SQLException ex) {
                    conn.rollback();
                    // 尝试更新
                    try (PreparedStatement upd = conn.prepareStatement(updateSql)) {
                        if (pageJsonStr != null) upd.setString(1, pageJsonStr); else upd.setNull(1, Types.VARCHAR);
                        upd.setString(2, pageText);
                        upd.setString(3, pptName);
                        upd.setInt(4, pageNumber);
                        int updated = upd.executeUpdate();
                        conn.commit();
                        if (updated == 0) {
                            throw new SQLException("插入与更新均失败: ppt=" + pptName + " page=" + pageNumber);
                        }
                    }
                }
            }
        }
    }

    private static int extractPageNumber(JsonNode pageNode, int defaultNum) {
        if (pageNode == null) return defaultNum;
        if (pageNode.has("page") && pageNode.get("page").canConvertToInt()) {
            return pageNode.get("page").asInt();
        }
        return defaultNum;
    }

    private static String extractPageText(JsonNode pageNode) {
        if (pageNode == null) return "";
        if (pageNode.isTextual()) return pageNode.asText();
        // 尝试从常见字段组合摘要
        StringBuilder sb = new StringBuilder();
        if (pageNode.has("title")) sb.append(pageNode.get("title").asText()).append(" ");
        if (pageNode.has("body")) sb.append(pageNode.get("body").asText()).append(" ");
        if (sb.length() > 0) return sb.toString().trim();
        // fallback: 返回整个 JSON 的紧凑字符串
        return pageNode.toString();
    }
}
