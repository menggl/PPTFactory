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
import java.util.Optional;

/**
 * 将 produce/ppt内容页.txt 中的数据保存到 MySQL 表 ppt_contents
 * Usage:
 *   java -cp <classpath-with-mysql-connector> com.pptfactory.util.SavePptPagesToMySQL <ppt_name> [path/to/produce/ppt内容页.txt]
 *
 * DB 配置通过环境变量：DB_URL, DB_USER, DB_PASS
 * 示例 DB_URL: jdbc:mysql://localhost:3306/pptfactory?useUnicode=true&characterEncoding=utf8mb4&serverTimezone=UTC
 */
public class SavePptPagesToMySQL {
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

            // 尝试解析为 JSON，如果不是有效 JSON，则把 raw_text 存入并将 contents 置为 [raw]
            JsonNode contentsNode;
            try {
                JsonNode parsed = MAPPER.readTree(raw);
                // 如果 parsed 是对象或数组，则直接使用；否则把其包装成数组
                if (parsed.isArray()) {
                    contentsNode = parsed;
                } else {
                    ArrayNode arr = MAPPER.createArrayNode();
                    arr.add(parsed);
                    contentsNode = arr;
                }
            } catch (IOException e) {
                // 不是 JSON，包装为数组，保存原始文本
                ArrayNode arr = MAPPER.createArrayNode();
                arr.add(raw);
                contentsNode = arr;
            }

            int pageCount = contentsNode.isArray() ? contentsNode.size() : 1;

            upsertToDb(dbUrl, dbUser, dbPass, pptName, contentsNode.toString(), raw, pageCount, inputPath);

            System.out.println("保存成功: " + pptName);
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
            System.exit(10);
        }
    }

    private static void upsertToDb(String url, String user, String pass,
                                   String pptName, String contentsJson, String rawText,
                                   int pageCount, String sourceFile) throws SQLException {
        // Ensure MySQL driver loaded (if available on classpath)
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
        } catch (ClassNotFoundException ignored) {
        }

        String insertSql = "INSERT INTO ppt_contents (ppt_name, contents, raw_text, page_count, source_file) VALUES (?,?,?,?,?)";
        String updateSql = "UPDATE ppt_contents SET contents = ?, raw_text = ?, page_count = ?, source_file = ?, updated_at = CURRENT_TIMESTAMP WHERE ppt_name = ?";

        try (Connection conn = DriverManager.getConnection(url, user, pass)) {
            conn.setAutoCommit(false);
            // Try insert
            try (PreparedStatement ins = conn.prepareStatement(insertSql)) {
                ins.setString(1, pptName);
                ins.setString(2, contentsJson);
                ins.setString(3, rawText);
                ins.setInt(4, pageCount);
                ins.setString(5, sourceFile);
                ins.executeUpdate();
                conn.commit();
                return;
            } catch (SQLException ex) {
                // 如果插入失败（例如主键冲突），执行更新
                conn.rollback();
                try (PreparedStatement upd = conn.prepareStatement(updateSql)) {
                    upd.setString(1, contentsJson);
                    upd.setString(2, rawText);
                    upd.setInt(3, pageCount);
                    upd.setString(4, sourceFile);
                    upd.setString(5, pptName);
                    int updated = upd.executeUpdate();
                    conn.commit();
                    if (updated == 0) {
                        throw new SQLException("更新失败且未插入新行");
                    }
                }
            }
        }
    }
}
