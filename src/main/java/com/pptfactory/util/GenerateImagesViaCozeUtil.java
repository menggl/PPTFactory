package com.pptfactory.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Duration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Utility to read produce/ppt内容映射.txt, call a Coze image-generation workflow for each
 * 图片提示词准备 entry and write back 图片链接映射 (label -> url) into the mapping file.
 *
 * Environment:
 * - COZE_ENDPOINT: HTTP endpoint to POST prompts (default: http://localhost:8000/generate)
 * - COZE_API_KEY: optional API key sent as Authorization: Bearer <key>
 */
public class GenerateImagesViaCozeUtil {
    private static final ObjectMapper M = new ObjectMapper();
    private static final HttpClient CLIENT = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(10))
            .build();

    public static void main(String[] args) throws Exception {
        String mappingPath = "produce/ppt内容映射.txt";
        boolean dryRun = false;
        String cliToken = null;
        boolean useConfigToken = false;
        for (int i = 0; i < args.length; i++) {
            String a = args[i];
            if ("--dry-run".equals(a) || "-n".equals(a)) {
                dryRun = true;
            } else if ("--token".equals(a) || "-t".equals(a)) {
                if (i + 1 < args.length) {
                    cliToken = args[++i];
                } else {
                    System.out.println("--token 需要一个值，请在参数后指定 token。");
                    return;
                }
            } else if ("--use-config-token".equals(a) || "-c".equals(a)) {
                useConfigToken = true;
            } else if (!a.startsWith("-")) {
                mappingPath = a;
            }
        }
        // Determine token to use. Precedence: CLI token > --use-config-token (forces DEFAULT_TOKEN) > CozeConfig.TOKEN
        String tokenToUse;
        if (cliToken != null) tokenToUse = cliToken;
        else if (useConfigToken) tokenToUse = CozeConfig.DEFAULT_TOKEN;
        else tokenToUse = CozeConfig.TOKEN;
        boolean verbose = false;
        for (String a : args) if ("--verbose".equals(a) || "-v".equals(a)) verbose = true;
        generateImages(mappingPath, CozeConfig.ENDPOINT, tokenToUse, dryRun, verbose);
    }

    public static void generateImages(String mappingFilePath, String cozeEndpoint, String apiKey, boolean dryRun, boolean verbose) throws Exception {
        Path p = Path.of(mappingFilePath);
        if (!Files.exists(p)) {
            System.out.println("映射文件不存在: " + mappingFilePath);
            return;
        }

        String raw = Files.readString(p, StandardCharsets.UTF_8);
        int start = raw.indexOf('[');
        if (start < 0) {
            System.out.println("未在映射文件中找到 JSON 数组。");
            return;
        }
        int count = 0, end = -1;
        for (int i = start; i < raw.length(); i++) {
            char c = raw.charAt(i);
            if (c == '[') count++;
            else if (c == ']') count--;
            if (count == 0) { end = i; break; }
        }
        if (end < 0) {
            System.out.println("无法解析映射文件中的 JSON 数组范围。");
            return;
        }

        String jsonArr = raw.substring(start, end + 1);
        List<Map<String, Object>> mappings = M.readValue(jsonArr, new TypeReference<>() {});

        boolean changed = false;
        for (int idx = 0; idx < mappings.size(); idx++) {
            Map<String, Object> entry = mappings.get(idx);
            Object picPrepObj = entry.get("图片提示词准备");
            if (!(picPrepObj instanceof Map)) continue;
            @SuppressWarnings("unchecked")
            Map<String, Object> picPrep = (Map<String, Object>) picPrepObj;

            @SuppressWarnings("unchecked")
            Map<String, String> existMap = (Map<String, String>) entry.get("图片链接映射");
            Map<String, String> urlMap = existMap != null ? new HashMap<>(existMap) : new HashMap<>();

            for (Map.Entry<String, Object> kv : picPrep.entrySet()) {
                String label = kv.getKey();
                String prompt = kv.getValue() == null ? "" : String.valueOf(kv.getValue());
                if (prompt.isBlank()) continue;
                if (urlMap.containsKey(label) && urlMap.get(label) != null && !urlMap.get(label).isBlank()) {
                    System.out.println("已存在图片链接，跳过: " + label);
                    continue;
                }
                System.out.println("生成图片：mappingIndex=" + (idx+1) + " label=" + label);
                String workflowId = CozeConfig.IMAGE_WORKFLOW_ID;
                String token = CozeConfig.TOKEN == null || CozeConfig.TOKEN.isBlank() ? (apiKey == null ? "" : apiKey) : CozeConfig.TOKEN;
                int type = CozeConfig.TYPE;
                if (dryRun) {
                    // 构造并打印将要发送的 payload，但不实际发送 HTTP 请求（安全干跑）
                    Map<String, Object> payload = new HashMap<>();
                    if (workflowId != null && !workflowId.isBlank()) {
                        payload.put("workflow_id", String.valueOf(workflowId));
                        payload.put("is_async", false);
                        Map<String, Object> parameters = new HashMap<>();
                        parameters.put("main_title", prompt);
                        parameters.put("type", type);
                        payload.put("parameters", parameters);
                    } else {
                        payload.put("main_title", prompt);
                        payload.put("type", type);
                    }
                    String body = M.writeValueAsString(payload);
                    System.out.println("DRY-RUN -> endpoint: " + cozeEndpoint);
                    System.out.println("DRY-RUN -> body: " + body);
                    System.out.println("DRY-RUN -> Authorization: " + (token == null || token.isBlank() ? "<no-token>" : CozeConfig.AUTH_PREFIX + "****"));
                } else {
                    String url = callCozeForImage(prompt, cozeEndpoint, token, workflowId, type, verbose);
                    if (url != null) {
                        urlMap.put(label, url);
                        changed = true;
                        System.out.println("  -> 获得图片链接: " + url);
                    } else {
                        System.out.println("  -> 未获得图片链接: " + label);
                    }
                    // 简单节流：小延迟，避免短时间内大量请求
                    try { Thread.sleep(CozeConfig.REQUEST_SLEEP_MS); } catch (InterruptedException ignored) {}
                }
            }

            if (!urlMap.isEmpty()) {
                entry.put("图片链接映射", urlMap);
            }
        }

        if (changed) {
            String newJson = M.writerWithDefaultPrettyPrinter().writeValueAsString(mappings);
            String out = raw.substring(0, start) + newJson + raw.substring(end + 1);
            Files.writeString(p, out, StandardCharsets.UTF_8, StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.CREATE);
            System.out.println("已更新映射文件并写回 图片链接映射 字段: " + mappingFilePath);
        } else {
            System.out.println("未检测到需要更新的图片链接。");
        }
    }

    /**
     * Convenience wrapper to generate a single image from a prompt using configured Coze settings.
     * Returns the image URL on success or null on failure.
     */
    public static String generateImageForPrompt(String prompt) {
        try {
            String endpoint = CozeConfig.ENDPOINT;
            String token = CozeConfig.TOKEN;
            String workflowId = CozeConfig.IMAGE_WORKFLOW_ID;
            int type = CozeConfig.TYPE;
            return callCozeForImage(prompt, endpoint, token, workflowId, type);
        } catch (Exception e) {
            System.out.println("生成图片失败: " + e.getMessage());
            return null;
        }
    }

    private static String callCozeForImage(String prompt, String endpoint, String token, String workflowId, int type, boolean verbose) {
        // Try with simple retries for transient errors
        int maxRetries = CozeConfig.MAX_RETRIES;
        int attempt = 0;
        while (attempt < maxRetries) {
            attempt++;
            try {
                Map<String, Object> payload = new HashMap<>();
                // For Coze run API, payload uses workflow_id + is_async + parameters
                if (workflowId != null && !workflowId.isBlank()) {
                    payload.put("workflow_id", String.valueOf(workflowId));
                    payload.put("is_async", false);
                    Map<String, Object> parameters = new HashMap<>();
                    parameters.put("main_title", prompt);
                    parameters.put("type", type);
                    payload.put("parameters", parameters);
                } else {
                    // fallback: keep previous behavior
                    payload.put("main_title", prompt);
                    payload.put("type", type);
                }

                String body = M.writeValueAsString(payload);
                // 打印请求体用于调试（如果需要可控制为干跑）
                System.out.println("请求 endpoint: " + endpoint);
                System.out.println("请求 body: " + body);
                HttpRequest.Builder reqb = HttpRequest.newBuilder()
                        .uri(URI.create(endpoint))
                        .timeout(Duration.ofSeconds(CozeConfig.TIMEOUT_SECONDS))
                        .header("Content-Type", "application/json")
                        .POST(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8));
                if (token != null && !token.isBlank()) {
                    reqb.header(CozeConfig.AUTH_HEADER, CozeConfig.AUTH_PREFIX + token);
                }
                HttpRequest req = reqb.build();
                HttpResponse<String> resp = CLIENT.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8));
                if (resp.statusCode() == 429 || resp.statusCode() >= 500) {
                    System.out.println("Coze 请求遇到临时错误 (status=" + resp.statusCode() + "), 重试: " + attempt);
                    Thread.sleep((long)CozeConfig.RETRY_BACKOFF_MS * attempt);
                    continue;
                }
                if (resp.statusCode() < 200 || resp.statusCode() >= 300) {
                    System.out.println("Coze 返回非 2xx 状态: " + resp.statusCode() + " -> " + resp.body());
                    System.out.println("使用的 token: " + (token == null ? "<none>" : token));
                    return null;
                }

                if (verbose) {
                    System.out.println("Coze 响应 (attempt=" + attempt + ", status=" + resp.statusCode() + "): " + resp.body());
                }

                JsonNode root = M.readTree(resp.body());

                // First: some Coze responses are simple wrappers with data as a JSON string:
                // {"data":"{\"output\":\"https://...\"}", ...}
                if (root.has("data") && root.get("data").isTextual()) {
                    String inner = root.get("data").asText();
                    try {
                        JsonNode innerNode = M.readTree(inner);
                        if (innerNode.has("output")) return innerNode.get("output").asText();
                        String foundInner = findFirstUrlInJson(innerNode);
                        if (foundInner != null) return foundInner;
                    } catch (Exception ignored) {
                        if (looksLikeUrl(inner)) return inner;
                    }
                }

                // prefer run-style response: {"code":0, "data": ... }
                if (root.has("code") && root.get("code").isInt() && root.get("code").asInt() == 0) {
                    JsonNode data = root.get("data");
                    if (data != null) {
                        // data may be a JSON string like '{"output":"..."}'
                        if (data.isTextual()) {
                            String inner = data.asText();
                            try {
                                JsonNode innerNode = M.readTree(inner);
                                if (innerNode.has("output")) return innerNode.get("output").asText();
                                String found = findFirstUrlInJson(innerNode);
                                if (found != null) return found;
                            } catch (Exception ignored) {
                                // not a JSON string, try to see if text itself looks like URL
                                if (looksLikeUrl(inner)) return inner;
                            }
                        } else if (data.isObject()) {
                            if (data.has("output")) return data.get("output").asText();
                            String found = findFirstUrlInJson(data);
                            if (found != null) return found;
                        }
                    }
                }

                // fallback: previous heuristics
                // first, try to find explicit result.urls (workflow wrapper)
                if (root.has("result") && root.get("result").has("urls")) {
                    JsonNode urls = root.get("result").get("urls");
                    if (urls.isArray() && urls.size() > 0) return urls.get(0).asText();
                }
                if (root.has("outputs") && root.get("outputs").isArray() && root.get("outputs").size() > 0) {
                    JsonNode first = root.get("outputs").get(0);
                    String found = findFirstUrlInJson(first);
                    if (found != null) return found;
                }
                if (root.has("data") && root.get("data").isArray() && root.get("data").size() > 0) {
                    String found = findFirstUrlInJson(root.get("data"));
                    if (found != null) return found;
                }
                String found = findFirstUrlInJson(root);
                if (found != null) return found;

                System.out.println("无法从 Coze 响应中解析图片链接（尝试过常见字段），响应体: " + resp.body());
                return null;
            } catch (IOException | InterruptedException e) {
                System.out.println("调用 Coze 出错: " + e.getMessage() + "，attempt=" + attempt);
                try { Thread.sleep((long)CozeConfig.RETRY_BACKOFF_MS * attempt); } catch (InterruptedException ignored) {}
            }
        }
        return null;
    }

    private static String findFirstUrlInJson(JsonNode node) {
        if (node == null) return null;
        if (node.isTextual()) {
            String s = node.asText();
            if (looksLikeUrl(s)) return s;
        }
        if (node.isArray()) {
            for (JsonNode n : node) {
                String r = findFirstUrlInJson(n);
                if (r != null) return r;
            }
        }
        if (node.isObject()) {
            for (var it = node.fields(); it.hasNext(); ) {
                var e = it.next();
                String r = findFirstUrlInJson(e.getValue());
                if (r != null) return r;
            }
        }
        return null;
    }

    private static boolean looksLikeUrl(String s) {
        if (s == null) return false;
        s = s.trim();
        if (s.startsWith("http://") || s.startsWith("https://")) return true;
        // also accept base64 data URLs
        if (s.startsWith("data:image/")) return true;
        // quick heuristic for image file extension
        String sl = s.toLowerCase();
        if (sl.contains(".png") || sl.contains(".jpg") || sl.contains(".jpeg") || sl.contains(".gif")) return true;
        return false;
    }
}
