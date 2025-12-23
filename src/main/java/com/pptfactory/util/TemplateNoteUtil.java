package com.pptfactory.util;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class TemplateNoteUtil {
    public static void main(String[] args) throws Exception {
        extractAllNotesToOneJson("/Users/menggl/workspace/PPTFactory/templates/type_purchase/pages/模板_master_template.pptx", "/Users/menggl/workspace/PPTFactory/templates/type_purchase/pages/模板_master_template.json");
    }
    public static void extractAllNotesToOneJson(String pptxPath, String outputJsonPath) throws Exception {
        Map<Integer, String> notes = PPTXNoteUtil.getSlideNotes(pptxPath);
        ObjectMapper mapper = new ObjectMapper();
        List<JsonNode> merged = new ArrayList<>();
        for (Map.Entry<Integer, String> e : notes.entrySet()) {
            int pageIndex = e.getKey();
            String raw = e.getValue();
            if (raw == null || raw.trim().isEmpty()) {
                continue;
            }
            JsonNode node = mapper.readTree(raw);
            if (node.isArray()) {
                for (JsonNode item : node) {
                    if (item instanceof ObjectNode) {
                        ((ObjectNode) item).put("page_index", pageIndex);
                        ((ObjectNode) item).put("template_id", String.format("T%03d", pageIndex));
                    }
                    merged.add(item);
                }
            } else if (node instanceof ObjectNode) {
                ObjectNode obj = (ObjectNode) node;
                obj.put("page_index", pageIndex);
                obj.put("template_id", String.format("T%03d", pageIndex));
                merged.add(obj);
            } else {
                continue;
            }
        }
        ArrayNode arrayNode = mapper.createArrayNode();
        arrayNode.addAll(merged);
        Path outPath = Paths.get(outputJsonPath).toAbsolutePath().normalize();
        Path parent = outPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        try (OutputStream os = new FileOutputStream(outPath.toFile())) {
            os.write(mapper.writerWithDefaultPrettyPrinter().writeValueAsString(arrayNode).getBytes(StandardCharsets.UTF_8));
        }
    }
}
