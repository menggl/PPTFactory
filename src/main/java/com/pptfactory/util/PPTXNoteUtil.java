package com.pptfactory.util;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class PPTXNoteUtil {
    public static void main(String[] args) throws IOException {
        Map<Integer, String> map = getSlideNotes("/Users/menggl/workspace/PPTFactory/templates/type_purchase/pages/模板_master_template.pptx");
        for (Integer keySet : map.keySet()) {
            System.out.println("---------------"+keySet+"-----------------");
            System.out.println(map.get(keySet));
        }
    }
    public static Map<Integer, String> getSlideNotes(String pptxPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(pptxPath);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {
            Map<Integer, String> notesMap = new LinkedHashMap<>();
            int index = 0;
            for (XSLFSlide slide : ppt.getSlides()) {
                index++;
                String text = extractNotesText(slide);
                notesMap.put(index, text == null ? "" : text);
            }
            return notesMap;
        }
    }

    public static List<String> getNotesList(String pptxPath) throws IOException {
        Map<Integer, String> map = getSlideNotes(pptxPath);
        return new ArrayList<>(map.values());
    }

    private static String extractNotesText(XSLFSlide slide) {
        XSLFNotes notes = slide.getNotes();
        if (notes == null) return null;
        StringBuilder sb = new StringBuilder();
        for (XSLFShape shape : notes.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                String t = ((XSLFTextShape) shape).getText();
                if (t != null && !t.trim().isEmpty()) {
                    if (sb.length() > 0) sb.append("\n");
                    sb.append(t.trim());
                }
            }
        }
        String s = sb.toString().trim();
        return s.isEmpty() ? null : s;
    }
}
