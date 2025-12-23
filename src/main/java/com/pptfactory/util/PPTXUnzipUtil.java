package com.pptfactory.util;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public final class PPTXUnzipUtil {
    public static void main(String[] args) throws IOException {
        unzip("/Users/menggl/workspace/PPTFactory/templates/type_purchase/pages/模板_master_template.pptx",
            "/Users/menggl/workspace/PPTFactory/templates/type_purchase/pages/模板_master_template");
    }
    private PPTXUnzipUtil() {}

    public static void unzip(String pptxPath, String outputDir) throws IOException {
        Path source = Paths.get(pptxPath).toAbsolutePath().normalize();
        Path targetDir = Paths.get(outputDir).toAbsolutePath().normalize();
        Files.createDirectories(targetDir);
        try (InputStream is = Files.newInputStream(source);
             ZipInputStream zis = new ZipInputStream(is)) {
            ZipEntry entry;
            byte[] buffer = new byte[8192];
            while ((entry = zis.getNextEntry()) != null) {
                String name = entry.getName();
                Path outPath = targetDir.resolve(name).normalize();
                if (!outPath.startsWith(targetDir)) {
                    zis.closeEntry();
                    continue;
                }
                if (entry.isDirectory()) {
                    Files.createDirectories(outPath);
                } else {
                    Path parent = outPath.getParent();
                    if (parent != null) {
                        Files.createDirectories(parent);
                    }
                    try (var os = Files.newOutputStream(outPath)) {
                        int read;
                        while ((read = zis.read(buffer)) != -1) {
                            os.write(buffer, 0, read);
                        }
                    }
                }
                zis.closeEntry();
            }
        }
    }
}

