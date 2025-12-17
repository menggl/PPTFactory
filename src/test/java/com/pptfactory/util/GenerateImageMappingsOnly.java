package com.pptfactory.util;

/**
 * 仅生成图片映射（不重新生成PPT）
 */
public class GenerateImageMappingsOnly {
    
    public static void main(String[] args) {
        try {
            // 找到最新的PPT文件
            String produceDir = "produce";
            java.io.File dir = new java.io.File(produceDir);
            java.io.File[] files = dir.listFiles((d, name) -> name.startsWith("new_ppt_") && name.endsWith(".pptx"));
            
            if (files == null || files.length == 0) {
                System.out.println("未找到PPT文件");
                return;
            }
            
            // 按文件名排序，获取最新的
            java.io.File latestFile = null;
            for (java.io.File f : files) {
                if (latestFile == null || f.getName().compareTo(latestFile.getName()) > 0) {
                    latestFile = f;
                }
            }
            
            System.out.println("=== 生成图片映射 ===");
            System.out.println("PPT文件: " + latestFile.getAbsolutePath());
            System.out.println();
            
            // 调用生成图片映射的方法
            ProduceUtil.generateImageMappings(latestFile.getAbsolutePath());
            
            System.out.println("\n✓ 图片映射生成完成！");
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
