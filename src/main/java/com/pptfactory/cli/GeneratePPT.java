package com.pptfactory.cli;

import com.pptfactory.template.engine.PPTTemplateEngine;
import com.pptfactory.style.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;

import java.io.File;
import java.util.Locale;
import java.util.Map;

/**
 * PPT生成应用入口
 * 
 * 该模块是PPT生成应用的主入口，整合了模板引擎、风格策略、大模型等功能。
 * 
 * 使用示例：
 * java -cp target/ppt-template-engine-1.0.0-jar-with-dependencies.jar \
 *      com.pptfactory.cli.GeneratePPT \
 *      input.json \
 *      -o output.pptx \
 *      --style safety \
 *      --template safety
 */
public class GeneratePPT {
    
    // 静态初始化块：在类加载时设置区域设置，确保在任何 Aspose.Slides 类加载之前设置
    static {
        Locale.setDefault(Locale.US);
    }
    
    /**
     * 根据风格名称获取风格策略对象
     * 
     * @param styleName 风格名称（default, chinese, math, finance, safety）
     * @return 风格策略对象
     */
    private static StyleStrategy getStyleStrategy(String styleName) {
        switch (styleName.toLowerCase()) {
            case "chinese":
                return new ChineseStyle();
            case "math":
                return new MathStyle();
            case "finance":
                return new FinanceStyle();
            case "safety":
                return new SafetyStyle();
            case "default":
            default:
                return new DefaultStyle();
        }
    }
    
    /**
     * 根据模板名称获取模板文件路径
     * 
     * @param templateName 模板名称（chinese, math, finance, safety）
     * @return 模板文件路径
     */
    private static String getTemplateFile(String templateName) {
        // 如果提供了完整路径，直接返回
        File template = new File(templateName);
        if (template.exists()) {
            return templateName;
        }
        
        // 否则在templates目录下查找
        String templatePath = "templates/" + templateName.toLowerCase() + "/theme.pptx";
        File templateFile = new File(templatePath);
        if (templateFile.exists()) {
            return templatePath;
        }
        
        // 如果找不到，返回null（将使用默认模板）
        return null;
    }
    
    /**
     * 主函数：PPT生成应用入口
     * 
     * @param args 命令行参数
     */
    public static void main(String[] args) {
        // 设置默认区域设置为 en-US，避免 Aspose.Slides 不支持的系统区域设置问题
        // 这必须在任何 Aspose.Slides 类加载之前设置
        Locale.setDefault(Locale.US);
        
        if (args.length == 0) {
            printUsage();
            System.exit(1);
        }
        
        try {
            // 解析命令行参数
            String inputFile = null;
            String outputFile = "output.pptx";
            String style = "default";
            String template = "default";
            
            for (int i = 0; i < args.length; i++) {
                String arg = args[i];
                if (arg.equals("-o") || arg.equals("--output")) {
                    if (i + 1 < args.length) {
                        outputFile = args[++i];
                    }
                } else if (arg.equals("--style")) {
                    if (i + 1 < args.length) {
                        style = args[++i];
                    }
                } else if (arg.equals("--template")) {
                    if (i + 1 < args.length) {
                        template = args[++i];
                    }
                } else if (!arg.startsWith("-")) {
                    inputFile = arg;
                }
            }
            
            if (inputFile == null) {
                System.err.println("错误：必须指定输入文件");
                printUsage();
                System.exit(1);
            }
            
            // 加载输入数据
            System.out.println("正在加载输入数据: " + inputFile);
            ObjectMapper mapper = new ObjectMapper();
            Map<String, Object> slidesData = mapper.readValue(
                new File(inputFile),
                new TypeReference<Map<String, Object>>() {}
            );
            System.out.println("✓ 输入数据加载成功");
            
            // 获取风格策略
            StyleStrategy styleStrategy = getStyleStrategy(style);
            System.out.println("✓ 使用风格策略: " + style);
            
            // 获取模板文件
            String templateFile = getTemplateFile(template);
            if (templateFile == null) {
                System.err.println("警告：找不到模板文件，将使用默认模板");
                // 创建一个默认的模板文件路径或使用内置默认模板
                templateFile = "templates/default/theme.pptx";
            }
            System.out.println("✓ 使用模板文件: " + templateFile);
            
            // 创建引擎
            PPTTemplateEngine engine = new PPTTemplateEngine(templateFile, styleStrategy);
            System.out.println("✓ 模板引擎创建成功");
            
            // 渲染PPT
            System.out.println("正在渲染PPT...");
            engine.renderFromJson(slidesData);
            System.out.println("✓ PPT渲染完成");
            
            // 保存PPT
            System.out.println("正在保存PPT: " + outputFile);
            engine.save(outputFile);
            System.out.println("✓ PPT保存成功: " + outputFile);
            
            // 关闭引擎
            engine.close();
            
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    /**
     * 打印使用说明
     */
    private static void printUsage() {
        System.out.println("PPT生成应用 - 使用大模型和模板引擎生成PPT");
        System.out.println();
        System.out.println("用法:");
        System.out.println("  java -cp <jar-file> com.pptfactory.cli.GeneratePPT <input.json> [选项]");
        System.out.println();
        System.out.println("参数:");
        System.out.println("  input.json              输入文件路径（JSON格式的slides数据）");
        System.out.println();
        System.out.println("选项:");
        System.out.println("  -o, --output <file>     输出的PPT文件名（默认: output.pptx）");
        System.out.println("  --style <style>         风格选择（default, chinese, math, finance, safety）");
        System.out.println("  --template <template>   模板选择（chinese, math, finance, safety）");
        System.out.println();
        System.out.println("示例:");
        System.out.println("  java -cp target/ppt-template-engine-1.0.0-jar-with-dependencies.jar \\");
        System.out.println("       com.pptfactory.cli.GeneratePPT \\");
        System.out.println("       examples/safety_slides_extended.json \\");
        System.out.println("       -o safety_demo_final.pptx \\");
        System.out.println("       --style safety \\");
        System.out.println("       --template safety");
    }
}

