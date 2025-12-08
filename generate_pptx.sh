#!/bin/bash
# PPTX 生成脚本

# 检查 Aspose.Slides JAR 文件是否存在
if [ ! -f "lib/aspose-slides-25.11-jdk16.jar" ]; then
    echo "错误：找不到 Aspose.Slides JAR 文件"
    echo "请将 aspose-slides-25.11-jdk16.jar 下载到 lib/ 文件夹"
    echo "下载地址：https://products.aspose.com/slides/java"
    exit 1
fi

# 编译项目
echo "正在编译项目..."
mvn clean compile

if [ $? -ne 0 ]; then
    echo "编译失败，请检查错误信息"
    exit 1
fi

# 运行生成命令
# 设置区域设置为 en-US，避免 Aspose.Slides 不支持的系统区域设置问题
echo "正在生成 PPTX..."
mvn exec:java -Dexec.mainClass="com.pptfactory.cli.GeneratePPT" \
    -Dexec.args="examples/example_slides.json -o test_output.pptx --style default --template safety" \
    -Dexec.classpathScope=compile \
    -Duser.language=en -Duser.country=US

echo "完成！"

