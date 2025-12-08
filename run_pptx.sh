#!/bin/bash
# 运行打包后的 JAR 文件生成 PPTX

# 检查 Aspose JAR 是否存在（打包时需要）
if [ ! -f "lib/aspose-slides-25.11-jdk16.jar" ]; then
    echo "错误：找不到 Aspose.Slides JAR 文件"
    echo "请将 aspose-slides-25.11-jdk16.jar 下载到 lib/ 文件夹"
    echo "下载地址：https://products.aspose.com/slides/java"
    exit 1
fi

# 检查 JAR 文件是否存在，如果不存在则打包
JAR_FILE="target/ppt-template-engine-1.0.0-jar-with-dependencies.jar"
if [ ! -f "$JAR_FILE" ]; then
    echo "正在打包项目（包含所有依赖）..."
    mvn package -DskipTests
    if [ $? -ne 0 ]; then
        echo "打包失败，请检查错误信息"
        exit 1
    fi
    echo "✓ 打包完成"
fi

# 运行生成命令
# 注意：system scope 的依赖不会被打包到 JAR 中，需要手动添加到 classpath
# 设置区域设置为 en-US，避免 Aspose.Slides 不支持的系统区域设置问题
# 使用 LC_ALL 环境变量强制设置区域设置，并清除所有区域设置相关的环境变量
echo "正在生成 PPTX..."
export LC_ALL=en_US.UTF-8
export LANG=en_US.UTF-8
export LANGUAGE=en_US
unset LC_CTYPE
unset LC_NUMERIC
unset LC_TIME
unset LC_COLLATE
unset LC_MONETARY
unset LC_MESSAGES
unset LC_PAPER
unset LC_NAME
unset LC_ADDRESS
unset LC_TELEPHONE
unset LC_MEASUREMENT
unset LC_IDENTIFICATION

java -Duser.language=en \
    -Duser.country=US \
    -Duser.variant="" \
    -Duser.script="" \
    -Duser.extensions="" \
    -Djava.locale.providers=COMPAT \
    -Dfile.encoding=UTF-8 \
    -cp "$JAR_FILE:lib/aspose-slides-25.11-jdk16.jar" \
    com.pptfactory.cli.GeneratePPT \
    examples/safety_slides_extended.json \
    -o output.pptx \
    --style safety \
    --template safety

if [ $? -eq 0 ]; then
    echo "✓ PPTX 生成成功: output.pptx"
else
    echo "✗ 生成失败，请检查错误信息"
    exit 1
fi

