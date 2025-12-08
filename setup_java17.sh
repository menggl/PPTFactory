#!/bin/bash
# 配置 Java 17 环境脚本

echo "配置 Java 17 环境..."

# 设置 PATH，让 openjdk@17 优先
export PATH="/usr/local/opt/openjdk@17/bin:$PATH"

# 设置 JAVA_HOME
export JAVA_HOME="/usr/local/opt/openjdk@17"

# 验证 Java 版本
echo ""
echo "当前 Java 版本："
java -version

echo ""
echo "JAVA_HOME: $JAVA_HOME"

echo ""
echo "Maven 使用的 Java 版本："
mvn -version | head -3

echo ""
echo "配置完成！现在可以运行: mvn clean compile"

