# lib 文件夹说明

此文件夹用于存放本地 JAR 依赖文件。

## Aspose.Slides

**重要**：项目已迁移到使用 Aspose.Slides 生成 PPTX，必须下载 JAR 文件才能运行。

### 需要的文件

- `aspose-slides-25.11-jdk16.jar` （当前配置的版本）

### 下载方式

1. 访问 Aspose 官网：https://products.aspose.com/slides/java
2. 下载对应 JDK 16 版本的 JAR 文件（版本 25.11）
3. 将 JAR 文件重命名为：`aspose-slides-25.11-jdk16.jar`
4. 放置到此 `lib/` 文件夹中

### 注意事项

- Aspose.Slides 是商业库，需要许可证
- 如果没有许可证，会有水印限制
- 文件名必须与 `pom.xml` 中配置的路径一致
- 如果使用不同版本，需要同时更新 `pom.xml` 中的配置

### 生成 PPTX

下载 JAR 文件后，运行：

```bash
./generate_pptx.sh
```

或者使用 Maven：

```bash
mvn exec:java -Dexec.mainClass="com.pptfactory.cli.GeneratePPT" \
    -Dexec.args="examples/example_slides.json -o test_output.pptx --style default --template safety"
```

