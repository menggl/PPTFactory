# VS Code 调试配置说明

如果遇到 `ClassNotFoundException` 错误，请按照以下步骤操作：

## 1. 确保测试代码已编译

```bash
mvn clean test-compile
```

## 2. 重新加载 VS Code Java 项目

1. 在 VS Code 中按 `Cmd+Shift+P` (macOS) 或 `Ctrl+Shift+P` (Windows/Linux)
2. 输入 `Java: Clean Java Language Server Workspace`
3. 选择并执行该命令
4. 重启 VS Code

## 3. 验证编译输出

确保以下目录存在编译后的类文件：
- `target/classes/` - 主代码
- `target/test-classes/` - 测试代码

## 4. 手动运行测试（验证配置）

```bash
java -cp "target/classes:target/test-classes:lib/aspose-slides-25.11-jdk16.jar" \
     -Duser.language=en -Duser.country=US \
     com.pptfactory.template.engine.TestTemplateExtraction
```

如果命令行运行正常，说明代码没问题，只需要重新加载VS Code Java项目。

## 5. 检查 Java 扩展

确保安装了以下 VS Code 扩展：
- **Extension Pack for Java** (Microsoft)
- 或者至少安装 **Language Support for Java by Red Hat**

## 6. 检查工作区设置

确保 `.vscode/settings.json` 中包含正确的 Java 路径：
- `java.home` 指向 Java 17 安装路径
- 项目路径配置正确

## 7. 如果仍然无法调试

### 方法A：使用 Maven 运行测试

在 VS Code 终端中运行：

```bash
mvn exec:java -Dexec.mainClass="com.pptfactory.template.engine.TestTemplateExtraction" \
              -Dexec.classpathScope=test
```

### 方法B：直接使用 Java 命令调试

在 VS Code 中创建自定义 launch 配置，使用外部终端：

```json
{
    "type": "java",
    "name": "TestTemplateExtraction (External Terminal)",
    "request": "launch",
    "mainClass": "com.pptfactory.template.engine.TestTemplateExtraction",
    "cwd": "${workspaceFolder}",
    "vmArgs": "-Duser.language=en -Duser.country=US",
    "console": "externalTerminal"
}
```

## 常见问题

### Q: 仍然显示 ClassNotFoundException

**A:** 尝试：
1. 删除 `.classpath` 和 `.project` 文件（如果存在）
2. 重新执行 `Java: Clean Java Language Server Workspace`
3. 重新编译：`mvn clean test-compile`

### Q: 找不到 Aspose.Slides 类

**A:** 确保 `lib/aspose-slides-25.11-jdk16.jar` 文件存在，并且 VS Code 能够识别它。
可以在 `.vscode/settings.json` 中手动添加：

```json
{
    "java.classPaths": [
        "lib/aspose-slides-25.11-jdk16.jar"
    ]
}
```

### Q: 调试时断点不生效

**A:** 
1. 确保编译时包含调试信息：`mvn clean test-compile -Dmaven.compiler.debug=true`
2. 检查断点是否设置在正确的行号
3. 确保使用 Debug 模式运行，而不是 Run 模式
