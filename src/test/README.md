# 测试代码使用说明

本目录包含所有功能模块的单步调试测试代码。每个测试类都可以独立运行，用于调试和验证对应功能模块的行为。

## 测试类列表

### 1. TestPPTTemplateEngine.java
**功能模块**: PPT模板引擎核心类  
**测试内容**:
- 模板引擎实例创建
- 模板文件加载
- 风格策略应用
- PPT渲染流程
- PPT保存

**使用方法**:
```bash
# 在IDE中打开文件，设置断点，以Debug模式运行
# 或使用命令行：
java -cp "target/classes:lib/aspose-slides-25.11-jdk16.jar" \
     com.pptfactory.template.engine.TestPPTTemplateEngine
```

### 2. TestStyleStrategy.java
**功能模块**: 风格策略类  
**测试内容**:
- 不同风格策略的创建（SafetyStyle, DefaultStyle等）
- 标题样式应用
- 副标题样式应用
- 正文样式应用
- 要点列表样式应用
- 字体大小和颜色设置

**使用方法**:
```bash
java -cp "target/classes:lib/aspose-slides-25.11-jdk16.jar" \
     com.pptfactory.style.TestStyleStrategy
```

### 3. TestWatermarkRemoval.java
**功能模块**: 水印移除功能  
**测试内容**:
- XML方式移除水印
- 水印检测
- 水印文本框删除
- 移除结果验证

**使用方法**:
```bash
java -cp "target/classes:lib/aspose-slides-25.11-jdk16.jar" \
     com.pptfactory.template.engine.TestWatermarkRemoval
```

### 4. TestTemplateExtraction.java
**功能模块**: 模板提取功能  
**测试内容**:
- 从源PPT提取模板幻灯片
- 模板文字替换（替换为"模板文字模板文字模板文字"）
- 模板图片替换（替换为"No Image"占位符）
- master_template.pptx生成

**使用方法**:
```bash
java -cp "target/classes:lib/aspose-slides-25.11-jdk16.jar" \
     com.pptfactory.template.engine.TestTemplateExtraction
```

### 5. TestContentGenerator.java
**功能模块**: 内容生成器（AI模块）  
**测试内容**:
- 内容生成器初始化
- 用户输入处理
- 结构化内容生成（JSON格式）

**使用方法**:
```bash
java -cp "target/classes" \
     com.pptfactory.ai.TestContentGenerator
```

### 6. TestLayoutClassifier.java
**功能模块**: 布局分类器（AI模块）  
**测试内容**:
- 布局分类器初始化
- 单个内容布局分类
- 批量幻灯片布局分类
- 不同布局类型的识别（标题页、内容页、图片+文字等）

**使用方法**:
```bash
java -cp "target/classes" \
     com.pptfactory.ai.TestLayoutClassifier
```

### 7. TestGeneratePPT.java
**功能模块**: CLI入口  
**测试内容**:
- 命令行参数解析
- 输入文件加载
- 风格策略选择
- 模板文件选择
- 完整的PPT生成流程

**使用方法**:
```bash
java -cp "target/classes:lib/aspose-slides-25.11-jdk16.jar" \
     com.pptfactory.cli.TestGeneratePPT
```

## 在IDE中使用

### IntelliJ IDEA / Eclipse

1. **打开测试文件**
   - 在IDE中打开对应的测试类（如 `TestPPTTemplateEngine.java`）

2. **设置断点**
   - 在需要调试的代码行左侧点击，设置断点
   - 建议在以下位置设置断点：
     - 方法入口处
     - 关键逻辑判断处
     - 数据转换处

3. **以Debug模式运行**
   - 右键点击测试类的 `main` 方法
   - 选择 "Debug 'TestXXX.main()'"
   - 或使用快捷键（IDEA: Shift+F9, Eclipse: F11）

4. **单步调试**
   - **Step Over (F8)**: 执行当前行，不进入方法内部
   - **Step Into (F7)**: 进入方法内部
   - **Step Out (Shift+F8)**: 跳出当前方法
   - **Resume (F9)**: 继续执行到下一个断点

5. **查看变量**
   - 在Debug窗口中查看变量值
   - 在代码中悬停鼠标查看变量值
   - 在Watch窗口中添加要监控的表达式

## 注意事项

1. **Java版本要求（重要）**
   - **项目要求Java 17或更高版本**
   - 如果IDE使用Java 8运行测试，会出现 `UnsupportedClassVersionError` 错误
   - **配置IDE使用Java 17**：
     - **IntelliJ IDEA**:
       1. File -> Project Structure -> Project
       2. 在 "SDK" 下拉菜单中选择 Java 17
       3. 在 "Language level" 中选择 17
       4. File -> Settings -> Build, Execution, Deployment -> Build Tools -> Maven -> Runner
       5. 在 "JRE" 中选择 Java 17
     - **Eclipse**:
       1. Window -> Preferences -> Java -> Installed JREs
       2. 点击 "Add..." 添加 Java 17
       3. 勾选 Java 17 作为默认JRE
       4. Window -> Preferences -> Java -> Compiler
       5. 将 "Compiler compliance level" 设置为 17
   - 所有测试类在启动时会自动检查Java版本，如果版本不符合要求会显示错误信息

2. **依赖文件**
   - 某些测试需要特定的文件存在（如 `master_template.pptx`、`1.2 安全生产方针政策.pptx`）
   - 如果文件不存在，测试会输出错误信息并退出

3. **输出文件**
   - 测试会生成输出文件（如 `test_output_engine.pptx`）
   - 这些文件用于验证功能是否正常工作
   - 可以手动打开这些文件检查结果

4. **环境设置**
   - 所有测试类都设置了 `Locale.setDefault(Locale.US)` 以避免Aspose.Slides的区域设置问题
   - 确保Java版本为17或更高

5. **Aspose.Slides依赖**
   - 使用Aspose.Slides的测试需要将JAR文件添加到classpath
   - 确保 `lib/aspose-slides-25.11-jdk16.jar` 文件存在

## 测试输出

每个测试类都会输出详细的执行日志，包括：
- 测试步骤说明
- 成功/失败标记（✓/✗）
- 关键数据信息
- 错误信息（如果有）

## 扩展测试

如果需要添加新的测试用例：

1. 在对应的测试类中添加新的测试方法
2. 在 `main` 方法中调用新的测试方法
3. 确保测试方法有清晰的输出说明
4. 在关键位置设置断点以便调试
