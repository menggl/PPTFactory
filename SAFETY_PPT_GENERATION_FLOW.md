# 安全生产类型PPTX生成详细流程

## 一、整体流程概览

```
命令行启动 → 加载JSON → 创建引擎 → 初始化（安全生产特殊处理） → 渲染幻灯片 → 保存PPTX
```

## 二、详细流程步骤

### 阶段1：程序启动和参数解析

**入口：`GeneratePPT.main(String[] args)`**

1. **解析命令行参数**
   - `input.json`：输入JSON文件路径
   - `-o output.pptx`：输出PPTX文件路径（默认：output.pptx）
   - `--style safety`：风格策略（safety）
   - `--template safety`：模板类型（safety）

2. **加载输入JSON文件**
   - 使用Jackson解析JSON文件
   - 获取 `slides` 数组，包含所有幻灯片数据

3. **获取风格策略对象**
   - 根据 `--style` 参数创建对应的 `StyleStrategy` 对象
   - 安全生产类型使用 `SafetyStyle` 类

4. **获取模板文件路径**
   - 根据 `--template` 参数查找模板文件
   - 安全生产类型：`templates/safety/theme.pptx`
   - **注意**：这个文件主要用于：
     - 获取幻灯片尺寸（宽度、高度、类型）等基础属性
     - 判断是否为安全生产类型（通过路径判断）
     - **实际上不直接用于渲染幻灯片内容**（渲染时使用的是从参考PPT提取的经典布局）

### 阶段2：模板引擎初始化

**入口：`PPTTemplateEngine(String templateFile, StyleStrategy styleStrategy)`**

#### 2.1 基础初始化

1. **加载模板文件**
   - 打开 `templates/safety/theme.pptx`
   - 创建新的空白演示文稿（`Presentation`）
   - **设置幻灯片尺寸**（从模板文件获取）
     - 这是 `templates/safety/theme.pptx` 的主要作用
     - 用于设置输出PPTX的幻灯片尺寸（宽度、高度、类型）
   - **注意**：这个模板文件的内容（幻灯片）不会被直接使用
     - 实际渲染时使用的是从参考PPT提取的经典布局（`templates/classic/`）
     - 或者使用硬编码的标准布局

2. **设置风格策略**
   - 保存 `StyleStrategy` 对象（`SafetyStyle`）

#### 2.2 安全生产类型特殊初始化

**判断条件：`isSafetyTemplate()`**
- 检查模板文件路径是否包含 "safety" 或 "安全生产"

**如果为安全生产类型，执行以下步骤：**

##### 步骤1：加载参考PPT
**方法：`loadSafetyReferencePresentation()`**

- 加载 `1.2 安全生产方针政策.pptx` 文件
- 保存为 `safetyReferencePresentation` 对象
- 用于后续提取布局和样式

##### 步骤2：加载布局配置文件
**方法：`loadLayoutConfig()`**

- 读取 `config/layouts.json` 文件
- 解析布局配置：
  - `layouts` 数组 → `layoutConfigMap`（布局名称 → 配置信息）
  - `styles` 对象 → `layoutStyleMap`（布局名称 → 样式信息）
- 如果配置文件不存在，`layoutConfigMap` 为空

##### 步骤3：加载经典布局映射
**方法：`loadClassicLayoutsFromConfig()` 或 `extractAndRegisterClassicLayouts()`**

**情况A：配置文件存在且不为空**
- 从 `layoutConfigMap` 中提取布局信息
- 构建 `classicLayoutMap`（布局名称 → 源PPT页码）
- 支持别名映射和类别映射

**情况B：配置文件不存在或为空**
- 从参考PPT的第5页到倒数第2页提取布局
- 自动注册为经典布局类型（`classic_image_text_5`, `classic_image_text_6`, ...）
- 构建 `classicLayoutMap`

##### 步骤4：提取模板文件和样式类
**方法：`extractTemplatesAndStyles()`**

- 遍历参考PPT的第5页到倒数第2页
- 对每一页执行：
  1. **保存为模板文件**：`saveSlideAsTemplate()`
     - 克隆幻灯片
     - 保存为临时PPTX文件
     - **去除水印**：通过XML方式移除Aspose.Slides水印
     - **替换文本**：将所有文本替换为"模板文字模板文字..."（保留样式）
     - **替换图片**：将所有图片（除顶部图标）替换为"No Image"占位图
     - 保存到 `templates/classic/classic_image_text_N.pptx`
  2. **生成样式类**：`createStyleClassForTemplate()`
     - 生成对应的Java样式类文件（如 `ImageText5Style.java`）
     - 保存到 `src/main/java/com/pptfactory/style/` 目录

### 阶段3：渲染PPT

**入口：`renderFromJson(Map<String, Object> slidesData)`**

#### 3.1 插入固定前四页（仅安全生产类型）

**方法：`addSafetyCoverSlidesIfNeeded()`**

- 从 `1.2 安全生产方针政策.pptx` 复制前4页（第1-4页）
- 按顺序插入到当前演示文稿的开头
- **完整保留**：包括动画、视频、所有格式和内容

#### 3.2 渲染JSON中的幻灯片

**循环处理JSON中的每一张幻灯片：**

```java
for (int i = 0; i < slides.size(); i++) {
    Map<String, Object> slideData = slides.get(i);
    ISlide slide = renderSlide(slideData);
}
```

#### 3.3 插入固定最后一页（仅安全生产类型）

**方法：`addSafetyLastSlideIfNeeded()`**

- 从 `1.2 安全生产方针政策.pptx` 复制最后1页
- 插入到当前演示文稿的末尾
- **完整保留**：包括动画、视频、所有格式和内容

### 阶段4：单张幻灯片渲染

**入口：`renderSlide(Map<String, Object> slideData)`**

根据JSON中的 `layout` 字段，按以下优先级匹配布局：

#### 优先级1：经典布局（从配置文件或自动提取）

**匹配方式：**
1. **直接匹配**：`layout` 字段等于布局名称（如 `"classic_image_text_5"`）
2. **小写匹配**：不区分大小写匹配
3. **别名匹配**：通过 `aliases` 字段匹配
4. **类别匹配**：通过 `category` 字段匹配（如 `"image_with_text"` 匹配到 `classic_image_text_5`）

**渲染方法：`renderClassicLayout(pageNumber, slideData, layoutName)`**

**流程：**
1. 从 `templates/classic/` 目录加载对应的模板文件（如 `classic_image_text_5.pptx`）
2. 如果模板文件不存在，使用默认内容页布局
3. 复制模板文件中的第一张幻灯片
4. **替换文本内容**：`replaceSlideTextContent()`
   - 收集所有包含"模板文字"的文本框（排除顶部标题栏，Y < 72点）
   - 按位置排序（先Y坐标，再X坐标）
   - 智能匹配：
     - **标题**：使用最上面的文本框（Y坐标最小）
     - **正文**：使用最大的文本框（面积最大）
     - **多列布局**：按X坐标位置匹配（左、中、右）
   - 替换文本，保留原有样式
5. **替换图片内容**：`replaceSlideImageContent()`
   - 如果JSON中提供了 `image_path`，替换模板中的图片
   - 如果没有提供，保留模板中的"No Image"占位图
6. **应用样式**：如果配置文件中定义了样式，应用样式
7. 返回渲染后的幻灯片

#### 优先级2：经典布局（classic_* 格式）

如果 `layout` 字段以 `"classic_"` 开头，直接使用该布局名称。

#### 优先级3：安全生产内容页布局（safety_content_N 格式）

**格式：`"layout": "safety_content_5"`**

**渲染方法：`renderSafetyContentPage(pageNumber, slideData)`**

- 直接从参考PPT复制指定页面（第5页到倒数第2页）
- 完整复制所有内容、格式、动画和视频

#### 优先级4：安全生产类型布局匹配

**方法：`findMatchingSafetyLayout(layoutName)`**

- 尝试从参考PPT中查找匹配的布局页面
- 使用参考PPT中的布局，然后替换文本内容

#### 优先级5：标准布局类型

**硬编码的标准布局：**
- `title_page`：标题页
- `content_page` / `title_with_content`：内容页
- `two_column`：两列布局
- `three_column`：三列布局
- `image_with_text` / `image_with_content`：图片+文字
- `image_left_text_right`：左图右文
- `image_right_text_left`：右图左文
- `pure_content`：纯内容页
- `quote_page`：引用页
- `chapter_cover`：章节封面页

**渲染方法：** 调用对应的硬编码渲染方法（如 `renderTitlePage()`, `renderContentPage()` 等）

### 阶段5：文本替换详细流程

**方法：`replaceSlideTextContent(ISlide slide, Map<String, Object> slideData)`**

#### 5.1 收集可替换文本框

**方法：`collectReplaceableTextShapesRecursive()`**

- 递归遍历所有形状（包括组合形状）
- 筛选条件：
  - 必须是 `IAutoShape`（自动形状/文本框）
  - 文本内容包含"模板文字"
  - Y坐标 >= 72点（排除顶部标题栏）

#### 5.2 排序文本框

- 先按Y坐标排序（从上到下）
- 再按X坐标排序（从左到右）

#### 5.3 智能匹配和替换

**单列布局：**
- **标题**：最上面的文本框（Y坐标最小）
- **正文**：最大的文本框（面积最大）
- **其他字段**：剩余的文本框按顺序替换

**多列布局（`left_content`, `middle_content`, `right_content`）：**
- 按X坐标位置分类：
  - 左列：X < 33% 幻灯片宽度
  - 中间列：33% <= X < 67%
  - 右列：X >= 67%
- 分别替换对应列的内容

**替换方法：`replaceTextInShape(IAutoShape shape, String text, boolean isTitle)`**
- 清空原有文本
- 按换行符分割文本，创建多个段落
- 应用样式（如果是标题，使用标题样式）

### 阶段6：保存PPTX

**方法：`save(String filename)`**

1. **保存PPTX文件**
   - 使用Aspose.Slides保存为PPTX格式

2. **去除水印（自动执行）**
   - **方法：`removeWatermarksFromXML(String filename)`**
   - 解压PPTX文件（ZIP格式）
   - 遍历所有XML文件（幻灯片、母版、布局）
   - 查找包含水印关键词的文本节点：
     - "Evaluation only."
     - "Created with Aspose.Slides for Java 25.11."
     - "Copyright 2004-2025 Aspose Pty Ltd."
   - 删除包含水印的整个形状节点（`<p:sp>`）
   - 重新打包为PPTX文件

3. **关闭资源**
   - 释放所有Presentation对象

## 三、关键数据结构

### 1. 布局映射

```java
Map<String, Integer> classicLayoutMap;
// 键：布局名称（如 "classic_image_text_5"）
// 值：源PPT页码（如 5）
```

### 2. 布局配置

```java
Map<String, Map<String, Object>> layoutConfigMap;
// 键：布局名称
// 值：布局配置信息（包含 pageNumber, displayName, category, aliases 等）
```

### 3. 布局样式

```java
Map<String, Map<String, Object>> layoutStyleMap;
// 键：布局名称
// 值：样式配置信息
```

## 四、文件结构

### 输入文件
- `examples/safety_slides_extended.json`：JSON输入文件
- `1.2 安全生产方针政策.pptx`：参考PPT（用于提取布局和固定页面）
- `config/layouts.json`：布局配置文件（可选）

### 输出文件
- `output.pptx`：生成的PPTX文件
- `templates/classic/classic_image_text_N.pptx`：提取的模板文件
- `src/main/java/com/pptfactory/style/ImageTextNStyle.java`：生成的样式类

### 模板文件
- `templates/safety/theme.pptx`：安全生产主题模板
  - **作用**：主要用于获取幻灯片尺寸等基础属性
  - **不直接用于渲染**：实际渲染使用的是从参考PPT提取的经典布局

## 五、特殊说明

### 1. 固定页面
- **前四页**：始终使用参考PPT的第1-4页，完整保留所有内容
- **最后一页**：始终使用参考PPT的最后一页，完整保留所有内容

### 2. 模板文件处理
- 所有模板文件中的文本都替换为"模板文字模板文字..."
- 所有图片（除顶部图标）都替换为"No Image"占位图
- 保留原有的样式、字体、大小、对齐方式等

### 3. 文本替换规则
- 只替换包含"模板文字"的文本框
- 跳过顶部标题栏（Y < 72点）
- 智能匹配：根据位置和大小判断文本框用途

### 4. 布局匹配优先级
1. 经典布局（配置文件/自动提取）
2. 经典布局（classic_* 格式）
3. 安全生产内容页（safety_content_N）
4. 安全生产类型布局匹配
5. 标准布局类型（硬编码）

## 六、流程图

```
开始
  ↓
解析命令行参数
  ↓
加载JSON文件
  ↓
创建模板引擎
  ↓
[是安全生产类型?]
  ├─ 是 → 加载参考PPT
  │      → 加载布局配置
  │      → 加载经典布局映射
  │      → 提取模板文件和样式类
  └─ 否 → 跳过
  ↓
插入固定前四页（仅安全生产）
  ↓
循环处理JSON中的每张幻灯片
  │
  ├─ 匹配布局类型
  │   ├─ 经典布局 → 从模板文件加载
  │   ├─ safety_content_N → 从参考PPT复制
  │   └─ 标准布局 → 硬编码渲染
  │
  ├─ 替换文本内容
  │   ├─ 收集可替换文本框
  │   ├─ 智能匹配（标题/正文/多列）
  │   └─ 替换文本，保留样式
  │
  ├─ 替换图片内容（如果提供）
  │
  └─ 应用样式（如果配置）
  ↓
插入固定最后一页（仅安全生产）
  ↓
保存PPTX文件
  ↓
去除水印（XML方式）
  ↓
结束
```

