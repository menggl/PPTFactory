# PPT模板引擎

PPT模板引擎是一个使用Java实现的PPT生成工具，使用大模型输出结构化内容（如 slides.json），而"怎么排版、如何布局"全部由模板引擎管理。

## 快速开始

```bash
# 1. 编译项目
mvn clean package -DskipTests

# 2. 运行生成脚本（默认生成安全生产类型PPT）
./run_pptx.sh

# 生成的文件：output.pptx
# 输入文件：examples/safety_slides_extended.json
# 默认风格和模板：safety（安全生产类型）
```

**默认配置说明**：
- **输入文件**：`examples/safety_slides_extended.json`（安全生产类型PPT的默认输入）
- **输出文件**：`output.pptx`（固定文件名，每次运行会覆盖）
- **风格**：`safety`（安全生产风格）
- **模板**：`safety`（安全生产模板）
- **注意**：每次代码变更后，运行 `./run_pptx.sh` 会自动更新 `output.pptx` 文件

**模板和风格自动抽离**：
- 系统会自动从《1.2 安全生产方针政策.pptx》的第5页到倒数第2页抽离模板文件和策略类风格
- **所有模板统一保存到 `templates/master_template.pptx` 一个文件中**（不再分散在多个文件中）
- 策略类风格保存在 `src/main/java/com/pptfactory/style/` 目录（如 `ImageText5Style.java`）
- 每个模板幻灯片对应一个策略类风格，保留了源PPT页面的样式特征
- **自动去除水印**：master_template.pptx 会自动去除 Aspose.Slides 水印（"Evaluation only."、"Created with Aspose.Slides for Java 25.11."、"Copyright 2004-2025 Aspose Pty Ltd."），包括水印使用的文本框（使用底层遍历XML的方式，直接操作PPTX文件的XML结构，通过两种方法检测和删除包含水印的形状节点：1) 通过文本框架收集完整文本进行检测；2) 直接遍历所有文本节点进行检测）
- **模板文字替换**：在删除水印后，master_template.pptx 中的所有文字内容（包括组合形状中的文字）会自动替换为"模板文字模板文字模板文字"，保持文本框中的字数与原《1.2 安全生产方针政策.pptx》文件中的文本框中的对应字数相同（字数多了截取，字数少了重复）、文字大小相同、字体样式相同、字体对齐方式相同
- **模板图片替换**：master_template.pptx 中的所有图片（顶部图标除外，包括组合形状中的图片）会自动替换为带有"No Image"文字标识的图片，顶部图标（Y坐标小于72点）会保留

## 项目结构

```
PPTFactory/
├── src/main/java/com/pptfactory/
│   ├── template/              # 模板相关
│   │   └── engine/           # 统一模板引擎（Aspose.Slides）
│   │       └── PPTTemplateEngine.java
│   ├── style/                # 风格策略类（多个）
│   │   ├── StyleStrategy.java
│   │   ├── DefaultStyle.java
│   │   ├── ChineseStyle.java
│   │   ├── MathStyle.java
│   │   ├── FinanceStyle.java
│   │   └── SafetyStyle.java
│   ├── ai/                   # AI/大模型相关
│   │   ├── ContentGenerator.java
│   │   └── LayoutClassifier.java
│   └── cli/                  # 命令行接口
│       └── GeneratePPT.java
├── templates/                 # PPT模板文件
│   ├── master_template.pptx  # 统一的模板文件（包含所有经典布局模板）
│   ├── chinese/
│   │   └── theme.pptx
│   ├── math/
│   │   └── theme.pptx
│   ├── finance/
│   │   └── theme.pptx
│   └── safety/
│       └── theme.pptx
├── examples/                  # 示例JSON文件
│   ├── example_slides.json
│   ├── safety_slides.json
│   └── safety_slides_extended.json
├── lib/                       # 第三方JAR文件
│   └── aspose-slides-25.11-jdk16.jar  # Aspose.Slides JAR（需手动下载）
├── run_pptx.sh               # 快速运行脚本
├── pom.xml                    # Maven配置文件
└── 1.2 安全生产方针政策.pptx  # 安全生产类型PPT的封面模板
```

## 设计理念

- **模板引擎保持统一**：所有风格通过策略类实现
- **模板文件决定布局**：页面长什么样、有哪些占位符、布局怎么摆
- **风格策略类决定样式**：颜色、字体、图片风格、行距间距等视觉效果

## 技术栈

- **Java 17+**：项目使用 JDK 17 编译和运行
- **Aspose.Slides for Java 25.11**：用于PPTX文件处理（支持动画、视频等高级功能）
- **Jackson 2.15.2**：用于JSON处理
- **Maven**：依赖管理和构建工具

## 测试和调试

项目提供了完整的单步调试测试代码，位于 `src/test/java/` 目录下。每个功能模块都有对应的测试类，可以直接在IDE中设置断点进行单步调试。

### 测试类列表

- **TestPPTTemplateEngine.java**: 测试模板引擎核心功能
- **TestStyleStrategy.java**: 测试风格策略应用
- **TestWatermarkRemoval.java**: 测试水印移除功能
- **TestTemplateExtraction.java**: 测试模板提取功能
- **TestContentGenerator.java**: 测试内容生成器
- **TestLayoutClassifier.java**: 测试布局分类器
- **TestGeneratePPT.java**: 测试CLI入口

### 使用方法

1. 在IDE中打开对应的测试类
2. 在需要调试的代码行设置断点
3. 以Debug模式运行测试类
4. 使用单步调试功能（Step Over/Into/Out）查看执行流程

详细说明请参考 [src/test/README.md](src/test/README.md)

## 构建项目

### 前置准备

1. **下载 Aspose.Slides JAR**
   ```bash
   # 从 https://products.aspose.com/slides/java 下载
   # 将 JAR 文件放到 lib/ 目录
   mkdir -p lib
   # 将下载的 aspose-slides-25.11-jdk16.jar 复制到 lib/ 目录
   ```

2. **确保 Java 17 已安装并配置**
   ```bash
   java -version  # 应该显示 17.x.x
   ```

### 构建命令

```bash
# 编译项目
mvn clean compile

# 打包项目（包含所有依赖，但不包含 system scope 的 Aspose.Slides）
mvn clean package

# 运行测试
mvn test

# 跳过测试打包
mvn clean package -DskipTests
```

### 运行项目

```bash
# 使用快速运行脚本（推荐）
# 默认使用 examples/safety_slides_extended.json 作为输入
# 默认输出文件为 output.pptx
# 默认使用 safety 风格和模板
./run_pptx.sh

# 或手动运行（安全生产类型PPT生成）
java -cp "target/ppt-template-engine-1.0.0-jar-with-dependencies.jar:lib/aspose-slides-25.11-jdk16.jar" \
     -Duser.language=en -Duser.country=US \
     com.pptfactory.cli.GeneratePPT \
     examples/safety_slides_extended.json \
     -o output.pptx \
     --style safety \
     --template safety
```

**注意**：
- 默认输入文件：`examples/safety_slides_extended.json`
- 默认输出文件：`output.pptx`（固定文件名）
- 默认风格和模板：`safety`（安全生产类型）

## 使用方法

### 1. 准备输入JSON文件

创建一个JSON文件，包含幻灯片数据：

```json
{
  "slides": [
    {
      "layout": "title_page",
      "title": "我的演示",
      "subtitle": "副标题"
    },
    {
      "layout": "content_page",
      "title": "内容标题",
      "bullets": [
        "要点1",
        "要点2",
        "要点3"
      ]
    }
  ]
}
```

### 2. 运行生成命令

**方式一：使用快速运行脚本（推荐）**

```bash
# 使用提供的运行脚本
# 默认配置：
# - 输入文件：examples/safety_slides_extended.json
# - 输出文件：output.pptx（固定文件名）
# - 风格：safety
# - 模板：safety
./run_pptx.sh
```

**方式二：使用Maven运行**

```bash
mvn exec:java -Dexec.mainClass="com.pptfactory.cli.GeneratePPT" \
    -Dexec.args="examples/safety_slides_extended.json -o output.pptx --style safety --template safety"
```

**方式三：使用打包后的JAR文件**

```bash
# 注意：需要将 Aspose.Slides JAR 添加到 classpath
java -cp "target/ppt-template-engine-1.0.0-jar-with-dependencies.jar:lib/aspose-slides-25.11-jdk16.jar" \
     -Duser.language=en -Duser.country=US \
     com.pptfactory.cli.GeneratePPT \
     examples/safety_slides_extended.json \
     -o output.pptx \
     --style safety \
     --template safety
```

### 3. 命令行参数

- `input.json`：输入文件路径（JSON格式的slides数据）
- `-o, --output <file>`：输出的PPT文件名（默认: output.pptx）
- `--style <style>`：风格选择（default, chinese, math, finance, safety）
- `--template <template>`：模板选择（chinese, math, finance, safety）

## 支持的布局类型

### 标准布局类型

1. **title_page** - 标题页
2. **content_page** / **title_with_content** - 标题+内容页
3. **image_with_text** / **image_with_content** - 图片+内容页
4. **image_left_text_right** - 左图右文
5. **image_right_text_left** - 右图左文
6. **pure_content** - 纯内容页
7. **two_column** - 两栏页
8. **three_column** - 三栏页
9. **quote_page** - 引用页
10. **chapter_cover** - 章节封面页

### 安全生产专用布局类型

当使用 `--template safety` 时，支持以下特殊布局类型：

- **safety_content_N** - 从《1.2 安全生产方针政策.pptx》复制指定页面作为布局
  - N 为页码（5 到倒数第2页）
  - 例如：`safety_content_5` 表示使用源PPT的第5页布局
  - 例如：`safety_content_10` 表示使用源PPT的第10页布局
  - 完整保留源页面的所有格式、动画、视频等内容

### 经典布局类型（通用，所有PPT类型可用）

从《1.2 安全生产方针政策.pptx》中提取的经典布局，**可在任何PPT类型中使用**：

- **classic_*** - 自动提取并注册的经典布局类型
  - 格式：`classic_<特征>_<页码>`
  - 特征类型：
    - `image_text` - 包含图片和文本的布局
    - `multi_text` - 包含3个或更多文本框的布局
    - `two_text` - 包含2个文本框的布局
    - `single_text` - 包含1个文本框的布局
    - `simple` - 简单布局
  - 示例：
    - `classic_image_text_5` - 源PPT第5页的图片+文本布局
    - `classic_multi_text_10` - 源PPT第10页的多文本框布局
  - **自动注册**：当使用 `--template safety` 时，系统会自动从源PPT的第5页到倒数第2页提取布局并注册
  - **通用使用**：注册后，这些布局可以在任何PPT类型中使用（default, chinese, math, finance, safety等）
  - **完整保留**：保留源页面的所有格式、样式、动画和视频
  - **文本替换**：自动将源页面中的文本替换为JSON中提供的内容

## 支持的风格

- **default** - 默认风格
- **chinese** - 中文风格（使用微软雅黑字体）
- **math** - 数学风格（使用Times New Roman字体）
- **finance** - 金融风格（深蓝色系，专业稳重）
- **safety** - 安全生产风格（红色、橙色警示色，字体较大）

## 特殊功能

### 1. 自动水印移除

项目实现了自动移除 Aspose.Slides 评估版水印的功能：
- **实现方式**：直接操作 PPTX 文件的 XML 结构（不使用 Aspose API）
- **移除范围**：所有幻灯片、母版幻灯片、布局幻灯片
- **移除内容**：包含以下关键词的文本形状及其编辑框
  - "Evaluation only."
  - "Created with Aspose.Slides for Java 25.11."
  - "Copyright 2004-2025 Aspose Pty Ltd."
- **自动执行**：在保存 PPTX 文件时自动执行

### 2. 安全生产类型PPT特殊处理

当使用 `--template safety` 或模板路径包含 "safety" 时：

#### 固定页面
- **前四页固定使用**：`1.2 安全生产方针政策.pptx` 的前四张幻灯片（第1-4页）
- **最后一页固定使用**：`1.2 安全生产方针政策.pptx` 的最后一张幻灯片
- **完整保留**：包括动画、视频、所有格式和内容
- **自动插入**：
  - 前四页在渲染其他幻灯片之前自动插入（按顺序插入第1-4页）
  - 最后一页在所有幻灯片渲染完成后自动插入到末尾

#### 内容页布局集成

**方式一：使用 `safety_content_N` 布局类型（直接指定页面）**
- **支持使用源PPT的内容页布局**：从第5页到倒数第2页的所有页面都可以作为布局使用
- **使用方法**：在JSON中使用 `"layout": "safety_content_N"` 格式，其中 N 是页码（5 到倒数第2页）
- **示例**：
  ```json
  {
    "layout": "safety_content_5",
    "title": "自定义标题"
  }
  ```
- **完整复制**：使用 `safety_content_N` 布局时，会完整复制源PPT对应页面的所有内容、格式、动画和视频

**方式二：使用经典布局类型（推荐，通用）**
- **自动提取**：系统会自动从源PPT的第5页到倒数第2页提取布局并注册为经典布局类型
- **通用使用**：经典布局类型可以在任何PPT类型中使用（不仅仅是安全生产类型）
- **布局命名**：根据布局特征自动命名（如 `classic_image_text_5`, `classic_multi_text_10` 等）
- **完整保留**：保留源页面的所有格式、样式、动画和视频
- **文本替换**：自动将源页面中的文本替换为JSON中提供的内容
- **示例**：
  ```json
  {
    "layout": "classic_image_text_5",  // 使用经典布局类型
    "title": "安全生产的重要性",
    "text": "这是正文内容"
  }
  ```
  - 使用源PPT第5页的图片+文本布局
  - 可以在任何PPT类型中使用（default, chinese, math, finance, safety等）
  - 自动替换文本内容，保留所有样式和格式

**方式三：自动使用源PPT布局样式（标准布局类型自动映射）**
- **自动集成**：当使用标准布局类型（如 `content_page`, `two_column`, `title_page` 等）时，系统会自动从源PPT的第5页到倒数第2页中选择匹配的布局样式
- **智能映射**：系统会根据布局类型名称自动匹配源PPT中相似的页面布局
- **完整保留**：自动使用源PPT页面的所有格式、样式、动画和视频
- **文本替换**：自动将源PPT页面中的文本替换为JSON中提供的内容
- **示例**：
  ```json
  {
    "layout": "content_page",  // 使用标准布局类型
    "title": "安全生产的重要性",
    "bullets": ["要点1", "要点2"]
  }
  ```
  - 系统会自动从源PPT中选择一个 `content_page` 类型的页面作为模板
  - 保留该页面的所有样式和格式
  - 替换其中的文本内容为JSON中提供的内容

## 布局配置文件

### 配置文件位置

布局和样式配置保存在 `config/layouts.json` 文件中。

### 配置文件格式

```json
{
  "layouts": [
    {
      "name": "classic_image_text_5",
      "displayName": "图片+文本布局（经典）",
      "type": "classic",
      "category": "image_with_text",
      "aliases": ["image_text", "image_content", "picture_text"],
      "sourceFile": "1.2 安全生产方针政策.pptx",
      "pageNumber": 5,
      "description": "左侧或上方包含图片，右侧或下方包含文本内容的经典布局",
      "detailedDescription": "这是一个经典的图片+文本布局，适合用于展示产品、案例、说明等内容。",
      "useCases": ["产品介绍", "案例展示", "图文说明"],
      "features": {
        "hasImage": true,
        "textBoxCount": 2,
        "imageCount": 1,
        "layoutStructure": "图片位于左侧或上方，文本位于右侧或下方",
        "textPlaceholders": ["标题", "正文内容"],
        "imagePlaceholder": "主图片"
      },
      "recommendedContent": {
        "title": "建议包含标题文本",
        "text": "建议包含正文内容",
        "image": "建议包含一张主图片"
      },
      "tags": ["图片", "文本", "图文混排", "经典布局"]
    }
  ],
  "styles": {
    "classic_image_text_5": {
      "titleFontSize": 32,
      "titleColor": "#DC143C",
      "textFontSize": 18,
      "textColor": "#333333",
      "spacing": 12
    }
  }
}
```

### 配置文件说明

- **layouts**: 布局定义数组
  - `name`: 布局名称（唯一标识，用于在JSON中引用）
  - `displayName`: 布局显示名称（用于日志和界面显示）
  - `type`: 布局类型（如 "classic"）
  - `category`: 布局类别（如 "image_with_text", "content_page"）
  - `aliases`: 布局别名数组（支持通过别名引用布局，如 ["image_text", "image_content"]）
  - `sourceFile`: 源PPT文件路径
  - `pageNumber`: 源PPT中的页码（从1开始）
  - `description`: 布局简短描述
  - `detailedDescription`: 布局详细描述（包含排版格式的详细说明）
  - `useCases`: 适用场景数组（如 ["产品介绍", "案例展示"]）
  - `features`: 布局特征
    - `hasImage`: 是否包含图片
    - `textBoxCount`: 文本框数量
    - `imageCount`: 图片数量
    - `layoutStructure`: 布局结构描述（如"图片位于左侧，文本位于右侧"）
    - `textPlaceholders`: 文本占位符列表（如 ["标题", "正文内容"]）
    - `imagePlaceholder`: 图片占位符（如 "主图片"）
  - `layoutFormat`: 详细的排版格式信息（自动生成）
    - `textBoxes`: 文本框详细信息数组，每个文本框包含：
      - `index`: 文本框序号
      - `x`, `y`: 文本框位置坐标（点）
      - `width`, `height`: 文本框大小（点）
      - `alignment`: 文本对齐方式（"left", "center", "right", "justify"）
      - `fontSize`: 字体大小（点）
    - `images`: 图片详细信息数组，每个图片包含：
      - `index`: 图片序号
      - `x`, `y`: 图片位置坐标（点）
      - `width`, `height`: 图片大小（点）
    - `layoutStructure`: 布局结构描述
    - `textBoxCount`: 文本框总数
    - `imageCount`: 图片总数
  - `tags`: 标签数组（用于分类和搜索）

- **styles**: 布局样式配置
  - 键为布局名称
  - 值为样式配置对象（字体大小、颜色、间距等）

### 使用配置文件中的布局

在生成PPT的JSON文件中，可以通过以下方式使用配置文件中的布局：

1. **使用布局名称**：
   ```json
   {
     "layout": "classic_image_text_5",
     "title": "我的标题",
     "text": "正文内容",
     "image_path": "path/to/image.jpg"
   }
   ```

2. **使用别名**：
   ```json
   {
     "layout": "image_text",  // 使用别名
     "title": "我的标题",
     "text": "正文内容",
     "image_path": "path/to/image.jpg"
   }
   ```

3. **使用类别**：
   ```json
   {
     "layout": "image_with_text",  // 使用类别
     "title": "我的标题",
     "text": "正文内容",
     "image_path": "path/to/image.jpg"
   }
   ```

系统会自动匹配配置文件中的布局，支持不区分大小写的匹配。

**布局匹配优先级**：
1. 直接匹配布局名称（如 `"layout": "classic_image_text_5"`）
2. 匹配别名（如 `"layout": "image_text"` 匹配到 `classic_image_text_5`）
3. 匹配类别（如 `"layout": "image_with_text"` 匹配到 category 为 `"image_with_text"` 的布局）
4. 如果都不匹配，使用标准布局类型（如 `title_page`, `content_page` 等）

### 模板和样式选择机制

**生成新PPT时的流程**：
1. **挑选合适的模板幻灯片**：根据JSON中的 `layout` 字段和配置文件中定义的页码（`pageNumber`），从 `templates/master_template.pptx` 中选择对应的模板幻灯片
2. **挑选合适的样式文件**：如果配置文件中定义了样式（`styles` 字段），会自动应用对应的样式
3. **替换模板中的文本内容**：根据JSON中的 `title`、`text`、`bullets` 等字段，替换模板中的文本内容
4. **图片处理**：
   - 如果JSON中提供了 `image_path` 或 `imagePath`，会替换模板中的图片
   - 如果没有提供图片，保留模板文件中的"No Image"图片占位符

### 内容替换说明

当使用配置文件中的布局时，系统会自动替换布局中的内容：

- **文本替换**：
  - `title`: 替换第一个文本框（标题）
  - `subtitle`: 替换第二个文本框（副标题）
  - `text`: 替换正文文本框
  - `bullets`: 替换为要点列表（每个要点一行，自动添加"• "前缀）

- **图片替换**：
  - `image_path` 或 `imagePath` 或 `image`: 替换布局中的图片
  - 支持常见图片格式（JPG、PNG等）
  - 如果JSON中指定了图片路径，会自动替换布局中的第一个图片框
  - **如果没有提供图片路径，保留模板文件中的"No Image"图片占位符**（后续会实现图片生成功能）

**示例**：
```json
{
  "layout": "classic_image_text_5",
  "title": "产品介绍",
  "text": "这是产品的详细介绍内容",
  "image_path": "images/product.jpg",
  "bullets": [
    "特点1：高性能",
    "特点2：易用性",
    "特点3：可靠性"
  ]
}
```

**模板文件说明**：
- **master_template.pptx**：统一的模板文件，包含所有从源PPT抽离的模板幻灯片
  - 源PPT第5页对应 master_template.pptx 的第1张幻灯片（索引0）
  - 源PPT第6页对应 master_template.pptx 的第2张幻灯片（索引1）
  - 以此类推，源PPT第N页对应 master_template.pptx 的第(N-4)张幻灯片（索引N-5）
  - 所有模板幻灯片的文字已替换为"模板文字模板文字模板文字"
  - 所有模板幻灯片的图片已替换为"No Image"占位图（顶部图标除外）

### 手动维护配置文件

**重要**：排版格式描述信息需要**手动**在 `config/layouts.json` 中维护。

**手动添加排版格式描述**：
- 每个模板的排版格式描述需要手动添加到配置文件中
- 排版格式描述包括：
  - 文本框的位置、大小、对齐方式、字体大小
  - 图片的位置、大小
  - 布局结构描述（图片和文本的相对位置）
- 这些信息用于在生成新PPT时挑选合适的模板
- 建议在抽离模板后，手动查看模板文件，然后添加详细的排版格式描述到配置文件中

### 手动编辑配置文件

可以手动编辑 `config/layouts.json` 文件来：
- 添加新的布局定义
- 修改布局样式
- 删除不需要的布局

## 环境要求

- **JDK 17+**：项目使用 Java 17 编译
- **Maven 3.6+**：用于构建和依赖管理
- **操作系统**：支持 macOS、Linux、Windows

### 配置 Java 环境

如果使用 `jenv` 管理 Java 版本：

```bash
# 添加 Java 17 到 jenv
jenv add /usr/local/opt/openjdk@17/libexec/openjdk.jdk/Contents/Home

# 设置全局版本为 17
jenv global openjdk64-17.0.17

# 启用 export 插件（自动设置 JAVA_HOME）
jenv enable-plugin export
```

## 开发说明

### 添加新的风格策略

1. 创建新的风格类，实现 `StyleStrategy` 接口
2. 在 `GeneratePPT.getStyleStrategy()` 方法中添加新风格的映射
3. 创建对应的模板文件（可选）

### 添加新的布局类型

1. 在 `PPTTemplateEngine` 类中添加新的渲染方法
2. 在 `renderSlide()` 方法的 switch 语句中添加新的 case
3. 实现具体的布局渲染逻辑

## 依赖说明

### 必需依赖

- **Aspose.Slides for Java 25.11**：PPTX文件处理（支持动画、视频等高级功能）
  - 下载地址：https://products.aspose.com/slides/java
  - 将下载的 JAR 文件（如 `aspose-slides-25.11-jdk16.jar`）放到 `lib/` 目录
  - 注意：Aspose.Slides 是商业软件，评估版会在生成的PPT中添加水印
  - **本项目已实现自动水印移除功能**（通过直接操作 XML，不使用 Aspose API）
- **Jackson 2.15.2**：JSON处理
- **SLF4J + Logback**：日志记录

### 依赖配置

Aspose.Slides 在 `pom.xml` 中配置为 `system` scope，指向 `lib/` 目录中的 JAR 文件。运行时需要手动将 JAR 添加到 classpath（`run_pptx.sh` 脚本已自动处理）。

## 许可证

本项目使用 Apache License 2.0 许可证。

## 贡献

欢迎提交 Issue 和 Pull Request！
