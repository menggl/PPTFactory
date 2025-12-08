# PPT模板使用指南

## 概述

PPT模板引擎现在支持从现有的PPT样例文件中提取样式和布局，让生成的PPT与您提供的样例风格保持一致。

## 两种使用方式

### 方式一：使用模板PPT文件（推荐）

如果您有现成的PPT样例文件，可以直接将其作为模板使用：

```bash
python main.py slides.json -o output.pptx --template your_template.pptx
```

**优点：**
- 完全保留原PPT的布局、样式、主题色等所有视觉元素
- 无需手动定义样式
- 生成的PPT与样例风格完全一致

**使用步骤：**
1. 准备一个PPT样例文件（包含您想要的样式和布局）
2. 运行命令，指定模板文件
3. 生成的PPT将使用模板的样式

### 方式二：提取样式并保存为JSON

如果您想先提取样式信息，然后可以手动调整：

```bash
# 步骤1：从PPT样例中提取样式信息
python template_extractor.py your_template.pptx extracted_styles.json

# 步骤2：查看并编辑extracted_styles.json（可选）

# 步骤3：使用提取的样式生成PPT
python main.py slides.json -o output.pptx --style extracted_styles.json
```

**优点：**
- 可以查看和编辑样式配置
- 可以混合使用多个模板的样式
- 样式配置可以版本控制

## 详细使用说明

### 1. 提取模板样式

从现有PPT文件中提取样式信息：

```bash
python template_extractor.py <模板PPT文件> [输出JSON文件]
```

**示例：**
```bash
python template_extractor.py my_template.pptx my_styles.json
```

**提取的信息包括：**
- 字体大小（标题、副标题、正文、要点）
- 字体名称
- 文本颜色
- 文本框位置和大小
- 对齐方式
- 幻灯片尺寸

### 2. 使用模板文件生成PPT

直接使用PPT模板文件：

```bash
python main.py slides.json -o output.pptx --template template.pptx
```

**示例：**
```bash
python main.py example_slides.json -o result.pptx --template demo.pptx
```

### 3. 使用提取的样式生成PPT

使用从模板中提取的样式JSON文件：

```bash
python main.py slides.json -o output.pptx --style styles.json
```

**样式JSON文件格式：**
```json
{
  "title_font_size": 48,
  "subtitle_font_size": 24,
  "content_font_size": 18,
  "bullet_font_size": 16
}
```

### 4. 组合使用模板和样式

可以同时使用模板文件和样式文件，样式文件会覆盖模板中的对应样式：

```bash
python main.py slides.json -o output.pptx --template template.pptx --style styles.json
```

## 编程接口使用

### 从模板文件创建引擎

```python
from ppt_template_engine import PPTTemplateEngine

# 使用模板文件
engine = PPTTemplateEngine.from_template("template.pptx")
engine.render_from_file("slides.json")
engine.save("output.pptx")
```

### 使用模板文件并自定义样式

```python
from ppt_template_engine import PPTTemplateEngine
from pptx.util import Pt

# 使用模板文件，但覆盖部分样式
custom_style = {
    "title_font_size": Pt(54),  # 更大的标题
    "content_font_size": Pt(20)
}
engine = PPTTemplateEngine(template_file="template.pptx", style=custom_style)
engine.render_from_file("slides.json")
engine.save("output.pptx")
```

### 提取模板样式

```python
from template_extractor import PPTTemplateExtractor

# 创建提取器
extractor = PPTTemplateExtractor("template.pptx")

# 提取所有样式信息
styles = extractor.extract_all_styles()

# 保存为JSON
extractor.save_styles_to_json("styles.json")

# 获取适合模板引擎使用的样式
engine_style = extractor.get_template_for_engine()
```

## 常见问题

### Q: 模板文件需要包含什么内容？

A: 模板文件可以是任何PPT文件。建议包含：
- 您想要的标题样式
- 您想要的正文样式
- 您想要的布局结构
- 您想要的颜色主题

### Q: 如果模板文件中的布局与我的JSON不匹配怎么办？

A: 模板引擎会使用模板文件的样式（字体、颜色等），但布局仍然由模板引擎根据JSON中的`layout`字段决定。如果您想完全使用模板的布局，需要修改模板引擎代码或使用模板文件的slide_layouts。

### Q: 可以同时使用多个模板吗？

A: 目前不支持直接使用多个模板。但您可以：
1. 从多个模板中提取样式
2. 手动合并样式JSON文件
3. 使用合并后的样式文件

### Q: 提取的样式JSON可以手动编辑吗？

A: 可以！提取的JSON文件是纯文本格式，您可以用任何文本编辑器打开并修改字体大小、颜色等值。

## 最佳实践

1. **准备模板文件**：创建一个包含您想要的所有样式元素的PPT样例
2. **测试提取**：先用`template_extractor.py`提取样式，查看是否符合预期
3. **使用模板**：直接使用`--template`参数，最简单快捷
4. **保存样式**：如果样式满意，保存提取的JSON文件，方便后续使用和版本控制

## 示例工作流

```bash
# 1. 准备您的PPT样例
# （在PowerPoint中创建一个包含您想要样式的PPT文件，保存为template.pptx）

# 2. 提取样式（可选，用于查看和编辑）
python template_extractor.py template.pptx styles.json

# 3. 准备内容JSON文件
# （大模型输出或手动编写slides.json）

# 4. 生成PPT
python main.py slides.json -o output.pptx --template template.pptx

# 5. 查看生成的PPT，如果满意就完成了！
```

