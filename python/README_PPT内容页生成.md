# PPT内容页信息列表生成工具使用说明

## 功能说明

根据 `requirement.txt` 和 `大纲目录.txt` 的内容，结合 `生成ppt_slide内容页的提示词.txt` 的要求，生成PPT内容页信息列表，保存到 `ppt内容页.txt` 文件中。

## 使用方法

### 使用OpenAI API生成（推荐）

```bash
# 设置API密钥
export OPENAI_API_KEY="your-api-key-here"

# 运行脚本
cd /Users/menggl/workspace/PPTFactory
python3 python/generate_ppt_slides.py --use-api
```

或者直接在命令行指定API密钥：

```bash
python3 python/generate_ppt_slides.py --use-api --api-key "your-api-key-here"
```

### 使用其他模型

```bash
# 使用gpt-3.5-turbo（更便宜，但质量可能稍低）
python3 python/generate_ppt_slides.py --use-api --model "gpt-3.5-turbo"

# 使用gpt-4（更高质量，推荐）
python3 python/generate_ppt_slides.py --use-api --model "gpt-4"
```

## 功能特点

1. **自动加载模板信息**: 自动读取所有模板文件（T001-T012.json），了解模板结构
2. **智能模板选择**: 根据内容类型和文本数量自动选择合适的模板
3. **内容结构化**: 严格按照模板要求组织key_points的结构
4. **自动拆分**: 如果内容过多，自动拆分成多个页面
5. **JSON格式输出**: 生成标准的JSON数组格式，便于后续处理

## 输入文件

- **requirement.txt**: 必须包含原始文本内容
- **大纲目录.txt**: 必须包含大纲目录
- **生成ppt_slide内容页的提示词.txt**: 必须包含生成提示词的要求
- **templates/metadata/T*.json**: 模板文件（自动加载）

## 输出结果

生成的PPT内容页信息列表保存在 `produce/ppt内容页.txt` 文件中，格式为JSON数组：

```json
[
  {
    "templateId": "T001",
    "slide_title": "一、煤矿从业人员主要权利",
    "slide_subtitle": "总览",
    "layout": "title_with_three_image_cards",
    "key_points": [
      "健康保障权",
      "拒绝违章指挥权",
      "停止作业避险权"
    ],
    "optional_graphics": "三张配图/图标，对应三要点",
    "notes": "caption 2-5 字，标题 10-20 字"
  },
  {
    "templateId": "T002",
    "slide_title": "一、煤矿从业人员主要权利",
    "slide_subtitle": "1. 健康保障权",
    "layout": "section_title_with_two_icon_paragraphs",
    "key_points": [
      {
        "heading": "企业需为确诊尘肺病司机提供治疗和康复服务",
        "body": "司机若确诊尘肺病，企业需提供治疗和康复服务。这是保障从业人员健康权益的重要措施，企业必须履行相关责任，确保患病司机得到及时有效的医疗救治和康复支持。"
      },
      {
        "heading": "离职时可索要历年体检报告，企业不得拒绝",
        "body": "从业人员在离职时有权索要历年体检报告，企业不得拒绝。这些体检报告是了解自身健康状况的重要依据，有助于后续的健康管理和职业病防治。"
      }
    ],
    "optional_graphics": "双圆形图标（健康/体检）",
    "notes": "每段 80-200 字，标题 10-25 字"
  }
]
```

## 生成规则

### 模板选择

1. **优先匹配文本数量**: 优先选择文本占位符数量与内容匹配的模板
2. **考虑内容类型**: 根据内容类型（概念、案例、制度等）选择合适模板
3. **文本容量分数**: 根据文本容量分数选择合适的模板
   - 0.3-0.7: 要点/标题页，精简
   - 0.7-0.85: 一般说明，每段 80-200 字
   - 0.85-0.92: 详细解读，每段 200-350/450 字

### 内容组织

1. **标题要求**:
   - slide_title: 10-25 或 15-40 字符，字体 30-44
   - slide_subtitle: 10-30 字符，字体 18-28
   - 如果副标题是"总览（上）"或"总览（下）"，设为空字符串

2. **正文要求**:
   - heading: 15-40 字符
   - body_text / paragraph: 80-200 或 160-300 字符
   - caption: 2-5 字符

3. **key_points结构**:
   - 仅要点列表类：数组字符串
   - heading+body/paragraph 类：数组对象 { "heading": "...", "body"/"paragraph": "..." }
   - multi_text_block / card_block / info_block：数组对象，字段名与模板占位一致

### 页面拆分

- 如果内容过多，无法在一个页面中展示，会自动拆分成多个页面
- 拆分后的页面使用相同的模板
- 每个页面都有完整的标题和副标题

## 注意事项

1. **API使用**: 使用OpenAI API需要安装 `openai` 库：
   ```bash
   pip install openai
   ```

2. **API密钥**: 必须设置正确的API密钥，否则无法生成

3. **模型选择**: 
   - gpt-4: 更高质量，但更贵
   - gpt-3.5-turbo: 更便宜，但质量可能稍低

4. **内容质量**: 生成的内容质量取决于：
   - 原始文本的质量
   - 大纲目录的清晰度
   - 提示词模板的详细程度
   - 使用的模型

5. **JSON验证**: 脚本会自动验证生成的JSON格式，如果格式有问题会提示

6. **手动调整**: 生成后可能需要手动调整一些内容，特别是：
   - 模板选择是否合适
   - 内容长度是否符合要求
   - 页面拆分是否合理

## 示例

### 输入

**requirement.txt**:
```
煤矿开采岗位安全操作与法律法规精炼文本
一、煤矿从业人员主要权利

健康保障权
司机若确诊尘肺病，企业需提供治疗和康复服务...
```

**大纲目录.txt**:
```
一、煤矿从业人员主要权利
1. 健康保障权
2. 拒绝违章指挥权
...
```

### 输出

生成JSON数组，包含所有内容页信息，每个大纲条目对应至少一页。

## 故障排除

1. **API调用失败**: 检查API密钥是否正确，网络是否正常
2. **JSON格式错误**: 检查生成的内容，可能需要手动修复
3. **模板选择不当**: 可以手动调整templateId
4. **内容过长**: 检查是否需要进一步拆分页面

## 后续步骤

生成PPT内容页信息列表后，可以：
1. 检查生成的内容是否符合要求
2. 手动调整模板选择或内容
3. 使用 `generate_ppt_content_mapping.py` 生成文本映射关系
4. 使用 `ProduceUtil.java` 生成PPT文件








