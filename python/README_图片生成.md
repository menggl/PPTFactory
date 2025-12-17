# 图片生成工具使用说明

## 功能说明

根据 `ppt内容映射.txt` 文件中的图片提示词生成对应的图片，并更新映射文件。

## 使用方法

### 方式1: 使用简化脚本（当前使用）

```bash
cd /Users/menggl/workspace/PPTFactory
python3 python/generate_images_simple.py
```

### 方式2: 使用完整脚本（支持多种图片生成方式）

```bash
cd /Users/menggl/workspace/PPTFactory
python3 python/generate_images.py
```

## 功能特点

1. **自动查找最新PPT文件**: 自动查找 `produce/` 目录下最新的 `new_ppt_*.pptx` 文件
2. **创建图片目录**: 在 `produce/images/new_ppt_[时间戳]/` 目录下保存图片
3. **图片命名规则**: `[第几页]_[第几张图片].png`，例如 `1_1.png`, `1_2.png`
4. **更新映射文件**: 在 `ppt内容映射.txt` 中添加 `新图片映射` 字段，保存图片标注与图片路径的映射关系

## 图片生成方式

### 当前实现（占位图片）

当前脚本生成的是最小PNG占位文件（1x1像素），用于测试流程。

### 替换为真实图片生成

要使用真实的AI图片生成，可以：

#### 1. 使用OpenAI DALL-E

```python
# 在脚本中设置
use_openai = True
openai_api_key = os.getenv("OPENAI_API_KEY")  # 或直接设置密钥
```

需要安装：
```bash
pip install openai
```

#### 2. 使用本地Stable Diffusion

```python
# 在脚本中设置
use_stable_diffusion = True
stable_diffusion_url = "http://localhost:7860"  # 你的SD API地址
```

#### 3. 使用其他图片生成API

修改 `generate_images_from_prompts` 函数中的图片生成逻辑，调用你选择的图片生成服务。

## 输出结果

- **图片文件**: 保存在 `produce/images/new_ppt_[时间戳]/` 目录
- **映射文件**: `produce/ppt内容映射.txt` 中每个页面添加了 `新图片映射` 字段

## 示例

映射文件中的结构：

```json
{
  "模板页编号": "T001",
  "图片映射": {
    "三我是文本": "健康保障权"
  },
  "图片提示词": {
    "三我是文本": "PPT演示风格的专业插图，主题：..."
  },
  "新图片映射": {
    "三我是文本": "images/new_ppt_20251212141750/1_1.png"
  }
}
```

## 注意事项

1. 当前使用的是占位图片，实际使用时需要替换为真实的图片生成API
2. 图片生成可能需要较长时间，建议批量生成时添加进度提示
3. 生成的图片路径是相对路径，相对于项目根目录








