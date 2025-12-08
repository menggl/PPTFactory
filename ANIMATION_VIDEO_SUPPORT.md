# 动画和视频支持说明

## 概述

根据 requirements.md 第82行的需求，安全生产的第一页需要支持复制动画和视频。已实现相关功能。

## 实现状态

### ✅ 已实现功能

1. **动画效果复制**
   - 实现位置：`template_engine/engine.py` 的 `_copy_slide_animations()` 方法
   - 实现方式：通过底层XML操作，查找并复制 `<p:timing>` 节点中的动画信息
   - 支持范围：幻灯片级别的动画效果

2. **视频/媒体文件复制**
   - 实现位置：`template_engine/engine.py` 的 `_copy_media_shape()` 方法
   - 实现方式：检测媒体形状类型（MSO_SHAPE_TYPE.MEDIA），尝试通过XML复制
   - 支持范围：视频、音频等媒体文件

### ⚠️ 限制说明

由于 `python-pptx` 库对动画和视频的支持有限，当前实现存在以下限制：

1. **动画支持**
   - ✅ 可以检测和复制动画XML结构
   - ⚠️ 复杂的动画效果可能需要进一步测试
   - ⚠️ 形状级别的动画可能需要额外处理

2. **视频支持**
   - ✅ 可以检测视频/媒体形状
   - ⚠️ 视频文件的完整复制可能需要访问底层ZIP结构
   - ⚠️ 视频文件的路径和引用关系需要正确维护

## 技术实现

### 动画复制

```python
def _copy_slide_animations(self, source_slide, target_slide):
    """复制幻灯片的动画效果到目标幻灯片"""
    # 查找 <p:timing> 节点
    timing_elements = source_element.xpath('.//p:timing', namespaces=source_element.nsmap)
    # 复制动画XML节点
    ...
```

### 视频复制

```python
def _copy_media_shape(self, source_shape, target_slide):
    """复制媒体形状（视频、音频等）到目标幻灯片"""
    # 检测媒体类型
    if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
        # 通过XML复制媒体信息
        ...
```

## 使用方法

使用方式与之前相同，系统会自动检测并复制动画和视频：

```bash
python app/generate_ppt.py examples/safety_slides_extended.json \
    -o output.pptx \
    --style safety \
    --template safety
```

## 故障排除

如果动画或视频没有正确复制：

1. **检查日志输出**：查看是否有警告信息
2. **验证源文件**：确认源PPT中确实包含动画和视频
3. **测试简单案例**：先用简单的动画/视频测试

## 未来改进

1. **更完整的动画支持**：支持形状级别的动画
2. **视频文件处理**：完整复制视频文件及其引用关系
3. **动画参数调整**：支持动画参数的修改和自定义

## 相关文件

- `template_engine/engine.py`: 核心实现
- `requirements.md`: 需求文档（第82行）

