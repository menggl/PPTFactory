# WPS Office 工具支持

## 概述

在 macOS 系统上，PPT模板引擎现在支持使用 WPS Office 工具来复制幻灯片。这可以作为 `python-pptx` 库的备选方案，在某些情况下可能提供更好的样式保留效果。

## 功能说明

### 当前实现状态

- ✅ **系统检测**：自动检测当前系统是否为 macOS
- ✅ **WPS 检测**：自动检测 WPS Office 是否已安装
- ✅ **命令行参数**：添加了 `--use-wps` 参数来启用 WPS 支持
- ⚠️ **WPS 工具调用**：框架已实现，但由于 WPS Office 的 AppleScript 支持可能有限，当前会回退到 `python-pptx` 方法

### 工作原理

1. **检测阶段**：
   - 检查系统是否为 macOS
   - 检查 WPS Office 是否安装在以下位置：
     - `/Applications/WPS Office.app`
     - `/Applications/Kingsoft Office.app`
     - `~/Applications/WPS Office.app`

2. **复制阶段**：
   - 如果启用 `--use-wps` 且 WPS 可用，尝试使用 WPS 工具
   - 如果 WPS 方法失败或不可用，自动回退到 `python-pptx` 方法
   - 确保无论哪种方法，都能成功复制幻灯片

## 使用方法

### 基本用法

```bash
# 使用 WPS 工具复制幻灯片（如果可用）
python app/generate_ppt.py examples/safety_slides_extended.json \
    -o output.pptx \
    --style safety \
    --template safety \
    --use-wps
```

### 不使用 WPS（默认）

```bash
# 默认使用 python-pptx 方法（跨平台）
python app/generate_ppt.py examples/safety_slides_extended.json \
    -o output.pptx \
    --style safety \
    --template safety
```

## 技术细节

### WPS 检测方法

```python
def _is_wps_available(self) -> bool:
    """检测 WPS Office 是否在 macOS 上可用"""
    if not self._is_macos():
        return False
    
    wps_paths = [
        '/Applications/WPS Office.app',
        '/Applications/Kingsoft Office.app',
        os.path.expanduser('~/Applications/WPS Office.app')
    ]
    
    for path in wps_paths:
        if os.path.exists(path):
            return True
    
    return False
```

### WPS 复制方法

当前实现尝试使用 AppleScript 控制 WPS Office，但由于 WPS 的 AppleScript 支持可能不完整，会自动回退到 `python-pptx` 方法。

如果将来 WPS Office 提供了更好的命令行工具或 API，可以在此处扩展实现。

## 未来改进

1. **WPS 命令行工具**：如果 WPS 提供命令行工具，可以直接调用
2. **更好的 AppleScript 支持**：如果 WPS 改进 AppleScript 支持，可以完善脚本
3. **其他工具支持**：可以添加对 LibreOffice、Keynote 等其他工具的支持

## 注意事项

1. **macOS 专用**：WPS 工具支持仅在 macOS 系统上可用
2. **自动回退**：如果 WPS 不可用或失败，会自动使用 `python-pptx` 方法
3. **样式保留**：两种方法都会尽力保留幻灯片的样式和格式
4. **性能**：`python-pptx` 方法通常更快，因为不需要启动外部应用程序

## 相关文件

- `template_engine/engine.py`：核心实现，包含 WPS 检测和调用逻辑
- `app/generate_ppt.py`：命令行接口，包含 `--use-wps` 参数

## 示例输出

```
使用参考PPT的第一页: 1.2 安全生产方针政策.pptx
已启用 WPS 工具支持（如果可用）
尝试使用 WPS Office 工具复制幻灯片...
WPS 工具不可用或失败，使用 python-pptx 方法
已使用参考PPT的第一页: 1.2 安全生产方针政策.pptx
```

