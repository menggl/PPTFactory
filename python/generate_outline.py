#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据requirement.txt生成大纲目录

功能：
1. 读取requirement.txt中的原始文本内容
2. 使用大模型API或手动解析生成大纲目录
3. 将生成的大纲目录保存到大纲目录.txt文件中
"""

import json
import os
import sys
import re
from pathlib import Path
from typing import Optional, List, Dict

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
REQUIREMENT_FILE = PROJECT_ROOT / "produce" / "requirement.txt"
OUTLINE_FILE = PROJECT_ROOT / "produce" / "大纲目录.txt"
OUTLINE_PROMPT_FILE = PROJECT_ROOT / "produce" / "生成大纲目录提示词.txt"


def load_text_file(file_path: Path) -> str:
    """加载文本文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception as e:
        print(f"错误：无法读取文件 {file_path}: {e}")
        sys.exit(1)


def save_text_file(file_path: Path, content: str):
    """保存文本文件"""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✓ 已保存文件: {file_path}")
    except Exception as e:
        print(f"错误：无法保存文件 {file_path}: {e}")
        sys.exit(1)


def generate_outline_with_openai(
    requirement_text: str,
    prompt_template: str = "",
    api_key: Optional[str] = None,
    model: str = "gpt-4"
) -> Optional[str]:
    """
    使用OpenAI API生成大纲目录
    
    Args:
        requirement_text: 需求文档内容
        prompt_template: 提示词模板（如果提供）
        api_key: OpenAI API密钥（如果为None，从环境变量获取）
        model: 使用的模型名称
    
    Returns:
        生成的大纲目录，失败返回None
    """
    try:
        import openai
        
        if api_key is None:
            api_key = os.getenv("OPENAI_API_KEY")
        
        if not api_key:
            print("错误：未设置OPENAI_API_KEY环境变量")
            return None
        
        client = openai.OpenAI(api_key=api_key)
        
        # 构建提示词
        if prompt_template:
            system_prompt = prompt_template
        else:
            system_prompt = """你是一个专业的PPT大纲目录生成专家。你的任务是根据给定的原始文本内容，生成一个清晰、结构化的PPT大纲目录。

要求：
1. 大纲目录应该层次清晰，包含主标题和子标题
2. 主标题使用中文大写序号（一、二、三、四...）
3. 子标题使用阿拉伯数字序号（1. 2. 3. ...）
4. 大纲应该覆盖原始文本的所有主要内容
5. 结构要合理，适合制作PPT演示文稿
6. 输出格式：直接输出大纲目录，每行一个条目，使用适当的缩进表示层级

输出格式示例：
标题名称

一、第一章标题
1. 子标题1
2. 子标题2

二、第二章标题
1. 子标题1
2. 子标题2

三、第三章标题
1. 子标题1
2. 子标题2

四、总结"""

        user_prompt = f"""请根据以下原始文本内容，生成一个清晰的PPT大纲目录：

【原始文本内容】
{requirement_text}

请生成一个结构化的PPT大纲目录，包含主标题和子标题，适合制作演示文稿。直接输出大纲目录，不要添加任何前缀或说明。"""

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=1000
        )
        
        return response.choices[0].message.content.strip()
        
    except ImportError:
        print("错误：未安装openai库，请运行: pip install openai")
        return None
    except Exception as e:
        print(f"OpenAI API调用失败: {e}")
        return None


def generate_outline_manually(requirement_text: str) -> str:
    """
    手动解析生成大纲目录（基于规则）
    
    Args:
        requirement_text: 需求文档内容
    
    Returns:
        生成的大纲目录
    """
    lines = requirement_text.split('\n')
    outline_lines = []
    current_section = None
    subsection_num = 0
    
    # 提取标题
    title = lines[0].strip() if lines else "PPT大纲目录"
    outline_lines.append(title)
    outline_lines.append("")
    
    i = 1
    while i < len(lines):
        line = lines[i].strip()
        
        # 跳过空行
        if not line:
            i += 1
            continue
        
        # 检测主标题（一、二、三、四...）
        main_title_match = re.match(r'^([一二三四五六七八九十]+)、(.+)$', line)
        if main_title_match:
            if current_section:
                outline_lines.append("")
            current_section = line
            outline_lines.append(line)
            subsection_num = 0
            i += 1
            continue
        
        # 检测子标题（数字序号）
        sub_title_match = re.match(r'^(\d+)\.\s*(.+)$', line)
        if sub_title_match:
            if current_section:
                subsection_num += 1
                outline_lines.append(f"{subsection_num}. {sub_title_match.group(2)}")
            i += 1
            continue
        
        # 检测可能的子标题（没有数字，但可能是关键词）
        # 判断标准：
        # 1. 行长度较短（通常子标题不超过30字）
        # 2. 不包含句号（子标题通常是短语，不是完整句子）
        # 3. 下一行是详细内容（包含句号或较长）
        # 4. 包含常见的关键词（权、义务、案例、总结等）
        is_potential_subtitle = (
            len(line) < 30 and 
            '。' not in line and
            current_section is not None
        )
        
        # 检查下一行是否是详细内容
        next_line_is_detail = False
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line and ('。' in next_line or len(next_line) > 30):
                next_line_is_detail = True
        
        # 检查是否包含常见的关键词
        has_keyword = any(keyword in line for keyword in [
            '权', '义务', '案例', '总结', '要点', '规程', '管理', 
            '用品', '隐患', '培训', '环境', '设备', '操作'
        ])
        
        if is_potential_subtitle and (next_line_is_detail or has_keyword):
            subsection_num += 1
            outline_lines.append(f"{subsection_num}. {line}")
            # 跳过下一行（详细内容）
            if next_line_is_detail:
                i += 2
            else:
                i += 1
            continue
        
        i += 1
    
    return '\n'.join(outline_lines)


def generate_outline(
    use_api: bool = False,
    api_key: Optional[str] = None,
    model: str = "gpt-4"
):
    """
    生成大纲目录并保存到文件
    
    Args:
        use_api: 是否使用API生成（True使用OpenAI API，False使用手动生成）
        api_key: OpenAI API密钥
        model: 使用的模型名称
    """
    print("=== 生成大纲目录 ===\n")
    
    # 1. 加载文件
    print("1. 加载文件...")
    requirement_text = load_text_file(REQUIREMENT_FILE)
    print(f"   ✓ 需求文档: {len(requirement_text)} 字符")
    
    # 尝试加载提示词模板
    prompt_template = ""
    if OUTLINE_PROMPT_FILE.exists():
        try:
            prompt_template = load_text_file(OUTLINE_PROMPT_FILE)
            print(f"   ✓ 加载了提示词模板: {len(prompt_template)} 字符")
        except:
            print("   ⚠ 无法加载提示词模板，使用默认提示词")
    
    print()
    
    # 2. 生成大纲目录
    print("2. 生成大纲目录...")
    
    # 检查是否使用API
    has_api_key = api_key is not None or os.getenv("OPENAI_API_KEY") is not None
    if use_api or has_api_key:
        print(f"   使用模式: OpenAI API ({model})")
        if not has_api_key:
            print("   ⚠ 警告: 未找到API密钥，将回退到手动生成模式")
            use_api = False
    else:
        print("   使用模式: 手动生成（基于规则）")
    print()
    
    if use_api:
        outline = generate_outline_with_openai(
            requirement_text,
            prompt_template,
            api_key,
            model
        )
        if not outline:
            print("   ⚠ API生成失败，使用手动生成")
            outline = generate_outline_manually(requirement_text)
    else:
        outline = generate_outline_manually(requirement_text)
    
    if not outline:
        print("   ✗ 生成失败")
        sys.exit(1)
    
    print("   ✓ 生成成功")
    print("\n生成的大纲目录预览：")
    print("-" * 50)
    preview_lines = outline.split('\n')[:20]  # 显示前20行
    for line in preview_lines:
        print(line)
    if len(outline.split('\n')) > 20:
        print("...")
    print("-" * 50)
    print()
    
    # 3. 保存大纲目录
    print("3. 保存大纲目录...")
    save_text_file(OUTLINE_FILE, outline)
    print("\n✓ 大纲目录生成完成！")


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="生成PPT大纲目录")
    parser.add_argument(
        "--use-api",
        action="store_true",
        help="使用OpenAI API生成（需要设置OPENAI_API_KEY环境变量）"
    )
    parser.add_argument(
        "--api-key",
        type=str,
        default=None,
        help="OpenAI API密钥（如果不提供，从环境变量OPENAI_API_KEY获取）"
    )
    parser.add_argument(
        "--model",
        type=str,
        default="gpt-4",
        help="使用的模型名称（默认: gpt-4）"
    )
    
    args = parser.parse_args()
    
    generate_outline(
        use_api=args.use_api,
        api_key=args.api_key,
        model=args.model
    )


if __name__ == "__main__":
    main()








