#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据requirement.txt和大纲目录.txt生成PPT内容页信息列表

功能：
1. 读取requirement.txt和大纲目录.txt
2. 读取所有模板文件，了解模板结构
3. 使用大模型API生成PPT内容页信息
4. 为每个大纲条目选择合适的模板
5. 生成JSON格式的内容页信息列表
6. 保存到ppt内容页.txt文件
"""

import json
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
REQUIREMENT_FILE = PROJECT_ROOT / "produce" / "requirement.txt"
OUTLINE_FILE = PROJECT_ROOT / "produce" / "大纲目录.txt"
PROMPT_FILE = PROJECT_ROOT / "produce" / "生成ppt_slide内容页的提示词.txt"
OUTPUT_FILE = PROJECT_ROOT / "produce" / "ppt内容页.txt"
TEMPLATES_DIR = PROJECT_ROOT / "templates" / "metadata"


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


def load_all_templates() -> List[Dict]:
    """加载所有模板文件"""
    templates = []
    template_files = sorted(TEMPLATES_DIR.glob("T*.json"))
    
    for template_file in template_files:
        try:
            with open(template_file, 'r', encoding='utf-8') as f:
                template = json.load(f)
                templates.append(template)
        except Exception as e:
            print(f"警告：无法读取模板文件 {template_file}: {e}")
    
    return templates


def format_templates_info(templates: List[Dict]) -> str:
    """格式化模板信息，用于提示词"""
    info_lines = []
    for template in templates:
        template_id = template.get("template_id", "")
        layout_type = template.get("layout_type", "")
        text_capacity = template.get("text_capacity_score", 0)
        recommended = template.get("recommended_usage", [])
        example = template.get("example_text", {})
        
        # 统计文本占位符数量
        placeholders = template.get("placeholders", [])
        text_placeholders = [p for p in placeholders if p.get("type") in 
                           ["title", "section_title", "subtitle", "heading", "body_text", 
                            "paragraph", "caption", "body_text_block", "multi_text_block"]]
        
        info_lines.append(f"模板 {template_id} ({layout_type}):")
        info_lines.append(f"  - 文本容量分数: {text_capacity}")
        info_lines.append(f"  - 文本占位符数量: {len(text_placeholders)}")
        info_lines.append(f"  - 推荐用途: {', '.join(recommended[:3])}")
        if example:
            info_lines.append(f"  - 示例: {json.dumps(example, ensure_ascii=False)[:100]}...")
        info_lines.append("")
    
    return "\n".join(info_lines)


def generate_slides_with_openai(
    requirement_text: str,
    outline_text: str,
    prompt_template: str,
    templates_info: str,
    api_key: Optional[str] = None,
    model: str = "gpt-4"
) -> Optional[str]:
    """
    使用OpenAI API生成PPT内容页信息列表
    
    Args:
        requirement_text: 需求文档内容
        outline_text: 大纲目录内容
        prompt_template: 提示词模板
        templates_info: 模板信息
        api_key: OpenAI API密钥
        model: 使用的模型名称
    
    Returns:
        生成的JSON格式内容页信息列表，失败返回None
    """
    try:
        import openai
        
        if api_key is None:
            api_key = os.getenv("OPENAI_API_KEY")
        
        if not api_key:
            print("错误：未设置OPENAI_API_KEY环境变量")
            return None
        
        client = openai.OpenAI(api_key=api_key)
        
        # 构建系统提示词
        system_prompt = f"""你是一个专业的PPT内容生成专家。你的任务是根据原始文本内容和大纲目录，为每个大纲条目生成对应的PPT内容页信息。

{prompt_template}

【可用模板信息】
{templates_info}

重要要求：
1. 必须为每个大纲条目生成对应的内容页信息
2. 必须选择合适的模板（templateId），优先匹配文本数量相同的模板
3. 必须包含templateId字段
4. 如果副标题是"总览（上）"或"总览（下）"，则slide_subtitle设为空字符串
5. 输出必须是有效的JSON数组格式
6. 每个对象必须包含：slide_title, slide_subtitle, layout, key_points, templateId
7. key_points的结构必须匹配所选模板的占位符类型
8. 如果内容太多，需要拆分成多个页面，使用相同的模板"""

        user_prompt = f"""请根据以下信息生成PPT内容页信息列表：

【原始文本内容】
{requirement_text}

【大纲目录】
{outline_text}

请为每个大纲条目生成对应的PPT内容页信息，选择合适的模板，输出JSON数组格式。确保：
1. 每个大纲条目至少对应一页
2. 内容过多时拆分成多页
3. 选择合适的模板（根据文本数量和内容类型）
4. 严格按照模板要求组织key_points的结构
5. 输出有效的JSON数组，不要添加任何前缀或说明文字"""

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=4000
        )
        
        result = response.choices[0].message.content.strip()
        
        # 尝试提取JSON（去除可能的markdown代码块标记）
        if result.startswith("```json"):
            result = result[7:]
        if result.startswith("```"):
            result = result[3:]
        if result.endswith("```"):
            result = result[:-3]
        result = result.strip()
        
        # 验证JSON格式
        try:
            json.loads(result)
            return result
        except json.JSONDecodeError as e:
            print(f"警告：生成的JSON格式可能有问题: {e}")
            print(f"生成的内容前500字符: {result[:500]}")
            return result
        
    except ImportError:
        print("错误：未安装openai库，请运行: pip install openai")
        return None
    except Exception as e:
        print(f"OpenAI API调用失败: {e}")
        return None


def generate_ppt_slides(
    use_api: bool = False,
    api_key: Optional[str] = None,
    model: str = "gpt-4"
):
    """
    生成PPT内容页信息列表并保存到文件
    
    Args:
        use_api: 是否使用API生成（True使用OpenAI API，False提示手动生成）
        api_key: OpenAI API密钥
        model: 使用的模型名称
    """
    print("=== 生成PPT内容页信息列表 ===\n")
    
    # 1. 加载文件
    print("1. 加载文件...")
    requirement_text = load_text_file(REQUIREMENT_FILE)
    outline_text = load_text_file(OUTLINE_FILE)
    prompt_template = load_text_file(PROMPT_FILE)
    
    print(f"   ✓ 需求文档: {len(requirement_text)} 字符")
    print(f"   ✓ 大纲目录: {len(outline_text)} 字符")
    print(f"   ✓ 提示词模板: {len(prompt_template)} 字符")
    
    # 2. 加载模板信息
    print("\n2. 加载模板信息...")
    templates = load_all_templates()
    templates_info = format_templates_info(templates)
    print(f"   ✓ 加载了 {len(templates)} 个模板")
    
    print()
    
    # 3. 生成内容页信息
    print("3. 生成PPT内容页信息...")
    
    # 检查是否使用API
    has_api_key = api_key is not None or os.getenv("OPENAI_API_KEY") is not None
    
    if not use_api and not has_api_key:
        print("   ⚠ 未设置API密钥，无法自动生成")
        print("   请使用以下方式之一：")
        print("   1. 设置环境变量: export OPENAI_API_KEY='your-api-key'")
        print("   2. 使用参数: --use-api --api-key 'your-api-key'")
        print("   3. 手动使用大模型生成内容，然后保存到 ppt内容页.txt 文件")
        sys.exit(1)
    
    if use_api or has_api_key:
        print(f"   使用模式: OpenAI API ({model})")
        if not has_api_key:
            print("   ⚠ 警告: 未找到API密钥")
            sys.exit(1)
    print()
    
    slides_json = generate_slides_with_openai(
        requirement_text,
        outline_text,
        prompt_template,
        templates_info,
        api_key,
        model
    )
    
    if not slides_json:
        print("   ✗ 生成失败")
        sys.exit(1)
    
    # 4. 验证和格式化JSON
    print("4. 验证和格式化JSON...")
    try:
        slides_data = json.loads(slides_json)
        formatted_json = json.dumps(slides_data, ensure_ascii=False, indent=2)
        print(f"   ✓ 生成了 {len(slides_data)} 个内容页")
    except json.JSONDecodeError as e:
        print(f"   ⚠ JSON格式错误: {e}")
        print("   保存原始内容，请手动修复")
        formatted_json = slides_json
    
    # 5. 保存结果
    print("\n5. 保存结果...")
    save_text_file(OUTPUT_FILE, formatted_json)
    print("\n✓ PPT内容页信息列表生成完成！")


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="生成PPT内容页信息列表")
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
    
    generate_ppt_slides(
        use_api=args.use_api,
        api_key=args.api_key,
        model=args.model
    )


if __name__ == "__main__":
    main()








