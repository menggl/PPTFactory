#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据图片提示词准备生成图片提示词

功能：
1. 读取ppt内容映射.txt中的图片提示词准备
2. 结合大纲目录.txt和requirement.txt的内容
3. 使用大模型API生成图片提示词
4. 将生成的图片提示词保存到ppt内容映射.txt文件中
"""

import json
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
MAPPING_FILE = PROJECT_ROOT / "produce" / "ppt内容映射.txt"
OUTLINE_FILE = PROJECT_ROOT / "produce" / "大纲目录.txt"
REQUIREMENT_FILE = PROJECT_ROOT / "produce" / "requirement.txt"


def load_json_file(file_path: Path) -> List[Dict]:
    """加载JSON文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            # 处理可能的BOM
            if content.startswith('\ufeff'):
                content = content[1:]
            return json.loads(content)
    except Exception as e:
        print(f"错误：无法读取文件 {file_path}: {e}")
        sys.exit(1)


def save_json_file(file_path: Path, data: List[Dict]):
    """保存JSON文件"""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"✓ 已保存文件: {file_path}")
    except Exception as e:
        print(f"错误：无法保存文件 {file_path}: {e}")
        sys.exit(1)


def load_text_file(file_path: Path) -> str:
    """加载文本文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception as e:
        print(f"警告：无法读取文件 {file_path}: {e}")
        return ""


def generate_prompt_with_openai(
    image_prompt_prep: str,
    outline: str,
    requirement: str,
    api_key: Optional[str] = None,
    model: str = "gpt-4"
) -> Optional[str]:
    """
    使用OpenAI API生成图片提示词
    
    Args:
        image_prompt_prep: 图片提示词准备内容（替换文本|其他信息|图片大小）
        outline: 大纲目录内容
        requirement: 需求文档内容
        api_key: OpenAI API密钥（如果为None，从环境变量获取）
        model: 使用的模型名称
    
    Returns:
        生成的图片提示词，失败返回None
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
        system_prompt = """你是一个专业的PPT图片提示词生成专家。你的任务是根据给定的内容生成适合PPT演示的图片提示词。

要求：
1. 生成的图片必须适合在PPT中演示，符合PPT的展示风格（专业、清晰、简洁的商务或教育风格）
2. 图片内容要与给定的文本内容高度相关，准确反映主题
3. 图片大小信息已经在图片提示词准备中提供，生成的图片必须严格按照该尺寸生成
4. 提示词应该详细但简洁，能够指导AI图片生成工具（如DALL-E、Midjourney等）生成高质量的演示图片
5. 使用中文描述为主，可以适当包含英文关键词以提高生成质量
6. 确保生成的图片风格统一，适合PPT演示场景

输出格式：直接输出图片提示词，不要添加任何前缀、说明或引号。"""

        user_prompt = f"""请根据以下信息生成图片提示词：

【大纲目录】
{outline}

【需求文档】
{requirement}

【图片提示词准备内容】
{image_prompt_prep}

请生成一个适合PPT演示的图片提示词。要求：
1. 图片要符合PPT展示风格，专业、清晰、简洁
2. 图片内容要准确反映图片提示词准备中的文本内容
3. 图片大小必须严格按照图片提示词准备中指定的像素尺寸生成
4. 如果图片提示词准备中包含风格描述（如"极简线条"、"圆形图片"等），请在提示词中体现这些风格要求

直接输出图片提示词，不要添加任何前缀或说明。"""

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=500
        )
        
        return response.choices[0].message.content.strip()
        
    except ImportError:
        print("错误：未安装openai库，请运行: pip install openai")
        return None
    except Exception as e:
        print(f"OpenAI API调用失败: {e}")
        return None


def generate_prompt_manually(
    image_prompt_prep: str,
    outline: str,
    requirement: str
) -> str:
    """
    手动生成图片提示词（不使用API，基于规则）
    
    Args:
        image_prompt_prep: 图片提示词准备内容
        outline: 大纲目录内容
        requirement: 需求文档内容
    
    Returns:
        生成的图片提示词
    """
    # 解析图片提示词准备
    parts = image_prompt_prep.split('|')
    
    # 提取替换文本（排除其他信息和图片大小）
    replacement_texts = []
    other_info = []
    image_size = ""
    
    for part in parts:
        part = part.strip()
        if '图片尺寸为' in part or '像素' in part:
            image_size = part
        elif part and not part.startswith('图片尺寸'):
            # 判断是否是其他信息（如"极简线条"、"圆形图片"等）
            if any(keyword in part for keyword in ['线条', '圆形', '矩形', '图标', '图标']):
                other_info.append(part)
            else:
                replacement_texts.append(part)
    
    # 构建基础提示词
    prompt_parts = []
    
    # 添加风格描述
    style_keywords = []
    if '极简' in image_prompt_prep or '线条' in image_prompt_prep:
        style_keywords.append("极简线条风格")
    if '圆形' in image_prompt_prep:
        style_keywords.append("圆形设计")
    if '矩形' in image_prompt_prep:
        style_keywords.append("矩形布局")
    
    if style_keywords:
        prompt_parts.append("，".join(style_keywords))
    
    # 添加内容描述
    if replacement_texts:
        content_desc = "，".join(replacement_texts[:2])  # 取前两个替换文本
        if len(content_desc) > 100:
            content_desc = content_desc[:100] + "..."
        prompt_parts.append(f"主题：{content_desc}")
    
    # 添加PPT风格要求
    prompt_parts.append("PPT演示风格，专业商务插图，清晰简洁，适合教育展示")
    
    # 添加尺寸信息
    if image_size:
        prompt_parts.append(f"尺寸：{image_size}")
    
    return "，".join(prompt_parts)


def generate_image_prompts(
    use_api: bool = False,
    api_key: Optional[str] = None,
    model: str = "gpt-4"
):
    """
    生成图片提示词并更新映射文件
    
    Args:
        use_api: 是否使用API生成（True使用OpenAI API，False使用手动生成）
        api_key: OpenAI API密钥
        model: 使用的模型名称
    """
    print("=== 生成图片提示词 ===\n")
    
    # 1. 加载文件
    print("1. 加载文件...")
    mappings = load_json_file(MAPPING_FILE)
    outline = load_text_file(OUTLINE_FILE)
    requirement = load_text_file(REQUIREMENT_FILE)
    print(f"   ✓ 加载了 {len(mappings)} 个页面映射")
    print(f"   ✓ 大纲目录: {len(outline)} 字符")
    print(f"   ✓ 需求文档: {len(requirement)} 字符\n")
    
    # 2. 遍历每个页面，生成图片提示词
    print("2. 生成图片提示词...")
    
    # 检查是否使用API
    has_api_key = api_key is not None or os.getenv("OPENAI_API_KEY") is not None
    if use_api or has_api_key:
        print(f"   使用模式: OpenAI API ({model})")
        if not has_api_key:
            print("   ⚠ 警告: 未找到API密钥，将回退到手动生成模式")
    else:
        print("   使用模式: 手动生成（基于规则）")
    print()
    
    has_new_prompts = False
    
    for i, mapping in enumerate(mappings):
        page_num = i + 1
        template_id = mapping.get("模板页编号", f"T{page_num:03d}")
        
        # 获取图片提示词准备（兼容旧字段“图片标注映射”）
        image_prompt_prep = mapping.get("图片提示词准备", {})
        if not image_prompt_prep:
            legacy = mapping.get("图片标注映射", {})
            if legacy:
                image_prompt_prep = legacy
                mapping["图片提示词准备"] = legacy  # 迁移到新字段
                has_new_prompts = True
            else:
                print(f"   跳过第 {page_num} 页（模板 {template_id}）：无图片提示词准备")
                continue
        
        # 初始化图片提示词映射（如果不存在）
        image_prompts = mapping.get("图片提示词", {})
        if not image_prompts:
            image_prompts = {}
            mapping["图片提示词"] = image_prompts
            has_new_prompts = True
        
        print(f"\n   处理第 {page_num} 页（模板 {template_id}）:")
        
        # 遍历每个图片提示词准备
        for annotation_key, annotation_value in image_prompt_prep.items():
            # 如果已经存在提示词，跳过
            if annotation_key in image_prompts:
                print(f"     跳过: {annotation_key[:50]}...（已存在提示词）")
                continue
            
            print(f"     生成提示词: {annotation_key[:50]}...")
            
            # 生成图片提示词
            # 如果use_api为True或API密钥可用，优先使用API生成
            has_api_key = api_key is not None or os.getenv("OPENAI_API_KEY") is not None
            should_use_api = use_api or has_api_key
            
            if should_use_api:
                prompt = generate_prompt_with_openai(
                    annotation_value,
                    outline,
                    requirement,
                    api_key,
                    model
                )
                if not prompt:
                    print(f"       ⚠ API生成失败，使用手动生成")
                    prompt = generate_prompt_manually(
                        annotation_value,
                        outline,
                        requirement
                    )
            else:
                prompt = generate_prompt_manually(
                    annotation_value,
                    outline,
                    requirement
                )
            
            if prompt:
                image_prompts[annotation_key] = prompt
                has_new_prompts = True
                print(f"       ✓ 生成成功: {prompt[:80]}...")
            else:
                print(f"       ⚠ 生成失败")
    
    # 3. 保存更新后的映射文件
    if has_new_prompts:
        print("\n3. 保存映射文件...")
        save_json_file(MAPPING_FILE, mappings)
        print("\n✓ 图片提示词生成完成！")
    else:
        print("\n3. 无需更新映射文件（所有图片提示词已存在）")


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="生成图片提示词")
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
    
    generate_image_prompts(
        use_api=args.use_api,
        api_key=args.api_key,
        model=args.model
    )


if __name__ == "__main__":
    main()
