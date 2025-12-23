#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据图片提示词生成图片并更新映射文件
"""

import json
import os
from pathlib import Path
from typing import Optional, Dict

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
MAPPING_FILE = PROJECT_ROOT / "produce" / "ppt内容映射.txt"
IMAGES_DIR = PROJECT_ROOT / "produce" / "images"


def find_latest_ppt_file() -> Optional[str]:
    """查找最新的PPT文件名"""
    produce_dir = PROJECT_ROOT / "produce"
    ppt_files = list(produce_dir.glob("new_ppt_*.pptx"))
    
    if not ppt_files:
        return None
    
    # 按文件名排序（文件名包含时间戳）
    ppt_files.sort(reverse=True)
    latest_file = ppt_files[0]
    
    # 提取文件名（不含扩展名）
    return latest_file.stem


def create_image_directory(ppt_filename: str) -> Path:
    """创建图片保存目录"""
    image_dir = IMAGES_DIR / ppt_filename
    image_dir.mkdir(parents=True, exist_ok=True)
    return image_dir


def generate_image_placeholder(prompt: str, output_path: Path) -> bool:
    """
    生成占位图片（创建一个最小的有效PNG文件）
    
    Args:
        prompt: 图片提示词
        output_path: 输出路径
    
    Returns:
        是否成功
    """
    try:
        # 创建一个1x1像素的透明PNG（最小有效PNG文件）
        png_data = (
            b'\x89PNG\r\n\x1a\n'  # PNG签名
            b'\x00\x00\x00\r'  # IHDR块长度
            b'IHDR'  # IHDR标识
            b'\x00\x00\x00\x01'  # 宽度1
            b'\x00\x00\x00\x01'  # 高度1
            b'\x08\x06\x00\x00\x00'  # 位深度8, 颜色类型RGBA
            b'\x1f\x15\xc4\x89'  # CRC
            b'\x00\x00\x00\n'  # IDAT块长度
            b'IDAT'  # IDAT标识
            b'x\x9c\x63\x00\x01\x00\x00\x05\x00\x01'  # 压缩数据
            b'\r\n-\xdc'  # CRC
            b'\x00\x00\x00\x00'  # IEND块长度
            b'IEND'  # IEND标识
            b'\xaeB`\x82'  # CRC
        )
        with open(output_path, 'wb') as f:
            f.write(png_data)
        return True
    except Exception as e:
        print(f"创建占位文件失败: {e}")
        return False


def generate_images_from_prompts(
    ppt_filename: str,
    use_placeholder: bool = True
) -> Dict[int, Dict[str, str]]:
    """
    根据图片提示词生成图片
    
    Args:
        ppt_filename: PPT文件名（不含扩展名）
        use_placeholder: 是否使用占位图片（用于测试）
    
    Returns:
        图片映射字典 {页面索引: {图片标注: 图片文件路径}}
    """
    # 读取映射文件
    with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 创建图片目录
    image_dir = create_image_directory(ppt_filename)
    
    # 图片映射结果 {页面索引: {图片标注: 图片路径}}
    new_image_mapping = {}
    
    # 遍历每个页面
    for page_index, page in enumerate(data, start=1):
        if '图片提示词' not in page or not page['图片提示词']:
            continue
        
        if page_index not in new_image_mapping:
            new_image_mapping[page_index] = {}
        
        image_index = 1
        
        # 遍历每个图片提示词
        for annotation, prompt in page['图片提示词'].items():
            # 生成图片文件名：第几页_第几张图片.png
            image_filename = f"{page_index}_{image_index}.png"
            image_path = image_dir / image_filename
            relative_path = f"images/{ppt_filename}/{image_filename}"
            
            print(f"正在生成图片: 第{page_index}页 第{image_index}张 - {annotation[:30]}...")
            
            # 生成占位图片
            if use_placeholder:
                if generate_image_placeholder(prompt, image_path):
                    print(f"  ✓ 占位图片已保存: {relative_path}")
                else:
                    print(f"  ✗ 图片生成失败")
                    continue
            else:
                print(f"  ✗ 图片生成失败（未启用占位图片）")
                continue
            
            # 保存映射关系
            new_image_mapping[page_index][annotation] = relative_path
            image_index += 1
    
    return new_image_mapping


def update_mapping_file(new_image_mapping: Dict[int, Dict[str, str]]):
    """
    更新映射文件，添加新图片映射
    
    Args:
        new_image_mapping: 图片映射字典，{页面索引: {图片标注: 图片路径}}
    """
    # 读取映射文件
    with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 为每个页面添加新图片映射
    for page_index, page in enumerate(data, start=1):
        if '图片提示词' not in page or not page['图片提示词']:
            continue
        
        # 初始化新图片映射
        if '新图片映射' not in page:
            page['新图片映射'] = {}
        
        # 获取该页面的图片映射
        if page_index in new_image_mapping:
            page['新图片映射'].update(new_image_mapping[page_index])
    
    # 保存更新后的映射文件
    with open(MAPPING_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n✓ 映射文件已更新: {MAPPING_FILE}")


def main():
    """主函数"""
    print("=== 图片生成工具 ===\n")
    
    # 查找最新的PPT文件
    ppt_filename = find_latest_ppt_file()
    if not ppt_filename:
        print("错误: 未找到PPT文件 (new_ppt_*.pptx)")
        return
    
    print(f"找到PPT文件: {ppt_filename}.pptx\n")
    
    # 使用占位图片（用于测试）
    # 实际使用时，可以替换为真实的图片生成API调用
    use_placeholder = True
    
    # 生成图片
    print("开始生成图片...\n")
    new_image_mapping = generate_images_from_prompts(
        ppt_filename=ppt_filename,
        use_placeholder=use_placeholder
    )
    
    if not new_image_mapping:
        print("\n错误: 未生成任何图片")
        return
    
    total_images = sum(len(mapping) for mapping in new_image_mapping.values())
    print(f"\n✓ 共生成 {total_images} 张图片")
    
    # 更新映射文件
    print("\n更新映射文件...")
    update_mapping_file(new_image_mapping)
    
    print("\n✓ 完成！")
    print("\n注意: 当前使用的是占位图片。")
    print("要使用真实的图片生成，请:")
    print("1. 配置OpenAI API密钥: export OPENAI_API_KEY=your_key")
    print("2. 或配置本地Stable Diffusion API")
    print("3. 修改脚本中的图片生成逻辑")


if __name__ == "__main__":
    main()











