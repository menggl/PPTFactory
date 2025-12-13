#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据图片提示词生成图片并更新映射文件
"""

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List

# 可选依赖
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

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


def generate_image_with_openai(prompt: str, api_key: str, model: str = "dall-e-3") -> Optional[bytes]:
    """
    使用OpenAI DALL-E生成图片
    
    Args:
        prompt: 图片提示词
        api_key: OpenAI API密钥
        model: 模型名称 (dall-e-2 或 dall-e-3)
    
    Returns:
        图片的二进制数据，失败返回None
    """
    try:
        import openai
        
        client = openai.OpenAI(api_key=api_key)
        
        response = client.images.generate(
            model=model,
            prompt=prompt,
            n=1,
            size="1024x1024" if model == "dall-e-2" else "1024x1792",  # dall-e-3支持1024x1792
            quality="standard",
            response_format="b64_json"
        )
        
        import base64
        image_data = base64.b64decode(response.data[0].b64_json)
        return image_data
        
    except Exception as e:
        print(f"OpenAI API调用失败: {e}")
        return None


def generate_image_with_stable_diffusion(prompt: str, api_url: str = "http://localhost:7860") -> Optional[bytes]:
    """
    使用本地Stable Diffusion API生成图片
    
    Args:
        prompt: 图片提示词
        api_url: Stable Diffusion API地址
    
    Returns:
        图片的二进制数据，失败返回None
    """
    if not HAS_REQUESTS:
        print("错误: 需要安装 requests 库才能使用 Stable Diffusion API")
        return None
    
    try:
        payload = {
            "prompt": prompt,
            "negative_prompt": "blurry, low quality, distorted, watermark",
            "steps": 20,
            "width": 1024,
            "height": 576,  # 16:9比例
            "cfg_scale": 7,
            "sampler_index": "DPM++ 2M Karras"
        }
        
        response = requests.post(
            f"{api_url}/sdapi/v1/txt2img",
            json=payload,
            timeout=300
        )
        
        if response.status_code == 200:
            result = response.json()
            import base64
            image_data = base64.b64decode(result["images"][0])
            return image_data
        else:
            print(f"Stable Diffusion API调用失败: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Stable Diffusion API调用失败: {e}")
        return None


def generate_image_placeholder(prompt: str, output_path: Path) -> bool:
    """
    生成占位图片（用于测试，实际使用时应该替换为真实的图片生成逻辑）
    
    Args:
        prompt: 图片提示词
        output_path: 输出路径
    
    Returns:
        是否成功
    """
    if HAS_PIL:
        try:
            from PIL import Image, ImageDraw, ImageFont
            
            # 创建16:9的图片
            width, height = 1920, 1080
            image = Image.new('RGB', (width, height), color='#E8E8E8')
            draw = ImageDraw.Draw(image)
            
            # 绘制边框
            draw.rectangle([10, 10, width-10, height-10], outline='#CCCCCC', width=3)
            
            # 绘制提示词文本（简化版）
            try:
                font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 40)
            except:
                font = ImageFont.load_default()
            
            # 文本换行
            words = prompt[:100].split('，')  # 简化显示
            y = height // 2 - 100
            for i, word in enumerate(words[:3]):
                draw.text((width // 2 - 200, y + i * 60), word, fill='#666666', font=font)
            
            # 保存图片
            image.save(output_path, 'PNG')
            return True
            
        except Exception as e:
            print(f"生成占位图片失败: {e}")
            return False
    else:
        # 如果没有PIL，创建一个最小的有效PNG文件作为占位符
        try:
            # 创建一个1x1像素的透明PNG
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
            print("  注意: 已创建最小PNG占位文件（建议安装Pillow生成更好的占位图片）")
            return True
        except Exception as e:
            print(f"创建占位文件失败: {e}")
            return False


def generate_images_from_prompts(
    ppt_filename: str,
    use_openai: bool = False,
    openai_api_key: Optional[str] = None,
    use_stable_diffusion: bool = False,
    stable_diffusion_url: str = "http://localhost:7860",
    use_placeholder: bool = True
) -> Dict[str, Dict[str, str]]:
    """
    根据图片提示词生成图片
    
    Args:
        ppt_filename: PPT文件名（不含扩展名）
        use_openai: 是否使用OpenAI API
        openai_api_key: OpenAI API密钥
        use_stable_diffusion: 是否使用Stable Diffusion
        stable_diffusion_url: Stable Diffusion API地址
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
            
            # 生成图片
            image_data = None
            
            if use_openai and openai_api_key:
                image_data = generate_image_with_openai(prompt, openai_api_key)
            elif use_stable_diffusion:
                image_data = generate_image_with_stable_diffusion(prompt, stable_diffusion_url)
            
            # 保存图片
            if image_data:
                with open(image_path, 'wb') as f:
                    f.write(image_data)
                print(f"  ✓ 图片已保存: {relative_path}")
            elif use_placeholder:
                # 使用占位图片
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
    
    # 配置图片生成方式
    # 方式1: 使用OpenAI DALL-E（需要API密钥）
    use_openai = False
    openai_api_key = os.getenv("OPENAI_API_KEY")
    
    # 方式2: 使用本地Stable Diffusion
    use_stable_diffusion = False
    stable_diffusion_url = "http://localhost:7860"
    
    # 方式3: 使用占位图片（用于测试）
    use_placeholder = True
    
    # 检查是否有可用的图片生成方式
    if not use_openai and not use_stable_diffusion and not use_placeholder:
        print("错误: 未配置任何图片生成方式")
        print("请设置以下环境变量之一:")
        print("  - OPENAI_API_KEY: 使用OpenAI DALL-E")
        print("或者启用 use_stable_diffusion 使用本地Stable Diffusion")
        print("或者启用 use_placeholder 使用占位图片（测试用）")
        return
    
    # 生成图片
    print("开始生成图片...\n")
    new_image_mapping = generate_images_from_prompts(
        ppt_filename=ppt_filename,
        use_openai=use_openai,
        openai_api_key=openai_api_key,
        use_stable_diffusion=use_stable_diffusion,
        stable_diffusion_url=stable_diffusion_url,
        use_placeholder=use_placeholder
    )
    
    if not new_image_mapping:
        print("\n错误: 未生成任何图片")
        return
    
    print(f"\n✓ 共生成 {len(new_image_mapping)} 张图片")
    
    # 更新映射文件
    print("\n更新映射文件...")
    update_mapping_file(new_image_mapping)
    
    print("\n✓ 完成！")


if __name__ == "__main__":
    main()

