#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import zlib
import struct
from pathlib import Path

def create_visible_placeholder_png(width, height, page_num, image_num, annotation=""):
    """创建一个可见的PNG占位图片，带有明显的颜色和图案"""
    # PNG文件头
    png_signature = b'\x89PNG\r\n\x1a\n'
    
    # 创建RGB图像数据
    pixels = []
    for y in range(height):
        row = []
        for x in range(width):
            # 使用明显的浅灰色背景（高对比度）
            r, g, b = 240, 240, 240
            
            # 添加明显的网格图案（每100像素一条粗线）
            if x % 100 == 0 or y % 100 == 0:
                r, g, b = 150, 150, 150  # 深灰色网格线
            elif x % 50 == 0 or y % 50 == 0:
                r, g, b = 200, 200, 200  # 浅灰色网格线
            
            # 中心区域添加明显的边框和内容区域
            center_x, center_y = width // 2, height // 2
            border_width = 800
            border_height = 600
            
            # 外边框（深灰色，高对比度）
            if (abs(x - center_x) < border_width and abs(y - center_y) < border_height):
                if (abs(x - center_x) > border_width - 30 or abs(y - center_y) > border_height - 30):
                    r, g, b = 100, 100, 100  # 深灰色边框（高对比度）
                elif (abs(x - center_x) < border_width - 60 and abs(y - center_y) < border_height - 60):
                    # 中心内容区域（纯白色）
                    r, g, b = 255, 255, 255
                    # 添加细网格线
                    if (x - center_x) % 50 == 0 or (y - center_y) % 50 == 0:
                        r, g, b = 245, 245, 245
            
            # 四角添加明显的标识块（深灰色）
            corner_size = 150
            if (x < corner_size and y < corner_size) or (x >= width - corner_size and y < corner_size) or \
               (x < corner_size and y >= height - corner_size) or (x >= width - corner_size and y >= height - corner_size):
                r, g, b = 120, 120, 120  # 深灰色四角标识
            
            row.extend([r, g, b])
        # PNG行过滤器：0 = None
        row_data = bytes([0] + row)
        pixels.append(row_data)
    
    # 压缩图像数据
    image_data = b''.join(pixels)
    compressed = zlib.compress(image_data, level=6)
    
    # 计算CRC32
    def crc32(data):
        return zlib.crc32(data) & 0xffffffff
    
    # IHDR块
    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)  # RGB, 8-bit
    ihdr_chunk = b'IHDR' + ihdr_data
    ihdr_crc = struct.pack('>I', crc32(ihdr_chunk))
    ihdr = struct.pack('>I', len(ihdr_data)) + ihdr_chunk + ihdr_crc
    
    # IDAT块
    idat_chunk = b'IDAT' + compressed
    idat_crc = struct.pack('>I', crc32(idat_chunk))
    idat = struct.pack('>I', len(compressed)) + idat_chunk + idat_crc
    
    # tEXt块 - 添加文本信息（可选，用于存储元数据）
    text_info = f"Page {page_num} Image {image_num}"
    text_data = b'Placeholder' + b'\x00' + text_info.encode('utf-8')
    text_chunk = b'tEXt' + text_data
    text_crc = struct.pack('>I', crc32(text_chunk))
    text_block = struct.pack('>I', len(text_data)) + text_chunk + text_crc
    
    # IEND块
    iend_chunk = b'IEND'
    iend_crc = struct.pack('>I', crc32(iend_chunk))
    iend = struct.pack('>I', 0) + iend_chunk + iend_crc
    
    return png_signature + ihdr + idat + text_block + iend

# 读取映射文件
mapping_file = Path('produce/ppt内容映射.txt')
with open(mapping_file, 'r', encoding='utf-8') as f:
    data = json.load(f)

# 查找最新PPT文件
produce_dir = Path('produce')
ppt_files = list(produce_dir.glob('new_ppt_*.pptx'))
if not ppt_files:
    print('未找到PPT文件')
    exit(1)

ppt_files.sort(reverse=True)
ppt_filename = ppt_files[0].stem
print(f'找到PPT文件: {ppt_filename}')

# 创建图片目录
image_dir = Path('produce/images') / ppt_filename
image_dir.mkdir(parents=True, exist_ok=True)

# 生成可见的占位图片并更新映射
new_image_mapping = {}
for page_index, page in enumerate(data, start=1):
    if '图片提示词' not in page or not page['图片提示词']:
        continue
    
    if page_index not in new_image_mapping:
        new_image_mapping[page_index] = {}
    
    image_index = 1
    for annotation, prompt in page['图片提示词'].items():
        image_filename = f'{page_index}_{image_index}.png'
        image_path = image_dir / image_filename
        relative_path = f'images/{ppt_filename}/{image_filename}'
        
        # 创建可见的占位图片 (1920x1080, 16:9比例)
        png_data = create_visible_placeholder_png(1920, 1080, page_index, image_index, annotation)
        
        with open(image_path, 'wb') as f:
            f.write(png_data)
        
        new_image_mapping[page_index][annotation] = relative_path
        print(f'生成图片: 第{page_index}页 第{image_index}张 - {annotation[:30]}... -> {relative_path}')
        image_index += 1

# 更新映射文件
for page_index, page in enumerate(data, start=1):
    if '图片提示词' not in page or not page['图片提示词']:
        continue
    
    if '新图片映射' not in page:
        page['新图片映射'] = {}
    
    if page_index in new_image_mapping:
        page['新图片映射'].update(new_image_mapping[page_index])

# 保存
with open(mapping_file, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

total = sum(len(m) for m in new_image_mapping.values())
print(f'\n✓ 完成！共生成 {total} 张可见占位图片')
print(f'✓ 映射文件已更新')
print(f'\n注意: 当前生成的是占位图片（带网格图案和边框）。')
print(f'要使用真实的AI图片生成，请修改脚本调用图片生成API。')

