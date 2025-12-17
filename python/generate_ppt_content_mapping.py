#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成PPT内容映射脚本

根据《produce/ppt内容页.txt》中的内容，生成PPT模板页的文本映射关系。
映射文件保存到《produce/ppt内容映射.txt》。

映射规则：
1. 待替换文本键名不包含【】符号，如"一我是主标题"而不是"【一我是主标题】"
2. 每个文本位置都同时提供"文本"和"长文本"两种映射关系（容错处理）
3. 包含模板页编号（T001, T002等）
"""

import json
import os
import sys

# 中文数字映射
chinese_nums = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']


def parse_json_objects(content):
    """
    解析JSON数组或JSON对象列表
    
    Args:
        content: 文件内容字符串
        
    Returns:
        JSON对象列表
    """
    try:
        # 尝试解析为JSON数组
        data = json.loads(content)
        if isinstance(data, list):
            return data
        elif isinstance(data, dict):
            return [data]
        else:
            return []
    except json.JSONDecodeError:
        # 如果解析失败，尝试逐个解析JSON对象
        json_objects = []
        current_json = ''
        brace_count = 0
        
        for char in content:
            current_json += char
            if char == '{':
                brace_count += 1
            elif char == '}':
                brace_count -= 1
                if brace_count == 0:
                    try:
                        json_objects.append(json.loads(current_json.strip()))
                        current_json = ''
                    except json.JSONDecodeError:
                        current_json = ''
                        pass
        
        return json_objects


def generate_mapping(obj):
    """
    根据单个内容页对象生成映射关系
    
    Args:
        obj: 内容页JSON对象
        
    Returns:
        映射字典
    """
    template_id = obj.get('templateId', '')
    slide_title = obj.get('slide_title', '')
    slide_subtitle = obj.get('slide_subtitle', '')
    key_points = obj.get('key_points', [])
    
    mapping = {
        '模板页编号': template_id,
        '文本映射': {}
    }
    
    # 主标题（同时提供文本和长文本映射）
    mapping['文本映射']['一我是主标题'] = slide_title
    
    # 副标题（如果存在，同时提供文本和长文本映射）
    if slide_subtitle:
        mapping['文本映射']['二我是副标题'] = slide_subtitle
    
    # 根据模板类型处理key_points
    if template_id == 'T001':
        # T001: 3个caption（短文本），同时提供文本和长文本映射
        for i, point in enumerate(key_points[:3], start=3):
            if i <= 10:
                mapping['文本映射'][f'{chinese_nums[i]}我是文本'] = point
                mapping['文本映射'][f'{chinese_nums[i]}我是长文本'] = point
                
    elif template_id in ['T002', 'T003', 'T004', 'T005', 'T009']:
        # 这些模板：2组(heading + body)
        text_idx = 3
        for point in key_points[:2]:
            if isinstance(point, dict):
                heading = point.get('heading', '')
                body = point.get('body', '') or point.get('paragraph', '')
                if heading and text_idx <= 10:
                    # heading同时提供文本和长文本映射
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = heading
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = heading
                    text_idx += 1
                if body and text_idx <= 10:
                    # body同时提供文本和长文本映射
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = body
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = body
                    text_idx += 1
                    
    elif template_id == 'T007':
        # T007: 2-3组(heading + paragraph)
        text_idx = 3
        for point in key_points[:3]:
            if isinstance(point, dict):
                heading = point.get('heading', '')
                paragraph = point.get('paragraph', '') or point.get('body', '')
                if heading and text_idx <= 10:
                    # heading同时提供文本和长文本映射
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = heading
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = heading
                    text_idx += 1
                if paragraph and text_idx <= 10:
                    # paragraph同时提供文本和长文本映射
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = paragraph
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = paragraph
                    text_idx += 1
                    
    elif template_id == 'T006':
        # T006: 只有标题，没有其他文本内容
        # 只映射主标题即可
        pass
        
    elif template_id == 'T008':
        # T008: 3组(heading + body)
        text_idx = 3
        for point in key_points[:3]:
            if isinstance(point, dict):
                heading = point.get('heading', '')
                body = point.get('body', '') or point.get('paragraph', '')
                if heading and text_idx <= 10:
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = heading
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = heading
                    text_idx += 1
                if body and text_idx <= 10:
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = body
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = body
                    text_idx += 1
                    
    elif template_id == 'T010':
        # T010: 特殊格式，包含summary和points
        # 根据实际内容结构处理
        text_idx = 3
        for point in key_points:
            if isinstance(point, dict):
                # 处理heading和body/paragraph
                heading = point.get('heading', '')
                body = point.get('body', '') or point.get('paragraph', '')
                if heading and text_idx <= 10:
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = heading
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = heading
                    text_idx += 1
                if body and text_idx <= 10:
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = body
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = body
                    text_idx += 1
            elif isinstance(point, str):
                # 处理字符串类型的要点
                if text_idx <= 10:
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是文本'] = point
                    mapping['文本映射'][f'{chinese_nums[text_idx]}我是长文本'] = point
                    text_idx += 1
                    
    elif template_id in ['T011', 'T012']:
        # T011/T012: 1个heading + 1个body/paragraph
        if key_points and isinstance(key_points[0], dict):
            point = key_points[0]
            heading = point.get('heading', '')
            body = point.get('paragraph', '') or point.get('body', '')
            if heading:
                # heading同时提供文本和长文本映射
                mapping['文本映射']['三我是文本'] = heading
                mapping['文本映射']['三我是长文本'] = heading
            if body:
                # body同时提供文本和长文本映射
                mapping['文本映射']['四我是文本'] = body
                mapping['文本映射']['四我是长文本'] = body
    
    return mapping


def main():
    """主函数"""
    # 获取脚本所在目录的父目录（项目根目录）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    
    # 输入和输出文件路径
    input_file = os.path.join(project_root, 'produce', 'ppt内容页.txt')
    output_file = os.path.join(project_root, 'produce', 'ppt内容映射.txt')
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"错误：输入文件不存在: {input_file}")
        sys.exit(1)
    
    # 读取内容页文件
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"错误：读取文件失败: {e}")
        sys.exit(1)
    
    # 解析JSON对象
    json_objects = parse_json_objects(content)
    if not json_objects:
        print("错误：未能解析到任何JSON对象")
        sys.exit(1)
    
    # 生成映射
    mappings = []
    for obj in json_objects:
        mapping = generate_mapping(obj)
        mappings.append(mapping)
    
    # 输出JSON到文件
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(mappings, f, ensure_ascii=False, indent=2)
        print(f"成功生成 {len(mappings)} 个页面的映射")
        print(f"输出文件: {output_file}")
        print(f"使用模板编号: {', '.join(sorted(set(m['模板页编号'] for m in mappings)))}")
    except Exception as e:
        print(f"错误：写入文件失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
