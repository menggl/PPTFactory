-- Create table to store outline (大纲目录) for generated PPT files
-- 主键: ppt_name (new_ppt_[年月日时分秒])
-- 使用说明: 在 MySQL 中运行: mysql -u <user> -p <database> < create_outline_table.sql

CREATE DATABASE IF NOT EXISTS pptfactory DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE pptfactory;

CREATE TABLE IF NOT EXISTS ppt_outline (
  ppt_name VARCHAR(64) NOT NULL COMMENT '生成的PPT文件名，示例 new_ppt_20251217123045',
  outline_json JSON NULL COMMENT '大纲目录的结构化表示（JSON）',
  outline_text LONGTEXT NULL COMMENT '大纲目录的原始文本或纯文本版本',
  slide_count INT NULL COMMENT '由大纲生成的幻灯片数量（可选）',
  source_file VARCHAR(255) NULL COMMENT '来源文件路径，例如 produce/大纲目录.txt',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (ppt_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- 示例：插入一条记录（请根据实际数据替换 JSON 与文本）
INSERT INTO ppt_outline (ppt_name, outline_json, outline_text, slide_count, source_file)
VALUES (
  'new_ppt_20251217123045',
  JSON_ARRAY('第一部分: 项目背景','第二部分: 技术方案','第三部分: 实施计划'),
  '第一部分: 项目背景\n第二部分: 技术方案\n第三部分: 实施计划',
  3,
  'produce/大纲目录.txt'
);

-- 可选：查询示例
-- SELECT ppt_name, JSON_PRETTY(outline_json), slide_count, created_at FROM ppt_outline ORDER BY created_at DESC;
