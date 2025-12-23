-- 建表：ppt_contents，用于存储 `produce/ppt内容页.txt` 的内容
-- 主键：ppt_name（例如 new_ppt_20251217123045）
-- 使用：mysql -u <user> -p <database> < create_ppt_contents_table.sql

CREATE DATABASE IF NOT EXISTS pptfactory DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE pptfactory;

CREATE TABLE IF NOT EXISTS ppt_contents (
  ppt_name VARCHAR(64) NOT NULL COMMENT '生成的PPT文件名，示例 new_ppt_20251217123045',
  contents JSON NULL COMMENT 'ppt内容页的结构化 JSON（通常是页面数组）',
  raw_text LONGTEXT NULL COMMENT 'ppt内容页.txt 的原始文本备份',
  page_count INT NULL COMMENT '页面数量（如果可计算）',
  source_file VARCHAR(255) NULL COMMENT '来源文件路径，例如 produce/ppt内容页.txt',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (ppt_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- 示例插入（请根据实际数据替换）
INSERT INTO ppt_contents (ppt_name, contents, raw_text, page_count, source_file)
VALUES (
  'new_ppt_20251217123045',
  JSON_ARRAY(
    JSON_OBJECT('page',1,'title','封面','body','...'),
    JSON_OBJECT('page',2,'title','目录','body','...')
  ),
  '第一页文本\n第二页文本',
  2,
  'produce/ppt内容页.txt'
);

-- 查询示例：
-- SELECT ppt_name, JSON_PRETTY(contents), page_count, created_at FROM ppt_contents ORDER BY created_at DESC;
