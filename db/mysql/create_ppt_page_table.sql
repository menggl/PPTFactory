-- 建表：ppt_page，用于按页存储 `produce/ppt内容页.txt` 中的每一页内容
-- 主键： (ppt_name, page_number)
-- 使用：mysql -u <user> -p <database> < create_ppt_page_table.sql

CREATE DATABASE IF NOT EXISTS pptfactory DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE pptfactory;

CREATE TABLE IF NOT EXISTS ppt_page (
  ppt_name VARCHAR(64) NOT NULL COMMENT '生成的PPT文件名，例如 new_ppt_20251217123045',
  page_number INT NOT NULL COMMENT '页码，从1开始',
  template_id VARCHAR(32) NULL COMMENT '模板编号，例如 T002',
  template_page_index INT NULL COMMENT '该页在模板 PPT 中的页序，从1开始',
  page_json JSON NULL COMMENT '页面的结构化 JSON（如果可用）',
  page_text LONGTEXT NULL COMMENT '页面的纯文本内容或摘要',
  image_prompts JSON NULL COMMENT '由该页生成的图片提示词（可选）',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (ppt_name, page_number)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- 示例插入（请根据实际数据替换）
INSERT INTO ppt_page (ppt_name, page_number, template_id, template_page_index, page_json, page_text)
VALUES (
  'new_ppt_20251217123045',
  1,
  'T003',
  1,
  JSON_OBJECT('title','封面','body','欢迎使用'),
  '封面: 欢迎使用'
);

-- 查询示例：
-- SELECT ppt_name, page_number, JSON_PRETTY(page_json) FROM ppt_page WHERE ppt_name = 'new_ppt_20251217123045' ORDER BY page_number;
