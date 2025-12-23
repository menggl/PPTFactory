-- 建表：ppt_text_map，用于按页存储 `文本映射` 的每一条键值对
-- 关联主表：ppt_mapping (ppt_name, page_number)

CREATE DATABASE IF NOT EXISTS pptfactory DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE pptfactory;

CREATE TABLE IF NOT EXISTS ppt_text_map (
  ppt_name VARCHAR(64) NOT NULL,
  page_number INT NOT NULL,
  map_key VARCHAR(512) NOT NULL COMMENT '文本映射中的 key，例如 一我是主标题、二我是副标题 等，或多键合并字符串',
  map_value LONGTEXT NULL COMMENT '对应的文本值',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (ppt_name, page_number, map_key),
  CONSTRAINT fk_ppt_text_map_page FOREIGN KEY (ppt_name, page_number) REFERENCES ppt_mapping (ppt_name, page_number) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- 示例：
-- INSERT INTO ppt_text_map (ppt_name,page_number,map_key,map_value) VALUES ('new_ppt_20251216180358',1,'一我是主标题','一、煤矿从业人员主要权利');
