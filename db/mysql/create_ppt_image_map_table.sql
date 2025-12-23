-- 建表：ppt_image_map，用于按页存储图片相关映射：
-- 包括 图片提示词准备、图片提示词（若有）、图片链接映射、图片路径映射
-- 关联主表：ppt_mapping (ppt_name, page_number)

CREATE DATABASE IF NOT EXISTS pptfactory DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE pptfactory;

CREATE TABLE IF NOT EXISTS ppt_image_map (
  ppt_name VARCHAR(64) NOT NULL,
  page_number INT NOT NULL,
  map_key VARCHAR(512) NOT NULL COMMENT '图片映射的 key（通常是用于生成图片的文本组合，例如 三我是文本|四我是长文本|...）',
  prompt_prepare LONGTEXT NULL COMMENT '图片提示词准备（原始串）',
  prompt LONGTEXT NULL COMMENT '（可选）图片提示词或最终用于生成的 prompt',
  image_url VARCHAR(2048) NULL COMMENT '图片链接映射（远端 URL）',
  image_path VARCHAR(1024) NULL COMMENT '下载到本地后的图片路径（相对仓库路径）',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (ppt_name, page_number, map_key),
  CONSTRAINT fk_ppt_image_map_page FOREIGN KEY (ppt_name, page_number) REFERENCES ppt_mapping (ppt_name, page_number) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- 示例：
-- INSERT INTO ppt_image_map (ppt_name,page_number,map_key,prompt_prepare,image_url,image_path) VALUES ('new_ppt_20251216180358',1,'三我是文本|四我是长文本|五我是文本|六我是长文本','确诊尘肺病需及时...','https://s.coze.cn/t/wrfQ8BurzJ8/','produce/images/new_ppt_20251216180358/1_1.png');
