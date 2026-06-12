# week2
## 一、项目概述

本周任务主要完成 8 家拟上市公司的招股说明书抽取、整理和校验工作，形成 JSONL、三表 Excel 以及关键页批注 PDF。整个流程包括：

1. PDF 下载与解析
2. 候选文本定位
3. JSONL 文件生成（`subscription_flow` 和 `equity_snapshot`）
4. 三表抽取 Excel（表1：认缴流量表，表2：股权结构存量表，表3：schema + cross-check 校验表）
5. 关键页批注 PDF 生成
6. 校验日志生成

---

## 二、目录结构

```
week2/
  README.md                     # 本文件
  company_list/week2_public_8.csv  # 公司基础信息
  outputs/week2_jsonl/           # JSONL 抽取结果
  outputs/week2_excel/           # 三表抽取 Excel
  scripts/                       # 脚本目录
    validate_jsonl.py
    jsonl_to_excel_demo_style_v2.py
    fix_demo_excel_review_status.py
    build_annotations_pdf.py
  logs/
    schema_validation_log.csv
    cross_check_summary.csv
  annotations_pdf/               # 关键页批注 PDF + 索引 CSV
  prompts/
    llm_usage.md                 # LLM 使用说明
    extraction_prompt_template.md # LLM 提示词模板
  weekly_reports/week2.md        # 本周周报
```

---

## 三、文件说明

- **JSONL 文件**：每条记录对应一个认缴或股权结构信息，包含 PDF 页码、股东/认购方、快照时间、原文证据等。  
- **三表 Excel**：
  - **表1**：增资扩股认缴流量表  
  - **表2**：股权结构存量表  
  - **表3**：schema 与 cross-check 校验表  
- **关键页批注 PDF**：根据 JSONL `pdf_page` 提取关键页，并加批注提示该页包含哪些记录。  
- **logs**：JSONL 校验日志及 cross-check 汇总结果。  
- **prompts**：LLM 使用说明及提示词模板，说明如何辅助抽取和复核。  
- **weekly_reports**：周报，记录本周各公司工作概况、问题处理、cross-check 结果。

---

## 四、LLM 使用说明

- **是否使用 LLM**：未使用 LLM API 批量抽取，主要用于辅助生成脚本、设计字段和复核特殊口径。  
- **辅助用途**：
  - 辅助生成三表抽取脚本
  - 理解复杂股权结构口径
  - 校验 cross-check 异常
- **最终结果依据**：JSONL、Excel、关键页批注 PDF、校验日志。

