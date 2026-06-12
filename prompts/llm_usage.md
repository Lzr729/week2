# LLM 使用说明

## 1. 是否使用 LLM
本项目未使用 LLM API 批量抽取，主要采用规则定位 + 人工复核。
LLM 仅用于：
- 生成或修改三表抽取脚本
- 理解复杂股权结构口径
- 辅助 cross-check 校验异常处理

## 2. 使用模型
- ChatGPT / GPT-5.5 Thinking

## 3. 使用方式
- 将候选文本、示范 Excel、JSONL 校验结果及 cross-check 结果输入 LLM，用于辅助判断字段口径和生成处理脚本
- 并非直接调用 API 自动抽取，最终结果以 JSONL、Excel、批注 PDF 为准

## 4. 输出依据
- JSONL 文件
- 三表 Excel 文件
- 关键页批注 PDF
- 校验日志
