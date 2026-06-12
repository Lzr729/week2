# JSONL 抽取提示词模板

## 任务说明
给定招股说明书候选文本，请抽取两类事实记录：
1. subscription_flow：认缴流量
2. equity_snapshot：股权结构存量

## 抽取要求
- 不要编造 PDF 未披露的信息
- 数字必须来源于原文
- PDF 未直接披露的字段可留空
- evidence_text 必须保留原文证据，不可改写为摘要
- 每家公司必须尽量识别 t0 股权结构
- 输出 JSONL，每行一个 JSON 对象，结构如下：

```json
{
  "company_code": "公司代码",
  "company_name": "公司名称",
  "record_type": "subscription_flow 或 equity_snapshot",
  "subscriber": "认缴方名称（仅 subscription_flow）",
  "shareholder_name": "股东名称（仅 equity_snapshot）",
  "increase_date": "认缴日期（仅 subscription_flow）",
  "snapshot_time": "快照时间（仅 equity_snapshot）",
  "subscription_amount": "认缴数量",
  "subscription_value": "认缴金额",
  "evidence_text": "原文证据段落",
  "pdf_page": "PDF 页码"
}
```

## 使用说明
- 将候选文本输入 LLM
- 使用上述 JSON 模板生成抽取结果
- 所有输出仅作为参考，最终结果以规则复核和校验通过的 JSONL 为准
