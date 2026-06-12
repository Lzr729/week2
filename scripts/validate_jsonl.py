import os
import json
import csv
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(BASE_DIR)

from schemas.extraction_models import SubscriptionFlow, EquitySnapshot


JSONL_DIR = os.path.join(BASE_DIR, "outputs", "week2_jsonl")
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_PATH = os.path.join(LOG_DIR, "schema_validation_log.csv")

os.makedirs(LOG_DIR, exist_ok=True)


EXPECTED_COMPANIES = {
    "001282": "三联锻造",
    "301563": "云汉芯城",
    "301581": "黄山谷捷",
    "603418": "友升股份",
    "688758": "赛分科技",
    "688775": "影石创新",
    "920100": "三协电机",
    "920116": "星图测控",
}


def is_blank(value):
    if value is None:
        return True
    if str(value).strip() == "":
        return True
    return False


def validate_one_record(record):
    record_type = record.get("record_type")

    if record_type == "subscription_flow":
        obj = SubscriptionFlow(**record)

        problems = []

        if is_blank(obj.company_code):
            problems.append("company_code为空")
        if is_blank(obj.company_name):
            problems.append("company_name为空")
        if is_blank(obj.pdf_page):
            problems.append("PDF页码为空")
        if is_blank(obj.subscriber):
            problems.append("认购方为空")
        if is_blank(obj.evidence_text):
            problems.append("原文证据为空")

        # 认缴流量至少要有一种数值信号
        if (
            is_blank(obj.subscribed_shares_wan)
            and is_blank(obj.subscription_amount_wan)
            and is_blank(obj.subscription_price_yuan_per_share)
        ):
            problems.append("认缴流量缺少数量/金额/价格数值信号")

        return problems

    elif record_type == "equity_snapshot":
        obj = EquitySnapshot(**record)

        problems = []

        if is_blank(obj.company_code):
            problems.append("company_code为空")
        if is_blank(obj.company_name):
            problems.append("company_name为空")
        if is_blank(obj.pdf_page):
            problems.append("PDF页码为空")
        if is_blank(obj.snapshot_time):
            problems.append("时点为空")
        if is_blank(obj.shareholder_name):
            problems.append("股东名称为空")
        if is_blank(obj.evidence_text):
            problems.append("原文证据为空")

        # 股权结构至少要有持股数、出资额、持股比例中的一种
        if (
            is_blank(obj.shares_wan)
            and is_blank(obj.capital_contribution_wan)
            and is_blank(obj.shareholding_ratio)
        ):
            problems.append("股权结构缺少持股数/出资额/持股比例")

        return problems

    else:
        return [f"record_type非法：{record_type}"]


def main():
    rows = []

    company_has_t0 = {code: False for code in EXPECTED_COMPANIES}
    company_record_count = {code: 0 for code in EXPECTED_COMPANIES}

    jsonl_files = [
        f for f in os.listdir(JSONL_DIR)
        if f.endswith(".jsonl")
    ]

    for jsonl_file in jsonl_files:
        jsonl_path = os.path.join(JSONL_DIR, jsonl_file)

        with open(jsonl_path, "r", encoding="utf-8") as f:
            for line_no, line in enumerate(f, start=1):
                line = line.strip()
                if not line:
                    continue

                try:
                    record = json.loads(line)
                except Exception as e:
                    rows.append({
                        "jsonl_file": jsonl_file,
                        "line_no": line_no,
                        "company_code": "",
                        "company_name": "",
                        "record_type": "",
                        "check_item": "JSON解析",
                        "status": "fail",
                        "message": str(e),
                    })
                    continue

                company_code = str(record.get("company_code", "")).strip()
                company_name = str(record.get("company_name", "")).strip()
                record_type = record.get("record_type")

                if company_code in company_record_count:
                    company_record_count[company_code] += 1

                if (
                    record_type == "equity_snapshot"
                    and "t0" in str(record.get("snapshot_time", "")).lower()
                    and company_code in company_has_t0
                ):
                    company_has_t0[company_code] = True

                try:
                    problems = validate_one_record(record)
                except Exception as e:
                    problems = [f"Pydantic解析失败：{e}"]

                if problems:
                    for p in problems:
                        rows.append({
                            "jsonl_file": jsonl_file,
                            "line_no": line_no,
                            "company_code": company_code,
                            "company_name": company_name,
                            "record_type": record_type,
                            "check_item": "字段校验",
                            "status": "fail",
                            "message": p,
                        })
                else:
                    rows.append({
                        "jsonl_file": jsonl_file,
                        "line_no": line_no,
                        "company_code": company_code,
                        "company_name": company_name,
                        "record_type": record_type,
                        "check_item": "字段校验",
                        "status": "pass",
                        "message": "",
                    })

    # 公司完整性检查
    for code, name in EXPECTED_COMPANIES.items():
        if company_record_count[code] == 0:
            rows.append({
                "jsonl_file": "",
                "line_no": "",
                "company_code": code,
                "company_name": name,
                "record_type": "",
                "check_item": "8家公司完整性",
                "status": "fail",
                "message": "该公司没有记录",
            })
        else:
            rows.append({
                "jsonl_file": "",
                "line_no": "",
                "company_code": code,
                "company_name": name,
                "record_type": "",
                "check_item": "8家公司完整性",
                "status": "pass",
                "message": f"记录数：{company_record_count[code]}",
            })

        if company_has_t0[code]:
            rows.append({
                "jsonl_file": "",
                "line_no": "",
                "company_code": code,
                "company_name": name,
                "record_type": "equity_snapshot",
                "check_item": "是否存在t0股权结构",
                "status": "pass",
                "message": "存在t0",
            })
        else:
            rows.append({
                "jsonl_file": "",
                "line_no": "",
                "company_code": code,
                "company_name": name,
                "record_type": "equity_snapshot",
                "check_item": "是否存在t0股权结构",
                "status": "fail",
                "message": "未发现snapshot_time包含t0的股权结构",
            })

    with open(LOG_PATH, "w", encoding="utf-8-sig", newline="") as f:
        fieldnames = [
            "jsonl_file",
            "line_no",
            "company_code",
            "company_name",
            "record_type",
            "check_item",
            "status",
            "message",
        ]

        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"schema validation log 已生成：{LOG_PATH}")

    fail_count = sum(1 for r in rows if r["status"] == "fail")
    pass_count = sum(1 for r in rows if r["status"] == "pass")

    print(f"通过：{pass_count}")
    print(f"失败：{fail_count}")


if __name__ == "__main__":
    main()