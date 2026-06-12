import os
import json
import csv
from collections import defaultdict
from pathlib import Path

import fitz  # PyMuPDF


BASE_DIR = Path(__file__).resolve().parents[1]

PDF_DIR = BASE_DIR / "data" / "pdfs"
JSONL_DIR = BASE_DIR / "outputs" / "week2_jsonl"
OUT_DIR = BASE_DIR / "annotations_pdf"

OUT_DIR.mkdir(parents=True, exist_ok=True)

INDEX_CSV = OUT_DIR / "关键页索引.csv"


def read_jsonl(path):
    records = []

    with open(path, "r", encoding="utf-8") as f:
        for line_no, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue

            try:
                r = json.loads(line)
                r["_line_no"] = line_no
                records.append(r)
            except Exception as e:
                print(f"JSONL解析失败：{path.name} 第{line_no}行：{e}")

    return records


def find_pdf_by_company_code(company_code):
    """
    根据公司代码在 data/pdfs/ 里找对应PDF。
    要求PDF文件名里包含公司代码，比如：
    920100_三协电机_IPO招股说明书.pdf
    """
    candidates = []

    for p in PDF_DIR.glob("*.pdf"):
        if company_code in p.name:
            candidates.append(p)

    if not candidates:
        return None

    # 如果有多个，优先选文件名较短的
    candidates = sorted(candidates, key=lambda x: len(x.name))
    return candidates[0]


def normalize_pdf_page(value):
    """
    把 JSONL 里的 pdf_page 转成整数页码。
    注意：PDF页码是从1开始，fitz内部页码从0开始。
    """
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    # 处理 35.0 这种
    try:
        return int(float(s))
    except Exception:
        return None


def collect_pages(records):
    """
    汇总每家公司每个PDF页码上有哪些记录。
    """
    page_map = defaultdict(list)

    for r in records:
        page = normalize_pdf_page(r.get("pdf_page"))
        if page is None:
            continue

        page_map[page].append(r)

    return page_map


def short_text(text, max_len=80):
    text = str(text or "").replace("\n", " ").replace("\r", " ").strip()
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def make_note_for_page(page_no, records):
    """
    生成批注说明。
    尽量用英文/数字，避免某些PDF字体环境下中文乱码。
    """
    record_types = sorted(set(str(r.get("record_type", "")) for r in records))
    line_nos = [str(r.get("_line_no")) for r in records if r.get("_line_no")]

    sub_count = sum(1 for r in records if r.get("record_type") == "subscription_flow")
    snap_count = sum(1 for r in records if r.get("record_type") == "equity_snapshot")

    note = (
        f"Extracted evidence page: PDF page {page_no}\n"
        f"record_types: {', '.join(record_types)}\n"
        f"subscription_flow records: {sub_count}\n"
        f"equity_snapshot records: {snap_count}\n"
        f"jsonl line_no: {', '.join(line_nos[:20])}"
    )

    return note


def build_annotation_pdf(company_code, company_name, records, source_pdf):
    page_map = collect_pages(records)

    if not page_map:
        print(f"{company_code}_{company_name} 没有找到有效pdf_page，跳过")
        return []

    src = fitz.open(source_pdf)
    out = fitz.open()

    index_rows = []

    for page_no in sorted(page_map.keys()):
        page_index = page_no - 1

        if page_index < 0 or page_index >= len(src):
            print(f"{company_code}_{company_name} PDF页码超出范围：{page_no}")
            continue

        # 复制原PDF对应页
        out.insert_pdf(src, from_page=page_index, to_page=page_index)

        new_page = out[-1]
        records_on_page = page_map[page_no]

        # 添加一个便签批注，不遮挡正文
        note_text = make_note_for_page(page_no, records_on_page)
        annot = new_page.add_text_annot((36, 36), note_text)
        annot.set_info(title="extraction evidence")
        annot.update()

        # 加一个红框提示“这是关键页”，不高亮具体文字，避免坐标误伤
        rect = new_page.rect
        border_rect = fitz.Rect(20, 20, rect.width - 20, rect.height - 20)
        border = new_page.add_rect_annot(border_rect)
        border.set_colors(stroke=(1, 0, 0))
        border.set_border(width=1)
        border.update()

        # 写索引
        for r in records_on_page:
            index_rows.append({
                "company_code": company_code,
                "company_name": company_name,
                "source_pdf": source_pdf.name,
                "pdf_page": page_no,
                "record_type": r.get("record_type", ""),
                "jsonl_line_no": r.get("_line_no", ""),
                "object_name": r.get("subscriber") or r.get("shareholder_name") or "",
                "snapshot_or_date": r.get("snapshot_time") or r.get("increase_date") or "",
                "evidence_text": short_text(r.get("evidence_text"), 120),
            })

    src.close()

    out_name = f"{company_code}_{company_name}_关键页批注.pdf"
    out_path = OUT_DIR / out_name
    out.save(out_path)
    out.close()

    print(f"已生成：{out_path}")

    return index_rows


def main():
    all_index_rows = []

    jsonl_files = sorted(JSONL_DIR.glob("*.jsonl"))

    if not jsonl_files:
        print(f"没有找到JSONL文件，请检查：{JSONL_DIR}")
        return

    for jsonl_path in jsonl_files:
        records = read_jsonl(jsonl_path)

        if not records:
            continue

        company_code = str(records[0].get("company_code", "")).strip()
        company_name = str(records[0].get("company_name", "")).strip()

        source_pdf = find_pdf_by_company_code(company_code)

        if source_pdf is None:
            print(f"未找到 {company_code}_{company_name} 的PDF，请检查 data/pdfs/ 文件名是否包含公司代码")
            continue

        rows = build_annotation_pdf(company_code, company_name, records, source_pdf)
        all_index_rows.extend(rows)

    # 输出总索引
    fieldnames = [
        "company_code",
        "company_name",
        "source_pdf",
        "pdf_page",
        "record_type",
        "jsonl_line_no",
        "object_name",
        "snapshot_or_date",
        "evidence_text",
    ]

    with open(INDEX_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_index_rows)

    print("\n全部关键页批注PDF生成完成。")
    print(f"关键页索引：{INDEX_CSV}")


if __name__ == "__main__":
    main()