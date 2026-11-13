#!/usr/bin/env python3
"""記入済みヒアリングシートを読み込み、回答と不足項目をJSONで出力する。

使い方:
    python scripts/read_sheet.py <記入済み.xlsx> [--files 添付1 添付2 ...] [--out result.json]

--files には、ユーザーがExcelと一緒に渡したファイルのパス（またはファイル名）を並べる。
シート内に書かれた添付ファイル名と照合し、渡されていないものを不足として報告する。

終了コード:
    0 = 必須項目がすべて揃っている（status: "ok"）
    2 = 不足あり（status: "incomplete"）
    1 = 読み込めない（status: "error"：テンプレート以外のファイル、シート欠落など）
"""
import argparse
import json
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SCHEMA = json.loads((ROOT / "references" / "schema.json").read_text(encoding="utf-8"))

ANSWER_COL = 4  # build_template.py の Q_COLS["answer"] と一致
FILE_LIKE = re.compile(r"[^\s/\\]+\.[A-Za-z0-9]{1,5}$")


def norm(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y/%m/%d")
    return str(v).strip()


def find_row(ws, value, start=1):
    for r in range(start, ws.max_row + 1):
        if norm(ws.cell(r, 1).value) == value:
            return r
    return None


def read_question(ws, item):
    r = find_row(ws, item["id"])
    if r is None:
        return None, f"{item['id']} の行が見つかりません（行が削除された可能性があります）"
    return norm(ws.cell(r, ANSWER_COL).value), None


def read_table(ws, item):
    r = find_row(ws, item["id"])
    if r is None:
        return None, f"{item['id']} の表が見つかりません"
    head = find_row(ws, "No", r + 1)
    if head is None:
        return None, f"{item['id']} の表の見出し行が見つかりません"
    rows = []
    rr = head + 1
    if norm(ws.cell(rr, 1).value) == "例":
        rr += 1
    while rr <= ws.max_row:
        no = ws.cell(rr, 1).value
        if not isinstance(no, (int, float)) and not norm(no).isdigit():
            break
        rec = {"no": int(no)}
        for i, col in enumerate(item["columns"], start=2):
            rec[col["key"]] = norm(ws.cell(rr, i).value)
        if any(rec[c["key"]] for c in item["columns"]):
            rows.append(rec)
        rr += 1
    return rows, None


def is_required(item, answers):
    if item["req"] == "required":
        return True
    if item["req"] == "conditional":
        # 条件の書き方は build_template.py の describe_condition と対応させること
        cond = item["condition"]
        value = answers.get(cond["id"], "")
        if cond.get("empty"):
            return not value
        return value == cond["equals"]
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--files", nargs="*", default=[])
    ap.add_argument("--out")
    args = ap.parse_args()

    result = {
        "status": "ok",
        "template_version": None,
        "answers": {},
        "tables": {},
        "missing": [],
        "warnings": [],
        "referenced_files": [],
        "missing_files": [],
        "errors": [],
    }

    try:
        wb = load_workbook(args.xlsx, data_only=True)
    except Exception as e:  # noqa: BLE001
        result["status"] = "error"
        result["errors"].append(f"Excelを開けません: {e}")
        return finish(result, args)

    if "_meta" not in wb.sheetnames or norm(wb["_meta"]["B1"].value) != "skill-seed":
        result["status"] = "error"
        result["errors"].append("skill-seed のヒアリングシートではありません（_meta シートがありません）")
        return finish(result, args)
    ver = norm(wb["_meta"]["B2"].value)
    result["template_version"] = ver
    if ver != SCHEMA["template_version"]:
        result["warnings"].append(
            f"テンプレートの版が異なります（シート: {ver} / 現行: {SCHEMA['template_version']}）。最新版での再記入を検討してください。"
        )

    # 1周目: 値の読み取り（条件付き必須の判定に回答が必要なため）
    items = []
    for sheet in SCHEMA["sheets"]:
        if sheet["name"] not in wb.sheetnames:
            result["errors"].append(f"シート「{sheet['name']}」がありません")
            continue
        ws = wb[sheet["name"]]
        for item in sheet["items"]:
            if item["kind"] == "q":
                val, err = read_question(ws, item)
                if err:
                    result["errors"].append(err)
                    continue
                result["answers"][item["id"]] = val
            else:
                rows, err = read_table(ws, item)
                if err:
                    result["errors"].append(err)
                    continue
                result["tables"][item["id"]] = rows
            items.append((sheet, item))
    if result["errors"]:
        result["status"] = "error"
        return finish(result, args)

    # 2周目: 必須チェック
    answers = result["answers"]
    for sheet, item in items:
        label = item.get("q") or item.get("title")
        required = is_required(item, answers)
        if item["kind"] == "q":
            if answers[item["id"]]:
                if item.get("file_ref"):
                    result["referenced_files"].append({"where": item["id"], "value": answers[item["id"]]})
                continue
            entry = {"id": item["id"], "sheet": sheet["name"], "item": label}
            if required:
                entry["problem"] = "未記入"
                result["missing"].append(entry)
            elif item["req"] == "recommended":
                entry["problem"] = "未記入（推奨項目）"
                result["warnings"].append(entry)
            continue

        rows = result["tables"][item["id"]]
        min_rows = item.get("min_rows", 0)
        for rec in rows:
            empty = [c["name"] for c in item["columns"] if c["req"] and not rec[c["key"]]]
            if empty and (required or rows):
                result["missing"].append({
                    "id": item["id"], "sheet": sheet["name"], "item": label,
                    "problem": f"No.{rec['no']} の「{'」「'.join(empty)}」が未記入",
                })
            for c in item["columns"]:
                if c.get("file_ref") and rec[c["key"]]:
                    result["referenced_files"].append({"where": f"{item['id']} No.{rec['no']}", "value": rec[c["key"]]})
        if len(rows) < min_rows:
            entry = {"id": item["id"], "sheet": sheet["name"], "item": label,
                     "problem": f"{min_rows}件以上必要（現在 {len(rows)}件）"}
            if required:
                result["missing"].append(entry)
            elif item["req"] == "recommended":
                entry["problem"] += "（推奨項目）"
                result["warnings"].append(entry)

    # 添付ファイルの照合（ファイル名らしい値のみ。保存場所やURLは対象外）
    given = {Path(f).name.lower() for f in args.files}
    for ref in result["referenced_files"]:
        for name in re.split(r"[、,，\n]+", ref["value"]):
            name = name.strip()
            if not name or "://" in name or not FILE_LIKE.fullmatch(name):
                continue
            if Path(name).name.lower() not in given:
                result["missing_files"].append({"where": ref["where"], "file": name})

    if result["missing"] or result["missing_files"]:
        result["status"] = "incomplete"
    return finish(result, args)


def finish(result, args):
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    print(text)
    return {"ok": 0, "incomplete": 2, "error": 1}[result["status"]]


if __name__ == "__main__":
    sys.exit(main())
