#!/usr/bin/env python3
"""schema.json からヒアリングシート（Excelテンプレート）を生成する。

使い方:
    python scripts/build_template.py [出力パス]
既定の出力先: assets/skill-seed_hearing_sheet.xlsx

テンプレートの質問を変えたいときは references/schema.json を編集してから
このスクリプトを再実行すること（読み込み側 read_sheet.py も同じ schema.json を使う）。
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
SCHEMA = json.loads((ROOT / "references" / "schema.json").read_text(encoding="utf-8"))

FONT = "Meiryo UI"
F_BASE = Font(name=FONT, size=10)
F_TITLE = Font(name=FONT, size=14, bold=True, color="FFFFFF")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_SECTION = Font(name=FONT, size=11, bold=True)
F_REQ = Font(name=FONT, size=10, bold=True, color="C00000")
F_REC = Font(name=FONT, size=10, bold=True, color="C65911")
F_OPT = Font(name=FONT, size=10, color="7F7F7F")
F_EXAMPLE = Font(name=FONT, size=10, italic=True, color="7F7F7F")
F_HINT = Font(name=FONT, size=9, color="595959")

FILL_TITLE = PatternFill("solid", fgColor="1F4E78")
FILL_HEAD = PatternFill("solid", fgColor="2F75B5")
FILL_SECTION = PatternFill("solid", fgColor="DDEBF7")
FILL_ANSWER = PatternFill("solid", fgColor="FFF2CC")
FILL_EXAMPLE = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

UNLOCK = Protection(locked=False)

# 質問ブロックの列配置（read_sheet.py と一致させること）
Q_COLS = {"id": 1, "q": 2, "req": 3, "answer": 4, "hint": 5, "example": 6}
Q_WIDTHS = {1: 8, 2: 42, 3: 9, 4: 50, 5: 40, 6: 36}


def describe_condition(cond):
    """条件付き必須の条件を、記入者向けの文に直す（read_sheet.py の is_required と対応）。"""
    if cond.get("empty"):
        return f"{cond['id']}が空欄の場合"
    return f"{cond['id']}が「{cond['equals']}」の場合"


def req_cell(cell, req):
    label = {
        "required": "★必須",
        "recommended": "推奨",
        "optional": "任意",
        "conditional": "★条件付",
    }[req]
    cell.value = label
    cell.font = {"required": F_REQ, "conditional": F_REQ, "recommended": F_REC}.get(req, F_OPT)
    cell.alignment = CENTER


def add_choice_validation(ws, choices, ref):
    # 「その他」を直接入力できるよう、エラーではなく警告スタイルにする
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(choices) + '"',
        allow_blank=True,
        errorStyle="warning",
        error="リストにない値です。このまま入力してよければ「はい」を押してください。",
        errorTitle="確認",
        showErrorMessage=True,
    )
    ws.add_data_validation(dv)
    dv.add(ref)


def write_guide(wb):
    ws = wb.active
    ws.title = "0_記入ガイド"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 80

    ws.merge_cells("B2:C2")
    ws["B2"] = "SKILL作成ヒアリングシート（skill-seed）"
    ws["B2"].font = F_TITLE
    ws["B2"].fill = FILL_TITLE
    ws["B2"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 30

    lines = [
        ("このシートの目的", "作りたいSKILL（Claudeに覚えさせる仕事のやり方）について、必要な情報を集めるためのシートです。記入後、このファイルをClaudeに渡すとSKILLを作成します。"),
        ("書く場所", "黄色のセルだけに書いてください。それ以外のセルは編集できないようにしてあります。"),
        ("★必須", "必ず書いてください。空欄があると、Claudeは作業を始めずに不足している項目の一覧を返します。"),
        ("★条件付", "前の質問の答えによって必須になります（例: H-01で「はい」を選んだらH-02が必須）。"),
        ("推奨", "書くとSKILLの精度が上がります。空欄でも作業は進みます。"),
        ("任意", "分かる範囲で構いません。"),
        ("灰色の「例」", "書き方の見本です。消したり書き換えたりする必要はありません（読み込み時は無視されます）。"),
        ("プルダウン", "▼が出るセルはリストから選んでください。「その他」を選ぶ代わりに直接入力もできます（警告が出たら「はい」）。"),
        ("表の行", "表は上の行から詰めて書いてください。1行に1件です。行が足りない場合は、表の最後の行に「（続きは別紙）」と書き、別ファイルを添付してください。"),
        ("セル内の改行", "Windows: Alt + Enter　／　Mac: Option + Return"),
        ("添付ファイル", "サンプルや理想の出力例があれば、シート内にファイル名を書いたうえで、このExcelと一緒にClaudeに渡してください。"),
        ("分からないとき", "「分からない」と書いて構いません。空欄よりもClaudeが正しく判断できます。"),
    ]
    r = 4
    ws.cell(r, 2, "記入のしかた").font = F_SECTION
    r += 1
    for k, v in lines:
        a = ws.cell(r, 2, k)
        b = ws.cell(r, 3, v)
        a.font = Font(name=FONT, size=10, bold=True)
        b.font = F_BASE
        a.alignment = WRAP
        b.alignment = WRAP
        a.border = b.border = BORDER
        a.fill = FILL_SECTION
        r += 1

    r += 1
    ws.cell(r, 2, "シート一覧").font = F_SECTION
    r += 1
    for sh in SCHEMA["sheets"]:
        a = ws.cell(r, 2, sh["name"])
        b = ws.cell(r, 3, sh["title"])
        a.font = Font(name=FONT, size=10, bold=True)
        b.font = F_BASE
        a.border = b.border = BORDER
        r += 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.protection.sheet = True


def write_q_header(ws, r):
    for col, name in [(1, "ID"), (2, "質問"), (3, "必須"), (4, "回答（ここに記入）"), (5, "説明・ヒント"), (6, "記入例")]:
        c = ws.cell(r, col, name)
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = CENTER
        c.border = BORDER


def write_question(ws, r, item):
    ws.cell(r, Q_COLS["id"], item["id"]).font = Font(name=FONT, size=10, bold=True)
    ws.cell(r, Q_COLS["q"], item["q"]).font = F_BASE
    req_cell(ws.cell(r, Q_COLS["req"]), item["req"])
    ans = ws.cell(r, Q_COLS["answer"])
    ans.fill = FILL_ANSWER
    ans.protection = UNLOCK
    ans.font = F_BASE
    hint = item.get("hint", "")
    if item.get("condition"):
        hint = describe_condition(item["condition"]) + "は必須。" + hint
    if item.get("choices"):
        hint = ("▼リストから選択。" + hint).strip()
        add_choice_validation(ws, item["choices"], ans.coordinate)
    ws.cell(r, Q_COLS["hint"], hint).font = F_HINT
    ws.cell(r, Q_COLS["example"], item.get("example", "")).font = F_EXAMPLE
    for col in range(1, 7):
        c = ws.cell(r, col)
        c.border = BORDER
        if col != Q_COLS["req"]:
            c.alignment = WRAP
    lines = max(str(item.get("example", "")).count("\n") + 1, 2)
    ws.row_dimensions[r].height = max(30, 15 * lines)


def write_table(ws, r, item, widths):
    ncol = len(item["columns"])
    # タイトル行: A=ID, B=タイトル, 最終列の次=必須
    ws.cell(r, 1, item["id"]).font = Font(name=FONT, size=10, bold=True)
    label = {"required": "★必須", "recommended": "推奨", "optional": "任意", "conditional": "★条件付"}[item["req"]]
    if item.get("condition"):
        label += "：" + describe_condition(item["condition"])
    ws.cell(r, 2, f"{item['title']}　【{label}】").font = F_SECTION
    for col in range(1, ncol + 2):
        ws.cell(r, col).fill = FILL_SECTION
    r += 1
    hint = item.get("hint", "")
    if item.get("min_rows"):
        hint = f"【{item['min_rows']}件以上】" + hint
    ws.cell(r, 2, hint).font = F_HINT
    ws.cell(r, 2).alignment = Alignment(wrap_text=False, vertical="top")
    r += 1
    # 見出し行
    heads = ["No"] + [c["name"] + ("★" if c["req"] else "") for c in item["columns"]]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(r, i, h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    # 例の行
    ws.cell(r, 1, "例").font = F_EXAMPLE
    for i, col in enumerate(item["columns"], start=2):
        c = ws.cell(r, i, item["example"].get(col["key"], ""))
        c.font = F_EXAMPLE
    for i in range(1, ncol + 2):
        ws.cell(r, i).fill = FILL_EXAMPLE
        ws.cell(r, i).border = BORDER
        ws.cell(r, i).alignment = WRAP
    ws.row_dimensions[r].height = 30
    r += 1
    # 記入行
    first = r
    for n in range(1, item["rows"] + 1):
        ws.cell(r, 1, n).alignment = CENTER
        ws.cell(r, 1).font = F_BASE
        ws.cell(r, 1).border = BORDER
        for i in range(2, ncol + 2):
            c = ws.cell(r, i)
            c.fill = FILL_ANSWER
            c.protection = UNLOCK
            c.border = BORDER
            c.alignment = WRAP
            c.font = F_BASE
        ws.row_dimensions[r].height = 30
        r += 1
    last = r - 1
    for i, col in enumerate(item["columns"], start=2):
        letter = get_column_letter(i)
        widths[i] = max(widths.get(i, 0), col["width"])
        if col.get("choices"):
            add_choice_validation(ws, col["choices"], f"{letter}{first}:{letter}{last}")
    return r


def write_sheet(wb, sheet):
    ws = wb.create_sheet(sheet["name"])
    ws.sheet_view.showGridLines = False
    widths = dict(Q_WIDTHS) if any(i["kind"] == "q" for i in sheet["items"]) else {1: 8}
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws["A1"] = sheet["title"]
    ws["A1"].font = F_TITLE
    ws["A1"].fill = FILL_TITLE
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    ws["A2"] = "黄色のセルに記入してください。灰色の「例」は見本です（書き換え不要）。"
    ws["A2"].font = F_HINT

    r = 4
    prev = None
    for item in sheet["items"]:
        if item["kind"] == "q":
            if prev != "q":
                write_q_header(ws, r)
                r += 1
            write_question(ws, r, item)
            r += 1
        else:
            if prev is not None:
                r += 1
            r = write_table(ws, r, item, widths)
            r += 1
        prev = item["kind"]

    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.protection.sheet = True
    ws.protection.formatCells = False
    ws.protection.formatRows = False
    ws.protection.formatColumns = False


def write_meta(wb):
    ws = wb.create_sheet("_meta")
    ws["A1"] = "template"
    ws["B1"] = "skill-seed"
    ws["A2"] = "template_version"
    ws["B2"] = SCHEMA["template_version"]
    ws.sheet_state = "hidden"
    ws.protection.sheet = True


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "assets" / "skill-seed_hearing_sheet.xlsx"
    wb = Workbook()
    write_guide(wb)
    for sheet in SCHEMA["sheets"]:
        write_sheet(wb, sheet)
    write_meta(wb)
    wb.active = 0
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"saved: {out}")


if __name__ == "__main__":
    main()
