---
name: skill-seed
description: Excelのヒアリングシートを使って、作りたいSKILLの要件を集め、skill-creatorで精度よくSKILLを作るためのSKILL。空のヒアリングシート（テンプレート）の配布と、記入済みシートの読み込み→不足チェック→skill-creator用プロンプトの作成→承認後にskill-creatorでSKILL生成、までを担当する。「SKILLを作りたい」「ヒアリングシートをください」「SKILLの要件シート」「記入したシートからSKILLを作って」「skill-seed」と言われたとき、またはskill-seedテンプレートのExcel（_metaシートにskill-seedと書かれたもの）が添付されたときは必ず使うこと。Excelを使わず会話だけでSKILLを作りたい場合や、既存SKILLの改善だけの場合は、skill-creatorを直接使う。
---

# skill-seed

非エンジニアを含む記入者に Excel のヒアリングシートを書いてもらい、その内容から
skill-creator に渡すプロンプトを作り、承認を得てから SKILL を生成する。

skill-creator は本来、会話でユーザーに要件を聞き取る。このSKILLはその聞き取りを
事前にシートで済ませることで、聞き漏れと「Claudeの推測」を減らすのが目的。
そのため全工程を通じて、**シートに書かれていないことを推測で補わない**。

## 同梱ファイル

| パス | 用途 |
|---|---|
| `assets/skill-seed_hearing_sheet.xlsx` | 配布用の空テンプレート |
| `scripts/read_sheet.py` | 記入済みシートを読み、回答・不足・添付ファイルの照合結果を JSON で出す |
| `scripts/build_template.py` | `references/schema.json` からテンプレートを再生成する（質問を変えたときのみ） |
| `references/schema.json` | 質問・表・必須条件の定義。テンプレートと読み込みの両方がこれを使う |
| `references/prompt_template.md` | skill-creator 用プロンプトの構成と組み立てルール |

スクリプトはこの SKILL のディレクトリを基準に実行する（`python <skill-dir>/scripts/...`）。
`openpyxl` が必要。import に失敗したときだけ `pip install openpyxl` する。

## どちらの依頼かを判断する

- **テンプレートが欲しい**（「ヒアリングシートをください」「SKILLを作りたいけど何を書けばいい？」など）
  → 「A. テンプレートを渡す」
- **記入済みシートが渡された**（xlsx が添付・パス指定された）→ 「B. シートからSKILLを作る」
- どちらか分からないときは聞く。

## 入出力の場所

| 環境 | 受け取るファイル | 出力先 |
|---|---|---|
| Claude.ai・アプリ | `/mnt/user-data/uploads/` 配下の添付 | `/mnt/user-data/outputs/` |
| Claude Code | ユーザーが示したパス（なければ聞く） | カレントディレクトリ（またはユーザー指定） |

`/mnt/user-data/` があれば Claude.ai、なければ Claude Code として扱う。

---

## A. テンプレートを渡す

1. `assets/skill-seed_hearing_sheet.xlsx` を出力先にコピーして渡す。
2. 次のことを短く伝える。記入者は非エンジニアの可能性があるので、専門用語を避ける。
   - 黄色のセルだけに書く。最初の「0_記入ガイド」シートに書き方がある
   - ★必須が空欄だと作業を始められない
   - サンプルや理想の出力例があれば、シートにファイル名を書き、Excel と一緒に渡す

## B. シートからSKILLを作る

### B-1. 読み込みと不足チェック

```bash
python <skill-dir>/scripts/read_sheet.py <記入済み.xlsx> --files <一緒に渡されたファイル...> --out <作業dir>/sheet.json
```

`--files` には Excel 以外の添付ファイルをすべて渡す（シート内のファイル名と照合するため）。

結果の `status` で分岐する。

- **`error`**：skill-seed のテンプレートではない、またはシートや行が消されている。
  `errors` を伝え、最新のテンプレートで書き直してもらうよう案内して**止まる**。
- **`incomplete`**：下の形式で不足一覧を返して**止まる**。不足を推測で埋めて先に進んではいけない。
  記入者が直して再提出するのが前提なので、記入者にそのまま転送できる文面にする。

  ```markdown
  ヒアリングシートに不足があるため、SKILLの作成を始められませんでした。
  以下を記入して、もう一度渡してください。

  | シート | ID | 項目 | 不足の内容 |
  |---|---|---|---|
  | （missing の各行） |

  **添付ファイルが見つかりません**（シートに書かれていますが、渡されていません）
  - （missing_files の各行: どこに書かれているか — ファイル名）

  参考（推奨項目・空欄でも進められます）
  - （warnings の各行）
  ```
  ファイルの未添付は、「ファイルを追加で渡す」か「シートから名前を消す」のどちらでも解消できることを添える。

- **`ok`**：B-2 へ進む。`warnings`（推奨項目の空欄、テンプレートの版違い）は B-3 で提示する。

### B-2. skill-creator 用プロンプトを作る

1. 作業ディレクトリ `<出力先>/<name候補>-seed/` を作る（name 候補は A-02、空欄なら A-01 から仮に作る）。
2. 添付ファイルを `<作業dir>/attachments/` にコピーする。
3. `references/prompt_template.md` を読み、その構成とルールに従ってプロンプトを作り、
   `<作業dir>/skill-creator_prompt.md` に保存する。

### B-3. ユーザーに確認する（必ずここで一度止まる）

次の内容をユーザーに見せ、承認を待つ。
- プロンプト全文（またはファイル）
- `warnings` の一覧（推奨項目の空欄など）と、「このまま進めると skill-creator が該当箇所を質問します」という一言
- 「この内容で skill-creator を使って SKILL を作ってよいですか？修正があれば教えてください」

修正を頼まれたら、プロンプトを直して再度確認する。シートの内容と矛盾する修正なら、
どちらを正とするかを聞く。

### B-4. skill-creator で SKILL を作る（承認後のみ）

1. skill-creator スキルを起動し、`skill-creator_prompt.md` の内容を依頼として渡す。
   添付ファイルは `<作業dir>/attachments/` のパスで伝える。
2. 以降は skill-creator の手順に従う（下書き、テスト、改善）。skill-creator から質問が出たら、
   シートに答えがあればそれを使い、なければユーザーに聞く。
3. 完成した SKILL は出力先に置き、`skill-creator_prompt.md` も一緒に残す
   （後で SKILL を直すとき、元の要件として使えるため）。

skill-creator が使えない環境の場合は、`skill-creator_prompt.md` を渡し、
skill-creator が使える環境でこのプロンプトを渡すよう案内する。

## テンプレートを変更するとき

質問を追加・変更するときは `references/schema.json` を編集し、`template_version` を上げてから
`python scripts/build_template.py` でテンプレートを再生成する。読み込みスクリプトも同じ
`schema.json` を使うため、両者は自動的に一致する。版が変わると、古い版で書かれたシートを
読んだときに警告が出る。
