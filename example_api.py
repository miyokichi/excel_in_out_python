"""Python APIとしての使用例。CLIを経由せずに直接モジュールを呼び出す。"""

import json
import extractor
import formula_parser
import ir_serializer
import executor

EXCEL_PATH = "sample.xlsx"

# ── 1. Excelから全数式を抽出してIRを作る ──────────────────────────────────────

extracted = extractor.extract(EXCEL_PATH)

asts = {}
for sheet_name, cells in extracted.items():
    asts[sheet_name] = {}
    for addr, info in cells.items():
        asts[sheet_name][addr] = formula_parser.parse(info["formula"])

ir = ir_serializer.build_ir(EXCEL_PATH, extracted, asts)

# JSONファイルに保存
ir_serializer.save_json(ir, "output.json")

# YAMLファイルに保存
ir_serializer.save_yaml(ir, "output.yaml")

print("=== 抽出されたシート・セル ===")
for sheet, cells in ir["sheets"].items():
    for addr, data in cells.items():
        print(f"  {sheet}!{addr}: {data['formula']}")

# ── 2. 単一の数式文字列をパースしてASTを得る ─────────────────────────────────

ast = formula_parser.parse("=IF(A1>0, SUM(B1:B10), 0)")
print("\n=== ASTのJSON表示 ===")
print(json.dumps(ast, ensure_ascii=False, indent=2))

# ── 3. IRファイルを読み込んで特定セルを実行する ───────────────────────────────

ir_loaded = ir_serializer.load("output.json")

cell_ast = ir_loaded["sheets"]["Sheet1"]["D1"]["ast"]
result = executor.run(cell_ast, EXCEL_PATH, sheet="Sheet1")
print(f"\n=== D1 (=SUM(A1:A4)) の実行結果: {result} ===")

# ── 4. 複数セルをまとめて実行する ─────────────────────────────────────────────

print("\n=== 全数式セルをPythonで実行 ===")
ctx = executor.WorkbookContext(EXCEL_PATH, default_sheet="Sheet1")

for sheet, cells in ir_loaded["sheets"].items():
    for addr, data in cells.items():
        if data["ast"] is None:
            continue
        try:
            value = executor.evaluate(data["ast"], ctx)
            print(f"  {sheet}!{addr}: {data['formula']}  =>  {value!r}")
        except Exception as e:
            print(f"  {sheet}!{addr}: エラー — {e}")

# ── 5. ASTを直接組み立てて実行する（Excelファイル不要）───────────────────────

print("\n=== ASTを手動で組み立てて実行 ===")

manual_ast = {
    "type": "Function",
    "name": "IF",
    "args": [
        {
            "type": "BinaryOp",
            "op": ">",
            "left":  {"type": "Literal", "kind": "number", "value": 10},
            "right": {"type": "Literal", "kind": "number", "value": 5},
        },
        {"type": "Literal", "kind": "string", "value": "大きい"},
        {"type": "Literal", "kind": "string", "value": "小さい"},
    ],
}

# Excelファイルなしで実行（リテラルのみのASTなのでコンテキスト不要）
result2 = executor.evaluate(manual_ast, ctx)
print(f"  IF(10>5, '大きい', '小さい')  =>  {result2!r}")
