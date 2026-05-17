"""CLI entry point for excel_utility.

Commands
--------
  extract  <excel>  -o <out>           Extract all formulas to JSON/YAML IR
  execute  <excel>  <ir>  <cell_ref>   Evaluate one cell's formula in Python
  info     <excel>                     Print all formula cells (no IR needed)
"""

import argparse
import json
import sys

import extractor
import formula_parser
import ir_serializer
import executor as exec_module


# ── helpers ───────────────────────────────────────────────────────────────────

def _parse_cell_ref(ref: str) -> tuple[str | None, str]:
    """Split 'Sheet1!B2' into ('Sheet1', 'B2'). If no sheet, return (None, ref)."""
    if "!" in ref:
        sheet, addr = ref.split("!", 1)
        return sheet, addr
    return None, ref


# ── sub-commands ──────────────────────────────────────────────────────────────

def cmd_extract(args: argparse.Namespace) -> None:
    print(f"Reading {args.excel} …")
    extracted = extractor.extract(args.excel)

    total = sum(len(cells) for cells in extracted.values())
    print(f"  Found {total} formula cell(s) across {len(extracted)} sheet(s)")

    print("Parsing formulas …")
    asts: dict[str, dict[str, dict]] = {}
    errors: list[str] = []
    for sheet_name, cells in extracted.items():
        asts[sheet_name] = {}
        for addr, info in cells.items():
            try:
                asts[sheet_name][addr] = formula_parser.parse(info["formula"])
            except Exception as e:
                errors.append(f"  {sheet_name}!{addr}: {e}")
                asts[sheet_name][addr] = None

    if errors:
        print("Parse errors:")
        for err in errors:
            print(err)

    ir = ir_serializer.build_ir(args.excel, extracted, asts)

    out_path = args.output or (args.excel.rsplit(".", 1)[0] + ".json")
    ir_serializer.save(ir, out_path)
    print(f"IR saved to {out_path}")


def cmd_execute(args: argparse.Namespace) -> None:
    sheet, addr = _parse_cell_ref(args.cell)

    ir = ir_serializer.load(args.ir)
    sheets = ir.get("sheets", {})

    target_sheet = sheet or next(iter(sheets), None)
    if target_sheet not in sheets:
        print(f"Sheet {target_sheet!r} not found in IR. Available: {list(sheets)}", file=sys.stderr)
        sys.exit(1)

    cell_data = sheets[target_sheet].get(addr)
    if cell_data is None:
        print(f"Cell {addr} not found in sheet {target_sheet!r}", file=sys.stderr)
        sys.exit(1)

    ast = cell_data.get("ast")
    if ast is None:
        print(f"No AST for {target_sheet}!{addr} (parse error during extract?)", file=sys.stderr)
        sys.exit(1)

    print(f"Formula : {cell_data['formula']}")
    print(f"Cached  : {cell_data['cached_value']}")

    try:
        result = exec_module.run(ast, args.excel, sheet=target_sheet)
        print(f"Python  : {result}")
    except Exception as e:
        print(f"Execution error: {e}", file=sys.stderr)
        sys.exit(1)


def cmd_info(args: argparse.Namespace) -> None:
    extracted = extractor.extract(args.excel)
    if not extracted:
        print("No formula cells found.")
        return

    for sheet_name, cells in extracted.items():
        print(f"\n[{sheet_name}]")
        for addr, info in sorted(cells.items()):
            cached = info["cached_value"]
            print(f"  {addr:>6}  {info['formula']:<50}  (cached={cached!r})")


def cmd_parse(args: argparse.Namespace) -> None:
    """Parse a single formula string and print the AST."""
    formula = args.formula
    if not formula.startswith("="):
        formula = "=" + formula
    try:
        ast = formula_parser.parse(formula)
        print(json.dumps(ast, ensure_ascii=False, indent=2))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


# ── argument parser ───────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel_utility",
        description="Convert Excel formulas to a portable intermediate representation (IR) and execute them in Python.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    # extract
    p_ext = sub.add_parser("extract", help="Extract formulas from an Excel file and save IR")
    p_ext.add_argument("excel", help="Path to .xlsx file")
    p_ext.add_argument("-o", "--output", help="Output path (.json or .yaml/.yml)")
    p_ext.set_defaults(func=cmd_extract)

    # execute
    p_exe = sub.add_parser("execute", help="Evaluate a formula cell using Python")
    p_exe.add_argument("excel", help="Path to .xlsx file (used for cell values)")
    p_exe.add_argument("ir", help="Path to IR file (.json or .yaml/.yml)")
    p_exe.add_argument("cell", help="Cell reference, e.g. Sheet1!B2 or B2")
    p_exe.set_defaults(func=cmd_execute)

    # info
    p_info = sub.add_parser("info", help="List all formula cells in an Excel file")
    p_info.add_argument("excel", help="Path to .xlsx file")
    p_info.set_defaults(func=cmd_info)

    # parse (single formula, no Excel file needed)
    p_parse = sub.add_parser("parse", help="Parse a single formula string and show its AST")
    p_parse.add_argument("formula", help='Formula string, e.g. "=IF(A1>0,SUM(B1:B10),0)"')
    p_parse.set_defaults(func=cmd_parse)

    return p


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
