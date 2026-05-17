"""Serialize and deserialize the IR (intermediate representation) to/from JSON and YAML."""

import json
import pathlib
from typing import Any

import yaml


def build_ir(source_path: str, extracted: dict[str, dict[str, dict[str, Any]]], asts: dict[str, dict[str, dict]]) -> dict:
    """Build the top-level IR dict from extraction results and parsed ASTs."""
    sheets: dict[str, dict] = {}
    for sheet_name, cells in extracted.items():
        sheet_cells: dict[str, dict] = {}
        for addr, info in cells.items():
            sheet_cells[addr] = {
                "formula": info["formula"],
                "cached_value": info["cached_value"],
                "ast": asts.get(sheet_name, {}).get(addr),
            }
        sheets[sheet_name] = sheet_cells

    return {
        "version": "1.0",
        "source": str(pathlib.Path(source_path).name),
        "sheets": sheets,
    }


def save_json(ir: dict, path: str, indent: int = 2) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(ir, f, ensure_ascii=False, indent=indent, default=str)


def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_yaml(ir: dict, path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(ir, f, allow_unicode=True, sort_keys=False, default_flow_style=False)


def load_yaml(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def save(ir: dict, path: str) -> None:
    """Save IR to JSON or YAML based on file extension."""
    ext = pathlib.Path(path).suffix.lower()
    if ext in (".yml", ".yaml"):
        save_yaml(ir, path)
    else:
        save_json(ir, path)


def load(path: str) -> dict:
    """Load IR from JSON or YAML based on file extension."""
    ext = pathlib.Path(path).suffix.lower()
    if ext in (".yml", ".yaml"):
        return load_yaml(path)
    return load_json(path)
