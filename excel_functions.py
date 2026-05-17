"""Python implementations of common Excel functions."""

import math
import re
import statistics
from typing import Any


# ── Helpers ───────────────────────────────────────────────────────────────────

def _flatten(values: Any) -> list:
    """Recursively flatten nested lists."""
    if isinstance(values, list):
        result = []
        for v in values:
            result.extend(_flatten(v))
        return result
    return [values]


def _numeric(values: list) -> list[float]:
    """Filter to numbers only, converting bools (Excel ignores text in ranges for SUM etc.)."""
    result = []
    for v in values:
        if isinstance(v, bool):
            result.append(float(v))
        elif isinstance(v, (int, float)):
            result.append(float(v))
    return result


def _coerce_num(v: Any) -> float:
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            raise ValueError(f"Cannot convert {v!r} to number")
    raise TypeError(f"Cannot convert {type(v).__name__} to number")


# ── Math / Statistics ─────────────────────────────────────────────────────────

def xl_sum(args: list) -> float:
    return sum(_numeric(_flatten(args)))


def xl_average(args: list) -> float:
    nums = _numeric(_flatten(args))
    if not nums:
        raise ZeroDivisionError("AVERAGE of empty range")
    return statistics.mean(nums)


def xl_min(args: list) -> float:
    nums = _numeric(_flatten(args))
    return min(nums)


def xl_max(args: list) -> float:
    nums = _numeric(_flatten(args))
    return max(nums)


def xl_count(args: list) -> int:
    return len(_numeric(_flatten(args)))


def xl_counta(args: list) -> int:
    return sum(1 for v in _flatten(args) if v is not None and v != "")


def xl_countblank(args: list) -> int:
    return sum(1 for v in _flatten(args) if v is None or v == "")


def xl_sumif(args: list) -> float:
    range_vals, criteria, sum_range = args[0], args[1], args[2] if len(args) > 2 else args[0]
    r = _flatten(range_vals)
    s = _flatten(sum_range)
    pred = _make_criteria_pred(criteria)
    return sum(float(s[i]) for i, v in enumerate(r) if pred(v) and isinstance(s[i], (int, float)))


def xl_countif(args: list) -> int:
    range_vals, criteria = args[0], args[1]
    pred = _make_criteria_pred(criteria)
    return sum(1 for v in _flatten(range_vals) if pred(v))


def _make_criteria_pred(criteria):
    if isinstance(criteria, str):
        m = re.match(r"^(>=|<=|<>|>|<|=)(.+)$", criteria.strip())
        if m:
            op, val = m.group(1), m.group(2)
            try:
                val = float(val)
            except ValueError:
                pass
            ops = {"=": lambda a, b: a == b, "<>": lambda a, b: a != b,
                   ">": lambda a, b: a > b, "<": lambda a, b: a < b,
                   ">=": lambda a, b: a >= b, "<=": lambda a, b: a <= b}
            return lambda v, _op=ops[op], _val=val: _op(v, _val)
        if "*" in criteria or "?" in criteria:
            pattern = re.compile(
                re.escape(criteria).replace(r"\*", ".*").replace(r"\?", "."), re.IGNORECASE
            )
            return lambda v, p=pattern: isinstance(v, str) and bool(p.fullmatch(v))
        return lambda v, c=criteria: str(v).casefold() == c.casefold() if isinstance(v, str) else v == c
    return lambda v, c=criteria: v == c


def xl_abs(args: list) -> float:
    return abs(_coerce_num(args[0]))


def xl_round(args: list) -> float:
    return round(_coerce_num(args[0]), int(args[1]))


def xl_roundup(args: list) -> float:
    n, d = _coerce_num(args[0]), int(args[1])
    factor = 10 ** d
    return math.ceil(n * factor) / factor


def xl_rounddown(args: list) -> float:
    n, d = _coerce_num(args[0]), int(args[1])
    factor = 10 ** d
    return math.floor(n * factor) / factor


def xl_int(args: list) -> int:
    return math.floor(_coerce_num(args[0]))


def xl_mod(args: list) -> float:
    n, d = _coerce_num(args[0]), _coerce_num(args[1])
    return n - d * math.floor(n / d)


def xl_power(args: list) -> float:
    return _coerce_num(args[0]) ** _coerce_num(args[1])


def xl_sqrt(args: list) -> float:
    return math.sqrt(_coerce_num(args[0]))


def xl_ln(args: list) -> float:
    return math.log(_coerce_num(args[0]))


def xl_log(args: list) -> float:
    base = _coerce_num(args[1]) if len(args) > 1 else 10
    return math.log(_coerce_num(args[0]), base)


def xl_exp(args: list) -> float:
    return math.exp(_coerce_num(args[0]))


# ── Logic ─────────────────────────────────────────────────────────────────────

def xl_if(args: list) -> Any:
    condition, true_val = args[0], args[1]
    false_val = args[2] if len(args) > 2 else False
    return true_val if condition else false_val


def xl_ifs(args: list) -> Any:
    if len(args) % 2 != 0:
        raise ValueError("IFS requires pairs of condition/value")
    for i in range(0, len(args), 2):
        if args[i]:
            return args[i + 1]
    raise ValueError("IFS: no condition was TRUE")


def xl_and(args: list) -> bool:
    return all(bool(v) for v in _flatten(args))


def xl_or(args: list) -> bool:
    return any(bool(v) for v in _flatten(args))


def xl_not(args: list) -> bool:
    return not bool(args[0])


def xl_iferror(args: list) -> Any:
    # Executor catches exceptions and passes sentinel; handled in executor
    return args[0]


def xl_ifna(args: list) -> Any:
    return args[0]


def xl_isblank(args: list) -> bool:
    return args[0] is None or args[0] == ""


def xl_isnumber(args: list) -> bool:
    return isinstance(args[0], (int, float)) and not isinstance(args[0], bool)


def xl_istext(args: list) -> bool:
    return isinstance(args[0], str)


def xl_iserror(args: list) -> bool:
    return isinstance(args[0], str) and args[0].startswith("#")


# ── String ────────────────────────────────────────────────────────────────────

def xl_concatenate(args: list) -> str:
    return "".join(str(v) if v is not None else "" for v in _flatten(args))


def xl_concat(args: list) -> str:
    return xl_concatenate(args)


def xl_textjoin(args: list) -> str:
    delimiter = str(args[0])
    ignore_empty = bool(args[1])
    parts = [str(v) for v in _flatten(args[2:]) if not (ignore_empty and (v is None or v == ""))]
    return delimiter.join(parts)


def xl_left(args: list) -> str:
    text = str(args[0]) if args[0] is not None else ""
    n = int(args[1]) if len(args) > 1 else 1
    return text[:n]


def xl_right(args: list) -> str:
    text = str(args[0]) if args[0] is not None else ""
    n = int(args[1]) if len(args) > 1 else 1
    return text[-n:] if n else ""


def xl_mid(args: list) -> str:
    text = str(args[0]) if args[0] is not None else ""
    start = int(args[1]) - 1  # Excel is 1-indexed
    length = int(args[2])
    return text[start: start + length]


def xl_len(args: list) -> int:
    return len(str(args[0])) if args[0] is not None else 0


def xl_trim(args: list) -> str:
    return re.sub(r" +", " ", str(args[0]).strip())


def xl_upper(args: list) -> str:
    return str(args[0]).upper()


def xl_lower(args: list) -> str:
    return str(args[0]).lower()


def xl_substitute(args: list) -> str:
    text = str(args[0])
    old = str(args[1])
    new = str(args[2])
    if len(args) > 3:
        n = int(args[3])
        count = 0
        result = []
        start = 0
        while True:
            idx = text.find(old, start)
            if idx == -1:
                result.append(text[start:])
                break
            count += 1
            if count == n:
                result.append(text[start:idx])
                result.append(new)
                result.append(text[idx + len(old):])
                break
            result.append(text[start:idx + len(old)])
            start = idx + len(old)
        return "".join(result)
    return text.replace(old, new)


def xl_find(args: list) -> int:
    needle = str(args[0])
    haystack = str(args[1])
    start = int(args[2]) - 1 if len(args) > 2 else 0
    idx = haystack.find(needle, start)
    if idx == -1:
        raise ValueError(f"FIND: {needle!r} not found")
    return idx + 1


def xl_search(args: list) -> int:
    needle = str(args[0]).replace("*", ".*").replace("?", ".")
    haystack = str(args[1])
    start = int(args[2]) - 1 if len(args) > 2 else 0
    m = re.search(needle, haystack[start:], re.IGNORECASE)
    if not m:
        raise ValueError(f"SEARCH: {args[0]!r} not found")
    return start + m.start() + 1


def xl_text(args: list) -> str:
    return str(args[0])


def xl_value(args: list) -> float:
    return _coerce_num(args[0])


# ── Lookup / Reference ────────────────────────────────────────────────────────

def xl_vlookup(args: list) -> Any:
    lookup_val = args[0]
    table = args[1]  # list of rows
    col_idx = int(args[2]) - 1
    exact = not bool(args[3]) if len(args) > 3 else False  # default: approximate

    if not isinstance(table, list):
        raise ValueError("VLOOKUP: table must be a 2-D list")

    rows = table if isinstance(table[0], list) else [table]

    if exact:
        for row in rows:
            if row[0] == lookup_val:
                return row[col_idx]
        raise ValueError(f"VLOOKUP: {lookup_val!r} not found")
    else:
        result = None
        for row in rows:
            if row[0] <= lookup_val:
                result = row[col_idx]
            else:
                break
        if result is None:
            raise ValueError(f"VLOOKUP: {lookup_val!r} below minimum")
        return result


def xl_hlookup(args: list) -> Any:
    lookup_val = args[0]
    table = args[1]
    row_idx = int(args[2]) - 1
    exact = not bool(args[3]) if len(args) > 3 else False

    if not isinstance(table, list) or not isinstance(table[0], list):
        raise ValueError("HLOOKUP: table must be 2-D")

    if exact:
        for col_i, header in enumerate(table[0]):
            if header == lookup_val:
                return table[row_idx][col_i]
        raise ValueError(f"HLOOKUP: {lookup_val!r} not found")
    else:
        result_col = None
        for col_i, header in enumerate(table[0]):
            if header <= lookup_val:
                result_col = col_i
            else:
                break
        if result_col is None:
            raise ValueError(f"HLOOKUP: {lookup_val!r} below minimum")
        return table[row_idx][result_col]


def xl_index(args: list) -> Any:
    array = args[0]
    row_num = int(args[1])
    col_num = int(args[2]) if len(args) > 2 else 0

    if isinstance(array[0], list):
        row = array[row_num - 1] if row_num else array
        if col_num:
            return (row[col_num - 1] if isinstance(row, list) else row)
        return row
    return array[row_num - 1]


def xl_match(args: list) -> int:
    lookup_val = args[0]
    lookup_array = _flatten(args[1])
    match_type = int(args[2]) if len(args) > 2 else 1

    if match_type == 0:
        for i, v in enumerate(lookup_array):
            if v == lookup_val:
                return i + 1
        raise ValueError(f"MATCH: {lookup_val!r} not found")
    elif match_type == 1:
        result = None
        for i, v in enumerate(lookup_array):
            if v <= lookup_val:
                result = i + 1
            else:
                break
        if result is None:
            raise ValueError(f"MATCH: {lookup_val!r} not found")
        return result
    else:  # -1
        result = None
        for i, v in enumerate(lookup_array):
            if v >= lookup_val:
                result = i + 1
            else:
                break
        if result is None:
            raise ValueError(f"MATCH: {lookup_val!r} not found")
        return result


def xl_choose(args: list) -> Any:
    idx = int(args[0])
    return args[idx]


def xl_offset(args: list) -> Any:
    raise NotImplementedError("OFFSET requires spreadsheet context; use executor with Excel data")


# ── Date (basic) ──────────────────────────────────────────────────────────────

def xl_today(args: list) -> str:
    import datetime
    return str(datetime.date.today())


def xl_now(args: list) -> str:
    import datetime
    return str(datetime.datetime.now())


# ── Registry ──────────────────────────────────────────────────────────────────

REGISTRY: dict[str, Any] = {
    # Math / Stats
    "SUM": xl_sum,
    "AVERAGE": xl_average,
    "MIN": xl_min,
    "MAX": xl_max,
    "COUNT": xl_count,
    "COUNTA": xl_counta,
    "COUNTBLANK": xl_countblank,
    "SUMIF": xl_sumif,
    "COUNTIF": xl_countif,
    "ABS": xl_abs,
    "ROUND": xl_round,
    "ROUNDUP": xl_roundup,
    "ROUNDDOWN": xl_rounddown,
    "INT": xl_int,
    "MOD": xl_mod,
    "POWER": xl_power,
    "SQRT": xl_sqrt,
    "LN": xl_ln,
    "LOG": xl_log,
    "EXP": xl_exp,
    # Logic
    "IF": xl_if,
    "IFS": xl_ifs,
    "AND": xl_and,
    "OR": xl_or,
    "NOT": xl_not,
    "IFERROR": xl_iferror,
    "IFNA": xl_ifna,
    "ISBLANK": xl_isblank,
    "ISNUMBER": xl_isnumber,
    "ISTEXT": xl_istext,
    "ISERROR": xl_iserror,
    # String
    "CONCATENATE": xl_concatenate,
    "CONCAT": xl_concat,
    "TEXTJOIN": xl_textjoin,
    "LEFT": xl_left,
    "RIGHT": xl_right,
    "MID": xl_mid,
    "LEN": xl_len,
    "TRIM": xl_trim,
    "UPPER": xl_upper,
    "LOWER": xl_lower,
    "SUBSTITUTE": xl_substitute,
    "FIND": xl_find,
    "SEARCH": xl_search,
    "TEXT": xl_text,
    "VALUE": xl_value,
    # Lookup
    "VLOOKUP": xl_vlookup,
    "HLOOKUP": xl_hlookup,
    "INDEX": xl_index,
    "MATCH": xl_match,
    "CHOOSE": xl_choose,
    # Date
    "TODAY": xl_today,
    "NOW": xl_now,
}
