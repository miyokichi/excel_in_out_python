"""Parse Excel formula strings into a JSON-serializable AST.

Grammar (simplified):
  expression     : concat_expr (comparison_op concat_expr)*
  concat_expr    : additive (& additive)*
  additive       : multiplicative ((+|-) multiplicative)*
  multiplicative : power ((*|/) power)*
  power          : unary (^ unary)*
  unary          : (OPERATOR-PREFIX)? percent
  percent        : atom (OPERATOR-POSTFIX %)?
  atom           : literal | cell_ref | range | function_call | '(' expression ')'
"""

import re
from openpyxl.formula import Tokenizer as _Tokenizer


# ── Token helpers ─────────────────────────────────────────────────────────────

_CELL_RE = re.compile(
    r"^(?:([A-Za-z_][\w.]*|'[^']+')\!)?(\$?[A-Za-z]{1,3}\$?\d+)$"
)
_RANGE_RE = re.compile(
    r"^(?:([A-Za-z_][\w.]*|'[^']+')\!)?(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3}\$?\d+)$"
)


def _is_range(value: str) -> bool:
    return bool(_RANGE_RE.match(value))


def _is_cell(value: str) -> bool:
    return bool(_CELL_RE.match(value))


# ── Token wrapper ─────────────────────────────────────────────────────────────

class _Token:
    __slots__ = ("type", "subtype", "value")

    def __init__(self, type_: str, subtype: str, value: str) -> None:
        self.type = type_
        self.subtype = subtype
        self.value = value

    def __repr__(self) -> str:
        return f"Token({self.type!r}, {self.subtype!r}, {self.value!r})"


def _tokenize(formula: str) -> list[_Token]:
    """Return a flat list of tokens (whitespace stripped)."""
    tok = _Tokenizer(formula)
    result: list[_Token] = []
    for t in tok.items:
        if t.type in ("WSPACE", "WHITE-SPACE"):
            continue
        result.append(_Token(t.type, t.subtype, t.value))
    return result


# ── Parser ────────────────────────────────────────────────────────────────────

class ParseError(ValueError):
    pass


class _Parser:
    def __init__(self, tokens: list[_Token]) -> None:
        self._tokens = tokens
        self._pos = 0

    # ── Peek / consume ────────────────────────────────────────────────────────

    def _peek(self) -> _Token | None:
        if self._pos < len(self._tokens):
            return self._tokens[self._pos]
        return None

    def _consume(self) -> _Token:
        t = self._tokens[self._pos]
        self._pos += 1
        return t

    def _expect_type_sub(self, type_: str, subtype: str) -> _Token:
        t = self._peek()
        if t is None or t.type != type_ or t.subtype != subtype:
            raise ParseError(f"Expected {type_}/{subtype}, got {t!r}")
        return self._consume()

    # ── Grammar rules ─────────────────────────────────────────────────────────

    def parse(self) -> dict:
        node = self._expression()
        if self._peek() is not None:
            raise ParseError(f"Unexpected token at end: {self._peek()!r}")
        return node

    def _expression(self) -> dict:
        """Comparison: = <> < > <= >="""
        left = self._concat()
        while True:
            t = self._peek()
            if t and t.type == "OPERATOR-INFIX" and t.value in ("=", "<>", "<", ">", "<=", ">="):
                op = self._consume().value
                right = self._concat()
                left = {"type": "BinaryOp", "op": op, "left": left, "right": right}
            else:
                break
        return left

    def _concat(self) -> dict:
        """& string concatenation"""
        left = self._additive()
        while True:
            t = self._peek()
            if t and t.type == "OPERATOR-INFIX" and t.value == "&":
                self._consume()
                right = self._additive()
                left = {"type": "BinaryOp", "op": "&", "left": left, "right": right}
            else:
                break
        return left

    def _additive(self) -> dict:
        left = self._multiplicative()
        while True:
            t = self._peek()
            if t and t.type == "OPERATOR-INFIX" and t.value in ("+", "-"):
                op = self._consume().value
                right = self._multiplicative()
                left = {"type": "BinaryOp", "op": op, "left": left, "right": right}
            else:
                break
        return left

    def _multiplicative(self) -> dict:
        left = self._power()
        while True:
            t = self._peek()
            if t and t.type == "OPERATOR-INFIX" and t.value in ("*", "/"):
                op = self._consume().value
                right = self._power()
                left = {"type": "BinaryOp", "op": op, "left": left, "right": right}
            else:
                break
        return left

    def _power(self) -> dict:
        left = self._unary()
        t = self._peek()
        if t and t.type == "OPERATOR-INFIX" and t.value == "^":
            self._consume()
            right = self._unary()
            return {"type": "BinaryOp", "op": "^", "left": left, "right": right}
        return left

    def _unary(self) -> dict:
        t = self._peek()
        if t and t.type == "OPERATOR-PREFIX" and t.value in ("-", "+"):
            op = self._consume().value
            operand = self._percent()
            return {"type": "UnaryOp", "op": op, "operand": operand}
        return self._percent()

    def _percent(self) -> dict:
        node = self._atom()
        t = self._peek()
        if t and t.type == "OPERATOR-POSTFIX" and t.value == "%":
            self._consume()
            return {"type": "UnaryOp", "op": "%", "operand": node}
        return node

    def _atom(self) -> dict:
        t = self._peek()
        if t is None:
            raise ParseError("Unexpected end of formula")

        # Parenthesised sub-expression
        if t.type == "PAREN" and t.subtype == "OPEN":
            self._consume()
            node = self._expression()
            self._expect_type_sub("PAREN", "CLOSE")
            return node

        # Array literal {1,2,3}
        if t.type == "ARRAY" and t.subtype == "OPEN":
            return self._array()

        # Function call
        if t.type == "FUNC" and t.subtype == "OPEN":
            return self._function()

        # Operand: number, string, boolean, error, cell ref, range
        if t.type == "OPERAND":
            self._consume()
            return self._operand_node(t)

        raise ParseError(f"Unexpected token: {t!r}")

    def _operand_node(self, t: _Token) -> dict:
        sub = t.subtype
        val = t.value

        if sub == "NUMBER":
            num = float(val) if ("." in val or "e" in val.lower()) else int(val)
            return {"type": "Literal", "kind": "number", "value": num}

        if sub == "TEXT":
            return {"type": "Literal", "kind": "string", "value": val[1:-1]}

        if sub == "LOGICAL":
            return {"type": "Literal", "kind": "bool", "value": val.upper() == "TRUE"}

        if sub == "ERROR":
            return {"type": "Literal", "kind": "error", "value": val}

        if sub == "RANGE":
            if _is_range(val):
                m = _RANGE_RE.match(val)
                return {"type": "Range", "sheet": m.group(1), "start": m.group(2), "end": m.group(3)}
            if _is_cell(val):
                m = _CELL_RE.match(val)
                return {"type": "CellRef", "sheet": m.group(1), "address": m.group(2)}
            return {"type": "NamedRange", "name": val}

        return {"type": "Operand", "subtype": sub, "value": val}

    def _function(self) -> dict:
        t = self._consume()  # FUNC/OPEN, e.g. "SUM("
        name = t.value.rstrip("(").upper()
        args: list[dict] = []

        # Empty argument list
        close = self._peek()
        if close and close.type == "FUNC" and close.subtype == "CLOSE":
            self._consume()
            return {"type": "Function", "name": name, "args": args}

        while True:
            t2 = self._peek()
            # Empty arg slot (e.g. IF(,,))
            if t2 and t2.type == "SEP" and t2.subtype == "ARG":
                args.append({"type": "Literal", "kind": "empty", "value": None})
                self._consume()
                continue
            if t2 and t2.type == "FUNC" and t2.subtype == "CLOSE":
                self._consume()
                break
            args.append(self._expression())
            t3 = self._peek()
            if t3 and t3.type == "SEP" and t3.subtype == "ARG":
                self._consume()
            elif t3 and t3.type == "FUNC" and t3.subtype == "CLOSE":
                self._consume()
                break
            else:
                raise ParseError(f"Expected ',' or ')' in function args, got {t3!r}")

        return {"type": "Function", "name": name, "args": args}

    def _array(self) -> dict:
        """Parse {1,2;3,4} array literals."""
        self._consume()  # ARRAY/OPEN '{'
        elements: list = []
        row: list = []
        while True:
            t = self._peek()
            if t is None:
                raise ParseError("Unterminated array literal")
            if t.type == "ARRAY" and t.subtype == "CLOSE":
                self._consume()
                if row:
                    elements.append(row)
                break
            if t.type == "SEP":
                self._consume()
                if t.value == ";":  # row separator
                    elements.append(row)
                    row = []
                # comma: column separator — just continue, expression already consumed
                continue
            row.append(self._expression())
        return {"type": "Array", "elements": elements}


# ── Public API ────────────────────────────────────────────────────────────────

def parse(formula: str) -> dict:
    """Parse an Excel formula string (with or without leading '=') into an AST dict."""
    if not formula.startswith("="):
        formula = "=" + formula
    tokens = _tokenize(formula)
    return _Parser(tokens).parse()
