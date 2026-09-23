"""Pure Excel formula syntax checks (no workbook I/O)."""

from __future__ import annotations

import re


def validate_formula(formula: str) -> tuple[bool, str]:
    """Validate Excel formula syntax and safety"""
    if not formula.startswith("="):
        return False, "Formula must start with '='"

    # Remove the '=' prefix for validation
    formula = formula[1:]

    # Check for balanced parentheses
    parens = 0
    for c in formula:
        if c == "(":
            parens += 1
        elif c == ")":
            parens -= 1
        if parens < 0:
            return False, "Unmatched closing parenthesis"

    if parens > 0:
        return False, "Unclosed parenthesis"

    # Basic function name validation
    func_pattern = r"([A-Z]+)\("
    funcs = re.findall(func_pattern, formula)
    unsafe_funcs = {"INDIRECT", "HYPERLINK", "WEBSERVICE", "DGET", "RTD"}

    for func in funcs:
        if func in unsafe_funcs:
            return False, f"Unsafe function: {func}"

    return True, "Formula is valid"
