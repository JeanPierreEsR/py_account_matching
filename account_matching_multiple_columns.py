# !/usr/bin/env python3
import win32clipboard
import re
from bisect import bisect_left
from typing import List, Tuple
import sys

# --------- Regex for numbers ---------
NUM_RE = re.compile(r"""
    ^\s*
    (?P<open_par>\()?        
    \s*
    (?P<sign>-)?             
    \s*
    [\$€£]*                  
    \s*
    (?:S/\.)?
    \s*
    (?P<int>\d{1,3}(?:[.,]\d{3})*|\d+) 
    (?P<dec>[.,]\d+)?        
    \s*
    %?                       
    \s*
    \)?                      
    \s*$
""", re.VERBOSE)

TRAILING_NUM_RE = re.compile(r"""
    (?P<numblock>
        \s*
        \(?\s*
        -?\s*
        [\$€£]*\s*
        (?:S/\.)?\s*
        (?P<int>\d{1,3}(?:[.,]\d{3})*|\d+)
        (?P<dec>[.,]\d+)? 
        \s*%?\s*
        \)?\s*
    )$
""", re.VERBOSE)

# --------- Clipboard ---------
def read_clipboard_text() -> str:
    win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()
    return data

def write_clipboard_text(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
    finally:
        win32clipboard.CloseClipboard()

# --------- Number parsing ---------
def parse_number(text: str) -> float:
    m = NUM_RE.match(text)
    if not m:
        raise ValueError(f"Could not parse number: {text!r}")
    int_part_raw = m.group("int") or ""
    dec_part_raw = m.group("dec") or ""
    is_paren_neg = bool(m.group("open_par"))
    is_minus = bool(m.group("sign"))
    int_digits = re.sub(r"[.,\s]", "", int_part_raw)
    if dec_part_raw:
        dec_digits = dec_part_raw[1:].replace(" ", "")
        norm = f"{int_digits}.{dec_digits}"
    else:
        norm = int_digits
    val = float(norm)
    if is_minus or is_paren_neg:
        val = -val
    return val

# --------- Parsing ---------
def parse_pasted_block(prompt: str, default_prefix: str, use_last_only: bool = True) -> List[Tuple[str, List[float]]]:
    """
    - If use_last_only=True → only last column is taken as scalar value (wrapped in list).
    - If use_last_only=False → all numeric columns except first are taken as a vector.
    Treats '-' as 0.0 in any numeric column.
    """
    input(prompt)
    text = read_clipboard_text()
    lines = [ln.rstrip("\r").strip() for ln in text.split("\n") if ln.strip()]
    rows: List[Tuple[str, List[float]]] = []
    idx = 1
    for line in lines:
        parts = [p.strip() for p in line.split("\t")]
        if not parts or all(p == "" for p in parts):
            continue

        name = parts[0] or f"{default_prefix} {idx}"

        if use_last_only:
            last = parts[-1].strip()
            if last == "-":
                cols = [0.0]
            else:
                try:
                    cols = [parse_number(last)]
                except Exception:
                    cols = [0.0]
        else:
            # All columns after the first name column are numeric features
            cols: List[float] = []
            for p in parts[1:]:
                pp = p.strip()
                if pp == "" or pp == "-":
                    cols.append(0.0)
                    continue
                try:
                    cols.append(parse_number(pp))
                except Exception:
                    # Be forgiving: treat unparseable numeric columns as 0.0
                    cols.append(0.0)
            # If there were no extra columns, keep at least a scalar 0.0 to avoid dimension errors
            if not cols:
                cols = [0.0]

        rows.append((name, cols))
        idx += 1

    return rows

# --------- Formatting ---------
def format_money(x: float) -> str:
    return f"{x:,.6f}"

# --------- Scalar MITM ---------
def all_subset_sums(values: List[float]) -> List[Tuple[float, int]]:
    n = len(values)
    out: List[Tuple[float, int]] = []
    for mask in range(1 << n):
        s = 0.0
        m = mask
        i = 0
        while m:
            if m & 1:
                s += values[i]
            i += 1
            m >>= 1
        out.append((s, mask))
    return out

def meet_in_the_middle_best(target: float, vals: List[float]) -> Tuple[float, List[int]]:
    n = len(vals)
    mid = n // 2
    left_vals = vals[:mid]
    right_vals = vals[mid:]
    left_sums = all_subset_sums(left_vals)
    right_sums = all_subset_sums(right_vals)
    right_sums.sort(key=lambda x: x[0])
    right_only = [s for s, _ in right_sums]
    best_err = float("inf")
    best_total = 0.0
    best_left_mask = 0
    best_right_mask = 0
    for ls, lmask in left_sums:
        need = target - ls
        pos = bisect_left(right_only, need)
        for j in (pos-1, pos, pos+1):
            if 0 <= j < len(right_sums):
                rs, rmask = right_sums[j]
                total = ls + rs
                err = abs(total - target)
                if err < best_err or (err == best_err and abs(total) < abs(best_total)):
                    best_err = err
                    best_total = total
                    best_left_mask = lmask
                    best_right_mask = rmask
    chosen_indices: List[int] = []
    i, m = 0, best_left_mask
    while i < len(left_vals):
        if m & 1:
            chosen_indices.append(i)
        m >>= 1; i += 1
    i, m = 0, best_right_mask
    while i < len(right_vals):
        if m & 1:
            chosen_indices.append(mid + i)
        m >>= 1; i += 1
    return best_total, sorted(chosen_indices)

# --------- Vector MITM ---------
def all_subset_sums_vec(values: List[List[float]]) -> List[Tuple[List[float], int]]:
    n = len(values)
    dim = len(values[0])
    out: List[Tuple[List[float], int]] = []
    for mask in range(1 << n):
        s = [0.0] * dim
        m = mask
        i = 0
        while m:
            if m & 1:
                vi = values[i]
                for d in range(dim):
                    s[d] += vi[d]
            i += 1
            m >>= 1
        out.append((s, mask))
    return out

def squared_error(vec1: List[float], vec2: List[float]) -> float:
    return sum((a - b) ** 2 for a, b in zip(vec1, vec2))

def meet_in_the_middle_best_vec(target: List[float], vals: List[List[float]]) -> Tuple[List[float], List[int]]:
    n = len(vals)
    dim = len(target)
    mid = n // 2
    left_vals = vals[:mid]
    right_vals = vals[mid:]

    left_sums = all_subset_sums_vec(left_vals)
    right_sums = all_subset_sums_vec(right_vals)

    # Optimization: Sort right_sums by the first dimension to allow pruning
    right_sums.sort(key=lambda x: x[0][0])
    right_first_vals = [x[0][0] for x in right_sums]

    best_err = float("inf")
    best_sum = [0.0] * dim
    best_lmask = 0
    best_rmask = 0

    for ls, lmask in left_sums:
        # We need: ls[0] + rs[0] ≈ target[0]  =>  rs[0] ≈ target[0] - ls[0]
        target_r0 = target[0] - ls[0]
        
        # Binary search for the best starting point in the first dimension
        idx = bisect_left(right_first_vals, target_r0)
        
        # Check candidates to the right (>= target_r0)
        for i in range(idx, len(right_sums)):
            rs, rmask = right_sums[i]
            diff0 = (ls[0] + rs[0]) - target[0]
            # Pruning: if error in dim 0 alone exceeds best_err, stop this branch
            if diff0 * diff0 >= best_err:
                break
            
            s = [ls[d] + rs[d] for d in range(dim)]
            err = squared_error(s, target)
            if err < best_err:
                best_err = err
                best_sum = s
                best_lmask = lmask
                best_rmask = rmask

        # Check candidates to the left (< target_r0)
        for i in range(idx - 1, -1, -1):
            rs, rmask = right_sums[i]
            diff0 = (ls[0] + rs[0]) - target[0]
            if diff0 * diff0 >= best_err:
                break
            
            s = [ls[d] + rs[d] for d in range(dim)]
            err = squared_error(s, target)
            if err < best_err:
                best_err = err
                best_sum = s
                best_lmask = lmask
                best_rmask = rmask

    # Reconstruct indices
    best_idxs: List[int] = []
    i, m = 0, best_lmask
    while i < len(left_vals):
        if m & 1:
            best_idxs.append(i)
        i += 1; m >>= 1
    i, m = 0, best_rmask
    while i < len(right_vals):
        if m & 1:
            best_idxs.append(mid + i)
        i += 1; m >>= 1

    return best_sum, best_idxs

# --------- Main solver ---------
def solve_for_targets(source_a, source_b, use_last_only: bool = True):
    if not source_b:
        raise ValueError("Source B is empty after parsing.")

    names_b = [n for n, _ in source_b]
    vals_b  = [v for _, v in source_b]

    # Validate dimensions
    if any(len(v) == 0 for v in vals_b):
        raise ValueError("Some Source B rows have no numeric columns. Ensure last column (scalar) or at least one numeric column exists.")

    dim = len(vals_b[0])
    if any(len(v) != dim for v in vals_b):
        raise ValueError("Inconsistent numeric column count across Source B rows.")

    lines: List[str] = []

    for a_name, a_vals in source_a:
        if len(a_vals) != dim:
            raise ValueError(f"Dimension mismatch between Source A row '{a_name}' (len={len(a_vals)}) and Source B (len={dim}).")

        lines.append(f"TARGET\t{a_name}\t" + "\t".join(f"{x:.6f}" for x in a_vals))

        if dim == 1:  # scalar mode
            best_sum, idxs = meet_in_the_middle_best(a_vals[0], [v[0] for v in vals_b])
            diff = best_sum - a_vals[0]
            lines.append(f"BEST_TOTAL\t{best_sum:.6f}\tDIFF\t{diff:.6f}")
            lines.append(f"COUNT\t{len(idxs)}")
            for i in idxs:
                lines.append(f"{names_b[i]}\t{vals_b[i][0]:.6f}")
        else:  # vector mode
            best_sum, idxs = meet_in_the_middle_best_vec(a_vals, vals_b)
            diffs = [best_sum[d] - a_vals[d] for d in range(dim)]
            lines.append("BEST_TOTAL\t" + "\t".join(f"{x:.6f}" for x in best_sum))
            lines.append("DIFF\t" + "\t".join(f"{x:.6f}" for x in diffs))
            lines.append(f"COUNT\t{len(idxs)}")
            for i in idxs:
                lines.append(f"{names_b[i]}\t" + "\t".join(f"{x:.6f}" for x in vals_b[i]))
        lines.append("")

    output_tsv = "\n".join(lines).rstrip("\n")
    input("Results ready. Press Enter to copy to clipboard...")
    write_clipboard_text(output_tsv)

# --------- Main ---------
def main():
    try:
        print("\n=== Tie Accounts Across Sources (Scalar or Vector Mode) ===\n")
        print("Use only the last column as value? (y/n) [y]: ", end="")
        use_last_only = (input().strip().lower() or "y").startswith("y")

        source_a = parse_pasted_block("Copy Source A (target values) now, then press Enter.", "Source A Item", use_last_only)
        if not source_a:
            print("No Source A (target values) data provided. Exiting.")
            return

        source_b = parse_pasted_block("Copy Source B (accounts to loop through) now, then press Enter.", "Source B Item", use_last_only)
        if not source_b:
            print("No Source B (accounts to loop through) data provided. Exiting.")
            return

        solve_for_targets(source_a, source_b, use_last_only=use_last_only)
        print("Results copied to clipboard.")

    except Exception as e:
        # Print a clear error to stderr so you see *why* it “stops”
        print(f"\n[ERROR] {e}", file=sys.stderr)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
