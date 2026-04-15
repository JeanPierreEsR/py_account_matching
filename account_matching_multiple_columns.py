# !/usr/bin/env python3
import win32clipboard
import re
import time
from bisect import bisect_left
from collections import defaultdict
from typing import Dict, List, Optional, Tuple
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


def meet_in_the_middle_top_k(
    target: float,
    vals: List[float],
    k: int = 1,
    required_count: Optional[int] = None,
) -> List[Tuple[float, List[int]]]:
    """
    Returns up to k best (total, sorted_indices) subsets ordered by |total - target|.
    If required_count is given, only subsets of exactly that size are considered.
    Uses a window search around the binary-search position; window = max(k*10, 50).
    """
    n = len(vals)
    mid = n // 2
    left_vals = vals[:mid]
    right_vals = vals[mid:]

    left_sums = all_subset_sums(left_vals)   # includes empty subset (0.0, 0)
    right_sums = all_subset_sums(right_vals)

    # Group right sums by popcount for fast required_count lookup
    right_by_count: Dict[int, List[Tuple[float, int]]] = defaultdict(list)
    for rs, rmask in right_sums:
        right_by_count[bin(rmask).count('1')].append((rs, rmask))
    for c in right_by_count:
        right_by_count[c].sort()

    # Sorted all right sums (for unconstrained case)
    right_sums_sorted = sorted(right_sums, key=lambda x: x[0])
    right_only = [s for s, _ in right_sums_sorted]

    window = max(k * 10, 50)
    candidates: List[Tuple[float, float, int, int]] = []  # (err, abs_total, lmask, rmask)

    for ls, lmask in left_sums:
        lcount = bin(lmask).count('1')
        need = target - ls

        if required_count is not None:
            needed_rcount = required_count - lcount
            if needed_rcount < 0 or needed_rcount > len(right_vals):
                continue
            rlist = right_by_count.get(needed_rcount, [])
            if not rlist:
                continue
            rvals = [s for s, _ in rlist]
            pos = bisect_left(rvals, need)
            for j in range(max(0, pos - window), min(len(rlist), pos + window + 1)):
                rs, rmask = rlist[j]
                total = ls + rs
                candidates.append((abs(total - target), abs(total), lmask, rmask))
        else:
            pos = bisect_left(right_only, need)
            for j in range(max(0, pos - window), min(len(right_sums_sorted), pos + window + 1)):
                rs, rmask = right_sums_sorted[j]
                total = ls + rs
                candidates.append((abs(total - target), abs(total), lmask, rmask))

    candidates.sort()

    results: List[Tuple[float, List[int]]] = []
    seen: set = set()
    for _, _, lmask, rmask in candidates:
        key = (lmask, rmask)
        if key in seen:
            continue
        seen.add(key)
        idxs: List[int] = []
        i, m = 0, lmask
        while i < len(left_vals):
            if m & 1:
                idxs.append(i)
            m >>= 1; i += 1
        i, m = 0, rmask
        while i < len(right_vals):
            if m & 1:
                idxs.append(mid + i)
            m >>= 1; i += 1
        total = sum(vals[idx] for idx in idxs)
        results.append((total, sorted(idxs)))
        if len(results) >= k:
            break

    return results


# --------- Vector MITM ---------
def all_subset_sums_vec(values: List[List[float]]) -> List[Tuple[List[float], int]]:
    n = len(values)
    if n == 0:
        return []
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


def meet_in_the_middle_top_k_vec(
    target: List[float],
    vals: List[List[float]],
    k: int = 1,
    required_count: Optional[int] = None,
) -> List[Tuple[List[float], List[int]]]:
    """
    Returns up to k best (sum_vec, sorted_indices) ordered by squared error from target.
    If required_count is given, only subsets of exactly that size are considered.
    """
    n = len(vals)
    dim = len(target)
    mid = n // 2
    left_vals = vals[:mid]
    right_vals = vals[mid:]

    # Handle empty left half (when n == 1)
    left_sums: List[Tuple[List[float], int]] = (
        all_subset_sums_vec(left_vals) if left_vals else [([0.0] * dim, 0)]
    )
    right_sums = all_subset_sums_vec(right_vals)

    # Group right sums by popcount
    right_by_count: Dict[int, List[Tuple[List[float], int]]] = defaultdict(list)
    for rv, rmask in right_sums:
        right_by_count[bin(rmask).count('1')].append((rv, rmask))
    for c in right_by_count:
        right_by_count[c].sort(key=lambda x: x[0][0])

    # Sorted all right sums by first dimension (for unconstrained case)
    right_sums_sorted = sorted(right_sums, key=lambda x: x[0][0])
    right_first_vals = [x[0][0] for x in right_sums_sorted]

    window = max(k * 10, 50)
    candidates: List[Tuple[float, int, int]] = []  # (err, lmask, rmask)

    for ls, lmask in left_sums:
        lcount = bin(lmask).count('1')
        target_r0 = target[0] - ls[0]

        if required_count is not None:
            needed_rcount = required_count - lcount
            if needed_rcount < 0 or needed_rcount > len(right_vals):
                continue
            rlist = right_by_count.get(needed_rcount, [])
            if not rlist:
                continue
            rfirst = [x[0][0] for x in rlist]
            pos = bisect_left(rfirst, target_r0)
            for j in range(max(0, pos - window), min(len(rlist), pos + window + 1)):
                rv, rmask = rlist[j]
                s = [ls[d] + rv[d] for d in range(dim)]
                err = squared_error(s, target)
                candidates.append((err, lmask, rmask))
        else:
            pos = bisect_left(right_first_vals, target_r0)
            for j in range(max(0, pos - window), min(len(right_sums_sorted), pos + window + 1)):
                rv, rmask = right_sums_sorted[j]
                s = [ls[d] + rv[d] for d in range(dim)]
                err = squared_error(s, target)
                candidates.append((err, lmask, rmask))

    candidates.sort(key=lambda x: x[0])

    results: List[Tuple[List[float], List[int]]] = []
    seen: set = set()
    for _, lmask, rmask in candidates:
        key = (lmask, rmask)
        if key in seen:
            continue
        seen.add(key)
        idxs: List[int] = []
        i, m = 0, lmask
        while i < len(left_vals):
            if m & 1:
                idxs.append(i)
            m >>= 1; i += 1
        i, m = 0, rmask
        while i < len(right_vals):
            if m & 1:
                idxs.append(mid + i)
            m >>= 1; i += 1
        s_vec = [sum(vals[idx][d] for idx in idxs) for d in range(dim)]
        results.append((s_vec, sorted(idxs)))
        if len(results) >= k:
            break

    return results


# --------- Main solver ---------
def solve_for_targets(
    source_a,
    source_b,
    use_last_only: bool = True,
    k_alternatives: int = 1,
    required_count: Optional[int] = None,
    best_per_count: bool = False,
    t_start: Optional[float] = None,
):
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

    n_b = len(vals_b)
    lines: List[str] = []

    for a_name, a_vals in source_a:
        if len(a_vals) != dim:
            raise ValueError(f"Dimension mismatch between Source A row '{a_name}' (len={len(a_vals)}) and Source B (len={dim}).")

        lines.append(f"TARGET\t{a_name}\t" + "\t".join(f"{x:.6f}" for x in a_vals))

        # Determine which counts to iterate over
        if best_per_count:
            max_cnt = required_count if required_count is not None else n_b
            count_range = list(range(max_cnt, 0, -1))
        elif required_count is not None:
            count_range = [required_count]
        else:
            count_range = [None]  # unconstrained

        for cnt in count_range:
            if best_per_count:
                lines.append(f"  COUNT {cnt}")

            if dim == 1:
                matches = meet_in_the_middle_top_k(
                    a_vals[0], [v[0] for v in vals_b],
                    k=k_alternatives, required_count=cnt,
                )
            else:
                matches = meet_in_the_middle_top_k_vec(
                    a_vals, vals_b,
                    k=k_alternatives, required_count=cnt,
                )

            if not matches:
                lines.append("  (no valid subset)")
                continue

            for match_idx, match in enumerate(matches, 1):
                if k_alternatives > 1:
                    lines.append(f"  MATCH {match_idx}")

                if dim == 1:
                    best_sum, idxs = match
                    used = set(idxs)
                    diff = best_sum - a_vals[0]
                    lines.append(f"BEST_TOTAL\t{best_sum:.6f}\tDIFF\t{diff:.6f}")
                    lines.append(f"COUNT\t{len(idxs)}")
                    for i in range(n_b):
                        flag = 1 if i in used else 0
                        lines.append(f"{names_b[i]}\t{vals_b[i][0]:.6f}\t{flag}")
                else:
                    best_sum, idxs = match
                    used = set(idxs)
                    diffs = [best_sum[d] - a_vals[d] for d in range(dim)]
                    lines.append("BEST_TOTAL\t\t" + "\t".join(f"{x:.6f}" for x in best_sum))
                    lines.append("DIFF\t\t" + "\t".join(f"{x:.6f}" for x in diffs))
                    lines.append(f"COUNT\t\t{len(idxs)}")
                    for i in range(n_b):
                        flag = 1 if i in used else 0
                        lines.append(f"{names_b[i]}\t\t" + "\t".join(f"{x:.6f}" for x in vals_b[i]) + f"\t{flag}")

        lines.append("")

    output_tsv = "\n".join(lines).rstrip("\n")
    if t_start is not None:
        elapsed = time.time() - t_start
        print(f"Computation time: {elapsed:.3f}s")
    input("Results ready. Press Enter to copy to clipboard...")
    write_clipboard_text(output_tsv)

# --------- Main ---------
def main():
    try:
        print("\n=== Tie Accounts Across Sources (Scalar or Vector Mode) ===\n")
        print("Use only the last column as value? (y/n) [y]: ", end="")
        use_last_only = (input().strip().lower() or "y").startswith("y")

        print("Best option per count (show best match for each subset size, highest to lowest)? (y/n) [n]: ", end="")
        best_per_count = (input().strip().lower() or "n").startswith("y")

        print("Number of alternatives [1]: ", end="")
        raw = input().strip()
        k_alternatives = int(raw) if raw.isdigit() and int(raw) >= 1 else 1

        print("Mandatory number of rows per match (0 = any) [0]: ", end="")
        raw = input().strip()
        required_count: Optional[int] = int(raw) if raw.isdigit() and int(raw) >= 1 else None

        source_a = parse_pasted_block("Copy Source A (target values) now, then press Enter.", "Source A Item", use_last_only)
        if not source_a:
            print("No Source A (target values) data provided. Exiting.")
            return

        source_b = parse_pasted_block("Copy Source B (accounts to loop through) now, then press Enter.", "Source B Item", use_last_only)
        if not source_b:
            print("No Source B (accounts to loop through) data provided. Exiting.")
            return
        t_start = time.time()

        solve_for_targets(
            source_a, source_b, use_last_only,
            k_alternatives=k_alternatives,
            required_count=required_count,
            best_per_count=best_per_count,
            t_start=t_start,
        )
        print("Results copied to clipboard.")

    except Exception as e:
        print(f"\n[ERROR] {e}", file=sys.stderr)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
