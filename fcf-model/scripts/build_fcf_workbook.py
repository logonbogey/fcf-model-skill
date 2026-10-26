#!/usr/bin/env python3
"""
Build a normalized Free Cash Flow workbook in a fixed owner's-earnings format.

Reads a JSON spec describing one company's statement-of-cash-flows data and
emits an .xlsx that matches one fixed template: three columns per year
(value, 1x?, FCF), a duplicated last-year "Test" block, the SBC / NWC
normalization rows, CapEx, Normalized FCF, and optional Average rows.

Design notes (match the reference workbooks):
  * The FCF flag column is the AUTHORITATIVE method. To subtract any line from
    normalized FCF (SBC, working capital, or a judgment item like "Other"),
    set "fcf": true on it. The SUMIFS column sums everything flagged "x".
  * The manual value column is a cross-check. The build wires the manual
    NWC row to subtract exactly the same set (all working-capital lines plus
    any non-SBC adjustment flagged fcf), so the two methods always agree.
  * Averages reference the FCF-flag (SUMIFS) columns, the way the hand-built reference copies do.
  * Verification: OCF (sum of all reported inputs) is compared in pure Python to
    the spec's reported_ocf for every year and printed PASS/FAIL. This is the
    master tie-out and needs no formula engine.

Usage:
    python build_fcf_workbook.py <spec.json> <output.xlsx>

See reference/spec-template.json for the input schema (a made-up company filled in as an example).
"""
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- colour code (blue input, green link, black formula) ----
# Full 8-digit ARGB: a 6-digit value gets its alpha padded with 00 rather than FF,
# which does not match the reference workbook.
BLUE = "FF0000FF"    # hard-coded input
GREEN = "FF008000"   # link to another cell in the same workbook (a pure mirror)
BLACK = "FF000000"   # computed formula

# ---- number formats ----
# Accounting, whole units. The "$" variant marks the first row of a section and
# every total row; the plain variant is used for the rows in between. This is the
# convention in the hand-built reference workbook.
ACCT_D = r'_([$$-409]* #,##0_);_([$$-409]* \(#,##0\);_([$$-409]* "-"??_);_(@_)'
ACCT = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
PCT = "0.00%"

THIN = Side(style="thin")

# Column widths, taken from the hand-built reference workbook.
W_NARROW = 1.7109375     # spacer cols A, C
W_B = 4.5703125
W_LABEL = 52.0           # col D
W_VALUE = 14.28515625    # year value + FCF cols
W_FLAG = 3.7109375       # the "1x?" cols
W_RATIONALE = 80.0


def col(idx):
    return get_column_letter(idx)


def style(cell, color=None, bold=False, italic=False, size=None, nf=None,
          wrap=None):
    """Apply font + number format in one shot (openpyxl fonts are immutable)."""
    cell.font = Font(color=color, bold=bold or None, italic=italic or None,
                     size=size)
    if nf:
        cell.number_format = nf
    if wrap is not None:
        cell.alignment = Alignment(wrap_text=wrap)
    return cell


def yearcols(i):
    """Return (value_col, flag1x_col, fcf_col) 1-based indices for year block i (0-based)."""
    base = 5 + i * 3          # first year value column = E (5)
    return base, base + 1, base + 2


def build(spec, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    years = spec["years"]
    n = len(years)
    # one extra "Test" block duplicating the final year
    n_blocks = n + 1
    rationale_col = 5 + n_blocks * 3        # one col after the last fcf col

    # ---- titles ----
    style(ws.cell(row=2, column=2, value=spec["company"]), bold=True, size=12)
    style(ws.cell(row=3, column=2, value="Historical Statement of Cash Flows"), italic=True)
    style(ws.cell(row=4, column=2, value=spec.get("units", "($ in Millions, except per share)")), italic=True)

    # ---- header row 6 ----
    test_col_v = yearcols(n)[0]
    ws.cell(row=5, column=test_col_v, value="Test (=FY%s check)" % str(years[-1])[-2:])
    for i, yr in enumerate(list(years) + [years[-1]]):     # real years + test
        v, f1, fc = yearcols(i)
        style(ws.cell(row=6, column=v, value=yr), bold=True)
        ws.cell(row=6, column=f1, value="1x?")
        ws.cell(row=6, column=fc, value="FCF")
    style(ws.cell(row=6, column=rationale_col, value="Rationale"), bold=True)

    block_indices = list(range(n)) + [n - 1]   # which source-year each block reads (test reuses last)

    # ---- starting line (row 7) ----
    # US GAAP starts from Net Income; IFRS indirect-method statements often start from
    # "Profit before tax" or "Operating profit". The label is spec-driven; the values
    # array is still `net_income`. Everything downstream (tie-out, NI Delta) is unaffected
    # by the label. Note: for a non-NI starting line, the footer "NI Delta" row compares
    # last-year Normalized FCF to that starting figure, not to bottom-line net income.
    r = 7
    style(ws.cell(row=r, column=4, value=spec.get("starting_label", "Net Income")), bold=True)
    for bi, src in enumerate(block_indices):
        v, _, _ = yearcols(bi)
        # section head -> the "$" accounting variant
        style(ws.cell(row=r, column=v, value=spec["net_income"][src]), color=BLUE, nf=ACCT_D)
    ni_row = r

    # ---- adjustments section ----
    style(ws.cell(row=9, column=4, value="Adjustments to reconcile net earnings to operating cash flow:"), italic=True)
    r = 10
    adj_rows = []
    sbc_row = None
    fcf_flag_rows = []          # rows flagged fcf (any section)
    for idx, item in enumerate(spec["adjustments"]):
        ws.cell(row=r, column=4, value=item["label"])
        is_sbc = item.get("is_sbc", False)
        nf = ACCT_D if idx == 0 else ACCT      # "$" on the first row of the section only
        for bi, src in enumerate(block_indices):
            v, f1, fc = yearcols(bi)
            yr = str(years[src])
            val = item["values"].get(yr, None)
            style(ws.cell(row=r, column=v, value=val), color=BLUE, nf=nf)
            if item.get("one_time"):
                ws.cell(row=r, column=f1, value="x")
            if item.get("fcf"):
                ws.cell(row=r, column=fc, value="x")
        if item.get("rationale"):
            c = ws.cell(row=r, column=rationale_col, value=item["rationale"])
            c.alignment = Alignment(wrap_text=False)
        adj_rows.append(r)
        if item.get("fcf"):
            fcf_flag_rows.append(r)
        if is_sbc:
            sbc_row = r
        r += 1

    adj_first, adj_last = adj_rows[0], adj_rows[-1]

    # ---- working capital section ----
    r += 1
    style(ws.cell(row=r, column=4, value="Changes in working capital:"), italic=True)
    r += 1
    wc_rows = []
    for idx, item in enumerate(spec["working_capital"]):
        ws.cell(row=r, column=4, value=item["label"])
        nf = ACCT_D if idx == 0 else ACCT
        for bi, src in enumerate(block_indices):
            v, f1, fc = yearcols(bi)
            yr = str(years[src])
            val = item["values"].get(yr, None)
            style(ws.cell(row=r, column=v, value=val), color=BLUE, nf=nf)
            # working capital is always one-time and always fcf-subtracted
            ws.cell(row=r, column=f1, value="x")
            ws.cell(row=r, column=fc, value="x")
        if item.get("rationale"):
            ws.cell(row=r, column=rationale_col,
                    value=item.get("rationale", "Working-capital timing swing, not recurring. Reverse out of normalized FCF."))
        wc_rows.append(r)
        fcf_flag_rows.append(r)
        r += 1
    wc_first, wc_last = wc_rows[0], wc_rows[-1]

    # ---- OCF row ----
    r += 1
    ocf_row = r
    style(ws.cell(row=r, column=4, value="Net Cash Provided by Operating Activities"), bold=True)
    for bi in range(n_blocks):
        v, _, _ = yearcols(bi)
        c = col(v)
        style(ws.cell(row=r, column=v, value="=SUM(%s%d:%s%d)" % (c, ni_row, c, wc_last)),
              bold=True, nf=ACCT_D)

    # ---- SBC row ----
    r += 2
    sbc_disp_row = r
    style(ws.cell(row=r, column=4, value="SBC"), bold=True)
    if sbc_row:
        for bi in range(n_blocks):
            v, _, _ = yearcols(bi)
            c = col(v)
            # pure mirror of another cell in this workbook -> green
            style(ws.cell(row=r, column=v, value="=-%s%d" % (c, sbc_row)), color=GREEN, nf=ACCT_D)
    ws.cell(row=r, column=rationale_col, value="Add SBC back out of OCF (real cost to owners).")

    # ---- NWC / one-time row ----
    r += 1
    nwc_row = r
    ws.cell(row=r, column=4, value="NWC / one-time benefits to subtract")
    # manual value column = -(all WC rows + non-SBC fcf-flagged adjustment rows)
    manual_rows = sorted(set(wc_rows) | {x for x in fcf_flag_rows if x in adj_rows and x != sbc_row})
    for bi in range(n_blocks):
        v, _, fc = yearcols(bi)
        cv, cf = col(v), col(fc)
        manual_terms = "+".join("%s%d" % (cv, rr) for rr in manual_rows)
        style(ws.cell(row=r, column=v, value="=-(%s)" % manual_terms), nf=ACCT)
        # SUMIFS over the whole adjustments+WC block on the fcf flag
        style(ws.cell(row=r, column=fc,
                      value='=-SUMIFS(%s%d:%s%d,%s%d:%s%d,"x")' % (cv, adj_first, cv, wc_last, cf, adj_first, cf, wc_last)),
              nf=ACCT)
    ws.cell(row=r, column=rationale_col,
            value="Reverse working capital and any item flagged 'x' in the FCF column. Value col = manual sum; FCF col = SUMIFS on the flags. The two must agree.")

    # ---- Normalized Cash from Ops ----
    r += 1
    ncfo_row = r
    style(ws.cell(row=r, column=4, value="Normalized Cash from Ops"), bold=True)
    for bi in range(n_blocks):
        v, _, fc = yearcols(bi)
        cv, cf = col(v), col(fc)
        style(ws.cell(row=r, column=v, value="=+%s%d+SUM(%s%d:%s%d)" % (cv, ocf_row, cv, sbc_disp_row, cv, nwc_row)),
              bold=True, nf=ACCT_D)
        style(ws.cell(row=r, column=fc, value="=+%s%d+%s%d" % (cv, ocf_row, cf, nwc_row)),
              bold=True, nf=ACCT_D)

    # ---- CapEx ----
    r += 2
    capex_row = r
    ws.cell(row=r, column=4, value="Capital Expenditures")   # an input line, not a total
    has_capex = "capex" in spec and any(x is not None for x in spec["capex"])
    for bi, src in enumerate(block_indices):
        v, _, fc = yearcols(bi)
        cv, cf = col(v), col(fc)
        if has_capex:
            style(ws.cell(row=r, column=v, value=spec["capex"][src]), color=BLUE, nf=ACCT_D)
            style(ws.cell(row=r, column=fc, value="=+%s%d" % (cv, r)), color=GREEN, nf=ACCT_D)
    ws.cell(row=r, column=rationale_col,
            value=spec.get("capex_rationale", "Purchases of plant assets (maintenance + growth capex)."))

    # ---- Normalized FCF ----
    r += 2
    fcf_row = r
    style(ws.cell(row=r, column=4, value="Normalized Free Cash Flow"), bold=True)
    for bi in range(n_blocks):
        v, _, fc = yearcols(bi)
        cv, cf = col(v), col(fc)
        if has_capex:
            style(ws.cell(row=r, column=v, value="=+%s%d+%s%d" % (cv, ncfo_row, cv, capex_row)), bold=True, nf=ACCT_D)
            style(ws.cell(row=r, column=fc, value="=+%s%d+%s%d" % (cf, ncfo_row, cf, capex_row)), bold=True, nf=ACCT_D)
        else:
            # no-capex business (e.g. Airbnb): Normalized FCF = Normalized Cash from Ops
            style(ws.cell(row=r, column=v, value="=+%s%d" % (cv, ncfo_row)), bold=True, nf=ACCT_D)
            style(ws.cell(row=r, column=fc, value="=+%s%d" % (cf, ncfo_row)), bold=True, nf=ACCT_D)

    # ---- averages + valuation footer (reference the FCF-flag columns, like the reference copy) ----
    # All footer values live in column G (7), lined up under the first year's FCF column,
    # exactly as in the hand-built reference copy.
    GV = 7
    gcol = col(GV)
    fcf_flag_year_cells = [col(yearcols(i)[2]) + str(fcf_row) for i in range(n)]
    last_fcf_cell = fcf_flag_year_cells[-1]                       # last real year's Normalized FCF (FCF-flag col)
    last_ni_cell = col(yearcols(n - 1)[0]) + str(ni_row)         # last real year's Net Income (value col)
    ev = spec.get("enterprise_value")
    mc = spec.get("market_cap")
    vdate = spec.get("valuation_date", "")
    dsuffix = (" - " + vdate) if vdate else ""

    if ev:
        # Full valuation footer: Market Cap, Enterprise Value, then averages with EV-based
        # yields, the FCF CAGR, last-year FCF and its yield, and NI Delta.
        # Feeds the FCF Comparison Table directly.
        #
        # The year-count labels and the CAGR are self-computing off COUNTA of the OCF row,
        # so they stay correct if the series is extended: COUNTA counts every year block
        # plus the Test block, so the number of real years is COUNTA-1 and the number of
        # compounding periods between the first and last year is COUNTA-2.
        first_fcf_cell = fcf_flag_year_cells[0]
        ocf_span = "%s%d:BB%d" % (col(5), ocf_row, ocf_row)
        nyr = 'TEXT(COUNTA(%s)-1,0)' % ocf_span
        r += 3
        style(ws.cell(row=r, column=4, value="Market Cap" + dsuffix), bold=True)
        style(ws.cell(row=r, column=GV, value=mc), color=BLUE, bold=True, nf=ACCT)
        mc_cell = gcol + str(r)
        r += 1
        style(ws.cell(row=r, column=4, value="Enterprise Value" + dsuffix), bold=True)
        style(ws.cell(row=r, column=GV, value=ev), color=BLUE, bold=True, nf=ACCT)
        ev_cell = gcol + str(r)
        # FCF-yield denominator = the smaller of market cap and enterprise value.
        # Net-debt names (EV > market cap) get measured against market cap; net-cash
        # names (EV < market cap) against EV. Falls back to EV alone if market cap missing.
        denom = ("MIN(%s,%s)" % (mc_cell, ev_cell)) if mc else ev_cell
        r += 1
        style(ws.cell(row=r, column=4, value="Average"), bold=True)
        style(ws.cell(row=r, column=GV, value="=AVERAGE(%s)" % ",".join(fcf_flag_year_cells)), nf=ACCT_D)
        avg_cell = gcol + str(r)
        r += 1
        style(ws.cell(row=r, column=4, value='=%s& "-Yr Average FCF yield"' % nyr), bold=True)
        style(ws.cell(row=r, column=GV, value="=%s/%s" % (avg_cell, denom)), italic=True, nf=PCT)
        r += 1
        style(ws.cell(row=r, column=4, value='=%s& "-Yr Average FCF Growth Rate"' % nyr), bold=True)
        style(ws.cell(row=r, column=GV,
                      value="=(%s/%s)^(1/(COUNTA(%s)-2))-1" % (last_fcf_cell, first_fcf_cell, ocf_span)),
              italic=True, nf=PCT)
        if n >= 5:
            r += 1
            style(ws.cell(row=r, column=4, value="Avg Last 5 years"), bold=True)
            style(ws.cell(row=r, column=GV, value="=AVERAGE(%s)" % ",".join(fcf_flag_year_cells[-5:])), nf=ACCT_D)
            l5_cell = gcol + str(r)
            r += 1
            style(ws.cell(row=r, column=4, value="5-Yr Average FCF Yield"), bold=True)
            style(ws.cell(row=r, column=GV, value="=%s/%s" % (l5_cell, denom)), italic=True, nf=PCT)
        r += 1
        style(ws.cell(row=r, column=4, value="FCF Last Year"), bold=True)
        style(ws.cell(row=r, column=GV, value="=+%s" % last_fcf_cell), color=GREEN, italic=True, nf=ACCT_D)
        r += 1
        style(ws.cell(row=r, column=4, value="FCF Yield Last year"), bold=True)
        style(ws.cell(row=r, column=GV, value="=%s/%s" % (last_fcf_cell, denom)), italic=True, nf=PCT)
        r += 1
        style(ws.cell(row=r, column=4, value="NI Delta"), bold=True)
        style(ws.cell(row=r, column=GV, value="=%s-%s" % (last_fcf_cell, last_ni_cell)), nf=ACCT_D)
        nid_cell = gcol + str(r)
        r += 1
        style(ws.cell(row=r, column=4, value="NI Delta %"), bold=True)
        style(ws.cell(row=r, column=GV, value="=%s/%s" % (nid_cell, last_ni_cell)), italic=True, nf=PCT)
    elif spec.get("averages", True):
        # Fallback footer: no enterprise value available, so no yields. Step 5 of the
        # workflow says not to ship this silently -- record why in _valuation_comment.
        r += 3
        style(ws.cell(row=r, column=4, value="Average"), bold=True)
        style(ws.cell(row=r, column=GV, value="=AVERAGE(%s)" % ",".join(fcf_flag_year_cells)), nf=ACCT_D)
        if n >= 5:
            r += 2
            style(ws.cell(row=r, column=4, value="Avg Last 5 years"), bold=True)
            style(ws.cell(row=r, column=GV, value="=AVERAGE(%s)" % ",".join(fcf_flag_year_cells[-5:])), nf=ACCT_D)
        if mc:
            r += 2
            style(ws.cell(row=r, column=4, value="Market Cap"), bold=True)
            style(ws.cell(row=r, column=5, value=mc), color=BLUE, bold=True, nf=ACCT)
            r += 2
            style(ws.cell(row=r, column=4, value="Market Cap / Normalized FCF"), bold=True)
            style(ws.cell(row=r, column=5, value="=+%s/%s" % (col(5) + str(r - 2), last_fcf_cell)), nf=ACCT)

    # ---- borders ----
    # Underline the end of each input section (last adjustment, last working-capital
    # row) and the NWC row, across every block. Fence the Test block off from the real
    # years with a vertical rule down the last real year's FCF column.
    last_block_col = yearcols(n_blocks - 1)[2]
    rule_col = yearcols(n - 1)[2]
    underline_rows = {adj_last, wc_last, nwc_row}
    for rr in range(6, fcf_row + 1):
        for cc in range(5, last_block_col + 1):
            bottom = THIN if rr in underline_rows else None
            right = THIN if cc == rule_col else None
            if bottom or right:
                ws.cell(row=rr, column=cc).border = Border(bottom=bottom, right=right)

    # ---- column widths ----
    ws.column_dimensions["A"].width = W_NARROW
    ws.column_dimensions["B"].width = W_B
    ws.column_dimensions["C"].width = W_NARROW
    ws.column_dimensions["D"].width = W_LABEL
    for bi in range(n_blocks):
        v, f1, fc = yearcols(bi)
        ws.column_dimensions[col(v)].width = W_VALUE
        ws.column_dimensions[col(f1)].width = W_FLAG
        ws.column_dimensions[col(fc)].width = W_VALUE
    ws.column_dimensions[col(rationale_col)].width = W_RATIONALE

    # keep the labels and the year headers on screen while scrolling right/down
    ws.freeze_panes = "E7"

    wb.save(out_path)
    return ocf_row


def verify(spec):
    """Master tie-out: sum every reported input per year, compare to reported_ocf."""
    years = spec["years"]
    reported = spec.get("reported_ocf")
    if not reported:
        print("  (no reported_ocf in spec -> skipping OCF tie-out; ADD IT before trusting the file)")
        return True
    ok = True
    for i, yr in enumerate(years):
        total = spec["net_income"][i]
        for item in spec["adjustments"] + spec["working_capital"]:
            total += item["values"].get(str(yr), 0) or 0
        match = abs(total - reported[i]) < 0.5
        ok = ok and match
        print("  %s  computed OCF %10.1f  reported %10.1f  %s"
              % (yr, total, reported[i], "PASS" if match else "*** FAIL ***"))
    return ok


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    spec_path, out_path = sys.argv[1], sys.argv[2]
    with open(spec_path, encoding="utf-8") as fh:
        spec = json.load(fh)

    print("Verifying OCF tie-out (master check)...")
    ok = verify(spec)
    if not ok:
        print("\n*** OCF tie-out FAILED. Fix the line items or signs before using this file. ***")
        # still build it so the user can inspect, but flag loudly
    build(spec, out_path)
    print("\nWrote %s" % out_path)
    print("Remaining gates to confirm in Excel: (1) value col = FCF col for Normalized FCF every year,")
    print("(2) overlap years match across filings, (3) signs sane (D&A/SBC/impairments +, CapEx -, divestiture gains -).")
    print("Then hand the judgment checkpoints (Other line, pension, CapEx scope) to a person before trusting the model.")


if __name__ == "__main__":
    main()
