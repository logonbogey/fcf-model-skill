---
name: fcf-model
description: Build a normalized free cash flow model for a public company from its filings. Use when the user says "build an FCF model for [ticker]", "normalize FCF for [company]", or wants an owner's-earnings calculation from the statement of cash flows. Pulls the statement from SEC EDGAR (US filers) or the annual report (non-US filers), applies the rules in reference/line-item-rules.md, and emits a formatted Excel workbook with a built-in tie-out check.
---

# FCF Model Builder

Builds a normalized free cash flow workbook for one company. The output is an `.xlsx` with three columns per year (value, `1x?`, `FCF`), a duplicated last-year "Test" block, stock comp and working-capital normalization rows, CapEx, Normalized FCF, average rows, and an optional valuation footer.

## Core idea

Start from the statement's starting line (Net Income under US GAAP; often *Profit before tax* or *Operating profit* under IFRS, set with `starting_label`) and walk down the reconciliation to an owner's-earnings figure. The `FCF` flag column is authoritative: set `"fcf": true` on anything that should come out of normalized FCF. The builder computes Normalized FCF two ways (a manual column and a SUMIFS over flagged lines) that must agree, and checks that every line re-adds to the operating cash flow the company reported.

**Manual first.** Build your first few models by hand, then compare. The tool is a cross-check, not a replacement for understanding the number.

## Files

- `reference/line-item-rules.md` is the judgment card. It ships almost empty. Fill it in with your own calls.
- `reference/spec-template.json` is the input schema, filled in with a made-up company.
- `scripts/build_fcf_workbook.py` builds the workbook and runs the tie-out. Needs `openpyxl`.

## Workflow

1. **Confirm scope.** Default to the most recent 10 years. Shorten only on purpose (not enough history, or a spin-off, fiscal-year change, or big acquisition makes earlier years non-comparable) and write down why.

2. **Get the statement of cash flows.**
   - **US filer, SEC EDGAR.** Look up the company's CIK (Central Index Key), then pull `https://data.sec.gov/submissions/CIK{cik padded to 10 digits}.json` and filter to `form == "10-K"`. Each 10-K shows three years, so about four filings cover a decade. For each one, read `FilingSummary.xml`, find the "Consolidated Statements of Cash Flows" entry, fetch its `R#.htm`, and strip it to a label/value table. Use `curl` with a `User-Agent: Your Name your@email.com` header. The SEC rejects requests without one.
   - **Non-US filer.** Pull the annual report PDF from the company's investor relations page and read the cash flow statement off it. An aggregator site can fill gaps, but an aggregator-only build still has to pass the tie-out.
   - **Overlap years.** Any year that shows up in two filings has to match in both.

3. **Build the spec.** Copy `reference/spec-template.json` and fill in the company, `units`, years, the starting line, every `adjustments` line, every `working_capital` line, `capex`, and `reported_ocf` for each year. Copy every value exactly how the company printed it. Keep each line's label the way the company reported it rather than renaming it to fit a template.

4. **Tag every line** per `reference/line-item-rules.md`, and write a one-sentence `rationale` for each call. When a line is not on the card, make a call, write the reasoning, and flag it for the user. When a line is genuinely subjective (a vague "Other," leases, pension swings), stop and ask instead of guessing.

5. **Optional valuation footer.** Set `market_cap`, `enterprise_value` (market cap + total debt - cash), and `valuation_date`. Every yield divides by `MIN(market cap, EV)`. Keep the whole model in the reporting currency.

6. **Run the builder.**
   ```
   python scripts/build_fcf_workbook.py spec.json "TICKER 2016-2025 FCF calculation.xlsx"
   ```
   It prints PASS/FAIL per year for the tie-out and writes the workbook either way, so a failure can be inspected.

## Verification gates (all must pass before you trust the model)

1. **Tie-out.** The starting line plus every adjustment and working-capital line equals reported operating cash flow, every year. A FAIL means a line got missed or copied wrong.
2. **Two-method agreement.** The manual Normalized FCF column equals the SUMIFS column every year.
3. **Overlap years** match across filings.
4. **Sign sanity.** D&A and stock comp positive, CapEx negative.
5. **Footer sanity** (if used). Market cap and EV are positive and dated, and the yields look reasonable.

Then reread the model once, adversarially. Is any line counted twice? Does normalized FCF sitting far above or far below net income make sense for this business? Would a different defensible call on one line move the answer a lot? If so, that line is a judgment call for a person, not one to settle silently.

## Conventions

- Excel color code: blue = hard-coded input, black = formula, green = link to another cell in the same workbook.
- Units match the filing. State the currency for non-US companies, e.g. `(EUR in Millions)`.
- File names: `<TICKER> <startyear>-<endyear> FCF calculation.xlsx`.
- The builder handles all formatting. If the format needs to change, change the script so every future model gets it.
