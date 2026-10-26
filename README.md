# fcf-model

A Claude Code skill that builds a normalized free cash flow model for a public company straight from its filings, and checks its own work before you trust it.

It goes with the write-up *I Taught an AI to Think Like an Analyst*, which explains what each file does and why the folder is split the way it is.

## What is in here

```
fcf-model/
├── SKILL.md                    the instructions
├── reference/
│   ├── line-item-rules.md      the judgment card (almost empty, fill it in)
│   └── spec-template.json      the tagged file, with a made-up example
└── scripts/
    └── build_fcf_workbook.py   builds the spreadsheet and checks it
```

The judgment card only has the two rules from the write-up: subtract stock-based compensation and reverse working-capital swings. The rest of the calls are yours to make.

## Setup

1. Install Python and openpyxl: `pip install openpyxl`
2. Copy the `fcf-model/` folder into `.claude/skills/` in whatever project you use Claude Code in.
3. Ask Claude to build an FCF model for a ticker.

Start with a company you have already done by hand, so you know what the right answer looks like.

## Try the builder without Claude

```bash
python fcf-model/scripts/build_fcf_workbook.py fcf-model/reference/spec-template.json "EXMP 2023-2025 FCF calculation.xlsx"
```

The example company is made up, but its numbers re-add to its reported operating cash flow, so every year prints PASS.

## Not investment advice

This builds a spreadsheet. It does not tell you what to buy.

## License

MIT
