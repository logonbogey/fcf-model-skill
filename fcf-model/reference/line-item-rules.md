# Line-Item Treatment Card

This is the judgment card: one row per type of line you find on a cash flow statement, the tag it gets in the spec, and why. **It is deliberately almost empty.** The full card I use belongs to the analyst whose judgment it encodes and is not mine to publish, so this version carries only the two rules explained in the write-up. Fill in the rest with your own calls.

Tags: `"fcf": true` subtracts a line from normalized free cash flow. `"is_sbc": true` marks stock compensation so it gets its own row. `"one_time": true` is an informational note and changes no math. Working-capital lines are tagged automatically.

| Line item | Spec flags | Treatment | Why |
|---|---|---|---|
| Net income (or the statement's starting line) | (none) | starting point | top of the build |
| Stock-based compensation | `is_sbc`, `fcf` | **subtract** | Real cost. The owners paid for it through dilution even though GAAP adds it back as non-cash. |
| Working-capital changes | (auto) | reverse all out | A good year for collections or for stretching payables flatters one year and does not repeat. |
| Capital expenditures | `capex` array | subtract | *Your call on scope.* |
| Depreciation & amortization | | | *Your call.* |
| Deferred income taxes | | | *Your call.* |
| Impairments, restructuring | | | *Your call.* |
| Gains / losses on sale of businesses | | | *Your call.* |
| Leases (US GAAP operating, IFRS 16) | | | *Your call.* |
| Pension / postretirement | | | *Your call.* |
| Deferred revenue | | | *Your call.* |
| "Other" | | | No fixed rule. Open the footnotes, see what is inside the number, and hand the call to a person. |

## How to grow this card

Every time you make a call on a line that is not here, add a row with the reasoning, not just the answer. "Reverse this line" helps on one company. "Reverse this line because it is a financing cost dressed up as an operating one" transfers to the next one.
