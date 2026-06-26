---
name: melt-pour-polishing
description: >-
  Transform a raw US "Melt & Pour" steel export (an ITA "Quantity in Metric Tons"
  spreadsheet keyed on Country of Melt & Pour) into the user's polished 8-column
  monthly dataset (a new .xlsx). Use this skill whenever the user mentions US melt
  and pour data, a "Quantity in Metric Tons" export, Country of Melt & Pour, or
  asks to "polish", "clean", or reformat a monthly melt/pour file into their
  standard layout for a given month — even if they don't name the skill. This is
  recurring monthly work and the user picks a different month each run, so trigger
  it readily for any melt/pour data refresh. Sister skill to us-steel-import.
---

# US Melt & Pour — monthly polishing

Turn the raw Melt & Pour monitor export into the user's polished **8-column**
dataset, written to a **new** `.xlsx` (the raw file is never modified). The
bundled scripts do the transform and the verification deterministically; this
file explains the rules so you can run them, check them, and adapt when the
calendar or the source layout shifts.

## What the user asks for each run

The user names **one month** (e.g. "polish May", "do the June data"). Each run
targets that month. Pin it to a **year** too — the raw file usually carries the
same month across several years, so "May" alone is ambiguous. Ask which year if
it isn't stated; default to the most recently completed month's year when the
context makes that obvious (this is last-week-of-the-month work, like the sister
`us-steel-import` skill).

The polished file is a **filtered projection** of the raw file: the selected
month's rows, in their original order, with the 8 columns the user wants. No rows
are added, dropped, or reordered — so the output row count must exactly equal the
number of raw rows for that month. Always count and reconcile (the verifier does
this automatically).

## Quick start

```powershell
& "<skill-dir>\scripts\Build-MeltPourPolished.ps1" `
    -RawPath "C:\path\to\Quantity in Metric Tons.xlsx" `
    -Month May -Year 2026
```

`-Month` accepts a name (`May`) or number (`5`). This filters to the chosen
month/year, applies the country renames, writes the 8-column layout, and saves
`Melt and Pour - Polished - May 2026.xlsx` next to the raw file. Use `-OutPath`
to choose the destination, or `-AllYears` to take every year's instance of the
month. It prints the target month, the row count, and the output path. **0 rows**
is flagged as a warning — check the month/year.

The scripts use Excel COM (standard on the user's Windows machine) and **must be
saved UTF-8 with BOM** — the headers `조강국` and `원산지` are Korean and Windows
PowerShell 5.1 reads a BOM-less `.ps1` as ANSI, which corrupts them and breaks
parsing. If you ever regenerate a script, re-save it with a BOM.

## Output format — exactly 8 columns (with a header row)

| # | Header | Contents |
|---|--------|----------|
| 1 | `Country of Melt & Pour` | Raw **`Country of Melt & Pour`**, with renames (below). |
| 2 | `조강국` | Header only — **every data cell blank** (a slot the user fills later). |
| 3 | `Product Category` | Raw **`Product Category`** — one of the 6 groups: Flat (Carbon and Alloy), Long (Carbon and Alloy), Pipe and Tube (Carbon and Alloy), Semi-Finished (Carbon and Alloy), Stainless, Other. |
| 4 | `Country of origin (Grouped)` | Raw **`Country of Origin (Grouped)`**, with renames (below). |
| 5 | `원산지` | Header only — **every data cell blank**. |
| 6 | `ExpectedDateOfImportation - Year` | Raw year (e.g. `2026`), as an integer. |
| 7 | `ExpectedDateOfImportation - Month` | Raw month **name** exactly as stored (e.g. `May`). |
| 8 | `Selected Measure` | Raw `Selected Measure` tonnage, full precision, displayed `#,##0.00`. |

The headers are reproduced verbatim, including the user's casing
(`Country of origin (Grouped)` — lowercase "origin" — differs from the raw
`Country of Origin (Grouped)`).

## Rules the scripts apply

- **Row selection** — keep a raw row only if its month equals the target month
  and (unless `-AllYears`) its year equals the target year. Everything else is
  dropped.
- **Country renames — applied to BOTH country columns (1 and 4)**: `South Korea →
  Korea`; `Türkiye → Turkey` (matched on `rkiye` so the special character can't
  cause a miss). All other names pass through unchanged. The raw often already
  uses `Korea`, which is fine — the rename is a no-op there.
- **Order** — the raw file's row order is preserved, so the output is a clean 1:1
  projection. (No sorting; this keeps the row-for-row reconciliation transparent.)
- **Blanks** — columns 2 (`조강국`) and 5 (`원산지`) get a header and otherwise
  genuinely empty cells.

## Verify before handing off — each column is its own reviewer

This feeds downstream analysis, so always reconcile the output against the raw
file. The bundled verifier runs **multi-persona verification: one independent
reviewer per output column**, plus structural reviewers. It re-derives the
expected month subset straight from the raw file (it does **not** trust the build
script's result), then checks each column position-by-position:

```powershell
& "<skill-dir>\scripts\Verify-MeltPourPolished.ps1" `
    -RawPath "C:\path\to\Quantity in Metric Tons.xlsx" `
    -OutPath "C:\path\to\Melt and Pour - Polished - May 2026.xlsx" `
    -Month May -Year 2026
```

It prints PASS/FAIL for 11 reviewers and an overall verdict. Expect **all-PASS**:

1. **Headers** — all 8 exact. **Row count** — raw month rows == polished data rows
   (the "same number of rows" guarantee; any mismatch means a row was dropped,
   doubled, or mis-filtered). **Col1/Col4** — match the renamed raw country.
   **Col2/Col5** — every cell blank. **Col3** — matches raw Product Category.
   **Col6** — year matches. **Col7** — month matches. **Col8** — tonnage matches
   exactly. **Rename integrity** — no `South Korea`/`Türkiye` leaks into col1, and
   the Korea/Turkey counts cover the raw originals.

Any FAIL names the offending column and shows sample mismatches with the row
number and both values.

## Edge cases & gotchas

- **Which year** — the raw file spans multiple years of the same month. Pin the
  year (`-Year`) or pass `-AllYears` deliberately. Don't guess silently.
- **Header not on row 1** — the raw export has a filter-banner row (and a blank
  row) above the real header. The scripts find the header row and every column
  **by name**, so they tolerate the banner and minor column reordering. If a
  source header is renamed the script errors clearly naming the missing column.
- **PowerShell array quirks** (why the scripts look the way they do): Windows
  PowerShell mis-handles 2-D array indexing when the index is an arithmetic
  expression (`$a[$i+1,0]` evaluates the comma as an array → `op_Addition`
  error) and can mis-read a COM 2-D array via an inline `([string]$a[$r,$c]).Trim()`
  inside a nested function. The scripts therefore use a plain index variable
  (`$ri = $i + 1`) and read each cell into a temp before comparing. Keep that
  pattern if you edit them.
- **Encoding** — re-save any edited script as **UTF-8 with BOM** (Korean headers).

See `references/data-format.md` for the full anatomy of the raw export.
