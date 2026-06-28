# Power Query M – Latest Forecast Month per Project (>999 threshold)
 
## Purpose
 
Source data is a monthly **Altus Power BI "Insights"** export (Project × Month × Forecast).
The business question this query answers:
 
> *For each project, what was the most recent month where the financial forecast exceeded $999 — and what was that forecast value?*
 
That output is designed to be compared against each project's **completion date** (in a separate register), to catch projects that are still forecasting meaningful spend in months *after* they were supposed to be finished.
 
Everything before the date-fix step is routine cleanup. The genuinely tricky part — and the reason this is worth keeping — is reliably turning a text month label into a real date, then picking the single "latest qualifying month" per project using a group-and-max pattern.
 
## Input shape
 
An Excel extract of the Insights report, with at least these columns after the initial load:
 
| Project | Month | Forecast |
|---|---|---|
| 1234 Park Upgrade | Jan-24 | 1500 |
| 1234 Park Upgrade | Feb-24 | 0 |
| ... | ... | ... |
| Total | ... | ... |
 
- `Project` combines an ID and a name (e.g. `"1234 Park Upgrade"`)
- `Month` is text in `Mon-YY` format
- The report includes a `Total` summary row that needs to be excluded
 
## Step-by-step walkthrough
 
**1. Source**
```m
Source = Excel.Workbook(File.Contents("C:\...\Source\Insights.xlsx"), null, true)
```
Loads the raw Excel export. Followed by the standard load chain: Navigation → Promoted Headers → Remove Other Columns → Changed Type.
 
**2. Filtered Totals** — drop the summary row early so it can't interfere with date parsing or sorting
```m
#"Filtered Totals" = Table.SelectRows(#"Changed Type", each ([Month] <> "Total"))
```
 
**3. Add Fixed Date** — the core logic. `Month` arrives as text like `"Jan-24"`. Power Query's automatic date parsing isn't reliable here (ambiguous 2-digit year handling), so the year and month are parsed manually and the year is explicitly forced into the 2000s:
```m
#"Add Fixed Date" = Table.AddColumn(#"Filtered Totals", "Fixed Date", each let
    parts = Text.Split([Month], "-"),
    mon = parts{0},
    yr = Number.From(parts{1}),
 
    // Convert month name to month number
    monthNum = Date.Month(Date.FromText("01-" & mon & "-2020")),
 
    // Force ALL 2-digit years into 20xx
    fixedYear = 2000 + yr
in
    #date(fixedYear, monthNum, 1))
```
- Splits `"Jan-24"` into `"Jan"` and `24`
- Builds a throwaway valid date (`"01-Jan-2020"`) purely to extract the month *number* from the month name
- Reconstructs the real date as `2024-01-01`, sidestepping any century ambiguity
 
**4. Changed Type1** — lock in proper types now that real dates exist
```m
#"Changed Type1" = Table.TransformColumnTypes(#"Add Fixed Date",
    {{"Fixed Date", type date}, {"Month", type date}})
```
 
**5. Forecast > 999** — the business filter: only months with a meaningful forecast count
```m
#"Forecast > 999" = Table.SelectRows(#"Changed Type1", each ([Forecast] > 999))
```
 
**6. Sort** — within each project, most recent month first
```m
#"Sort" = Table.Sort(#"Forecast > 999",
    {{"Project", Order.Ascending}, {"Month", Order.Descending}, {"Fixed Date", Order.Descending}})
```
 
**7. Grouped by Project** — bundle each project's rows into a sub-table so the latest one can be picked
```m
#"Grouped by Project" = Table.Group(#"Sort", {"Project"},
    {{"LatestRecord", each _, type table [Project=nullable text, Month=nullable date, Forecast=number, Fixed Date=nullable date]}})
```
 
**8. Add Table Max** — pull out the row with the maximum `Fixed Date` per project (i.e. the latest qualifying month)
```m
#"Add Table Max" = Table.AddColumn(#"Grouped by Project", "Custom",
    each Table.Max([LatestRecord], "Fixed Date"))
```
 
**9. Expanded Custom** — surface that row's `Forecast` / `Fixed Date` back into normal columns
```m
#"Expanded Custom" = Table.ExpandRecordColumn(#"Add Table Max", "Custom",
    {"Forecast", "Fixed Date"}, {"Forecast", "Fixed Date"})
```
 
**10. Removed Columns** — the sub-table is no longer needed
```m
#"Removed Columns" = Table.RemoveColumns(#"Expanded Custom", {"LatestRecord"})
```
 
**11. Changed Type2** — final typing for output
```m
#"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns",
    {{"Fixed Date", type date}, {"Forecast", Int64.Type}})
```
 
**12. Renamed Columns** — give the result business-friendly names
```m
#"Renamed Columns" = Table.RenameColumns(#"Changed Type2",
    {{"Fixed Date", "Latest Month >999"}, {"Forecast", "Latest Forecast"}})
```
 
**13. Inserted Text Before Delimiter** — split the combined `Project` field to isolate the ID
```m
#"Inserted Text Before Delimiter" = Table.AddColumn(#"Renamed Columns", "Text Before Delimiter",
    each Text.BeforeDelimiter([Project], " "), type text)
```
 
**14. Renamed Columns1** — clean naming
```m
#"Renamed Columns1" = Table.RenameColumns(#"Inserted Text Before Delimiter",
    {{"Text Before Delimiter", "Project ID"}, {"Project", "Project Name"}})
```
 
**15. Reordered Columns**
```m
#"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns1",
    {"Project ID", "Project Name", "Latest Forecast", "Latest Month >999"})
```
 
**16. Final filter** — drop any leftover `Total` row by ID
```m
#"Filtered Rows" = Table.SelectRows(#"Reordered Columns", each ([Project ID] <> "Total"))
```
 
## Output
 
One row per project:
 
| Project ID | Project Name | Latest Forecast | Latest Month >999 |
|---|---|---|---|
| 1234 | Park Upgrade | 1500 | 2024-01-01 |
 
This is the table that gets merged against the project completion-date register to flag projects forecasting spend past completion.
 
## Full consolidated M code
 
```m
let
    Source = Excel.Workbook(File.Contents("C:\Users\hadismirzajani\...\Source\Insights.xlsx"), null, true),
    #"Navigation" = Source{[Item="Insights", Kind="Sheet"]}[Data], // adjust sheet reference as needed
    #"Promoted Headers" = Table.PromoteHeaders(#"Navigation", [PromoteAllScalars=true]),
    #"Removed Other Columns" = Table.SelectColumns(#"Promoted Headers", {"Project", "Month", "Forecast"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns", {{"Project", type text}, {"Month", type text}, {"Forecast", type number}}),
    #"Filtered Totals" = Table.SelectRows(#"Changed Type", each ([Month] <> "Total")),
    #"Add Fixed Date" = Table.AddColumn(#"Filtered Totals", "Fixed Date", each let
            parts = Text.Split([Month], "-"),
            mon = parts{0},
            yr = Number.From(parts{1}),
            // Convert month name to month number
            monthNum = Date.Month(Date.FromText("01-" & mon & "-2020")),
            // Force ALL 2-digit years into 20xx
            fixedYear = 2000 + yr
        in
            #date(fixedYear, monthNum, 1)),
    #"Changed Type1" = Table.TransformColumnTypes(#"Add Fixed Date", {{"Fixed Date", type date}, {"Month", type date}}),
    #"Forecast > 999" = Table.SelectRows(#"Changed Type1", each ([Forecast] > 999)),
    #"Sort" = Table.Sort(#"Forecast > 999", {{"Project", Order.Ascending}, {"Month", Order.Descending}, {"Fixed Date", Order.Descending}}),
    #"Grouped by Project" = Table.Group(#"Sort", {"Project"}, {{"LatestRecord", each _, type table [Project=nullable text, Month=nullable date, Forecast=number, Fixed Date=nullable date]}}),
    #"Add Table Max" = Table.AddColumn(#"Grouped by Project", "Custom", each Table.Max([LatestRecord], "Fixed Date")),
    #"Expanded Custom" = Table.ExpandRecordColumn(#"Add Table Max", "Custom", {"Forecast", "Fixed Date"}, {"Forecast", "Fixed Date"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Custom", {"LatestRecord"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Columns", {{"Fixed Date", type date}, {"Forecast", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type2", {{"Fixed Date", "Latest Month >999"}, {"Forecast", "Latest Forecast"}}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Renamed Columns", "Text Before Delimiter", each Text.BeforeDelimiter([Project], " "), type text),
    #"Renamed Columns1" = Table.RenameColumns(#"Inserted Text Before Delimiter", {{"Text Before Delimiter", "Project ID"}, {"Project", "Project Name"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns1", {"Project ID", "Project Name", "Latest Forecast", "Latest Month >999"}),
    #"Filtered Rows" = Table.SelectRows(#"Reordered Columns", each ([Project ID] <> "Total"))
in
    #"Filtered Rows"
```
 
> **Note:** the `Navigation` / `Promoted Headers` / `Removed Other Columns` lines above are reconstructed to make this paste-ready into the Advanced Editor — adjust the sheet/table name and column selection to match your actual source. Everything from `Filtered Totals` onward is exactly the logic from the original query.
 
## Reusing this for a new month/source
 
To reuse this pattern on a different report:
 
1. Update the `Source` file path / sheet reference.
2. Confirm the `Month` column format still matches `Mon-YY` text (e.g. `"Jan-24"`). If the format differs, only the **Add Fixed Date** step needs adjusting.
3. Confirm the forecast threshold (`> 999`) is still the right cutoff for "meaningful" — this was a judgement call to filter out near-zero noise.
4. Confirm `Project` still follows the `"<ID> <Name>"` pattern with a single space before the name — if the ID has variable-length formats, `Text.BeforeDelimiter` may need a stricter split (e.g. on the first space only, which it already does, or a regex-like approach if IDs ever contain spaces).
 
## Why this approach (group + max) instead of just sorting and taking the top row
 
`Table.Group` + `Table.Max` on the sub-table is more robust than "sort then keep first row per group" because it doesn't depend on the sort being perfectly stable across ties, and it makes the "pick the latest" intent explicit and self-documenting in the step name — useful both for handing this off and for future-you six months from now.
 
