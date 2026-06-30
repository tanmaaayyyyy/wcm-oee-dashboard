# Microsoft Forms Template: "Daily Production Entry Form"

This form is filled in by a floor operator at the end of every shift. When
connected to Power Automate, each submission automatically appends a row to
`production_data.xlsx` in OneDrive, removing manual data entry and the
transcription errors that come with it.

> Build this at https://forms.office.com. Click **New Form**, name it **Daily
> Production Entry Form**, then add the fields below in order.

---

## Form settings

- **Title:** Daily Production Entry Form
- **Description:** "Complete at the end of every shift. One submission per line per shift."
- **Setting:** Turn ON "Record name" only if you want operator accountability;
  otherwise leave anonymous. Turn ON "One response per person" = OFF (operators
  submit multiple lines per shift).

---

## Fields

| # | Field Name | Field Type | Required | Options / Notes |
|---|------------|-----------|----------|-----------------|
| 1 | Production Date | Date | Yes | Defaults to today; operator confirms. |
| 2 | Shift | Choice (dropdown) | Yes | Morning, Afternoon, Night |
| 3 | Line | Choice (dropdown) | Yes | Line A, Line B, Line C |
| 4 | Planned Production Time (hrs) | Number | Yes | Standard is 8. Restrict 0–8. |
| 5 | Downtime (hrs) | Number | Yes | Total stop time this shift. Restrict 0–8, allow decimals. |
| 6 | Downtime Category | Choice (dropdown) | Yes | Mechanical Failure, Changeover, Material Shortage, Planned Maintenance, Quality Hold |
| 7 | Units Planned | Number | Yes | Target units for the shift. Whole number. |
| 8 | Units Produced | Number | Yes | Actual good + reject units packed. Whole number. |
| 9 | Units Rejected | Number | Yes | Units failing QC this shift. Whole number, <= Units Produced. |
| 10 | Rejection Category | Choice (dropdown) | Yes | Sealing Defect, Label Misalignment, Weight Variance, Contamination, Dimension Error |
| 11 | Operator Notes | Long answer (text) | No | Optional context: root cause, corrective action taken, observations. |

---

## Field type configuration details

### Choice fields (dropdowns)
For fields 2, 3, 6 and 10, click **Add option** for each value listed above and
enable the **Dropdown** display style (click the "..." on the question →
**Dropdown**). This keeps responses clean and consistent so the downstream
Pareto analysis groups correctly.

### Number fields
For fields 4, 5, 7, 8, 9, click the "..." → **Restrictions** → **Number** to
enforce numeric entry. Add min/max bounds as noted so an operator cannot type
a downtime of 50 hours by mistake.

### Date field
Field 1 uses the built-in **Date** picker. Power Automate maps this directly to
the `Date` column in Excel.

---

## Why these exact fields

The form columns map 1:1 to the raw input columns in `production_data.xlsx`. The
four OEE component columns (Availability, Performance, Quality, OEE) are **not**
on the form — they are calculated columns. You can compute them either:

- in the Excel table with formulas, or
- inside Power BI as DAX measures (recommended).

Keeping calculated fields out of the form means operators only enter what they
actually observe on the floor, which is the WCM principle of capturing loss data
at the source.
