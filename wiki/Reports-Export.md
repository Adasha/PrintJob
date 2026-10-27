# Exporting a report

Save the jobs Reports is showing as a printable Excel file, for a student, a
department or a manager.

## Before you start

- Set the filters so Reports shows exactly the jobs the report should
  cover. See [Filtering](Reports-Filtering).
- Optionally set a **Report heading** on [Settings](Setup-Settings).

## Export the report

1. On **Reports**, click **Export report...**.
2. The workbook asks whether to include student names:
   - **Yes** includes them.
   - **No** replaces each name with dashes. Student numbers are kept.
   - **Cancel** stops.
3. A message says where the file was saved.

## Result

An `.xlsx` file named like `PrintCosts-Report-UNI-20261027-1430.xlsx`. It
holds the values, not formulas, so it opens the same on any computer.

- At the top: the heading, the site, the dates covered, when it was made and
  how many jobs it holds.
- If every job shares one value, such as one student or one printer, that
  value is stated once at the top instead of in every row.
- It is set up to print on A4 landscape, one page wide, with the column
  headings repeated on each page.

## Things to know

- The file is saved in the **Export folder** from Settings, or beside the
  workbook if that is blank. See [Exporting a print room](Data-Export-Room).
- An exported report contains personal data. Store and send it as your
  university's data protection rules require.
- This report is for reading. To move jobs between workbooks, use
  [Exporting a print room](Data-Export-Room) instead.

## Platform notes

> **Mac:** Excel may ask for access to the folder and then the file. Allow
> both.
>
> If the workbook is in OneDrive and **Export folder** is blank, the
> file may be saved in your OneDrive folder or in Documents. The message says
> where.
>
> **Windows:** not yet checked.

## Related pages

- [Filtering](Reports-Filtering)
- [Settings](Setup-Settings)

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
