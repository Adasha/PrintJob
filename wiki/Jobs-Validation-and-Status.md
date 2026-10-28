# Validation and status

Find out whether a job is complete and correct, and what to fix if not.

## How jobs are checked

The workbook checks jobs in three ways:

- **Drop-downs** offer only active technicians, the room's printers, and
  papers that fit the chosen printer. If you change the printer and the paper
  no longer fits, the paper is cleared.
- **The Status column** on each job shows **OK** or what is wrong. It updates
  as you type.
- **The check buttons** go through a whole sheet or the whole workbook.

Problems never stop you saving. They are there to be fixed before you rely
on the totals.

## Read a job's status

1. Look at the job's **Status** cell. Orange means a problem; yellow means a
   notice.
2. Hover over the cell to see every problem in its note.
3. Fix each one. Status changes to **OK** when the job is complete.

| Status says | What to do |
| --- | --- |
| Date and time required | Click **Now**, or double-click the Date/Time cell |
| Student name or number required | Type at least one of them |
| Technician, Printer or Paper stock required | Choose one from the drop-down |
| Quantity required, or must be greater than zero | Type the Qty |
| Sheet quantity must be a whole number | Sheets are counted whole; type a whole number |
| Print width does not apply to sheet stock | Clear the Print Width mm |
| Print width exceeds stock width | Type a width no wider than the roll |
| Print width required for student-supplied roll stock | Type the Print Width mm |
| Sheet size required for student-supplied sheet stock | Choose the nearest Sheet size |
| At least one colour pass is required | Add a pass. See [Colour passes](Jobs-Colour-Passes) |

## Check a whole sheet or the workbook

- **Check this sheet**, on the right of a print room sheet, checks every job
  on it and tells you what it found. The result is shown beside it under
  **Sheet status**.
- **Check workbook**, on Summary or Settings, checks every print room. It
  also lists rooms with jobs that have not been exported.

## Things to know

- If a student's name and number do not match earlier jobs, in any room, the
  workbook warns you as you type. It never blocks the job. If it is the same
  person under a new name, carry on.
- A notice, such as a pass colour this workbook does not list, is not
  counted as a problem.
- Job `UNI-EXAMPLE-00012` in the sample room has problems on purpose, so you
  can see how this looks.

## Platform notes

> **Mac:** not yet checked.
>
> **Windows:** not yet checked.

## Related pages

- [Recording jobs](Jobs-Recording)
- [Troubleshooting](Troubleshooting)

---

*Draft for v0.11.4, without screenshots yet. Not yet verified on Windows.*
