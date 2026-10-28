# Paid tracking

Record which jobs have been paid for, one at a time or in bulk.

## Mark one job on its print room sheet

1. Find the job on its print room sheet.
2. Set its **Paid** cell to Yes.

Every new job starts with Paid set to No.

## Mark jobs from Reports

1. On **Reports**, use the filters to find the job. See
   [Filtering](Reports-Filtering).
2. Click any cell in its row.
3. Click **Toggle Paid**, under **Selected record:**.

A Yes becomes No and a No becomes Yes. To change several jobs, select cells
in up to 200 rows first. The row the cursor is on decides the new value, and
the workbook asks before changing them.

## Mark everything a filter shows

1. On **Reports**, set the filters so the list shows exactly the jobs to
   change, for example one student's unpaid jobs.
2. Under **Mark all as...**, click **Paid** or **Unpaid**.
3. Check the count shown and click **Yes**.

Only the jobs the list shows are changed. Your filters are left as they were.

## Result

The **Paid** and **Unpaid** totals on **Summary** and **Reports** update at
once.

## Things to know

- **Mark all as...** will not run until at least one filter is set, so a
  stray click cannot mark every job.
- The Paid cells on Reports cannot be typed in. Use **Toggle Paid**.
- **Unpaid** is the chargeable total of jobs not marked Paid. Jobs with
  nothing to charge do not add to it.

## Platform notes

> **Mac:** not yet checked.
>
> **Windows:** not yet checked.

## Related pages

- [Filtering](Reports-Filtering)
- [Summary](Reports-Summary)

---

*Draft for v0.11.4, without screenshots yet. Not yet verified on Windows.*
