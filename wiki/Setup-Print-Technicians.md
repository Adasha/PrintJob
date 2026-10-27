# Print technicians

List the people who print jobs, so each job can record who handled it.

## Before you start

- Set the **Site ID** on [Settings](Setup-Settings) first. Each technician's
  ID starts with it.

## Add a technician

1. Open the **Print Technicians** sheet.
2. Click **Add row**. A new row appears at the bottom of the table with its
   **TechID** filled in.
3. Type the technician's **Name**, as you want it to appear in the drop-down
   on a print room sheet.
4. Type their **Department**, if you want to record it.
5. Leave **Active** as Yes.

You can also type a name in the first empty row under the table. The row gets
its TechID as soon as it has a name.

## Result

The name appears in the **Technician** drop-down on every print room sheet.

## When someone leaves

Set their **Active** to No. Their name stops appearing in the drop-down for
new jobs, and the jobs they handled keep their name.

Avoid renaming a technician to reuse the row for someone else. Add a new row
instead, so old jobs still show who really handled them.

## Things to know

- Names must be unique. The workbook warns if you type a name that is already
  in the table.
- Nobody types a TechID. It is filled in for you, locked, and never reused,
  even after its row is removed.
- **Remove row** removes the row the cursor is in, after showing you what it
  holds. Jobs already recorded are not changed. It cannot be undone.

> **Warning:** **Clear table** empties the whole table. Recorded jobs are not
> changed, but it cannot be undone. Run **Backup workbook...** on Settings
> first. See [Deleting and clearing](Data-Deleting-and-Clearing).

## Platform notes

> **Mac:** not yet checked.
>
> **Windows:** not yet checked.

## Related pages

- [Recording jobs](Jobs-Recording): setting a room's default technician
- [Printers](Setup-Printers)

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
