# FAQ

## Does changing a price change jobs already recorded?

No. Each job keeps the prices it was recorded with. The one exception is
**Currency decimal places** on Settings, which recalculates every job. See
[How it works](How-It-Works).

## How do I charge a department instead of a student?

Type the department's name in **Student Name**. To track departments
separately, or to record their jobs without charging them, list them on
[Departments](Setup-Departments).

## Can I undo a command?

No. Commands such as **Delete**, **Clear All** and **Restore workbook...**
cannot be undone, which is why each one asks first. Back up before anything
large. See [Backup and restore](Data-Backup-and-Restore).

## Why can't I sort or filter a print room sheet?

So that colour passes always stay under their job. Use
[Reports](Reports-Filtering) to sort and filter jobs from every room.

## Why is a Job ID number missing?

That job was deleted. Numbers are never reused, so every Job ID stays unique.

## Can I combine the jobs from several sites?

By hand, yes: export the rooms from each site's workbook with **Export All
Locations...**, then import each file into one copy of the workbook. See
[Importing jobs](Data-Import). There is no automatic way yet.

## Can I export a report as a PDF?

Not directly. Export the report, open the file, and print it to PDF. See
[Exporting a report](Reports-Export).

## Is student data protected?

Sheet protection stops accidents, not people. Exports, reports and backups
are ordinary files that contain names and numbers, so store and share them
as your university's data protection rules require. Exports and reports can
leave names out.

## What is the difference between the workbook version and the schema?

The workbook version, such as 0.11.3, changes with every release. The schema
version, such as 1.4, changes only when the shape of the job records changes.
Both are on **Settings** and in **About**.

## Where does the university subsidy apply?

Only to the reports. Reports shows the chargeable amount after the subsidy
beside the full figures. Jobs themselves are not changed. See
[Settings](Setup-Settings).

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
