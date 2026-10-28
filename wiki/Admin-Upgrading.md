# Upgrading to a new version

Move your setup and jobs from an old workbook into a new version.

## Before you start

- Get the new workbook. See [Getting started](Getting-Started).
- Check which version you have: **Settings**, **About**.

## Upgrade

1. Save the new workbook where you will keep it, and open it with macros
   turned on.
2. On **Settings**, click **Restore workbook...**.
3. Choose your old workbook file.
4. On a Mac, if Excel asks about the old workbook's macros, choose **Disable
   Macros**. The workbook's notice says the same: they are not needed to
   read the old workbook.
5. Read the summary of what will come across and click **Yes**.

## Result

The new workbook has:

- your setup lists and settings
- each of your print rooms, with its defaults, printers and roll length unit
- every job, with the prices it was costed at

The sample room is removed once your own rooms are in. Check the totals on
**Summary** against the old workbook, then use the new one from now on.

## Things to know

- Nothing needs backing up first, and the old workbook is not changed. Keep
  it as an archive.
- You can also restore from a backup made in the old workbook instead. See
  [Backup and restore](Data-Backup-and-Restore).
- The version and data schema are shown on **Settings** and in **About**.
  The workbook version changes with each release; the schema only when the
  shape of the job records changes.
- What changed in each version is in [Release notes](Release-Notes).

## Platform notes

> **Mac:** use 0.11.1 or later. In earlier versions, Export and Backup on a
> Mac could delete the workbook.
>
> **Windows:** not yet checked.

## Related pages

- [Backup and restore](Data-Backup-and-Restore)
- [Release notes](Release-Notes)

---

*Draft for v0.11.4, without screenshots yet. Not yet verified on Windows.*
