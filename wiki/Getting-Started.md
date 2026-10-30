# Getting started

Open the workbook for the first time, turn its macros on and find your way
round.

## Before you start

- **Excel for Microsoft 365**, on Windows or a Mac. Excel 2019 and earlier
  cannot run the workbook, because its totals use formulas that only
  Microsoft 365 has.
- **The workbook file**, named like `PrintJob-v0.11.5.xlsm`. Ask whoever looks
  after the workbook for your bureau. Released versions are also in the
  [`dist` folder](https://github.com/Adasha/PrintJob/tree/master/dist) of the
  project on GitHub.
- **A place to keep it.** A folder on your computer or in OneDrive both work. Using a OneDrive location is recommended as Excel will autosave as print jobs are added.
  Keep one copy per bureau; each workbook is one site.

## Open the workbook

1. Save the file to the folder where you will keep it, then open it from
   there.
2. Turn on macros when Excel asks (see **Platform notes** below). Without
   macros the buttons do nothing.
3. The workbook opens on the **Summary** sheet.
4. To check the version, open the **Settings** sheet and click **About**.

## What you will see

The tabs along the bottom, from left to right:

| Tab | What it is for |
| --- | --- |
| **Summary** | Totals by print room, printer and paper; the job planner |
| **Reports** | Every job from every room, with filters and paid tracking |
| **Example Print Room** | A sample print room with 12 jobs |
| **Print Technicians**, **Printers**, **Papers**, **Consumables**, **Departments** | The lists that jobs are priced from |
| **Settings** | Site details, currency, and the print room and data commands |

The sample jobs show how a print room works. Job `UNI-EXAMPLE-00012` is wrong
on purpose, so you can see how the **Status** column reports a problem.

## Next steps

Set the workbook up in this order, so each list is ready before the next one
needs it:

1. [Settings](Setup-Settings): site name, currency and rounding.
2. [Print technicians](Setup-Print-Technicians), [Printers](Setup-Printers),
   [Papers](Setup-Papers) and, for multi-pass printers,
   [Consumables](Setup-Consumables).
3. [Departments](Setup-Departments), if some departments print for free.
4. [Print rooms](Setup-Print-Rooms): add your rooms, then remove the
   **Example Print Room**.

## Things to know

- Always save as an **Excel Macro-Enabled Workbook (.xlsm)**. Saving as
  `.xlsx` removes the macros.
- The sheets are protected without a password. This stops cells being changed
  by accident. See [Protection and hidden sheets](Admin-Protection-and-Hidden-Sheets).
- Renaming the workbook file is fine. The buttons repair themselves the next
  time it opens.

## Platform notes

> **Windows:** a workbook downloaded or received by email may open with a red
> bar saying Microsoft has blocked macros. Close it, right-click the file,
> choose **Properties**, tick **Unblock** and click **OK**, then open it
> again. Click **Enable Content** on the yellow bar.
>
> **Mac:** Excel asks whether to enable macros as the workbook opens. Click
> **Enable Macros**. If you chose **Disable Macros** by mistake, close the
> workbook and open it again.

## Related pages

- [How it works](How-It-Works)
- [Upgrading to a new version](Admin-Upgrading)
- [Troubleshooting](Troubleshooting)

---

*Draft for v0.11.5, without screenshots yet. Not yet verified on Windows.*
