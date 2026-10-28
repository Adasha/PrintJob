# Shot list

Every screenshot the wiki uses, and how to make it again. Images go in this
folder. This file is for whoever takes the screenshots and is not published
to the wiki.

## Names

`<Page>-<n>-mac.png` and `<Page>-<n>-win.png`, where `<Page>` is the wiki
page and `<n>` counts from 1 down the page. A shot whose two versions look
the same is saved once as `<Page>-<n>.png`. A shot marked Mac or Windows
below exists only on that platform.

## Image rules

- At most 1200 px wide; downscale Retina captures.
- PNG, optimised to about 150 KB or less. The wiki repository has no Git
  LFS.
- Alt text on every image, saying what it shows.
- Callouts, where a step needs one: numbered circles in one colour, matching
  the numbered steps on the page.

## The demo workbook

Take every shot from a fresh build, never from a workbook with real data, in
a copy (on a Mac, inside Excel's container). Set up these on top of the
shipped sample, in this order:

1. Settings: Site name `University Print Bureau`, Org `Example University`,
   currency £.
2. Printers: add a multi-pass printer `Riso SF9450`, Colour mode multi-pass,
   Max sheet size A3, Template cost 0.50, Cost per m2 0.
3. Consumables: add the colours Black, Red and Blue for it.
4. Add print room `Photography`, code `PHOTO`, with HP DesignJet Z9+ and
   Riso SF9450 ticked in **Select printers...**.
5. Photography: record three jobs. One on roll paper with a Print Width mm,
   one sheet job for `School of Art`, and one Riso job with two passes (Black,
   Red).
6. Mark the School of Art job Paid.

Leave job `UNI-EXAMPLE-00012` as shipped: it is the job with problems.

**Window:** Excel at 1440 x 900 points, zoom 100%, ribbon collapsed, the
cell cursor on A1 unless stated. Light appearance on the Mac.

## Capture

- **Range**: a cell range exported as a picture by script: `docs-screens.ps1`
  on Windows, a Python twin under `mac/` on a Mac. Both copy the range with
  `CopyPicture` (screen appearance, picture format), paste it into a chart
  twice the range's size with the picture stretched to fill it, and export
  the chart as PNG. The picture is vector, so the 2x chart renders sharp, at
  about 2 px per point. Probed on Excel for Mac 16.114 on 2026-10-28: every
  step worked, inside Excel's container, with no prompts. Exact ranges are
  found with `probe.ps1 -Sheet <name>` or its Mac equivalent once the demo
  workbook is built, and written into the script, not here.
- **Window**: the whole Excel window, by hand (Shift+Cmd+4 then Space on a
  Mac, Alt+PrtScn on Windows).
- **Dialog**: the dialog alone, by hand, the same way.

## Shots

| File | Page section | Shows | State | Capture | Platform |
| --- | --- | --- | --- | --- | --- |
| Getting-Started-1 | What you will see | The window with Summary open and every tab visible | Summary selected | Window | Both |
| Getting-Started-2 | Open the workbook | The macro prompt: Enable Macros on a Mac, the yellow Security Warning bar on Windows | Opening the demo workbook with macro security at its default | Dialog | Both, different |
| Setup-Settings-1 | Set up the global settings | The top of the settings table: site, org, department, currency, rounding, export folder | Demo settings | Range | Both |
| Setup-Settings-2 | The buttons | The Settings button rows: Print rooms and Data | | Range | Both |
| Setup-Print-Technicians-1 | Add a technician | Add row, Remove row, Clear table and the table with its TechIDs | | Range | Both |
| Setup-Printers-1 | Add a printer | The printers table: a roll printer, a sheet printer and the Riso | | Range | Both |
| Setup-Papers-1 | Add a paper | The papers table, showing Measure, Size mode, Std. size, Width and Height | | Range | Both |
| Setup-Consumables-1 | Add a colour | The colours table with Black, Red and Blue | | Range | Both |
| Setup-Departments-1 | Add a department | The departments table with School of Art (Free = Yes) and its aliases | | Range | Both |
| Setup-Lookup-Lists-1 | The lists | The lookup lists on Settings with their + and - buttons | | Range | Both |
| Setup-Print-Rooms-1 | Add a print room | The prompt asking for the new room's code | **Add print room...**, name `Photography` typed, at the code prompt | Dialog | Both |
| Setup-Print-Rooms-2 | Add a print room | The **Select printers...** picker with OK and Cancel | Photography, picker open with two printers ticked | Range | Both |
| Jobs-Recording-1 | Record a job | The top of a room sheet: settings block, the Print jobs panel (New, Repeat, Delete), and the first jobs | Example Print Room, scrolled to the top | Range | Both |
| Jobs-Recording-2 | Delete a job | The Delete confirmation showing the job's details | **Delete** on a Photography job, before answering | Dialog | Both |
| Jobs-Validation-and-Status-1 | Read a job's status | `UNI-EXAMPLE-00012`'s Status cell with its note showing | Note shown by script (the comment made visible), not by hovering | Range | Both |
| Jobs-Validation-and-Status-2 | Check a whole sheet or the workbook | The side panel: **Check this sheet** and **Sheet status** after a check | Example Print Room after **Check this sheet** | Range | Both |
| Jobs-Views-and-Units-1 | Show fewer columns | A room sheet at the Minimal view | Photography, view set to Minimal | Range | Both |
| Jobs-Colour-Passes-1 | Record a multi-pass job | The Riso job expanded with its two pass rows, and the Colour passes panel | Photography, passes shown | Range | Both |
| Jobs-Paid-Tracking-1 | Mark jobs from Reports | The Reports button strip: **Toggle Paid** under Selected record, and **Mark all as...** | Reports, filtered to Photography | Range | Both |
| Reports-Summary-1 | Read the totals | The Totals row and the first rows of the table | | Range | Both |
| Reports-Summary-2 | Price a job before printing | The planner filled in, with the Estimated cost and **Add to print room** | HP DesignJet Z9+, Matte 180gsm 1067mm roll, roll length 2 | Range | Both |
| Reports-Filtering-1 | The filters | The filter group expanded, one filter set, and the results below | Reports, Location = Photography | Range | Both |
| Data-Export-Room-1 | Export one room | The include-names question (Yes, No, Cancel) | **Export...** on Photography, before answering | Dialog | Both |
| Data-Export-Room-2 | Where the files go | Excel asking for access to the export folder | A copy of the demo workbook in an ordinary folder (not Excel's container), first export | Dialog | Mac |
| Data-Import-1 | Import into the room you are on | The counts to confirm before an import | Import of Data-Export-Room-1's file back into Photography | Dialog | Both |
| Data-Backup-and-Restore-1 | Restore from a backup | The summary of what a restore will change | **Restore workbook...** with a backup of the demo workbook, into a fresh build | Dialog | Both |
| Data-Deleting-and-Clearing-1 | Clear a room | The Clear All warning | **Clear All** on Photography, before answering | Dialog | Both |
| Admin-Protection-and-Hidden-Sheets-1 | Hide the setup sheets | The tab bar with the setup sheets hidden, and the **Show settings sheets** button | Summary after **Hide settings sheets** | Window | Both |

Pages with no shot: How-It-Works (diagram), Reports-Export, Admin-Upgrading,
and the help and reference pages.

## After a layout change

Retake every shot of a sheet whose layout changed, then check the footers of
the pages that use them.
