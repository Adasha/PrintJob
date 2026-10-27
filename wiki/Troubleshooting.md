# Troubleshooting

What to do when something does not work as expected. Find the symptom, then
follow the fix.

## Opening the workbook

| Symptom | Fix |
| --- | --- |
| The buttons do nothing | Macros are off. Close the workbook and open it again, turning macros on. See [Getting started](Getting-Started) |
| A red bar says Microsoft has blocked macros (Windows) | Close the workbook, right-click the file, **Properties**, tick **Unblock**, **OK**, then open it again |
| You saved and the buttons stopped working | It was saved as `.xlsx`, which removes macros. Go back to your `.xlsm` copy, and always save as **Excel Macro-Enabled Workbook** |

## Recording jobs

| Symptom | Fix |
| --- | --- |
| A printer is missing from the drop-down | Tick it in the room's **Select printers...**, and check it is Active on Printers |
| A paper is missing from the drop-down | It does not fit the job's printer, or it is inactive. See [Printers](Setup-Printers) |
| A technician is missing from the drop-down | Set them to Active on Print Technicians |
| The paper cleared when you changed the printer | The paper does not fit the new printer. Choose one that does |
| Status is orange | The job has a problem. Hover over Status to read it, then see [Validation and status](Jobs-Validation-and-Status) |
| You cannot type in a cell | It is calculated or locked, such as a cost, an ID or the room's name. See [Protection and hidden sheets](Admin-Protection-and-Hidden-Sheets) |
| **Add** and **Remove** in the Colour passes panel do nothing | Select a job on a multi-pass printer, or one of its passes, first |
| A job's printer cannot be changed | It has colour passes. Remove them first, or record a new job |

## Reports and totals

| Symptom | Fix |
| --- | --- |
| A new room is missing from Summary and Reports | Click **Refresh Locations** on Settings or Summary |
| Reports shows nothing | A filter excludes everything. Click **Clear all filters** |
| A Type on Summary shows **(not in Papers)** | That paper was renamed or removed after its jobs were recorded. Costs are unaffected |
| **Mark all as...** or **Delete visible records...** refuses to run | Set at least one filter first |

## Files

| Symptom | Fix |
| --- | --- |
| An export is not where you expected | The message after exporting says where it went. Set an **Export folder** on Settings to choose. See [Exporting a print room](Data-Export-Room) |
| A print room sheet was deleted by mistake | Click **Refresh Locations**, add the room again, then import its last export or restore the last backup |
| Restoring from an older workbook asks about macros (Mac) | Choose **Disable Macros** |

## Mac only

These were fixed in the versions shown. If you see them, upgrade. See
[Upgrading](Admin-Upgrading).

| Symptom | Fixed in |
| --- | --- |
| The workbook disappeared after Export, Export report or Backup | 0.11.1 |
| **Remove print room...** skipped asking for the room's name | 0.11.2 |
| A new job's Unit, Area and Paper Cost stayed blank | 0.11.2 |

## Still stuck

Note the exact message and what you clicked, then click **Check workbook** on
Settings and pass both to whoever looks after your workbook.

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
