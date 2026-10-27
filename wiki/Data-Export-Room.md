# Exporting a print room

Save a print room's jobs to a CSV file, as a safe copy or to move them to
another workbook.

## Export one room

1. Open the print room's sheet.
2. Click **Export...** on the right of the sheet.
3. The workbook asks whether to include student names. Choose **Yes** for a
   complete copy, **No** to leave names out, or **Cancel**.
4. A message says how many jobs were exported and where the file is.

## Export every room

1. On **Settings**, click **Export All Locations...**.
2. Answer the names question as above. It applies to every room.

Each room gets its own file.

## Result

A CSV file named like `PrintCosts-UNI-PHOTO-20261027-1430.csv`: site, room
code, date and time. It holds every job with all its columns, including the
prices each job was costed at, so it can be imported again with its costs
unchanged.

The **Export** line on the room's side panel then shows that the room has
been exported. Once jobs change, it says how many changes there are since.

## Where the files go

- To the **Export folder** set on [Settings](Setup-Settings), if that folder
  exists on this computer.
- Otherwise, beside the workbook.
- If neither works, to your Documents folder. The message always says where.

## Things to know

- Export a room before removing it, and regularly as a safe copy. Excel
  gives no warning when a sheet is deleted by hand.
- A file without names is not a complete copy, so it does not count as an
  export on the side panel. Importing it leaves names already in the
  workbook as they are.
- Export files contain personal data. Store them as your university's data
  protection rules require.
- [Backup workbook...](Data-Backup-and-Restore) exports every room and every
  setup list together, always with names.

## Platform notes

> **Mac:** the first time you export from a folder, Excel asks for access
> twice: once for the folder, then for the file. Allow both. With **Export
> folder** blank, the file goes beside the workbook.
>
> When the workbook is in OneDrive, Excel reports a web address for
> its folder. The workbook looks for the matching folder on your Mac; if it
> cannot find it, the file goes to Documents. Setting an **Export folder**
> avoids this.
>
> **Windows:** not yet checked.

## Related pages

- [Importing jobs](Data-Import)
- [Backup and restore](Data-Backup-and-Restore)
- [Print rooms](Setup-Print-Rooms)

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
