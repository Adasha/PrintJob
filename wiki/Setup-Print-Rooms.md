# Print rooms

Add a sheet for each print room, choose its printers, and rename or remove
rooms later.

## Before you start

- Add your printers on [Printers](Setup-Printers) first, so you can choose
  them for the room.

## Add a print room

1. On the **Settings** sheet, click **Add print room...**.
2. Type the room's name, such as `Photography`. It becomes the sheet's tab.
3. Type the room's department, or leave it blank.
4. Type a short code for the room's Job IDs, up to six letters or numbers.
   The workbook suggests one from the name; `PHOTO` gives Job IDs like
   `UNI-PHOTO-00001`. Each room needs its own code: if another room already
   has it, the workbook says so and suggests a free one, such as `PHOTO2`.
5. The new sheet opens, empty. Click **Select printers...** on it.
6. Tick every printer the room has, then click **OK**.

## Result

The room has its own sheet and appears on **Summary** and **Reports**. Its
**Printer** drop-down offers only the printers you ticked.

Next, set the room's defaults and roll length unit. See [Recording
jobs](Jobs-Recording) and [Views and units](Jobs-Views-and-Units).

## Rename a print room

1. On **Settings**, click **Rename print room...**.
2. Type the code of the room to rename, from the list shown.
3. Type the new name. The sheet tab changes to match.

The room keeps its code and its Job IDs. The name cell at the top of the room
sheet is locked once set, so use this command rather than typing over it.

## Remove a print room

1. Export the room first if you may want its jobs. See [Exporting a print
   room](Data-Export-Room).
2. On **Settings**, click **Remove print room...**.
3. Type the code of the room to remove.
4. Read what will be lost (the number of jobs, their dates and whether they
   were exported) and click **Yes**.
5. Type the room's name exactly, then click **OK**.

> **Warning:** removing a print room deletes its sheet and every job on it.
> It cannot be undone.

## Things to know

- The workbook needs at least one print room, so it will not remove the last
  one. To replace the **Example Print Room**, add your own room first, then
  remove it.
- A room's code is never reused or changed, because its Job IDs carry it.
- You can also add a room by copying a room's tab (**Move or Copy**, **Create
  a copy**). Then click **Clear All** and **Clear defaults** on the copy, and
  **Refresh Locations** on Settings. Refresh gives the copy its own code and
  adds it to the reports.
- Deleting a room's tab by hand also works, but nothing asks first. Click
  **Refresh Locations** afterwards.

## Platform notes

> **Mac:** since 0.11.2, **Remove print room...** asks Yes or No first and
> then asks for the name. In earlier versions on a Mac the name box did not
> appear.
>
> **Windows:** not yet checked.

## Related pages

- [Settings](Setup-Settings)
- [Recording jobs](Jobs-Recording)
- [Backup and restore](Data-Backup-and-Restore): restoring also creates rooms

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
