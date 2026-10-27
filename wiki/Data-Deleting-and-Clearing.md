# Deleting and clearing

Remove jobs or whole lists. Every command here asks first, shows what will
go, and cannot be undone.

## Choose the right command

| To remove | Use | Where |
| --- | --- | --- |
| One job | **Delete** | Print jobs panel on the room sheet. See [Recording jobs](Jobs-Recording) |
| Every job in one room | **Clear All** | Right of the room sheet |
| The jobs a Reports filter shows, across rooms | **Delete visible records...** | Reports |
| Every row of one setup list | **Clear table** | Print Technicians, Printers, Papers, Consumables or Departments |
| A whole print room | **Remove print room...** | Settings. See [Print rooms](Setup-Print-Rooms) |

> **Warning:** none of these can be undone, including with Ctrl+Z / Cmd+Z.
> Run **Backup workbook...** on Settings first. See
> [Backup and restore](Data-Backup-and-Restore).

## Clear a room

1. Open the print room's sheet.
2. Click **Clear All**.
3. Read the number of jobs and their dates, and whether the room has
   unexported changes, then click **Yes**.

The room keeps its settings, printers and code. Other rooms are not touched.

## Delete the jobs a filter shows

1. On **Reports**, set filters so the list shows exactly the jobs to delete.
2. Click **Delete visible records...**.
3. Read the count for each room, then confirm twice.

It will not run without at least one filter set. If no report of these jobs
has been exported, it says so: use **Export report...** first if you need a
record. See [Exporting a report](Reports-Export).

## Clear a setup list

1. Open the setup sheet, such as **Printers**.
2. Click **Clear table**.
3. Read the rows listed and click **Yes**.

Jobs already recorded keep their prices. The other setup sheets are not
touched.

## Things to know

- Before a big clear-out, export the rooms involved. See
  [Exporting a print room](Data-Export-Room).
- The workbook keeps a log of what was removed, when and by whom, on a
  hidden sheet.
- To stop using a technician, printer or paper, set it to inactive rather
  than clearing it.

## Platform notes

> **Mac:** not yet checked.
>
> **Windows:** not yet checked.

## Related pages

- [Backup and restore](Data-Backup-and-Restore)
- [Filtering](Reports-Filtering)

---

*Draft for v0.11.3. Not yet verified on Windows or Mac.*
