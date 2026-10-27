# Working on PrintJob

Notes for Claude Code sessions on any device. What the code does is in
`docs/ARCHITECTURE.md`; building and testing are in `SETUP.md`; release history
is in `docs/CHANGELOG.md`.

## Layout

- `src/PrintCosts.xlsx` (template) + `src/PrintCosts-VBA/*` (VBA source) build
  `src/PrintJob.xlsm`. Build output in `src/` is git-ignored.
- `dist/PrintJob-v<version>.xlsm` is the committed release copy of each point
  release (`dist/README.md`).
- Windows: `build.ps1`, `run-tests.ps1`, `test-*.ps1`. Mac: `mac/build.py`,
  `mac/run-tests.py`, `mac/test-*.py` (ports of the .ps1 tests, same checks and
  messages). A new or changed test gets both versions.
- Work in a clone outside OneDrive (`~/dev/PrintJob` on the Mac). OneDrive
  half-syncs `.git` and locks workbooks.

## How the owner works

- Branch from `master` and commit there; the owner pushes and opens PRs from
  VS Code. The Mac terminal has no GitHub auth and no `gh`.
- Once a version has been sent out, any further change needs a version bump
  (`APP_VERSION` in `modVersion.bas`) and its own changelog entry.
- Nothing has been verified on Windows since the Mac work (0.11.1 onward);
  say so whenever a change is only tested on the Mac.
- The wiki plan is the Claude Doc "PrintJob Wiki Plan"
  (<https://claude.ai/artifact/Xkwezyp5wbB7eUDrPU9Da3>). Keep its status and
  baseline current as work lands.

## Excel test runs on the Mac

- **Announce every batch first** (what runs, roughly how long) and wait for the
  go-ahead: the owner sets Excel's macro security to enable all macros for the
  run, else every workbook asks "Enable macros?" (hundreds of prompts). Remind
  them to restore it afterwards.
- Batch the work: write several changes, then run them together, so dialogs
  arrive in one sitting. Never probe anything that could crash Excel or raise a
  dialog without saying so first.
- The owner sometimes presents from this Mac in the morning: ask about time
  limits, and stop runs and quit Excel well before.
- Close any built workbook in Excel before `mac/build.py`; it refuses
  otherwise, and tests fail at once.
- Commands: `python3 mac/build.py` after VBA changes, then
  `python3 mac/run-tests.py --no-retry [--only name ...] [--stop-by HH:MM]`.
  The full suite takes about an hour.

## Driving Excel for Mac (16.114) - traps

Most of this is also in the headers of `mac/xlmac.py` and `mac/build.py`.

- `run VB Macro` (capital M, app level) runs a macro and returns text.
  Lowercase `run VB macro` is a range command and silently does nothing.
- Qualify and quote macro names: `'Book.xlsm'!Macro`. A missing macro returns
  empty text with no error, so treat empty as "did not run".
- Pass objects through `modScript.ScriptRun` (`wb.call()`), never as many
  optional AppleScript arguments. Crashes (Excel quits, then recovery and
  Grant Access prompts): calling a function with many omitted optionals, and
  `Application.Run` of a nonexistent macro inside VBA.
- Open workbooks with AppleScript `open POSIX file` inside
  `ignoring application responses`, then poll (`xlmac.open_in_excel`). Not
  `open -a`. Strip `com.apple.quarantine` from built files and copies.
- Tests and builds run inside Excel's sandbox container
  (`~/Library/Containers/com.microsoft.Excel/Data/PrintCosts{Test,Build}`),
  where Excel asks for no file access. Not `/tmp`, `$TMPDIR` or the Office
  group container.
- A workbook created while `ScreenUpdating = False` saves empty and Excel
  deletes it on close: use `modUtils.NewScratchBook`/`CloseScratchBook`.
- `Application.InputBox` with a prompt over about 255 characters shows nothing
  and returns Error 2015. All typed prompts go through `modUtils.InputText`.
  Tests cannot answer dialogs, so check prompt lengths by reading the code.
- VBE insert quirks: one `InsertLines` per module; `AddFromString` on a large
  module crashed Excel. Identifier case is normalised project-wide, so avoid a
  local that differs only in case from a parameter elsewhere.
- Setting a cell via AppleScript fires `Worksheet_Change`.
- Probe untried Excel features in isolation first, trapped, through
  `xlmac.builder_helper`.
