# Contributing

PrintJob is a small internal project, so changes are welcome but reviews may
be slow. This page says where things live and what a pull request needs.

## Where things are

| What | Where |
| --- | --- |
| VBA source | `src/PrintCosts-VBA/`, built with `src/PrintCosts.xlsx` into `src/PrintJob.xlsm` (git-ignored) |
| Building and testing | [SETUP.md](SETUP.md): `build.ps1` and `run-tests.ps1` on Windows, `mac/build.py` and `mac/run-tests.py` on a Mac |
| How it works | [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md), the reference for behaviour |
| What changed | [docs/CHANGELOG.md](docs/CHANGELOG.md), one entry per version |
| Released workbooks | [dist/](dist), one `PrintJob-v<version>.xlsm` per release |
| User guide | [wiki/](wiki), published to the [GitHub wiki](https://github.com/Adasha/PrintJob/wiki) |

## Making a change

1. Branch from `master`.
2. Change the VBA in `src/PrintCosts-VBA/`, not in a built workbook: the
   build replaces the workbook's code.
3. Build and run the tests on the machine you have. A new or changed test
   needs both versions: `test-<name>.ps1` (Windows) and `mac/test-<name>.py`
   (Mac), with the same checks and messages.
4. If the workbook changed and the last version has been sent out, bump
   `APP_VERSION` in `modVersion.bas` and add a changelog entry. End the entry
   with a `Docs:` line naming the wiki pages affected, or `Docs: none`.
5. Open a pull request and fill in the checklist. Say which platform you
   tested on; Windows and Mac behave differently in places.

## Changing the wiki

Edit the pages in `wiki/`, never on the GitHub wiki itself: each merge to
`master` copies `wiki/` over the published wiki, so edits made on the web
are lost.

- A new button or setting needs a row in `wiki/_map.md` with the page that
  documents it. The map is for reviewers and is not published.
- A new page needs a link in `wiki/_Sidebar.md`.
- Links between pages use the page name without `.md`:
  `[Printers](Setup-Printers)`.
- `ARCHITECTURE.md` wins when the two disagree: fix it first, then the wiki.
  The wiki explains how to use the workbook, not how the code works.
- Each page's footer says which version it was written for.

## Checks on every pull request

The **Docs check** workflow runs markdownlint (settings in
`.markdownlint.jsonc`) and `.github/scripts/check-docs.py`. The script fails
on a broken link or missing image, a wiki page missing from the sidebar, or a
`btn*` macro or `SET_*` setting with no row in the map. Run both locally
before pushing:

```text
npx markdownlint-cli2 "**/*.md" "#node_modules"
python3 .github/scripts/check-docs.py
```

GitHub's runners have no Excel, so the workbook tests run only on your
machine.
