# For developers

This wiki is for people using the workbook. How it is built and tested lives
in the project repository:

| Document | Covers |
| --- | --- |
| `docs/ARCHITECTURE.md` | The design: data model, sheets, formulas, VBA modules, decisions and open items |
| `SETUP.md` | Building the workbook on Windows and on a Mac, and running the tests |
| `docs/CHANGELOG.md` | Every version, in technical detail |
| `dist/` | The released workbook for each version |
| [`CONTRIBUTING.md`](https://github.com/Adasha/PrintJob/blob/master/CONTRIBUTING.md) | Making a change: branches, tests, versions, the wiki, and the checks each pull request runs |

The workbook is built from `src/PrintCosts.xlsx` and the VBA in
`src/PrintCosts-VBA/`. Each feature's page is listed in `wiki/_map.md`,
which reviewers use to spot undocumented features.

## Editing this wiki

Edit the pages in the repository's `wiki/` folder, not on GitHub's wiki
pages: a change merged into `master` is copied to the wiki by the
**Publish wiki** workflow (`.github/workflows/publish-wiki.yml`), and that
copy replaces anything edited on the web.

---

*Draft for v0.11.5.*
