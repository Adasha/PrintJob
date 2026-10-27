# For developers

This wiki is for people using the workbook. How it is built and tested lives
in the project repository:

| Document | Covers |
| --- | --- |
| `docs/ARCHITECTURE.md` | The design: data model, sheets, formulas, VBA modules, decisions and open items |
| `SETUP.md` | Building the workbook on Windows and on a Mac, and running the tests |
| `docs/CHANGELOG.md` | Every version, in technical detail |
| `dist/` | The released workbook for each version |

The workbook is built from `src/PrintCosts.xlsx` and the VBA in
`src/PrintCosts-VBA/`. Each feature's page is listed in `wiki/_map.md`,
which reviewers use to spot undocumented features.

---

*Draft for v0.11.3.*
