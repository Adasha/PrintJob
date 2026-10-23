# Releases

One workbook per point release, `PrintJob-v<version>.xlsm`, ready to hand out.
Each is the newest build of that version made from a committed tree, so it
matches the code at the commit that added or last changed it.

To make or refresh one, on a clean checkout of the release commit:

- Mac: `python3 mac/build.py --dist`
- Windows: `powershell -ExecutionPolicy Bypass -File build.ps1 -Dist`

Then commit the file with a message naming the version. What changed in each
version is in [docs/CHANGELOG.md](../docs/CHANGELOG.md).

Older versions are in git history only when they were built here; 0.11.1 is
the first.
