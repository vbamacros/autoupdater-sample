# autoupdater-sample

A simple way to remotely update macro-enabled Excel files.

## See it in action

1. Download [v1.0.0](https://github.com/vbamacros/autoupdater-sample/releases/download/v1.0.0/autoupdater-sample.xlsm) or [v2.0.0](https://github.com/vbamacros/autoupdater-sample/releases/download/v2.0.0/autoupdater-sample.xlsm).
2. Alt + click and check the Unblock box under Security. (All downloaded macros are flagged as "potentially dangerous")
3. Open the file; you should be prompted to update to v3.0.0.

## ⚠️ Warnings

- This will **irreversibly delete** any changes the user may have made to the old version.
- The repo must be PUBLIC, for releases to be visible via api.
- Don't forget to update the "LOCAL_VERSION_TAG" variable to match the latest release tag.
- To use with a Word file, you must change app-specific statements (e.g. ThisWorkbook => ThisDocument)

## How it works

1. Trigger `Autoupdater.Autoupdate` on `Workbook_Open`.
2. Fetch latest release using Github's api:
   `https://api.github.com/repos/{USER/ORG}/{REPO}/releases/latest`

3. Compare `tag_name` vs `LOCAL_VERSION`
4. Download binary from `browser_download_url` to temp folder
5. Create and call BAT file that:
   - Waits a few seconds for application to close
   - Deletes old file
   - Copies new file from temp, and opens it.

### TODO

[ ] Give user option to set a date to ask again, or never update.

[ ] Support add-ins (Require closing all instances of Word/Excel before replacing add-in file)

[ ] Support multiple downloads (i.e. Excel macro and companion Word template)

[ ] Support PowerPoint, which does not have equivalent to ThisDocument/ThisWorkbook.
https://stackoverflow.com/questions/1472418/how-to-simulate-thispresentation-in-powerpoint-vba

[ ] Import/export only the code modules, leaving the rest of the data untouched. Challenges: Properly importing document (event) modules and forms, and dealing with macro security for VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3).
