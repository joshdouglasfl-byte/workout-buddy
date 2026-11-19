# workout-buddy

Josh's home gym workout app for the iPad, plus the connection script for the Workout Buddy Google Sheet.

- `index.html`: the app (served by GitHub Pages)
- `apps-script/Code.gs`: the Google Apps Script web app that reads and writes the sheet. The deployed copy holds the access key; this copy does not.
- `voicetest/`: the iPad voice recognition test page
- The app reloads itself when a newer version is published (APP_VERSION is stamped on each commit by a local pre-commit hook).
