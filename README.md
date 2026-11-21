# workout-buddy

Josh's home gym workout app for the iPad, plus the connection script for the Workout Buddy Google Sheet.

- `josh/`, `michelle/`: each person's page; `app.js` and `app.css` hold the shared app (served by GitHub Pages). `index.html` redirects to `josh/`.
- `apps-script/Code.gs`: the Google Apps Script web app that reads and writes the sheet. The deployed copy holds the access key; this copy does not.
- `voicetest/`: the iPad voice recognition test page
- The app reloads itself when `version.json` changes (bumped on every push). See CLAUDE.md for how people are kept apart.
