# Workout Buddy

Home-gym workout app for Josh and Michelle. Follow the `josh-app-dev` skill for shared conventions (same setup as Food Log).

- App: static files on GitHub Pages. `josh/` and `michelle/` are per-person entry pages (they set `window.WB_PERSON`); shared code in `app.js` and `app.css`. The root `index.html` redirects to `josh/` so Josh's original home-screen icon keeps working.
- Device storage is per person (`wb.<Person>.<key>`); Josh's pre-split `wb.<key>` data is copied over once on first load.
- Data: Google Sheet "Workout Buddy" (id `1hmU_XKOhSxIM_eculF_hw0agBoI6TKY7EnNb10x5go0`). Per-person tabs have a Person column (Sets, Events, Progress, Plan, Feedback, Exercise Overrides); a blank Person means Josh (rows from before Nov 21, 2026). Settings: blank Person = everyone, a Person row overrides for that person. Exercises is the shared library; per-person tempo/pace/on-off/form tips go in Exercise Overrides.
- API: `apps-script/Code.gs`, deployed by Josh. `state` returns only the caller's rows (`person` in the request; missing = Josh) and echoes `person`, which tells the app the people-aware script is live. Michelle's page refuses to save until it is. Generic `read` / `append` (with `tab`) / `update` / `delete` match Food Log; build new features on them so the script doesn't need redeploying.
- Feedback: "Tell trainer" / "Tell developer" save to the Feedback tab (Date, Time, Person, For = Trainer or Developer, Message, Status, EntryID, Context). Developer notes from both people come to Josh's developer Claude; review New ones at the start of dev work and mark them Done. Older notes are Events rows with Type note_developer / note_trainer.
- Michelle's setup link: Josh's page, Settings > Michelle's setup link (builds `michelle/#setup=...` from his saved connection).
- Voice packs in `voice/<pack>/`; Michelle defaults to the `josh` pack.
- Every push: bump `version.json` so open apps reload to the new version.
- Test before pushing: Playwright with the API stubbed via page.route at 390x844 and 820x1180 (see the josh-app-dev skill), and `node --check` plus a SpreadsheetApp mock for Code.gs.
