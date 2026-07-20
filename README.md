# Timed candidate assessment using Google Apps Script

This package creates a Google Apps Script web app that:

- gives each candidate a unique tokenised link;
- presents one question at a time;
- records server-side question display and submission times;
- records active browser time, first interaction, paste events, tab visibility changes, focus losses and key events;
- records aggregate answer-construction statistics, including typed and pasted characters, largest insertion, deletions, revisions, undo and redo activity, and editing span;
- records periodic answer snapshots in a separate `AnswerSnapshots` tab;
- buffers snapshots in the browser and sends them when the question is submitted, rather than making a server call every few seconds;
- enforces one attempt per token;
- applies an overall assessment timer;
- saves all data in Google Sheets.

It does not prove that AI was used. Treat all activity data as review indicators only.

## Files

- `Code.gs`: server-side Apps Script code.
- `Index.html`: candidate web page and browser-side telemetry.
- `FormImporter.gs`: importer for copying compatible questions and sections from an existing Google Form.
- `README.md`: this guide; it remains outside Apps Script.

## Upgrading an existing installation

This version upgrades the earlier timer package without clearing existing data.

1. Open the Google Sheet that contains the assessment.
2. Select **Extensions > Apps Script**.
3. Replace the complete contents of the existing `Code.gs` with the new `Code.gs`.
4. Replace the complete contents of the existing `Index.html` with the new `Index.html`.
5. Keep the existing `FormImporter.gs`, or replace it with the package copy; its behaviour is unchanged.
6. Save all files.
7. Run `setupAssessment` once from the Apps Script editor.
8. Return to the spreadsheet and confirm that:
   - an `AnswerSnapshots` tab was created;
   - new telemetry columns were appended to `Responses` and `Sessions`;
   - `SnapshotIntervalSeconds` and `StoreSnapshotText` were added to `Settings`.
9. Review and update your existing `PrivacyNotice`. The upgrade deliberately does not overwrite a notice that you have already customised.
10. Select **Deploy > Manage deployments**, edit the current web-app deployment, choose **New version**, and deploy it.
11. Test with a fresh candidate token.

Running `setupAssessment` preserves existing imported questions, candidates, sessions and responses. It appends missing sheets, columns and settings.

## New settings

The following settings control snapshots:

- `SnapshotIntervalSeconds`: frequency of periodic snapshots. Default: `15`. Supported range: 5–300 seconds.
- `StoreSnapshotText`: `TRUE` stores the evolving answer text. `FALSE` stores only timings, lengths and construction statistics.

Full snapshot text creates considerably more personal data. Set `StoreSnapshotText` to `FALSE` when the timeline metrics are sufficient.

## New response metrics

The `Responses` tab includes:

- `TypedCharacters`: characters inserted through ordinary text-input events.
- `PastedCharacters`: characters inserted during detected paste or drag/drop operations.
- `PastedCharacterSharePercent`: pasted characters as a percentage of detected typed plus pasted insertions.
- `LargestInsertionCharacters`: largest net insertion in one edit event.
- `DeletedCharacters`: total number of removed characters.
- `DeletionEvents`: number of answer-changing events that removed text.
- `RevisionEvents`: number of detected answer changes.
- `UndoCount` and `RedoCount`: browser undo and redo input events.
- `LastEditSecondsClient`: time from question display to the final detected edit.
- `EditingSpanSecondsClient`: time between first interaction and final edit.
- `SnapshotCount`: number of snapshot rows stored for that response.

These values are browser-reported and can be affected by dictation, accessibility software, autofill, spellcheck, browser behaviour and deliberate manipulation.

## AnswerSnapshots tab

Each row is one timeline point for one question. Important columns include:

- `Reason`: `question_shown`, `restored`, `interval`, `paste`, or `submit`.
- `CapturedSecondsClient`: time since the question was displayed.
- `ActiveSecondsClient`: estimated active time accumulated by that point.
- `AnswerCharacters` and `AnswerWords`.
- cumulative typed, pasted, deleted and revision values.
- `AnswerSnapshot`: evolving answer text when `StoreSnapshotText` is `TRUE`.
- `QuestionShownAtServer` and `ServerReceivedAt`.

Snapshots are collected locally and written in one batch when the candidate saves the question. If the candidate closes the page before saving, the current question's telemetry may remain only in that browser's local storage.

## New installation

### 1. Create the data Sheet

1. Create a new blank Google Sheet.
2. Give it a name such as `Candidate Assessment - Responses`.
3. Open **Extensions > Apps Script**.

### 2. Add the files

1. Replace the default `Code.gs` with this package's `Code.gs`.
2. Add an HTML file named `Index` and paste `Index.html` into it.
3. Add a script file named `FormImporter` and paste `FormImporter.gs` into it.
4. Save the project.

The Apps Script file list should contain:

- `Code.gs`
- `FormImporter.gs`
- `Index.html`

### 3. Run setup

Run `setupAssessment`. It creates or upgrades:

- `Settings`
- `Questions`
- `Candidates`
- `Sessions`
- `Responses`
- `AnswerSnapshots`

### 4. Configure Settings

Review:

- `AssessmentTitle`
- `DurationMinutes`
- `WebAppUrl`
- `PrivacyNotice`
- `SourceFormUrl`
- `SnapshotIntervalSeconds`
- `StoreSnapshotText`

### 5. Import the existing Google Form

1. Paste the full Form edit URL ending in `/edit` beside `SourceFormUrl`.
2. Run `importQuestionsFromExistingForm`.
3. Review `Questions` and `ImportLog`.

The importer supports short text, paragraph text, multiple choice, dropdowns and linear scales. Sections become `GroupID` values. Unsupported items are listed in `ImportLog`.

### 6. Add candidates and deploy

1. Add candidates to `Candidates`.
2. Run `validateAssessmentSetup`.
3. Deploy as a web app, executing as the script owner.
4. Paste the `/exec` URL beside `WebAppUrl`.
5. Run `generateCandidateTokens`.

### 7. Test

Use a new test token and:

1. type text normally;
2. delete and rewrite part of it;
3. paste a sizeable paragraph;
4. use undo and redo;
5. remain on a question long enough for several interval snapshots;
6. submit and inspect `Responses` and `AnswerSnapshots`.

Confirm that pasted characters and largest insertion increase, snapshots include `interval`, `paste` and `submit`, and the final candidate status is `COMPLETED`.

## Privacy and interpretation

Update the candidate notice so it explains that answer-construction statistics and periodic answer snapshots are collected. Define access, purpose and retention. Restrict the Sheet to staff who genuinely need it.

Do not automatically reject candidates based on these measurements. Use correlated indicators to select answers for human review or a short verification interview.
