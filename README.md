# Timed candidate assessment using Google Apps Script

This package creates a Google Apps Script web app that:

- gives each candidate a unique tokenised link;
- presents one question at a time;
- records server-side question display and submission times;
- records active browser time, first interaction, paste events, tab visibility changes, focus losses and key events;
- records aggregate answer-construction statistics, including typed and pasted characters, largest insertion, deletions, revisions, undo and redo activity, and editing span;
- records periodic answer snapshots in a separate `AnswerSnapshots` tab;
- uploads snapshots incrementally and safely retries any that remain pending in the browser;
- enforces one attempt per token;
- applies an overall assessment timer;
- saves all data in Google Sheets.

It does not prove that AI was used. Treat all activity data as review indicators only.

## Files

- `Code.js`: server-side Apps Script code (paste its contents into the Apps Script `Code.gs` file).
- `Index.html`: candidate web page and browser-side telemetry.
- `FormImporter.js`: importer for copying compatible questions and sections from an existing Google Form.
- `README.md`: this guide; it remains outside Apps Script.

## Upgrading an existing installation

This version upgrades the earlier timer package without clearing existing data.

Before upgrading, make a copy of the spreadsheet. The setup routine is
non-destructive, but a backup is recommended before every production migration.

1. Open the Google Sheet that contains the assessment.
2. Select **Extensions > Apps Script**.
3. Replace the complete contents of the existing Apps Script `Code.gs` with this package's `Code.js`.
4. Replace the complete contents of the existing `Index.html` with the new `Index.html`.
5. Replace the existing importer with this package's `FormImporter.js` contents.
6. Save all files.
7. Run `setupAssessment` once from the Apps Script editor.
8. Return to the spreadsheet and confirm that:
   - `AnswerSnapshots` and `SessionQuestions` tabs exist;
   - `Active` was appended to `Questions` and `SessionNonce` was appended to `Sessions`;
   - new identifier and telemetry-interpretation columns were appended to `Responses`, `Sessions`, and `AnswerSnapshots`;
   - `CompanyName`, `CompanyLogoUrl`, `SnapshotIntervalSeconds`, `StoreSnapshotText`, `MaxAnswerCharacters`, `DeadlineGraceSeconds`, `DraftRetentionHours`, `MaxSnapshotsPerQuestion`, and `TelemetryInterpretation` exist in `Settings`.
9. Review and update your existing `PrivacyNotice`. The upgrade deliberately does not overwrite a notice that you have already customised.
10. Select **Deploy > Manage deployments**, edit the current web-app deployment, choose **New version**, and deploy it.
11. Test one existing token and one fresh candidate token. Mixed-case legacy tokens remain case-sensitive.
12. If the deployed `/exec` URL changed, update `WebAppUrl` and rerun `generateCandidateTokens` to refresh links. Existing tokens themselves are not replaced.

Running `setupAssessment` preserves existing imported questions, candidates, sessions and responses. It appends missing sheets, columns and settings. It also freezes the currently available questions for legacy in-progress sessions that predate `SessionQuestions`, preventing later imports from changing those attempts.

## Settings added by the upgrade

The following settings control snapshots:

- `SnapshotIntervalSeconds`: frequency of periodic snapshots. Default: `15`. Supported range: 5–300 seconds.
- `StoreSnapshotText`: `TRUE` stores the evolving answer text. `FALSE` stores only timings, lengths and construction statistics.
- `MaxAnswerCharacters`: maximum answer length accepted by both the page and server. Default: `40000`; supported range: 1,000–49,000.
- `DeadlineGraceSeconds`: server-controlled allowance for network and lock delay after the displayed deadline. Default: `15`; supported range: 0–120.
- `DraftRetentionHours`: expiry for unfinished answers stored in the candidate browser. Default: `24`; supported range: 1–168.
- `MaxSnapshotsPerQuestion`: total snapshot sequence allowance. Default: `250`; supported range: 50–1,000. The browser reserves the final entries for paste and submit events.

Full snapshot text creates considerably more personal data. Set `StoreSnapshotText` to `FALSE` when the timeline metrics are sufficient.

## New response metrics

The `Responses` tab includes:

- `TypedCharacters`: characters inserted through ordinary text-input events.
- `PastedCharacters`: characters inserted during detected paste operations.
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
Rows label these fields as `UNVERIFIED_CLIENT_REPORTED`; they are supporting information, not proof of misconduct.

## AnswerSnapshots tab

Each row is one timeline point for one question. Important columns include:

- `Reason`: `question_shown`, `restored`, `interval`, `paste`, or `submit`.
- `CapturedSecondsClient`: time since the question was displayed.
- `ActiveSecondsClient`: estimated active time accumulated by that point.
- `AnswerCharacters` and `AnswerWords`.
- cumulative typed, pasted, deleted and revision values.
- `AnswerSnapshot`: evolving answer text when `StoreSnapshotText` is `TRUE`.
- `QuestionShownAtServer` and `ServerReceivedAt`.

Snapshots are uploaded as the candidate works. A stable snapshot ID prevents duplicates when a request is retried. Snapshots that have not yet been acknowledged remain in expiring browser draft storage and are resent at the next interval or with the final answer.

Candidate drafts are stored only in the browser and expire according to
`DraftRetentionHours`. Completing the assessment removes all drafts for that
session from the device.

## New installation

### 1. Create the data Sheet

1. Create a new blank Google Sheet.
2. Give it a name such as `Candidate Assessment - Responses`.
3. Open **Extensions > Apps Script**.

### 2. Add the files

1. Replace the default Apps Script `Code.gs` contents with this package's `Code.js`.
2. Add an HTML file named `Index` and paste `Index.html` into it.
3. Add a script file named `FormImporter` and paste `FormImporter.js` into it.
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
- `SessionQuestions`

### 4. Configure Settings

Review:

- `AssessmentTitle`
- `CompanyName`
- `CompanyLogoUrl`: optional direct, publicly accessible HTTPS image URL. A built-in placeholder is shown when blank or unavailable. Ordinary Google Drive sharing-page URLs do not work as direct images. Loading an external logo lets that image host receive the browser request, although the page sends no referrer.
- `DurationMinutes`
- `WebAppUrl`
- `PrivacyNotice`
- `SourceFormUrl`
- `SnapshotIntervalSeconds`
- `StoreSnapshotText`
- `MaxAnswerCharacters`
- `DeadlineGraceSeconds`
- `DraftRetentionHours`
- `MaxSnapshotsPerQuestion`
- `TelemetryInterpretation`

### 5. Import the existing Google Form

1. Paste the full Form edit URL ending in `/edit` beside `SourceFormUrl`.
2. Run `importQuestionsFromExistingForm`.
3. Review `Questions` and `ImportLog`.

The importer supports short text, paragraph text, multiple choice, dropdowns and linear scales. Sections become `GroupID` values. Matching `GF_…` question IDs are updated without changing their established order, new questions are appended, unrelated questions are preserved, and each run is appended to `ImportLog`. Previously imported questions that no longer exist in the source Form are retained but set to `Active = FALSE`, so historical and in-progress sessions remain intact. Importing does not replace a non-blank `AssessmentTitle`. Unsupported items are listed in the log.

### 6. Add candidates and deploy

1. Add candidates to `Candidates`.
2. Run `validateAssessmentSetup`.
3. Deploy as a web app, executing as the script owner and allowing anyone to access it. Your Google Workspace policy must permit anonymous web apps.
4. Paste the `/exec` URL beside `WebAppUrl`.
5. Run `generateCandidateTokens`.

The administrative functions `setupAssessment`, `validateAssessmentSetup`,
`importQuestionsFromExistingForm`, and `generateCandidateTokens` must be run by
an authorised editor. Anonymous web-app visitors cannot invoke them. If an
authorised editor must use a different account, add a comma-separated
`ADMIN_EMAILS` script property in **Project Settings > Script properties**.

### 7. Test

Use a new test token and:

1. type text normally;
2. delete and rewrite part of it;
3. paste a sizeable paragraph;
4. use undo and redo;
5. remain on a question long enough for several interval snapshots;
6. submit and inspect `Responses` and `AnswerSnapshots`.

Confirm that pasted characters and largest insertion increase, snapshots include `interval`, `paste` and `submit`, and the final candidate status is `COMPLETED`.
Also confirm that the landing-page company branding appears and that the completion page offers a `Close form` button. Browsers that do not permit script-opened tab closure display a safe manual-close message instead.

## Privacy and interpretation

Update the candidate notice so it explains that answer-construction statistics and periodic answer snapshots are collected. Define access, purpose and retention. Restrict the Sheet to staff who genuinely need it.

Do not automatically reject candidates based on these measurements. Use correlated indicators to select answers for human review or a short verification interview.
