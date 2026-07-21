# Candidate Assessment Timer for Google Apps Script

A timed, one-question-at-a-time assessment web application built with Google
Apps Script and backed by Google Sheets.

The application creates candidate-specific links, enforces one attempt per
token, records answers and server-side timings, and captures browser-reported
answer-construction telemetry for human review.

> **Important:** Browser telemetry is unverified and may be incomplete,
> inaccurate, or manipulated. It must be treated as supporting information—not
> as proof of AI use or candidate misconduct.

## Contents

- [Capabilities](#capabilities)
- [Project files](#project-files)
- [Upgrade an existing installation](#upgrade-an-existing-installation)
- [Create a new installation](#create-a-new-installation)
- [Settings reference](#settings-reference)
- [Import questions from Google Forms](#import-questions-from-google-forms)
- [Spreadsheet data dictionary](#spreadsheet-data-dictionary)
- [Generate candidate links](#generate-candidate-links)
- [Test the deployment](#test-the-deployment)
- [Privacy, security, and interpretation](#privacy-security-and-interpretation)

## Capabilities

- Generates a unique token and link for each candidate.
- Presents one question at a time under an overall assessment timer.
- Records server-side question display and submission times.
- Records client-reported active time, first interaction, paste events, tab
  visibility changes, focus losses, key events, and editing statistics.
- Stores periodic answer snapshots in `AnswerSnapshots`.
- Retries unacknowledged snapshots using stable snapshot identifiers.
- Uses stable submission identifiers to prevent duplicate response rows.
- Freezes the question set for each session so later question edits do not
  change an assessment already in progress.
- Preserves legacy candidate tokens and deployed links.
- Stores assessment records in the configured Google Sheet.

## Project files

| File | Purpose |
| --- | --- |
| `Code.js` | Server-side Apps Script code, sheet setup, token handling, session management, response storage, and telemetry storage. Paste it into the Apps Script `Code.gs` file. |
| `FormImporter.js` | Imports supported questions and sections from an existing Google Form. Paste it into an Apps Script file named `FormImporter.gs`. |
| `Index.html` | Candidate-facing assessment page, timer, answer controls, browser telemetry, local draft recovery, and completion screen. |
| `appsscript.json` | Apps Script manifest, including the web-app access and execution settings. |
| `README.md` | Installation, configuration, operation, migration, and spreadsheet reference guide. |

## Upgrade an existing installation

The upgrade process preserves existing questions, candidates, sessions,
responses, snapshots, tokens, and links.

Before upgrading, make a backup copy of the spreadsheet.

1. Open the assessment spreadsheet.
2. Select **Extensions > Apps Script**.
3. Replace the existing `Code.gs` contents with `Code.js`.
4. Replace the existing `FormImporter.gs` contents with `FormImporter.js`.
5. Replace the existing `Index.html` contents with the new `Index.html`.
6. Save the Apps Script project.
7. Run `setupAssessment()` from the Apps Script editor.
8. Run `validateAssessmentSetup()` and resolve any reported problems.
9. Review the Settings sheet, particularly branding, privacy, timing, answer
   limits, draft retention, and snapshot settings.
10. Select **Deploy > Manage deployments**, edit the existing web-app
    deployment, select **New version**, and deploy it.
11. Test one existing token and one newly generated token.

`setupAssessment()` is safe to rerun. It appends missing columns and settings;
it does not clear existing assessment data or reorder existing columns.

The current migration adds:

- `Active` at the end of `Questions`;
- `SessionNonce` at the end of `Sessions`;
- any missing telemetry or identifier columns;
- any missing settings listed in [Settings reference](#settings-reference);
- frozen questions for legacy in-progress sessions that predate
  `SessionQuestions`.

Blank `Questions.Active` values remain active for backward compatibility.
Legacy sessions with a blank `SessionNonce` remain valid.

If the deployed `/exec` URL changes, update `WebAppUrl` and rerun
`generateCandidateTokens()` to refresh candidate links. Existing tokens are not
replaced.

## Create a new installation

### 1. Create the spreadsheet

1. Create a blank Google Sheet.
2. Give it a descriptive name, such as `Candidate Assessment - Responses`.
3. Select **Extensions > Apps Script**.

### 2. Add the project files

The Apps Script project should contain:

- `Code.gs`, containing `Code.js`;
- `FormImporter.gs`, containing `FormImporter.js`;
- `Index.html`, containing `Index.html`;
- `appsscript.json`, containing the supplied manifest.

### 3. Run setup

Run `setupAssessment()` from the Apps Script editor. It creates or upgrades:

- `Settings`;
- `Questions`;
- `Candidates`;
- `Sessions`;
- `Responses`;
- `AnswerSnapshots`;
- `SessionQuestions`.

The setup function also stores the spreadsheet ID in Script Properties so the
deployed web app can open the correct spreadsheet.

### 4. Configure and deploy

1. Review the Settings sheet.
2. Add or import assessment questions.
3. Add candidates to `Candidates`.
4. Run `validateAssessmentSetup()`.
5. Deploy the project as a web app:
   - execute as the script owner;
   - allow anyone to access it, if permitted by your Google Workspace policy.
6. Paste the deployed `/exec` URL into the `WebAppUrl` setting.
7. Run `generateCandidateTokens()`.

Administrative functions can only be run by an authorised editor. Anonymous
web-app visitors cannot run `setupAssessment()`, `validateAssessmentSetup()`,
`importQuestionsFromExistingForm()`, or `generateCandidateTokens()`.

If an authorised account needs to be explicitly allowed, add an `ADMIN_EMAILS`
Script Property containing comma-separated email addresses under
**Project Settings > Script properties**.

## Settings reference

| Setting | Description | Default or constraint |
| --- | --- | --- |
| `AssessmentTitle` | Assessment title displayed on the landing page and browser tab. | `Candidate Assessment` |
| `CompanyName` | Organisation name displayed in the landing-page branding area. | `Your Company` |
| `CompanyLogoUrl` | Optional direct URL for the company logo. It must be a publicly accessible HTTPS image without embedded credentials. | Blank; the page displays a built-in placeholder. |
| `DurationMinutes` | Overall time allowed for the assessment. | `60`; must be greater than zero. |
| `WebAppUrl` | Deployed Apps Script `/exec` URL used when generating candidate links. | Must be a deployed `https://script.google.com/...` URL. |
| `SourceFormUrl` | Full edit URL of the Google Form used by the optional importer. | Blank |
| `SnapshotIntervalSeconds` | Frequency of periodic client snapshots. | `15`; supported range: 5–300 seconds. |
| `StoreSnapshotText` | Controls whether evolving answer text is stored in `AnswerSnapshots`. When false, only metrics are stored. | `TRUE` |
| `MaxAnswerCharacters` | Maximum answer length enforced by both the browser and server. | `40000`; supported range: 1,000–49,000. |
| `DeadlineGraceSeconds` | Small server-controlled allowance for network and lock delay after the displayed deadline. | `15`; supported range: 0–120. |
| `DraftRetentionHours` | Expiry period for unfinished drafts stored in the candidate's browser. | `24`; supported range: 1–168. |
| `MaxSnapshotsPerQuestion` | Maximum snapshot sequence per question. The browser reserves the final entries for important paste and submit events. | `250`; supported range: 50–1,000. |
| `TelemetryInterpretation` | Internal statement explaining that client telemetry is unverified supporting information. | Added by setup and not overwritten when already present. |
| `PrivacyNotice` | Candidate-facing notice describing what activity and answer data are collected. | Added by setup and not overwritten when already present. |

An externally hosted logo can receive the candidate's image request. The page
sends no referrer, but the image host may still see ordinary request metadata.
Google Drive sharing-page URLs are not direct image URLs and generally do not
work for `CompanyLogoUrl`.

## Import questions from Google Forms

1. Run `setupAssessment()`.
2. Paste the full Google Form edit URL ending in `/edit` into `SourceFormUrl`.
3. Run `importQuestionsFromExistingForm()` from the Apps Script editor.
4. Review `Questions` and `ImportLog`.
5. Run `validateAssessmentSetup()` before deploying.

The importer supports:

- short text;
- paragraph text;
- multiple choice;
- dropdown lists, rendered as radio-button choices;
- linear scales, rendered as numbered choices;
- page breaks and section headers, mapped to `GroupID`.

Matching `GF_…` question IDs are updated without changing their established
order. New questions are appended. Unrelated questions and prior import logs are
preserved. Previously imported questions that no longer exist in the source
Form remain in the sheet but are set to `Active = FALSE`.

The importer does not replace a non-blank `AssessmentTitle`. Unsupported or
converted Form item types are recorded in `ImportLog` for review.

## Spreadsheet data dictionary

The following tables document the standard column order created by
`setupAssessment()`. Missing columns are appended during setup; existing columns
are not removed or reordered.

### AnswerSnapshots

Each row represents one point in the answer-construction timeline for a single
question. Snapshot metrics are cumulative within that question.

| # | Column heading | Description |
| ---: | --- | --- |
| 1 | `SessionID` | Server-generated identifier for the assessment attempt. Links the snapshot to `Sessions` and `Responses`. |
| 2 | `Token` | Candidate token associated with the session. |
| 3 | `CandidateName` | Candidate name copied from the session when the snapshot is stored. |
| 4 | `CandidateEmail` | Candidate email copied from the session when the snapshot is stored. |
| 5 | `QuestionOrder` | Configured display order of the frozen session question. |
| 6 | `QuestionID` | Stable question identifier. |
| 7 | `GroupID` | Optional question section or group identifier. |
| 8 | `SnapshotSequence` | Client-reported sequence number within the current question. |
| 9 | `Reason` | Event that created the snapshot: `question_shown`, `restored`, `interval`, `paste`, or `submit`. |
| 10 | `CapturedSecondsClient` | Client-reported seconds from the server-recorded question display time to snapshot capture. |
| 11 | `ActiveSecondsClient` | Estimated seconds the assessment tab was visible and focused by this point. |
| 12 | `AnswerCharacters` | Number of characters in the answer at snapshot capture. |
| 13 | `AnswerWords` | Estimated number of whitespace-delimited words at snapshot capture. |
| 14 | `TypedCharacters` | Cumulative characters attributed to ordinary typing events. |
| 15 | `PastedCharacters` | Cumulative characters attributed to detected paste events. |
| 16 | `LargestInsertionCharacters` | Largest net character insertion observed in one edit event. |
| 17 | `DeletedCharacters` | Cumulative characters removed from the answer. |
| 18 | `RevisionEvents` | Cumulative answer-changing edit events. |
| 19 | `PasteCount` | Cumulative detected paste events. |
| 20 | `TabLeaveCount` | Cumulative times the page became hidden. |
| 21 | `BlurCount` | Cumulative focus-loss events while the page remained visible. |
| 22 | `AnswerSnapshot` | Evolving answer text when `StoreSnapshotText` is true; otherwise blank. Stored text is capped at 20,000 characters per snapshot. |
| 23 | `SnapshotTextTruncated` | Boolean indicating that the snapshot text exceeded the stored-text limit. |
| 24 | `QuestionShownAtServer` | Server timestamp recorded when the current question was presented. |
| 25 | `ServerReceivedAt` | Server timestamp when the snapshot request was received. |
| 26 | `SnapshotID` | Stable client-generated identifier used to make snapshot retries idempotent. |
| 27 | `SubmissionID` | Stable identifier for the current question submission, allowing snapshots and the final response to be correlated. |
| 28 | `ClientTelemetryStatus` | Reliability label. Current browser-reported rows use `UNVERIFIED_CLIENT_REPORTED`. |

### Sessions

Each row represents one candidate assessment attempt and its aggregate state.

| # | Column heading | Description |
| ---: | --- | --- |
| 1 | `SessionID` | Server-generated unique identifier for the assessment attempt. |
| 2 | `Token` | Candidate token used to start or resume the session. |
| 3 | `CandidateName` | Candidate name copied from `Candidates` when the session starts. |
| 4 | `CandidateEmail` | Candidate email copied from `Candidates` when the session starts. |
| 5 | `StartedAt` | Server timestamp when the assessment attempt began. |
| 6 | `DeadlineAt` | Server-calculated assessment deadline. |
| 7 | `LastActivityAt` | Server timestamp of the most recent successfully processed answer activity. |
| 8 | `CurrentQuestionIndex` | Zero-based index of the current or next question in the frozen session question set. |
| 9 | `CurrentQuestionShownAt` | Server timestamp when the current question was presented. |
| 10 | `Status` | Session state, normally `IN_PROGRESS`, `COMPLETED`, or `TIME_EXPIRED`. |
| 11 | `CompletedAt` | Server timestamp when the assessment completed or expired. Blank while in progress. |
| 12 | `BrowserInfo` | JSON summary of browser-reported environment information captured at session start, truncated to 2,000 characters. |
| 13 | `TotalPasteCount` | Total detected paste events across submitted questions. |
| 14 | `TotalTabLeaveCount` | Total times the page became hidden across submitted questions. |
| 15 | `TotalBlurCount` | Total visible-page focus losses across submitted questions. |
| 16 | `TotalTypedCharacters` | Total characters attributed to ordinary typing across submitted questions. |
| 17 | `TotalPastedCharacters` | Total characters attributed to paste events across submitted questions. |
| 18 | `TotalDeletedCharacters` | Total removed characters across submitted questions. |
| 19 | `LargestInsertionCharacters` | Largest insertion observed in any submitted question. |
| 20 | `TotalRevisionEvents` | Total answer-changing edit events across submitted questions. |
| 21 | `TotalUndoCount` | Total detected browser undo events across submitted questions. |
| 22 | `TotalRedoCount` | Total detected browser redo events across submitted questions. |
| 23 | `TotalSnapshotCount` | Number of stored snapshots associated with the session. |
| 24 | `ClientTelemetryStatus` | Reliability label for aggregate browser-reported telemetry. |
| 25 | `SessionNonce` | Additional per-session secret required by new sessions when saving answers or progress. Blank legacy values remain supported. |

### Responses

Each row represents the submitted answer for one question. A stable
`SubmissionID`, together with the session and question identifiers, prevents a
retry from creating a duplicate response row.

| # | Column heading | Description |
| ---: | --- | --- |
| 1 | `SessionID` | Session identifier linking the response to `Sessions`. |
| 2 | `Token` | Candidate token associated with the session. |
| 3 | `CandidateName` | Candidate name copied from the session. |
| 4 | `CandidateEmail` | Candidate email copied from the session. |
| 5 | `QuestionOrder` | Configured display order of the frozen session question. |
| 6 | `QuestionID` | Stable identifier of the answered question. |
| 7 | `GroupID` | Optional question section or group identifier. |
| 8 | `QuestionText` | Frozen question text shown during the session. |
| 9 | `Answer` | Candidate's submitted answer, protected against spreadsheet formula injection. |
| 10 | `QuestionShownAtServer` | Server timestamp when the question was presented. |
| 11 | `SubmittedAtServer` | Server timestamp when the answer submission was processed. |
| 12 | `ElapsedSecondsServer` | Seconds between `QuestionShownAtServer` and `SubmittedAtServer`, calculated by the server. |
| 13 | `ActiveSecondsClient` | Client-reported seconds for which the page was visible and focused. |
| 14 | `FirstInteractionSecondsClient` | Client-reported seconds from question display to the first detected interaction. Blank when no interaction was detected. |
| 15 | `PasteCount` | Number of detected paste events for the question. |
| 16 | `TabLeaveCount` | Number of times the page became hidden during the question. |
| 17 | `BlurCount` | Number of visible-page focus-loss events during the question. |
| 18 | `KeyCount` | Number of detected editing key events. This is not necessarily equal to typed characters. |
| 19 | `AnswerCharacters` | Character count of the submitted answer. |
| 20 | `AnswerWords` | Estimated whitespace-delimited word count of the submitted answer. |
| 21 | `Required` | Whether the frozen question was configured as required. |
| 22 | `OverTime` | Boolean indicating that the answer was processed as a time-expired submission. |
| 23 | `BrowserInfo` | Browser environment summary copied from the session. |
| 24 | `TypedCharacters` | Characters attributed to ordinary typing events. |
| 25 | `PastedCharacters` | Characters attributed to detected paste events. |
| 26 | `PastedCharacterSharePercent` | `PastedCharacters` as a percentage of detected typed plus pasted insertions. |
| 27 | `LargestInsertionCharacters` | Largest net character insertion observed in one edit event. |
| 28 | `DeletedCharacters` | Total characters removed from the answer. |
| 29 | `DeletionEvents` | Number of answer-changing events that removed one or more characters. |
| 30 | `RevisionEvents` | Number of detected answer-changing edit events. |
| 31 | `UndoCount` | Number of detected browser undo input events. |
| 32 | `RedoCount` | Number of detected browser redo input events. |
| 33 | `LastEditSecondsClient` | Client-reported seconds from question display to the final detected edit. Blank when no edit was detected. |
| 34 | `EditingSpanSecondsClient` | Client-reported time between first interaction and final edit. Blank when either event was unavailable. |
| 35 | `SnapshotCount` | Number of stored `AnswerSnapshots` rows for this session and question. |
| 36 | `SubmissionID` | Stable client-generated identifier used to make answer submission retries idempotent. It is validated against the session and question. |
| 37 | `ClientTelemetryStatus` | Reliability label. Current browser-reported rows use `UNVERIFIED_CLIENT_REPORTED`. |

## Generate candidate links

1. Add each candidate to `Candidates` with a name or email.
2. Set `Active` to `TRUE` or leave it blank for token generation to activate the
   row.
3. Confirm that `WebAppUrl` contains the current deployed `/exec` URL.
4. Run `generateCandidateTokens()` from the Apps Script editor.
5. Distribute each candidate's generated `Link` privately.

Existing non-blank tokens are preserved. Token comparison is case-sensitive, so
legacy mixed-case tokens remain distinct.

Candidate links are bearer credentials: anyone who possesses an unused active
link can open that candidate's assessment. Do not publish links in shared or
public locations.

## Test the deployment

Use a dedicated test candidate and verify the following:

1. Open the generated `/exec?token=...` link in a private browser window.
2. Confirm that the company name, logo or placeholder, title, and privacy notice
   render correctly.
3. Start the assessment and verify that the timer and question progress appear.
4. Type, delete, paste, undo, redo, leave the tab, and return to it.
5. Wait long enough for at least one periodic snapshot.
6. Confirm that the page shows local and server save-status feedback.
7. Submit each question and confirm that duplicate clicking or retrying does not
   create duplicate `Responses` rows.
8. Confirm that `Sessions`, `Responses`, and `AnswerSnapshots` contain the
   expected identifiers and timestamps.
9. Confirm that the final candidate status is `COMPLETED`, or `TIME_EXPIRED` when
   testing the deadline.
10. Confirm that the completion page offers **Close form** and displays the
    manual-close fallback when the browser blocks scripted tab closure.
11. Resume an interrupted assessment from the original link and confirm that the
    current question and unexpired browser draft are restored.

For concurrency testing, submit two test candidates at nearly the same time and
verify that each receives the correct next question with no duplicate response
or snapshot rows.

## Privacy, security, and interpretation

Update `PrivacyNotice` so candidates understand what is collected, why it is
collected, who can review it, and how long it is retained.

Consider setting `StoreSnapshotText` to `FALSE` when timeline metrics are
sufficient. Full answer snapshots materially increase the amount of personal
data stored.

Restrict spreadsheet and Apps Script access to staff who need it. Review your
organisation's retention, recruitment, data-protection, and candidate-access
requirements before using the application in production.

Browser-reported telemetry can be affected by:

- dictation and accessibility software;
- autofill, spellcheck, and browser behaviour;
- mobile keyboards and input methods;
- network failures or blocked browser storage;
- deliberate client-side manipulation.

Do not automatically reject a candidate based on telemetry. Use it only as one
source of context for human review or a proportionate verification interview.
