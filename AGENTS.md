# Project instructions

This is a Google Apps Script web application attached to a Google Sheet.

## Main files

- Code.js contains server-side Apps Script code.
- FormImporter.js imports questions from an existing Google Form.
- Index.html contains the candidate-facing assessment UI.
- appsscript.json is the Apps Script manifest.

## Compatibility requirements

- Preserve all existing public function names.
- Do not rename existing Google Sheet tabs.
- Do not remove or reorder existing spreadsheet columns.
- New columns must be added through setupAssessment().
- setupAssessment() must be safe to run on an existing assessment.
- Do not clear existing Questions, Candidates, Sessions, Responses, or
  AnswerSnapshots data.
- Maintain compatibility with existing candidate tokens and deployed links.
- Do not store candidate telemetry anywhere except the configured Google Sheet.
- Treat telemetry as review information, not proof of AI use.

## Google Apps Script requirements

- Server-side code must remain compatible with Apps Script V8.
- Do not use Node.js-only APIs in server-side Apps Script files.
- Browser code may use modern JavaScript supported by current browsers.
- Remember that google.script.run works only inside the deployed Apps Script
  HTML service environment.
- Escape and validate candidate-supplied content before displaying it.

## Before finishing a task

- Check JavaScript syntax.
- Review backward compatibility.
- Explain which files changed.
- Explain whether setupAssessment() needs to be run.
- Explain whether a new web-app deployment version is required.
- Do not run clasp push unless explicitly requested.