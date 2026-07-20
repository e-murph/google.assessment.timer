/**
 * Candidate Assessment Timer
 * Google Apps Script server-side code.
 *
 * Bind this script to the Google Sheet that will store the assessment data.
 */

const APP = Object.freeze({
  DEFAULT_TITLE: 'Candidate Assessment',
  DEFAULT_DURATION_MINUTES: 60,
  SHEETS: Object.freeze({
    SETTINGS: 'Settings',
    QUESTIONS: 'Questions',
    CANDIDATES: 'Candidates',
    SESSIONS: 'Sessions',
    RESPONSES: 'Responses',
    SNAPSHOTS: 'AnswerSnapshots',
  }),
});

const HEADERS = Object.freeze({
  SETTINGS: ['Key', 'Value'],
  QUESTIONS: [
    'Order',
    'QuestionID',
    'GroupID',
    'QuestionText',
    'AnswerType',
    'Options',
    'Required',
    'MaxWords',
  ],
  CANDIDATES: [
    'CandidateName',
    'CandidateEmail',
    'Token',
    'Active',
    'Status',
    'Link',
    'Notes',
  ],
  SESSIONS: [
    'SessionID',
    'Token',
    'CandidateName',
    'CandidateEmail',
    'StartedAt',
    'DeadlineAt',
    'LastActivityAt',
    'CurrentQuestionIndex',
    'CurrentQuestionShownAt',
    'Status',
    'CompletedAt',
    'BrowserInfo',
    'TotalPasteCount',
    'TotalTabLeaveCount',
    'TotalBlurCount',
    'TotalTypedCharacters',
    'TotalPastedCharacters',
    'TotalDeletedCharacters',
    'LargestInsertionCharacters',
    'TotalRevisionEvents',
    'TotalUndoCount',
    'TotalRedoCount',
    'TotalSnapshotCount',
  ],
  RESPONSES: [
    'SessionID',
    'Token',
    'CandidateName',
    'CandidateEmail',
    'QuestionOrder',
    'QuestionID',
    'GroupID',
    'QuestionText',
    'Answer',
    'QuestionShownAtServer',
    'SubmittedAtServer',
    'ElapsedSecondsServer',
    'ActiveSecondsClient',
    'FirstInteractionSecondsClient',
    'PasteCount',
    'TabLeaveCount',
    'BlurCount',
    'KeyCount',
    'AnswerCharacters',
    'AnswerWords',
    'Required',
    'OverTime',
    'BrowserInfo',
    'TypedCharacters',
    'PastedCharacters',
    'PastedCharacterSharePercent',
    'LargestInsertionCharacters',
    'DeletedCharacters',
    'DeletionEvents',
    'RevisionEvents',
    'UndoCount',
    'RedoCount',
    'LastEditSecondsClient',
    'EditingSpanSecondsClient',
    'SnapshotCount',
  ],
  SNAPSHOTS: [
    'SessionID',
    'Token',
    'CandidateName',
    'CandidateEmail',
    'QuestionOrder',
    'QuestionID',
    'GroupID',
    'SnapshotSequence',
    'Reason',
    'CapturedSecondsClient',
    'ActiveSecondsClient',
    'AnswerCharacters',
    'AnswerWords',
    'TypedCharacters',
    'PastedCharacters',
    'LargestInsertionCharacters',
    'DeletedCharacters',
    'RevisionEvents',
    'PasteCount',
    'TabLeaveCount',
    'BlurCount',
    'AnswerSnapshot',
    'SnapshotTextTruncated',
    'QuestionShownAtServer',
    'ServerReceivedAt',
  ],
});

/**
 * Run this once from the Apps Script editor.
 * It creates the required sheets and sample rows.
 */
function setupAssessment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Open this script from a Google Sheet using Extensions > Apps Script.');
  }

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  const settings = ensureSheet_(ss, APP.SHEETS.SETTINGS, HEADERS.SETTINGS);
  const questions = ensureSheet_(ss, APP.SHEETS.QUESTIONS, HEADERS.QUESTIONS);
  const candidates = ensureSheet_(ss, APP.SHEETS.CANDIDATES, HEADERS.CANDIDATES);
  const sessions = ensureSheet_(ss, APP.SHEETS.SESSIONS, HEADERS.SESSIONS);
  const responses = ensureSheet_(ss, APP.SHEETS.RESPONSES, HEADERS.RESPONSES);
  const snapshots = ensureSheet_(ss, APP.SHEETS.SNAPSHOTS, HEADERS.SNAPSHOTS);

  ensureSetting_(settings, 'AssessmentTitle', APP.DEFAULT_TITLE);
  ensureSetting_(settings, 'DurationMinutes', APP.DEFAULT_DURATION_MINUTES);
  ensureSetting_(settings, 'WebAppUrl', 'PASTE_DEPLOYED_WEB_APP_URL_HERE');
  ensureSetting_(settings, 'SourceFormUrl', '');
  ensureSetting_(settings, 'SnapshotIntervalSeconds', 15);
  ensureSetting_(settings, 'StoreSnapshotText', true);
  ensureSetting_(
    settings,
    'PrivacyNotice',
    'This assessment records response times, active browser time, answer-construction statistics, periodic answer snapshots, paste events, and occasions when the assessment tab loses visibility or focus. This information is used to support human review of assessment integrity. These indicators are not treated as conclusive proof that artificial intelligence or another unauthorised tool was used.'
  );

  if (questions.getLastRow() <= 1) {
    questions.getRange(2, 1, 3, HEADERS.QUESTIONS.length).setValues([
      [1, 'Q1', 'A', 'Explain your approach to the example task.', 'long_text', '', true, 300],
      [2, 'Q2', 'A', 'Which option is most appropriate?', 'multiple_choice', 'Option A|Option B|Option C', true, ''],
      [3, 'Q3', 'B', 'Add any final comments.', 'short_text', '', false, 100],
    ]);
  }

  if (candidates.getLastRow() <= 1) {
    candidates.getRange(2, 1, 1, HEADERS.CANDIDATES.length).setValues([
      ['Example Candidate', 'candidate@example.com', '', true, '', '', 'Delete this example row before use.'],
    ]);
  }

  [settings, questions, candidates, sessions, responses, snapshots].forEach((sheet) => {
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  sessions.getRange('E:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sessions.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sessions.getRange('K:K').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  responses.getRange('J:K').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  snapshots.getRange('X:Y').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  return [
    'Setup and telemetry upgrade complete.',
    'Existing rows were preserved.',
    'Review PrivacyNotice, SnapshotIntervalSeconds and StoreSnapshotText in Settings, then deploy a new version of the web app.',
  ].join('\n');
}

/**
 * Run after adding candidate names/emails and after pasting the deployed URL
 * into Settings!B for the WebAppUrl row.
 */
function generateCandidateTokens() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const table = getTable_(APP.SHEETS.CANDIDATES);
    if (table.rows.length === 0) {
      throw new Error('Add at least one candidate to the Candidates sheet.');
    }

    const settings = getSettings_();
    const webAppUrl = String(settings.WebAppUrl || '').trim();
    const hasValidUrl = /^https:\/\/script\.google\.com\//i.test(webAppUrl);

    table.rows.forEach((row) => {
      const name = String(row[table.map.CandidateName] || '').trim();
      const email = String(row[table.map.CandidateEmail] || '').trim();
      if (!name && !email) return;

      let token = String(row[table.map.Token] || '').trim();
      if (!token) {
        token = Utilities.getUuid().replace(/-/g, '');
        row[table.map.Token] = token;
      }

      if (row[table.map.Active] === '') {
        row[table.map.Active] = true;
      }

      if (hasValidUrl) {
        const separator = webAppUrl.includes('?') ? '&' : '?';
        row[table.map.Link] = `${webAppUrl}${separator}token=${encodeURIComponent(token)}`;
      }
    });

    table.sheet
      .getRange(2, 1, table.rows.length, table.headers.length)
      .setValues(table.rows);

    return hasValidUrl
      ? 'Candidate tokens and links generated.'
      : 'Tokens generated. Paste the deployed web-app URL into Settings, then run generateCandidateTokens again to create links.';
  } finally {
    lock.releaseLock();
  }
}

/** Serve the assessment page. */
function doGet(e) {
  let settings = {};
  try {
    settings = getSettings_();
  } catch (error) {
    // Allows the page to render a useful message before setup is complete.
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.initialToken = e && e.parameter ? String(e.parameter.token || '') : '';
  template.assessmentTitle = String(settings.AssessmentTitle || APP.DEFAULT_TITLE);
  template.privacyNotice = String(settings.PrivacyNotice || 'Assessment activity is logged.');

  return template
    .evaluate()
    .setTitle(template.assessmentTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Starts a new attempt or resumes an existing in-progress attempt.
 * Called asynchronously from Index.html.
 */
function startAssessment(token, clientInfo) {
  const cleanToken = String(token || '').trim();
  if (!cleanToken) {
    return failure_('A candidate token is required. Open the unique link supplied to you.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const candidate = findCandidateByToken_(cleanToken);
    if (!candidate) return failure_('This candidate link is not valid.');
    if (!candidate.active) return failure_('This candidate link is inactive.');

    const questions = getQuestions_();
    if (questions.length === 0) return failure_('No assessment questions have been configured.');

    const settings = getSettings_();
    const durationMinutes = Math.max(
      1,
      Number(settings.DurationMinutes) || APP.DEFAULT_DURATION_MINUTES
    );
    const now = new Date();

    let session = findLatestSessionByToken_(cleanToken);
    if (session) {
      if (session.status === 'IN_PROGRESS') {
        if (now.getTime() > session.deadlineAt.getTime()) {
          completeSession_(session, 'TIME_EXPIRED', now);
          updateCandidateStatus_(candidate.rowNumber, 'TIME_EXPIRED');
          return failure_('The time allowed for this assessment has expired.');
        }
        return buildClientState_(session, questions, now);
      }

      return failure_('This candidate link has already been used.');
    }

    const startedAt = now;
    const deadlineAt = new Date(startedAt.getTime() + durationMinutes * 60 * 1000);
    const sessionId = Utilities.getUuid();
    const browserInfo = truncate_(JSON.stringify(clientInfo || {}), 2000);

    appendMappedRow_(APP.SHEETS.SESSIONS, {
      SessionID: sessionId,
      Token: cleanToken,
      CandidateName: safeForSheet_(candidate.name),
      CandidateEmail: safeForSheet_(candidate.email),
      StartedAt: startedAt,
      DeadlineAt: deadlineAt,
      LastActivityAt: now,
      CurrentQuestionIndex: 0,
      CurrentQuestionShownAt: now,
      Status: 'IN_PROGRESS',
      BrowserInfo: safeForSheet_(browserInfo),
      TotalPasteCount: 0,
      TotalTabLeaveCount: 0,
      TotalBlurCount: 0,
      TotalTypedCharacters: 0,
      TotalPastedCharacters: 0,
      TotalDeletedCharacters: 0,
      LargestInsertionCharacters: 0,
      TotalRevisionEvents: 0,
      TotalUndoCount: 0,
      TotalRedoCount: 0,
      TotalSnapshotCount: 0,
    });

    updateCandidateStatus_(candidate.rowNumber, 'IN_PROGRESS');
    session = findSessionById_(sessionId);
    return buildClientState_(session, questions, now);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves the current answer and returns the next question.
 * The server, not the browser, decides which question is current.
 */
function saveAndNext(payload) {
  payload = payload || {};
  const sessionId = String(payload.sessionId || '').trim();
  if (!sessionId) return failure_('The assessment session is missing. Reload your unique link.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const session = findSessionById_(sessionId);
    if (!session) return failure_('The assessment session could not be found.');
    if (session.status !== 'IN_PROGRESS') {
      return {
        ok: true,
        completed: true,
        status: session.status,
        message: 'This assessment has already been submitted.',
      };
    }

    const questions = getQuestions_();
    const index = Number(session.currentQuestionIndex);
    const question = questions[index];
    if (!question) {
      const nowWithoutQuestion = new Date();
      completeSession_(session, 'COMPLETED', nowWithoutQuestion);
      updateCandidateStatusByToken_(session.token, 'COMPLETED');
      return {
        ok: true,
        completed: true,
        status: 'COMPLETED',
        message: 'Assessment submitted.',
      };
    }

    // If a network retry submits an old question, do not create a duplicate row.
    if (String(payload.questionId || '') !== question.id) {
      return buildClientState_(session, questions, new Date());
    }

    const now = new Date();
    const overTime = now.getTime() > session.deadlineAt.getTime();
    const forceComplete = Boolean(payload.forceComplete);
    const answer = String(payload.answer == null ? '' : payload.answer).trim();
    const answerWords = countWords_(answer);

    if (!overTime && !forceComplete) {
      if (question.required && !answer) {
        return failure_('An answer is required before continuing.', 'VALIDATION');
      }
      if (question.maxWords && answerWords > question.maxWords) {
        return failure_(
          `This answer is ${answerWords} words. The maximum is ${question.maxWords}.`,
          'VALIDATION'
        );
      }
      if (
        question.answerType === 'multiple_choice' &&
        answer &&
        !question.options.includes(answer)
      ) {
        return failure_('Select one of the available options.', 'VALIDATION');
      }
    }

    const shownAt = session.currentQuestionShownAt;
    const elapsedMs = Math.max(0, now.getTime() - shownAt.getTime());
    const activeMs = Math.min(elapsedMs, nonNegativeNumber_(payload.activeMs));
    const firstInteractionMs = nullableNonNegativeNumber_(payload.firstInteractionMs);
    const pasteCount = nonNegativeInteger_(payload.pasteCount);
    const tabLeaveCount = nonNegativeInteger_(payload.tabLeaveCount);
    const blurCount = nonNegativeInteger_(payload.blurCount);
    const keyCount = nonNegativeInteger_(payload.keyCount);
    const typedCharacters = nonNegativeInteger_(payload.typedCharacters);
    const pastedCharacters = nonNegativeInteger_(payload.pastedCharacters);
    const largestInsertionCharacters = nonNegativeInteger_(payload.largestInsertionCharacters);
    const deletedCharacters = nonNegativeInteger_(payload.deletedCharacters);
    const deletionEvents = nonNegativeInteger_(payload.deletionEvents);
    const revisionEvents = nonNegativeInteger_(payload.revisionEvents);
    const undoCount = nonNegativeInteger_(payload.undoCount);
    const redoCount = nonNegativeInteger_(payload.redoCount);
    const lastEditMs = nullableNonNegativeNumber_(payload.lastEditMs);
    const editingSpanMs =
      firstInteractionMs == null || lastEditMs == null
        ? null
        : Math.max(0, lastEditMs - firstInteractionMs);
    const insertedCharacters = typedCharacters + pastedCharacters;
    const pastedSharePercent = insertedCharacters
      ? round_((pastedCharacters / insertedCharacters) * 100, 2)
      : 0;
    const snapshots = normaliseSnapshots_(payload.snapshots, {
      session,
      question,
      shownAt,
      receivedAt: now,
      elapsedMs,
    });

    appendMappedRow_(APP.SHEETS.RESPONSES, {
      SessionID: session.sessionId,
      Token: session.token,
      CandidateName: safeForSheet_(session.candidateName),
      CandidateEmail: safeForSheet_(session.candidateEmail),
      QuestionOrder: question.order,
      QuestionID: question.id,
      GroupID: safeForSheet_(question.groupId),
      QuestionText: safeForSheet_(question.text),
      Answer: safeForSheet_(answer),
      QuestionShownAtServer: shownAt,
      SubmittedAtServer: now,
      ElapsedSecondsServer: round_(elapsedMs / 1000, 2),
      ActiveSecondsClient: round_(activeMs / 1000, 2),
      FirstInteractionSecondsClient:
        firstInteractionMs == null ? '' : round_(firstInteractionMs / 1000, 2),
      PasteCount: pasteCount,
      TabLeaveCount: tabLeaveCount,
      BlurCount: blurCount,
      KeyCount: keyCount,
      AnswerCharacters: answer.length,
      AnswerWords: answerWords,
      Required: question.required,
      OverTime: overTime,
      BrowserInfo: safeForSheet_(session.browserInfo),
      TypedCharacters: typedCharacters,
      PastedCharacters: pastedCharacters,
      PastedCharacterSharePercent: pastedSharePercent,
      LargestInsertionCharacters: largestInsertionCharacters,
      DeletedCharacters: deletedCharacters,
      DeletionEvents: deletionEvents,
      RevisionEvents: revisionEvents,
      UndoCount: undoCount,
      RedoCount: redoCount,
      LastEditSecondsClient: lastEditMs == null ? '' : round_(lastEditMs / 1000, 2),
      EditingSpanSecondsClient: editingSpanMs == null ? '' : round_(editingSpanMs / 1000, 2),
      SnapshotCount: snapshots.length,
    });

    appendSnapshotRows_(snapshots);

    const nextIndex = index + 1;
    const isLastQuestion = nextIndex >= questions.length;
    const shouldComplete = overTime || forceComplete || isLastQuestion;

    if (shouldComplete) {
      const finalStatus = overTime ? 'TIME_EXPIRED' : 'COMPLETED';
      updateSession_(session.rowNumber, {
        LastActivityAt: now,
        CurrentQuestionIndex: nextIndex,
        Status: finalStatus,
        CompletedAt: now,
        TotalPasteCount: session.totalPasteCount + pasteCount,
        TotalTabLeaveCount: session.totalTabLeaveCount + tabLeaveCount,
        TotalBlurCount: session.totalBlurCount + blurCount,
        TotalTypedCharacters: session.totalTypedCharacters + typedCharacters,
        TotalPastedCharacters: session.totalPastedCharacters + pastedCharacters,
        TotalDeletedCharacters: session.totalDeletedCharacters + deletedCharacters,
        LargestInsertionCharacters: Math.max(
          session.largestInsertionCharacters,
          largestInsertionCharacters
        ),
        TotalRevisionEvents: session.totalRevisionEvents + revisionEvents,
        TotalUndoCount: session.totalUndoCount + undoCount,
        TotalRedoCount: session.totalRedoCount + redoCount,
        TotalSnapshotCount: session.totalSnapshotCount + snapshots.length,
      });
      updateCandidateStatusByToken_(session.token, finalStatus);

      return {
        ok: true,
        completed: true,
        status: finalStatus,
        message:
          finalStatus === 'TIME_EXPIRED'
            ? 'Time expired. Your current answer and earlier answers were submitted.'
            : 'Assessment submitted successfully.',
      };
    }

    updateSession_(session.rowNumber, {
      LastActivityAt: now,
      CurrentQuestionIndex: nextIndex,
      CurrentQuestionShownAt: now,
      TotalPasteCount: session.totalPasteCount + pasteCount,
      TotalTabLeaveCount: session.totalTabLeaveCount + tabLeaveCount,
      TotalBlurCount: session.totalBlurCount + blurCount,
      TotalTypedCharacters: session.totalTypedCharacters + typedCharacters,
      TotalPastedCharacters: session.totalPastedCharacters + pastedCharacters,
      TotalDeletedCharacters: session.totalDeletedCharacters + deletedCharacters,
      LargestInsertionCharacters: Math.max(
        session.largestInsertionCharacters,
        largestInsertionCharacters
      ),
      TotalRevisionEvents: session.totalRevisionEvents + revisionEvents,
      TotalUndoCount: session.totalUndoCount + undoCount,
      TotalRedoCount: session.totalRedoCount + redoCount,
      TotalSnapshotCount: session.totalSnapshotCount + snapshots.length,
    });

    const updatedSession = findSessionById_(session.sessionId);
    return buildClientState_(updatedSession, questions, now);
  } finally {
    lock.releaseLock();
  }
}

/** Optional diagnostic function to run from the editor. */
function validateAssessmentSetup() {
  const problems = [];
  const settings = getSettings_();
  const questions = getQuestions_();
  const candidates = getTable_(APP.SHEETS.CANDIDATES);

  if (!settings.AssessmentTitle) problems.push('AssessmentTitle is blank.');
  if (!(Number(settings.DurationMinutes) > 0)) problems.push('DurationMinutes must be greater than zero.');
  const snapshotIntervalSeconds = Number(settings.SnapshotIntervalSeconds);
  if (!(snapshotIntervalSeconds >= 5 && snapshotIntervalSeconds <= 300)) {
    problems.push('SnapshotIntervalSeconds must be between 5 and 300.');
  }
  if (!/^https:\/\/script\.google\.com\//i.test(String(settings.WebAppUrl || ''))) {
    problems.push('WebAppUrl has not been set to the deployed Apps Script URL.');
  }
  if (questions.length === 0) problems.push('No questions are configured.');
  if (!getSpreadsheet_().getSheetByName(APP.SHEETS.SNAPSHOTS)) {
    problems.push('AnswerSnapshots sheet is missing. Run setupAssessment.');
  }

  const ids = new Set();
  questions.forEach((question) => {
    if (ids.has(question.id)) problems.push(`Duplicate QuestionID: ${question.id}`);
    ids.add(question.id);
    if (!['short_text', 'long_text', 'multiple_choice'].includes(question.answerType)) {
      problems.push(`Unsupported AnswerType for ${question.id}: ${question.answerType}`);
    }
    if (question.answerType === 'multiple_choice' && question.options.length < 2) {
      problems.push(`Multiple-choice question ${question.id} needs at least two options.`);
    }
  });

  if (candidates.rows.length === 0) problems.push('No candidates are configured.');
  return problems.length ? problems.join('\n') : 'Setup looks valid.';
}

// -----------------------------------------------------------------------------
// Private helpers
// -----------------------------------------------------------------------------

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('Run setupAssessment once from the Sheet-bound Apps Script project.');
  }
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  const lastColumn = sheet.getLastColumn();
  const existingHeaders =
    sheet.getLastRow() > 0 && lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((value) => String(value).trim())
      : [];

  if (existingHeaders.length === 0 || existingHeaders.every((header) => !header)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const missingHeaders = headers.filter((header) => !existingHeaders.includes(header));
    if (missingHeaders.length) {
      sheet
        .getRange(1, existingHeaders.length + 1, 1, missingHeaders.length)
        .setValues([missingHeaders]);
    }
  }

  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  return sheet;
}

function ensureSetting_(settingsSheet, key, defaultValue) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow > 1) {
    const values = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const exists = values.some((row) => String(row[0] || '').trim() === key);
    if (exists) return;
  }
  settingsSheet.appendRow([key, defaultValue]);
}

function appendMappedRow_(sheetName, valuesByHeader) {
  const table = getTable_(sheetName);
  const row = Array(table.headers.length).fill('');
  Object.keys(valuesByHeader).forEach((header) => {
    if (!(header in table.map)) {
      throw new Error(`Required ${sheetName} column is missing: ${header}. Run setupAssessment.`);
    }
    row[table.map[header]] = valuesByHeader[header];
  });
  table.sheet.appendRow(row);
}

function appendSnapshotRows_(snapshots) {
  if (!snapshots.length) return;
  const table = getTable_(APP.SHEETS.SNAPSHOTS);
  const rows = snapshots.map((snapshot) => {
    const row = Array(table.headers.length).fill('');
    Object.keys(snapshot).forEach((header) => {
      if (header in table.map) row[table.map[header]] = snapshot[header];
    });
    return row;
  });

  table.sheet
    .getRange(table.sheet.getLastRow() + 1, 1, rows.length, table.headers.length)
    .setValues(rows);
}

function normaliseSnapshots_(rawSnapshots, context) {
  if (!Array.isArray(rawSnapshots)) return [];

  const settings = getSettings_();
  const storeSnapshotText = toBoolean_(settings.StoreSnapshotText);
  const maximumSnapshots = 250;
  const maximumTextCharacters = 20000;
  const permittedReasons = new Set(['question_shown', 'interval', 'paste', 'submit', 'restored']);

  return rawSnapshots.slice(0, maximumSnapshots).map((raw, index) => {
    const snapshot = raw && typeof raw === 'object' ? raw : {};
    const capturedMs = Math.min(
      context.elapsedMs,
      nonNegativeNumber_(snapshot.capturedMs)
    );
    const activeMs = Math.min(capturedMs, nonNegativeNumber_(snapshot.activeMs));
    const rawText = String(snapshot.answerSnapshot == null ? '' : snapshot.answerSnapshot);
    const textWasTruncated =
      toBoolean_(snapshot.snapshotTextTruncated) || rawText.length > maximumTextCharacters;
    const storedText = storeSnapshotText
      ? truncate_(rawText, maximumTextCharacters)
      : '';
    const reasonValue = String(snapshot.reason || 'interval').trim().toLowerCase();
    const reason = permittedReasons.has(reasonValue) ? reasonValue : 'interval';

    return {
      SessionID: context.session.sessionId,
      Token: context.session.token,
      CandidateName: safeForSheet_(context.session.candidateName),
      CandidateEmail: safeForSheet_(context.session.candidateEmail),
      QuestionOrder: context.question.order,
      QuestionID: context.question.id,
      GroupID: safeForSheet_(context.question.groupId),
      SnapshotSequence: index + 1,
      Reason: reason,
      CapturedSecondsClient: round_(capturedMs / 1000, 2),
      ActiveSecondsClient: round_(activeMs / 1000, 2),
      AnswerCharacters: nonNegativeInteger_(snapshot.answerCharacters),
      AnswerWords: nonNegativeInteger_(snapshot.answerWords),
      TypedCharacters: nonNegativeInteger_(snapshot.typedCharacters),
      PastedCharacters: nonNegativeInteger_(snapshot.pastedCharacters),
      LargestInsertionCharacters: nonNegativeInteger_(snapshot.largestInsertionCharacters),
      DeletedCharacters: nonNegativeInteger_(snapshot.deletedCharacters),
      RevisionEvents: nonNegativeInteger_(snapshot.revisionEvents),
      PasteCount: nonNegativeInteger_(snapshot.pasteCount),
      TabLeaveCount: nonNegativeInteger_(snapshot.tabLeaveCount),
      BlurCount: nonNegativeInteger_(snapshot.blurCount),
      AnswerSnapshot: safeForSheet_(storedText),
      SnapshotTextTruncated: storeSnapshotText && textWasTruncated,
      QuestionShownAtServer: context.shownAt,
      ServerReceivedAt: context.receivedAt,
    };
  });
}

function getTable_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Required sheet not found: ${sheetName}`);

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map((value) => String(value).trim());
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index;
  });

  return {
    sheet,
    headers,
    map,
    rows: values.slice(1),
  };
}

function getSettings_() {
  const table = getTable_(APP.SHEETS.SETTINGS);
  const settings = {};
  table.rows.forEach((row) => {
    const key = String(row[table.map.Key] || '').trim();
    if (key) settings[key] = row[table.map.Value];
  });
  return settings;
}

function getQuestions_() {
  const table = getTable_(APP.SHEETS.QUESTIONS);
  const questions = [];

  table.rows.forEach((row, rowIndex) => {
    const text = String(row[table.map.QuestionText] || '').trim();
    if (!text) return;

    const orderValue = Number(row[table.map.Order]);
    const order = Number.isFinite(orderValue) ? orderValue : rowIndex + 1;
    const id = String(row[table.map.QuestionID] || `Q${order}`).trim();
    const answerType = String(row[table.map.AnswerType] || 'long_text')
      .trim()
      .toLowerCase();
    const options = String(row[table.map.Options] || '')
      .split('|')
      .map((option) => option.trim())
      .filter(Boolean);
    const maxWordsValue = Number(row[table.map.MaxWords]);

    questions.push({
      order,
      id,
      groupId: String(row[table.map.GroupID] || '').trim(),
      text,
      answerType,
      options,
      required: toBoolean_(row[table.map.Required]),
      maxWords: Number.isFinite(maxWordsValue) && maxWordsValue > 0 ? maxWordsValue : null,
    });
  });

  questions.sort((a, b) => a.order - b.order);
  return questions;
}

function findCandidateByToken_(token) {
  const table = getTable_(APP.SHEETS.CANDIDATES);
  for (let i = 0; i < table.rows.length; i += 1) {
    const row = table.rows[i];
    if (String(row[table.map.Token] || '').trim() === token) {
      return {
        rowNumber: i + 2,
        name: String(row[table.map.CandidateName] || '').trim(),
        email: String(row[table.map.CandidateEmail] || '').trim(),
        active: toBoolean_(row[table.map.Active]),
        status: String(row[table.map.Status] || '').trim().toUpperCase(),
      };
    }
  }
  return null;
}

function updateCandidateStatus_(rowNumber, status) {
  const table = getTable_(APP.SHEETS.CANDIDATES);
  table.sheet.getRange(rowNumber, table.map.Status + 1).setValue(status);
}

function updateCandidateStatusByToken_(token, status) {
  const candidate = findCandidateByToken_(token);
  if (candidate) updateCandidateStatus_(candidate.rowNumber, status);
}

function findLatestSessionByToken_(token) {
  const table = getTable_(APP.SHEETS.SESSIONS);
  for (let i = table.rows.length - 1; i >= 0; i -= 1) {
    if (String(table.rows[i][table.map.Token] || '').trim() === token) {
      return sessionFromRow_(table.rows[i], i + 2, table.map);
    }
  }
  return null;
}

function findSessionById_(sessionId) {
  const table = getTable_(APP.SHEETS.SESSIONS);
  for (let i = table.rows.length - 1; i >= 0; i -= 1) {
    if (String(table.rows[i][table.map.SessionID] || '').trim() === sessionId) {
      return sessionFromRow_(table.rows[i], i + 2, table.map);
    }
  }
  return null;
}

function sessionFromRow_(row, rowNumber, map) {
  return {
    rowNumber,
    sessionId: String(row[map.SessionID] || ''),
    token: String(row[map.Token] || ''),
    candidateName: String(row[map.CandidateName] || ''),
    candidateEmail: String(row[map.CandidateEmail] || ''),
    startedAt: asDate_(row[map.StartedAt]),
    deadlineAt: asDate_(row[map.DeadlineAt]),
    lastActivityAt: asDate_(row[map.LastActivityAt]),
    currentQuestionIndex: Number(row[map.CurrentQuestionIndex]) || 0,
    currentQuestionShownAt: asDate_(row[map.CurrentQuestionShownAt]),
    status: String(row[map.Status] || '').trim().toUpperCase(),
    completedAt: row[map.CompletedAt] ? asDate_(row[map.CompletedAt]) : null,
    browserInfo: String(row[map.BrowserInfo] || ''),
    totalPasteCount: Number(row[map.TotalPasteCount]) || 0,
    totalTabLeaveCount: Number(row[map.TotalTabLeaveCount]) || 0,
    totalBlurCount: Number(row[map.TotalBlurCount]) || 0,
    totalTypedCharacters: Number(row[map.TotalTypedCharacters]) || 0,
    totalPastedCharacters: Number(row[map.TotalPastedCharacters]) || 0,
    totalDeletedCharacters: Number(row[map.TotalDeletedCharacters]) || 0,
    largestInsertionCharacters: Number(row[map.LargestInsertionCharacters]) || 0,
    totalRevisionEvents: Number(row[map.TotalRevisionEvents]) || 0,
    totalUndoCount: Number(row[map.TotalUndoCount]) || 0,
    totalRedoCount: Number(row[map.TotalRedoCount]) || 0,
    totalSnapshotCount: Number(row[map.TotalSnapshotCount]) || 0,
  };
}

function updateSession_(rowNumber, updates) {
  const table = getTable_(APP.SHEETS.SESSIONS);
  const row = table.sheet
    .getRange(rowNumber, 1, 1, table.headers.length)
    .getValues()[0];

  Object.keys(updates).forEach((header) => {
    if (!(header in table.map)) throw new Error(`Unknown Sessions column: ${header}`);
    row[table.map[header]] = updates[header];
  });

  table.sheet.getRange(rowNumber, 1, 1, table.headers.length).setValues([row]);
}

function completeSession_(session, status, completedAt) {
  updateSession_(session.rowNumber, {
    LastActivityAt: completedAt,
    Status: status,
    CompletedAt: completedAt,
  });
}

function buildClientState_(session, questions, now) {
  const index = Number(session.currentQuestionIndex);
  const question = questions[index];
  const settings = getSettings_();
  const snapshotIntervalSeconds = Math.min(
    300,
    Math.max(5, Number(settings.SnapshotIntervalSeconds) || 15)
  );

  if (!question) {
    return {
      ok: true,
      completed: true,
      status: session.status,
      message: 'Assessment submitted.',
    };
  }

  return {
    ok: true,
    completed: false,
    sessionId: session.sessionId,
    candidateName: session.candidateName,
    candidateEmail: session.candidateEmail,
    startedAtMs: session.startedAt.getTime(),
    deadlineAtMs: session.deadlineAt.getTime(),
    serverNowMs: now.getTime(),
    currentNumber: index + 1,
    totalQuestions: questions.length,
    questionShownAtMs: session.currentQuestionShownAt.getTime(),
    snapshotIntervalSeconds,
    storeSnapshotText: toBoolean_(settings.StoreSnapshotText),
    question: {
      id: question.id,
      groupId: question.groupId,
      text: question.text,
      answerType: question.answerType,
      options: question.options,
      required: question.required,
      maxWords: question.maxWords,
    },
  };
}

function failure_(message, code) {
  return {
    ok: false,
    code: code || 'ERROR',
    message: String(message || 'An error occurred.'),
  };
}

function toBoolean_(value) {
  if (value === true) return true;
  const text = String(value == null ? '' : value).trim().toLowerCase();
  return ['true', 'yes', 'y', '1'].includes(text);
}

function asDate_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) throw new Error(`Invalid date value: ${value}`);
  return date;
}

function countWords_(text) {
  const trimmed = String(text || '').trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}

function nonNegativeNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function nullableNonNegativeNumber_(value) {
  if (value === '' || value == null) return null;
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function nonNegativeInteger_(value) {
  return Math.min(1000000, Math.floor(nonNegativeNumber_(value)));
}

function round_(number, decimalPlaces) {
  const factor = 10 ** decimalPlaces;
  return Math.round(number * factor) / factor;
}

function truncate_(text, maxLength) {
  const value = String(text || '');
  return value.length <= maxLength ? value : value.slice(0, maxLength);
}

/** Protect the Sheet from formula injection through free-text answers. */
function safeForSheet_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}
