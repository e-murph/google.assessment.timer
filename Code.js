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
    SESSION_QUESTIONS: 'SessionQuestions',
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
    'Active',
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
    'ClientTelemetryStatus',
    'SessionNonce',
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
    'SubmissionID',
    'ClientTelemetryStatus',
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
    'SnapshotID',
    'SubmissionID',
    'ClientTelemetryStatus',
  ],
  SESSION_QUESTIONS: [
    'SessionID',
    'QuestionIndex',
    'Order',
    'QuestionID',
    'GroupID',
    'QuestionText',
    'AnswerType',
    'Options',
    'Required',
    'MaxWords',
  ],
});

/**
 * Run this once from the Apps Script editor.
 * It creates the required sheets and sample rows.
 */
function setupAssessment() {
  assertAdministrator_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('The assessment is busy. Wait for active submissions to finish, then run setup again.');
  }
  try {
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
  const sessionQuestions = ensureSheet_(
    ss,
    APP.SHEETS.SESSION_QUESTIONS,
    HEADERS.SESSION_QUESTIONS
  );

  ensureSetting_(settings, 'AssessmentTitle', APP.DEFAULT_TITLE);
  ensureSetting_(settings, 'CompanyName', 'Your Company');
  ensureSetting_(settings, 'CompanyLogoUrl', '');
  ensureSetting_(settings, 'DurationMinutes', APP.DEFAULT_DURATION_MINUTES);
  ensureSetting_(settings, 'WebAppUrl', 'PASTE_DEPLOYED_WEB_APP_URL_HERE');
  ensureSetting_(settings, 'SourceFormUrl', '');
  ensureSetting_(settings, 'SnapshotIntervalSeconds', 15);
  ensureSetting_(settings, 'StoreSnapshotText', true);
  ensureSetting_(settings, 'MaxAnswerCharacters', 40000);
  ensureSetting_(settings, 'DeadlineGraceSeconds', 15);
  ensureSetting_(settings, 'DraftRetentionHours', 24);
  ensureSetting_(settings, 'MaxSnapshotsPerQuestion', 250);
  ensureSetting_(
    settings,
    'TelemetryInterpretation',
    'Client telemetry is unverified, may be incomplete or manipulated, and must be treated only as supporting information for human review—not as proof of AI use or misconduct.'
  );
  ensureSetting_(
    settings,
    'PrivacyNotice',
    'This assessment records response times, active browser time, answer-construction statistics, periodic answer snapshots, paste events, and occasions when the assessment tab loses visibility or focus. This information is used to support human review of assessment integrity. These indicators are not treated as conclusive proof that artificial intelligence or another unauthorised tool was used.'
  );

  if (questions.getLastRow() <= 1) {
    appendMappedRowToSheet_(questions, {
      Order: 1, QuestionID: 'Q1', GroupID: 'A',
      QuestionText: 'Explain your approach to the example task.',
      AnswerType: 'long_text', Required: true, MaxWords: 300, Active: true,
    });
    appendMappedRowToSheet_(questions, {
      Order: 2, QuestionID: 'Q2', GroupID: 'A',
      QuestionText: 'Which option is most appropriate?',
      AnswerType: 'multiple_choice', Options: 'Option A|Option B|Option C',
      Required: true, Active: true,
    });
    appendMappedRowToSheet_(questions, {
      Order: 3, QuestionID: 'Q3', GroupID: 'B',
      QuestionText: 'Add any final comments.', AnswerType: 'short_text',
      Required: false, MaxWords: 100, Active: true,
    });
  }

  if (candidates.getLastRow() <= 1) {
    appendMappedRowToSheet_(candidates, {
      CandidateName: 'Example Candidate', CandidateEmail: 'candidate@example.com',
      Active: true, Notes: 'Delete this example row before use.',
    });
  }

  const backfilledLegacySessions = backfillLegacySessionQuestions_();

  [
    settings,
    questions,
    candidates,
    sessions,
    responses,
    snapshots,
    sessionQuestions,
  ].forEach((sheet) => {
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  formatColumnsByHeader_(sessions, [
    'StartedAt', 'DeadlineAt', 'LastActivityAt', 'CurrentQuestionShownAt', 'CompletedAt',
  ], 'yyyy-mm-dd hh:mm:ss');
  formatColumnsByHeader_(responses, [
    'QuestionShownAtServer', 'SubmittedAtServer',
  ], 'yyyy-mm-dd hh:mm:ss');
  formatColumnsByHeader_(snapshots, [
    'QuestionShownAtServer', 'ServerReceivedAt',
  ], 'yyyy-mm-dd hh:mm:ss');
  SpreadsheetApp.flush();

  return [
    'Setup and telemetry upgrade complete.',
    'Existing rows were preserved.',
    `Frozen questions were backfilled for ${backfilledLegacySessions} legacy in-progress session(s).`,
    'Review branding, privacy, timing, answer-limit, draft-retention and snapshot settings, then deploy a new version of the web app.',
  ].join('\n');
  } finally {
    flushAndReleaseLock_(lock);
  }
}

/**
 * Run after adding candidate names/emails and after pasting the deployed URL
 * into Settings!B for the WebAppUrl row.
 */
function generateCandidateTokens() {
  assertAdministrator_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const table = getTable_(APP.SHEETS.CANDIDATES);
    if (table.rows.length === 0) {
      throw new Error('Add at least one candidate to the Candidates sheet.');
    }
    const duplicateTokens = findDuplicateCandidateTokens_(table);
    if (duplicateTokens.length) {
      throw new Error(
        `Duplicate candidate token(s) found in rows ${duplicateTokens
          .map((duplicate) => duplicate.rows.join(' and '))
          .join(', ')}. Give each candidate a unique token before generating links.`
      );
    }

    const settings = getSettings_();
    const webAppUrl = String(settings.WebAppUrl || '').trim();
    const hasValidUrl = /^https:\/\/script\.google\.com\//i.test(webAppUrl);
    const usedTokens = new Set(
      table.rows
        .map((row) => String(row[table.map.Token] || '').trim())
        .filter(Boolean)
    );

    table.rows.forEach((row, index) => {
      const name = String(row[table.map.CandidateName] || '').trim();
      const email = String(row[table.map.CandidateEmail] || '').trim();
      if (!name && !email) return;

      let token = String(row[table.map.Token] || '').trim();
      if (!token) {
        do {
          token = Utilities.getUuid().replace(/-/g, '');
        } while (usedTokens.has(token));
        usedTokens.add(token);
      }

      const updates = { Token: token };
      if (row[table.map.Active] === '') updates.Active = true;

      if (hasValidUrl) {
        const separator = webAppUrl.includes('?') ? '&' : '?';
        updates.Link = `${webAppUrl}${separator}token=${encodeURIComponent(token)}`;
      }
      setMappedValuesInRow_(table, index + 2, updates);
    });

    return hasValidUrl
      ? 'Candidate tokens and links generated.'
      : 'Tokens generated. Paste the deployed web-app URL into Settings, then run generateCandidateTokens again to create links.';
  } finally {
    flushAndReleaseLock_(lock);
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
  template.companyName = String(settings.CompanyName || 'Your Company');
  const configuredLogoUrl = String(settings.CompanyLogoUrl || '').trim();
  template.companyLogoUrl = isSafeHttpsImageUrl_(configuredLogoUrl) ? configuredLogoUrl : '';
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
  if (!isValidIdentifier_(cleanToken, 250)) return failure_('This candidate link is not valid.');
  if (isRateLimited_('start', cleanToken, 10, 60)) {
    return failure_('Too many start attempts were received. Wait a minute and try again.', 'RATE_LIMIT');
  }

  // Reject bad bearer tokens before they can occupy the global write lock.
  const preflightCandidate = findCandidateByToken_(cleanToken);
  if (!preflightCandidate) return failure_('This candidate link is not valid.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return failure_('The assessment is busy saving another response. Please try again.', 'BUSY');
  }

  try {
    const candidate = findCandidateByToken_(cleanToken);
    if (!candidate) return failure_('This candidate link is not valid.');
    if (!candidate.active) return failure_('This candidate link is inactive.');

    const questions = getQuestions_();

    const settings = getSettings_();
    const durationMinutes = Math.max(
      1,
      Number(settings.DurationMinutes) || APP.DEFAULT_DURATION_MINUTES
    );
    const now = new Date();

    let session = findLatestSessionByToken_(cleanToken);
    if (session) {
      if (session.status === 'IN_PROGRESS') {
        // Return the current question even after the deadline. The browser can
        // then restore its locally saved answer and submit it as overtime.
        return buildClientState_(session, getQuestionsForSession_(session), now);
      }

      return failure_('This candidate link has already been used.');
    }

    if (questions.length === 0) return failure_('No assessment questions have been configured.');
    const questionProblems = getQuestionConfigurationProblems_(questions);
    if (questionProblems.length) {
      return failure_(
        `The assessment question configuration is invalid: ${questionProblems.join(' ')}`
      );
    }

    const startedAt = now;
    const deadlineAt = new Date(startedAt.getTime() + durationMinutes * 60 * 1000);
    const sessionId = Utilities.getUuid();
    const sessionNonce = Utilities.getUuid().replace(/-/g, '');
    const browserInfo = truncate_(JSON.stringify(clientInfo || {}), 2000);

    // Freeze the exact question set before creating the session so later edits
    // or imports cannot change an assessment that is already in progress.
    saveSessionQuestions_(sessionId, questions);

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
      ClientTelemetryStatus: 'UNVERIFIED_CLIENT_REPORTED',
      SessionNonce: sessionNonce,
    });

    updateCandidateStatus_(candidate.rowNumber, 'IN_PROGRESS');
    session = findSessionById_(sessionId);
    return buildClientState_(session, questions, now);
  } finally {
    flushAndReleaseLock_(lock);
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
  if (!isValidIdentifier_(sessionId, 200)) return failure_('The assessment session is invalid.');
  if (isRateLimited_('submit', sessionId, 30, 60)) {
    return failure_('Too many submission attempts were received. Wait a moment and try again.', 'RATE_LIMIT');
  }

  const preflightSession = findSessionById_(sessionId);
  if (!preflightSession) return failure_('The assessment session could not be found.');
  if (!sessionNonceMatches_(preflightSession, payload.sessionNonce)) {
    return failure_('The assessment session could not be authenticated. Reload your unique link.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return failure_('Your answer is still stored in this browser. Please try submitting again.', 'BUSY');
  }

  try {
    const session = findSessionById_(sessionId);
    if (!session) return failure_('The assessment session could not be found.');
    if (!sessionNonceMatches_(session, payload.sessionNonce)) {
      return failure_('The assessment session could not be authenticated. Reload your unique link.');
    }
    if (session.status !== 'IN_PROGRESS') {
      updateCandidateStatusByToken_(session.token, session.status);
      return {
        ok: true,
        completed: true,
        status: session.status,
        message: 'This assessment has already been submitted.',
      };
    }

    const questions = getQuestionsForSession_(session);
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
    const requestSettings = getSettings_();
    const deadlineGraceSeconds = Math.min(
      120,
      Math.max(0, Number(requestSettings.DeadlineGraceSeconds) || 0)
    );
    const deadlineExpired = now.getTime() > session.deadlineAt.getTime();
    const overTime = now.getTime() > session.deadlineAt.getTime() + deadlineGraceSeconds * 1000;
    const forceComplete = Boolean(payload.forceComplete);
    const timeExpiredSubmission = overTime || (forceComplete && deadlineExpired);
    const submissionId = truncate_(String(payload.submissionId || Utilities.getUuid()), 200);
    const answer = String(payload.answer == null ? '' : payload.answer).trim();
    const answerWords = countWords_(answer);
    const maxAnswerCharacters = getMaxAnswerCharacters_(requestSettings);

    if (answer.length > maxAnswerCharacters) {
      return failure_(
        `This answer is ${answer.length} characters. The maximum is ${maxAnswerCharacters}.`,
        'VALIDATION'
      );
    }

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
    const firstInteractionValue = nullableNonNegativeNumber_(payload.firstInteractionMs);
    const firstInteractionMs = firstInteractionValue == null
      ? null
      : Math.min(elapsedMs, firstInteractionValue);
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
    const lastEditValue = nullableNonNegativeNumber_(payload.lastEditMs);
    const lastEditMs = lastEditValue == null ? null : Math.min(elapsedMs, lastEditValue);
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
      submissionId,
      settings: requestSettings,
    });

    appendSnapshotRows_(snapshots);
    const snapshotCounts = getSnapshotCounts_(session.sessionId, question.id);

    upsertResponseForQuestion_(session.sessionId, question.id, {
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
      OverTime: timeExpiredSubmission,
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
      SnapshotCount: snapshotCounts.question,
      SubmissionID: safeForSheet_(submissionId),
      ClientTelemetryStatus: 'UNVERIFIED_CLIENT_REPORTED',
    });

    const nextIndex = index + 1;
    const isLastQuestion = nextIndex >= questions.length;
    const shouldComplete = overTime || forceComplete || isLastQuestion;

    if (shouldComplete) {
      const finalStatus = timeExpiredSubmission ? 'TIME_EXPIRED' : 'COMPLETED';
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
        TotalSnapshotCount: snapshotCounts.session,
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
      TotalSnapshotCount: snapshotCounts.session,
    });

    session.currentQuestionIndex = nextIndex;
    session.currentQuestionShownAt = now;
    session.lastActivityAt = now;
    session.totalSnapshotCount = snapshotCounts.session;
    return buildClientState_(session, questions, now, requestSettings);
  } finally {
    flushAndReleaseLock_(lock);
  }
}

/**
 * Persists incremental answer snapshots while a question is still open.
 * Snapshot IDs make retries safe, and the final submission can resend any
 * snapshot whose acknowledgement was lost.
 */
function saveProgress(payload) {
  payload = payload || {};
  const sessionId = String(payload.sessionId || '').trim();
  if (!sessionId) return failure_('The assessment session is missing.');
  if (!isValidIdentifier_(sessionId, 200)) return failure_('The assessment session is invalid.');
  if (isRateLimited_('progress', sessionId, 120, 60)) {
    return failure_('Progress saving is temporarily rate limited.', 'RATE_LIMIT');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) {
    return {
      ok: false,
      retry: true,
      message: 'Snapshot save deferred while another assessment submission is being processed.',
    };
  }

  try {
    const session = findSessionById_(sessionId);
    if (!session || session.status !== 'IN_PROGRESS') {
      return failure_('The assessment session is no longer in progress.');
    }
    if (!sessionNonceMatches_(session, payload.sessionNonce)) {
      return failure_('The assessment session could not be authenticated.');
    }

    const questions = getQuestionsForSession_(session);
    const question = questions[Number(session.currentQuestionIndex)];
    if (!question || String(payload.questionId || '') !== question.id) {
      return failure_('The assessment has moved to another question.');
    }

    const receivedAt = new Date();
    const shownAt = session.currentQuestionShownAt;
    const elapsedMs = Math.max(0, receivedAt.getTime() - shownAt.getTime());
    const submissionId = truncate_(String(payload.submissionId || ''), 200);
    const snapshots = normaliseSnapshots_(payload.snapshots, {
      session,
      question,
      shownAt,
      receivedAt,
      elapsedMs,
      submissionId,
    });

    appendSnapshotRows_(snapshots);
    return {
      ok: true,
      savedSnapshotIds: snapshots.map((snapshot) => snapshot.SnapshotID),
    };
  } finally {
    flushAndReleaseLock_(lock);
  }
}

/** Optional diagnostic function to run from the editor. */
function validateAssessmentSetup() {
  assertAdministrator_();
  const problems = [];
  const settings = getSettings_();
  const questions = getQuestions_();
  const candidates = getTable_(APP.SHEETS.CANDIDATES);

  if (!settings.AssessmentTitle) problems.push('AssessmentTitle is blank.');
  if (!String(settings.CompanyName || '').trim()) problems.push('CompanyName is blank.');
  const companyLogoUrl = String(settings.CompanyLogoUrl || '').trim();
  if (companyLogoUrl && !isSafeHttpsImageUrl_(companyLogoUrl)) {
    problems.push('CompanyLogoUrl must be a valid HTTPS URL without spaces or embedded credentials.');
  }
  if (!(Number(settings.DurationMinutes) > 0)) problems.push('DurationMinutes must be greater than zero.');
  const snapshotIntervalSeconds = Number(settings.SnapshotIntervalSeconds);
  if (!(snapshotIntervalSeconds >= 5 && snapshotIntervalSeconds <= 300)) {
    problems.push('SnapshotIntervalSeconds must be between 5 and 300.');
  }
  if (!(Number(settings.MaxAnswerCharacters) >= 1000 && Number(settings.MaxAnswerCharacters) <= 49000)) {
    problems.push('MaxAnswerCharacters must be between 1000 and 49000.');
  }
  if (!(Number(settings.DeadlineGraceSeconds) >= 0 && Number(settings.DeadlineGraceSeconds) <= 120)) {
    problems.push('DeadlineGraceSeconds must be between 0 and 120.');
  }
  if (!(Number(settings.DraftRetentionHours) >= 1 && Number(settings.DraftRetentionHours) <= 168)) {
    problems.push('DraftRetentionHours must be between 1 and 168.');
  }
  if (!(Number(settings.MaxSnapshotsPerQuestion) >= 50 && Number(settings.MaxSnapshotsPerQuestion) <= 1000)) {
    problems.push('MaxSnapshotsPerQuestion must be between 50 and 1000.');
  }
  if (!/^https:\/\/script\.google\.com\//i.test(String(settings.WebAppUrl || ''))) {
    problems.push('WebAppUrl has not been set to the deployed Apps Script URL.');
  }
  if (questions.length === 0) problems.push('No questions are configured.');
  if (!getSpreadsheet_().getSheetByName(APP.SHEETS.SNAPSHOTS)) {
    problems.push('AnswerSnapshots sheet is missing. Run setupAssessment.');
  }
  if (!getSpreadsheet_().getSheetByName(APP.SHEETS.SESSION_QUESTIONS)) {
    problems.push('SessionQuestions sheet is missing. Run setupAssessment.');
  }

  [
    [APP.SHEETS.RESPONSES, ['SubmissionID']],
    [APP.SHEETS.SNAPSHOTS, ['SnapshotID', 'SubmissionID']],
  ].forEach(([sheetName, requiredHeaders]) => {
    const sheet = getSpreadsheet_().getSheetByName(sheetName);
    if (!sheet) return;
    const table = getTable_(sheetName);
    requiredHeaders.forEach((header) => {
      if (!(header in table.map)) {
        problems.push(`${sheetName} is missing ${header}. Run setupAssessment.`);
      }
    });
  });

  problems.push(...getQuestionConfigurationProblems_(questions));

  if (candidates.rows.length === 0) problems.push('No candidates are configured.');
  findDuplicateCandidateTokens_(candidates).forEach((duplicate) => {
    problems.push(`Duplicate candidate token in rows ${duplicate.rows.join(' and ')}.`);
  });
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

  const nonBlankHeaders = existingHeaders.filter(Boolean);
  const duplicateHeaders = nonBlankHeaders.filter(
    (header, index) => nonBlankHeaders.indexOf(header) !== index
  );
  if (duplicateHeaders.length) {
    throw new Error(
      `${name} contains duplicate header(s): ${[...new Set(duplicateHeaders)].join(', ')}. Correct row 1 before running setupAssessment.`
    );
  }
  if (existingHeaders.some((header) => !header) && nonBlankHeaders.length) {
    throw new Error(`${name} contains a blank header in row 1. Correct it before running setupAssessment.`);
  }

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
  const structure = getSheetStructure_(settingsSheet);
  if (!(HEADERS.SETTINGS[0] in structure.map) || !(HEADERS.SETTINGS[1] in structure.map)) {
    throw new Error('Settings must contain Key and Value headers.');
  }
  const lastRow = settingsSheet.getLastRow();
  if (lastRow > 1) {
    const values = settingsSheet
      .getRange(2, structure.map.Key + 1, lastRow - 1, 1)
      .getValues();
    const exists = values.some((row) => String(row[0] || '').trim() === key);
    if (exists) return;
  }
  appendMappedRowToSheet_(settingsSheet, { Key: key, Value: defaultValue });
}

function getSheetStructure_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getValues()[0]
    .map((value) => String(value || '').trim());
  const map = {};
  headers.forEach((header, index) => {
    if (header) map[header] = index;
  });
  return { sheet, headers, map };
}

function appendMappedRowToSheet_(sheet, valuesByHeader) {
  const structure = getSheetStructure_(sheet);
  const row = Array(structure.headers.length).fill('');
  Object.keys(valuesByHeader).forEach((header) => {
    if (!(header in structure.map)) throw new Error(`${sheet.getName()} is missing ${header}.`);
    row[structure.map[header]] = valuesByHeader[header];
  });
  sheet.appendRow(row);
}

function formatColumnsByHeader_(sheet, headers, numberFormat) {
  const structure = getSheetStructure_(sheet);
  headers.forEach((header) => {
    if (!(header in structure.map)) throw new Error(`${sheet.getName()} is missing ${header}.`);
    sheet
      .getRange(2, structure.map[header] + 1, Math.max(sheet.getMaxRows() - 1, 1), 1)
      .setNumberFormat(numberFormat);
  });
}

/** Update only named cells, preserving formulas and custom columns elsewhere in the row. */
function setMappedValuesInRow_(table, rowNumber, valuesByHeader) {
  const cells = Object.keys(valuesByHeader)
    .map((header) => {
      if (!(header in table.map)) throw new Error(`${table.sheet.getName()} is missing ${header}.`);
      return { column: table.map[header] + 1, value: valuesByHeader[header] };
    })
    .sort((a, b) => a.column - b.column);

  let group = [];
  const writeGroup = () => {
    if (!group.length) return;
    table.sheet
      .getRange(rowNumber, group[0].column, 1, group.length)
      .setValues([group.map((cell) => cell.value)]);
    group = [];
  };

  cells.forEach((cell) => {
    if (group.length && cell.column !== group[group.length - 1].column + 1) writeGroup();
    group.push(cell);
  });
  writeGroup();
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

function upsertResponseForQuestion_(sessionId, questionId, valuesByHeader) {
  const table = getTableStructure_(APP.SHEETS.RESPONSES);
  const lastRow = table.sheet.getLastRow();
  let sheetRow = -1;
  const submissionId = String(valuesByHeader.SubmissionID || '').trim();

  if (lastRow > 1 && submissionId) {
    const matches = table.sheet
      .getRange(2, table.map.SubmissionID + 1, lastRow - 1, 1)
      .createTextFinder(submissionId)
      .matchCase(true)
      .matchEntireCell(true)
      .findAll();
    if (matches.length > 1) {
      throw new Error('Duplicate SubmissionID values exist in Responses. Correct them before continuing.');
    }
    if (matches.length === 1) {
      const candidateRow = matches[0].getRow();
      const storedIdentity = table.sheet
        .getRange(
          candidateRow,
          Math.min(table.map.SessionID, table.map.QuestionID) + 1,
          1,
          Math.abs(table.map.QuestionID - table.map.SessionID) + 1
        )
        .getValues()[0];
      const firstColumn = Math.min(table.map.SessionID, table.map.QuestionID);
      const storedSessionId = String(storedIdentity[table.map.SessionID - firstColumn] || '').trim();
      const storedQuestionId = String(storedIdentity[table.map.QuestionID - firstColumn] || '').trim();
      if (storedSessionId !== sessionId || storedQuestionId !== questionId) {
        throw new Error('The submission identifier conflicts with another response. Reload and try again.');
      }
      sheetRow = candidateRow;
    }
  }

  // Compatibility fallback for submissions made before SubmissionID existed.
  if (sheetRow < 0 && lastRow > 1) {
    const sessionMatches = table.sheet
      .getRange(2, table.map.SessionID + 1, lastRow - 1, 1)
      .createTextFinder(sessionId)
      .matchCase(true)
      .matchEntireCell(true)
      .findAll();
    for (let index = sessionMatches.length - 1; index >= 0; index -= 1) {
      const candidateRow = sessionMatches[index].getRow();
      const storedQuestionId = table.sheet
        .getRange(candidateRow, table.map.QuestionID + 1)
        .getValue();
      if (String(storedQuestionId || '').trim() === questionId) {
        sheetRow = candidateRow;
        break;
      }
    }
  }

  Object.keys(valuesByHeader).forEach((header) => {
    if (!(header in table.map)) {
      throw new Error(`Required Responses column is missing: ${header}. Run setupAssessment.`);
    }
  });

  if (sheetRow >= 0) {
    setMappedValuesInRow_(table, sheetRow, valuesByHeader);
  } else {
    const row = Array(table.headers.length).fill('');
    Object.keys(valuesByHeader).forEach((header) => {
      row[table.map[header]] = valuesByHeader[header];
    });
    table.sheet.appendRow(row);
  }
}

function getSnapshotCounts_(sessionId, questionId) {
  const table = getTableStructure_(APP.SHEETS.SNAPSHOTS);
  const lastRow = table.sheet.getLastRow();
  if (lastRow <= 1) return { question: 0, session: 0 };

  const firstColumn = Math.min(table.map.SessionID, table.map.QuestionID);
  const lastColumn = Math.max(table.map.SessionID, table.map.QuestionID);
  const values = table.sheet
    .getRange(2, firstColumn + 1, lastRow - 1, lastColumn - firstColumn + 1)
    .getValues();
  const sessionOffset = table.map.SessionID - firstColumn;
  const questionOffset = table.map.QuestionID - firstColumn;
  let sessionCount = 0;
  let questionCount = 0;
  values.forEach((row) => {
    if (String(row[sessionOffset] || '').trim() !== sessionId) return;
    sessionCount += 1;
    if (String(row[questionOffset] || '').trim() === questionId) questionCount += 1;
  });
  return { question: questionCount, session: sessionCount };
}

function appendSnapshotRows_(snapshots) {
  if (!snapshots.length) return;
  const table = getTableStructure_(APP.SHEETS.SNAPSHOTS);
  const lastRow = table.sheet.getLastRow();
  const snapshotIdRange = lastRow > 1
    ? table.sheet.getRange(2, table.map.SnapshotID + 1, lastRow - 1, 1)
    : null;
  const existingIds = snapshots.length > 10 && snapshotIdRange
    ? new Set(
        snapshotIdRange
          .getValues()
          .map((row) => String(row[0] || '').trim())
          .filter(Boolean)
      )
    : new Set();
  const batchIds = new Set();
  const newSnapshots = snapshots.filter((snapshot) => {
    const snapshotId = String(snapshot.SnapshotID || '').trim();
    if (!snapshotId || batchIds.has(snapshotId) || existingIds.has(snapshotId)) return false;
    if (
      snapshotIdRange &&
      snapshots.length <= 10 &&
      snapshotIdRange
        .createTextFinder(snapshotId)
        .matchCase(true)
        .matchEntireCell(true)
        .findNext()
    ) {
      return false;
    }
    batchIds.add(snapshotId);
    existingIds.add(snapshotId);
    return true;
  });
  if (!newSnapshots.length) return;

  const rows = newSnapshots.map((snapshot) => {
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

  const settings = context.settings || getSettings_();
  const storeSnapshotText = toBoolean_(settings.StoreSnapshotText);
  const maximumSnapshots = Math.min(
    1000,
    Math.max(50, Number(settings.MaxSnapshotsPerQuestion) || 250)
  );
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
    const sequence = Math.max(1, nonNegativeInteger_(snapshot.sequence) || index + 1);
    const suppliedSnapshotId = truncate_(String(snapshot.snapshotId || '').trim(), 200);
    const snapshotId = suppliedSnapshotId || [
      context.session.sessionId,
      context.question.id,
      sequence,
      Math.round(capturedMs),
    ].join(':');

    return {
      SessionID: context.session.sessionId,
      Token: context.session.token,
      CandidateName: safeForSheet_(context.session.candidateName),
      CandidateEmail: safeForSheet_(context.session.candidateEmail),
      QuestionOrder: context.question.order,
      QuestionID: context.question.id,
      GroupID: safeForSheet_(context.question.groupId),
      SnapshotSequence: sequence,
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
      SnapshotID: safeForSheet_(snapshotId),
      SubmissionID: safeForSheet_(context.submissionId || ''),
      ClientTelemetryStatus: 'UNVERIFIED_CLIENT_REPORTED',
    };
  });
}

function getTableStructure_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Required sheet not found: ${sheetName}`);

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map((value) => String(value).trim());
  const blankHeaderIndex = headers.findIndex((header) => !header);
  if (blankHeaderIndex >= 0) {
    throw new Error(
      `${sheetName} has a blank header in column ${blankHeaderIndex + 1}. Correct row 1 before continuing.`
    );
  }
  const duplicateHeaders = headers.filter(
    (header, index) => headers.indexOf(header) !== index
  );
  if (duplicateHeaders.length) {
    throw new Error(
      `${sheetName} contains duplicate header(s): ${[...new Set(duplicateHeaders)].join(', ')}.`
    );
  }
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index;
  });

  const sheetKey = Object.keys(APP.SHEETS).find((key) => APP.SHEETS[key] === sheetName);
  const expectedHeaders = sheetKey ? HEADERS[sheetKey] : null;
  if (expectedHeaders) {
    const missingHeaders = expectedHeaders.filter((header) => !(header in map));
    if (missingHeaders.length) {
      throw new Error(
        `${sheetName} is missing required column(s): ${missingHeaders.join(', ')}. Run setupAssessment.`
      );
    }
  }

  return {
    sheet,
    headers,
    map,
  };
}

function getTable_(sheetName) {
  const table = getTableStructure_(sheetName);
  const lastRow = table.sheet.getLastRow();
  const rows = lastRow > 1
    ? table.sheet.getRange(2, 1, lastRow - 1, table.headers.length).getValues()
    : [];
  return {
    ...table,
    rows,
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
    const activeValue = row[table.map.Active];
    if (String(activeValue == null ? '' : activeValue).trim() && !toBoolean_(activeValue)) return;
    const text = String(row[table.map.QuestionText] || '').trim();
    if (!text) return;

    const rawOrder = row[table.map.Order];
    const orderValue = Number(rawOrder);
    const hasExplicitOrder = String(rawOrder == null ? '' : rawOrder).trim() !== '';
    const hasValidOrder = hasExplicitOrder && Number.isFinite(orderValue) && orderValue > 0;
    const order = hasValidOrder
      ? orderValue
      : rowIndex + 1;
    const configuredId = String(row[table.map.QuestionID] || '').trim();
    const id = configuredId || `Q${order}`;
    const answerType = String(row[table.map.AnswerType] || 'long_text')
      .trim()
      .toLowerCase();
    const options = String(row[table.map.Options] || '')
      .split('|')
      .map((option) => option.trim())
      .filter(Boolean);
    const maxWordsValue = Number(row[table.map.MaxWords]);
    const hasMaxWords = String(row[table.map.MaxWords] == null ? '' : row[table.map.MaxWords]).trim() !== '';

    questions.push({
      order,
      id,
      hasExplicitOrder,
      hasValidOrder,
      hasExplicitId: Boolean(configuredId),
      invalidMaxWords: hasMaxWords && !(Number.isFinite(maxWordsValue) && maxWordsValue > 0),
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

function getQuestionConfigurationProblems_(questions) {
  const problems = [];
  const ids = new Set();
  const orders = new Set();
  const supportedTypes = new Set(['short_text', 'long_text', 'multiple_choice']);

  questions.forEach((question) => {
    if (question.hasExplicitId === false) {
      problems.push(`Question at order ${question.order} needs a QuestionID.`);
    }
    if (question.hasValidOrder === false) {
      problems.push(`Question ${question.id} needs a positive numeric Order.`);
    }
    if (ids.has(question.id)) problems.push(`Duplicate QuestionID: ${question.id}.`);
    ids.add(question.id);
    if (orders.has(question.order)) problems.push(`Duplicate question Order: ${question.order}.`);
    orders.add(question.order);
    if (!supportedTypes.has(question.answerType)) {
      problems.push(`Unsupported AnswerType for ${question.id}: ${question.answerType}.`);
    }
    if (question.answerType === 'multiple_choice') {
      if (question.options.length < 2) {
        problems.push(`Multiple-choice question ${question.id} needs at least two options.`);
      }
      if (new Set(question.options).size !== question.options.length) {
        problems.push(`Multiple-choice question ${question.id} contains duplicate options.`);
      }
    }
    if (question.invalidMaxWords) {
      problems.push(`MaxWords for ${question.id} must be a positive number or blank.`);
    }
  });
  return problems;
}

function saveSessionQuestions_(sessionId, questions) {
  const table = getTable_(APP.SHEETS.SESSION_QUESTIONS);
  const rows = questions.map((question, index) => {
    const row = Array(table.headers.length).fill('');
    const values = {
      SessionID: sessionId,
      QuestionIndex: index,
      Order: question.order,
      QuestionID: question.id,
      GroupID: safeForSheet_(question.groupId),
      QuestionText: safeForSheet_(question.text),
      AnswerType: question.answerType,
      Options: safeForSheet_(JSON.stringify(question.options)),
      Required: question.required,
      MaxWords: question.maxWords == null ? '' : question.maxWords,
    };
    Object.keys(values).forEach((header) => {
      if (!(header in table.map)) {
        throw new Error(
          `Required ${APP.SHEETS.SESSION_QUESTIONS} column is missing: ${header}. Run setupAssessment.`
        );
      }
      row[table.map[header]] = values[header];
    });
    return row;
  });

  if (rows.length) {
    table.sheet
      .getRange(table.sheet.getLastRow() + 1, 1, rows.length, table.headers.length)
      .setValues(rows);
  }
  cacheSessionQuestions_(sessionId, questions);
}

function backfillLegacySessionQuestions_() {
  const sessions = getTable_(APP.SHEETS.SESSIONS);
  const frozenQuestions = getTable_(APP.SHEETS.SESSION_QUESTIONS);
  const existingSessionIds = new Set(
    frozenQuestions.rows
      .map((row) => String(row[frozenQuestions.map.SessionID] || '').trim())
      .filter(Boolean)
  );
  const legacySessionIds = sessions.rows
    .filter((row) => String(row[sessions.map.Status] || '').trim().toUpperCase() === 'IN_PROGRESS')
    .map((row) => String(row[sessions.map.SessionID] || '').trim())
    .filter((sessionId) => sessionId && !existingSessionIds.has(sessionId));
  if (!legacySessionIds.length) return 0;

  const questions = getQuestions_();
  if (!questions.length) return 0;
  const rows = [];
  legacySessionIds.forEach((sessionId) => {
    questions.forEach((question, index) => {
      const row = Array(frozenQuestions.headers.length).fill('');
      const values = {
        SessionID: sessionId,
        QuestionIndex: index,
        Order: question.order,
        QuestionID: question.id,
        GroupID: safeForSheet_(question.groupId),
        QuestionText: safeForSheet_(question.text),
        AnswerType: question.answerType,
        Options: safeForSheet_(JSON.stringify(question.options)),
        Required: question.required,
        MaxWords: question.maxWords == null ? '' : question.maxWords,
      };
      Object.keys(values).forEach((header) => {
        row[frozenQuestions.map[header]] = values[header];
      });
      rows.push(row);
    });
    cacheSessionQuestions_(sessionId, questions);
  });

  frozenQuestions.sheet
    .getRange(
      frozenQuestions.sheet.getLastRow() + 1,
      1,
      rows.length,
      frozenQuestions.headers.length
    )
    .setValues(rows);
  return legacySessionIds.length;
}

function getQuestionsForSession_(session, fallbackQuestions) {
  const cachedQuestions = readCachedSessionQuestions_(session.sessionId);
  if (cachedQuestions) return cachedQuestions;

  const sheet = getSpreadsheet_().getSheetByName(APP.SHEETS.SESSION_QUESTIONS);
  if (!sheet) return fallbackQuestions || getQuestions_();

  const table = getTableStructure_(APP.SHEETS.SESSION_QUESTIONS);
  const lastRow = table.sheet.getLastRow();
  if (lastRow <= 1) return fallbackQuestions || getQuestions_();

  const sessionIdValues = table.sheet
    .getRange(2, table.map.SessionID + 1, lastRow - 1, 1)
    .getValues();
  const matchingRows = [];
  sessionIdValues.forEach((row, index) => {
    if (String(row[0] || '').trim() === session.sessionId) matchingRows.push(index + 2);
  });
  if (!matchingRows.length) return fallbackQuestions || getQuestions_();

  const firstRow = matchingRows[0];
  const finalRow = matchingRows[matchingRows.length - 1];
  const matchingRowSet = new Set(matchingRows);
  const questions = table.sheet
    .getRange(firstRow, 1, finalRow - firstRow + 1, table.headers.length)
    .getValues()
    .map((row, offset) => ({ row, sheetRow: firstRow + offset }))
    .filter((entry) => matchingRowSet.has(entry.sheetRow))
    .map((entry) => entry.row)
    .map((row) => {
      let options = [];
      try {
        const parsed = JSON.parse(String(row[table.map.Options] || '[]'));
        options = Array.isArray(parsed) ? parsed.map(String) : [];
      } catch (error) {
        options = String(row[table.map.Options] || '')
          .split('|')
          .map((option) => option.trim())
          .filter(Boolean);
      }
      const maxWordsValue = Number(row[table.map.MaxWords]);
      return {
        index: Number(row[table.map.QuestionIndex]) || 0,
        order: Number(row[table.map.Order]),
        id: String(row[table.map.QuestionID] || '').trim(),
        groupId: String(row[table.map.GroupID] || ''),
        text: String(row[table.map.QuestionText] || ''),
        answerType: String(row[table.map.AnswerType] || 'long_text'),
        options,
        required: toBoolean_(row[table.map.Required]),
        maxWords:
          Number.isFinite(maxWordsValue) && maxWordsValue > 0 ? maxWordsValue : null,
      };
    });

  questions.sort((a, b) => a.index - b.index);
  cacheSessionQuestions_(session.sessionId, questions);
  return questions;
}

function cacheSessionQuestions_(sessionId, questions) {
  try {
    CacheService.getScriptCache().put(
      `session-questions:${sessionId}`,
      JSON.stringify(questions),
      21600
    );
  } catch (error) {
    // Large question sets can exceed cache limits; the Sheet remains authoritative.
  }
}

function readCachedSessionQuestions_(sessionId) {
  try {
    const value = CacheService.getScriptCache().get(`session-questions:${sessionId}`);
    if (!value) return null;
    const questions = JSON.parse(value);
    return Array.isArray(questions) && questions.length ? questions : null;
  } catch (error) {
    return null;
  }
}

function findCandidateByToken_(token) {
  const table = getTableStructure_(APP.SHEETS.CANDIDATES);
  const lastRow = table.sheet.getLastRow();
  if (lastRow <= 1) return null;
  const matches = table.sheet
    .getRange(2, table.map.Token + 1, lastRow - 1, 1)
    .createTextFinder(token)
    .matchCase(true)
    .matchEntireCell(true)
    .findAll();
  if (!matches.length) return null;
  if (matches.length > 1) {
    throw new Error(
      `This token is assigned to more than one candidate (rows ${matches
        .map((match) => match.getRow())
        .join(' and ')}). Correct the Candidates sheet before continuing.`
    );
  }
  const rowNumber = matches[0].getRow();
  const row = table.sheet.getRange(rowNumber, 1, 1, table.headers.length).getValues()[0];
  return {
    rowNumber,
    name: String(row[table.map.CandidateName] || '').trim(),
    email: String(row[table.map.CandidateEmail] || '').trim(),
    active: toBoolean_(row[table.map.Active]),
    status: String(row[table.map.Status] || '').trim().toUpperCase(),
  };
}

function findDuplicateCandidateTokens_(table) {
  const rowsByToken = {};
  table.rows.forEach((row, index) => {
    const token = String(row[table.map.Token] || '').trim();
    if (!token) return;
    if (!rowsByToken[token]) rowsByToken[token] = [];
    rowsByToken[token].push(index + 2);
  });
  return Object.keys(rowsByToken)
    .filter((token) => rowsByToken[token].length > 1)
    .map((token) => ({ token, rows: rowsByToken[token] }));
}

function updateCandidateStatus_(rowNumber, status) {
  const table = getTableStructure_(APP.SHEETS.CANDIDATES);
  table.sheet.getRange(rowNumber, table.map.Status + 1).setValue(status);
}

function updateCandidateStatusByToken_(token, status) {
  const candidate = findCandidateByToken_(token);
  if (candidate) updateCandidateStatus_(candidate.rowNumber, status);
}

function findLatestSessionByToken_(token) {
  const table = getTableStructure_(APP.SHEETS.SESSIONS);
  const lastRow = table.sheet.getLastRow();
  if (lastRow <= 1) return null;
  const matches = table.sheet
    .getRange(2, table.map.Token + 1, lastRow - 1, 1)
    .createTextFinder(token)
    .matchCase(true)
    .matchEntireCell(true)
    .findAll();
  if (!matches.length) return null;
  const rowNumber = Math.max(...matches.map((match) => match.getRow()));
  const row = table.sheet.getRange(rowNumber, 1, 1, table.headers.length).getValues()[0];
  return sessionFromRow_(row, rowNumber, table.map);
}

function findSessionById_(sessionId) {
  const table = getTableStructure_(APP.SHEETS.SESSIONS);
  const lastRow = table.sheet.getLastRow();
  if (lastRow <= 1) return null;
  const match = table.sheet
    .getRange(2, table.map.SessionID + 1, lastRow - 1, 1)
    .createTextFinder(sessionId)
    .matchCase(true)
    .matchEntireCell(true)
    .findNext();
  if (!match) return null;
  const rowNumber = match.getRow();
  const row = table.sheet.getRange(rowNumber, 1, 1, table.headers.length).getValues()[0];
  return sessionFromRow_(row, rowNumber, table.map);
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
    sessionNonce: String(row[map.SessionNonce] || ''),
  };
}

function updateSession_(rowNumber, updates) {
  const table = getTableStructure_(APP.SHEETS.SESSIONS);
  Object.keys(updates).forEach((header) => {
    if (!(header in table.map)) throw new Error(`Unknown Sessions column: ${header}`);
  });
  setMappedValuesInRow_(table, rowNumber, updates);
}

function completeSession_(session, status, completedAt) {
  updateSession_(session.rowNumber, {
    LastActivityAt: completedAt,
    Status: status,
    CompletedAt: completedAt,
  });
}

function buildClientState_(session, questions, now, suppliedSettings) {
  const index = Number(session.currentQuestionIndex);
  const question = questions[index];
  const settings = suppliedSettings || getSettings_();
  const snapshotIntervalSeconds = Math.min(
    300,
    Math.max(5, Number(settings.SnapshotIntervalSeconds) || 15)
  );
  const maxSnapshotsPerQuestion = Math.min(
    1000,
    Math.max(50, Number(settings.MaxSnapshotsPerQuestion) || 250)
  );
  const draftRetentionHours = Math.min(
    168,
    Math.max(1, Number(settings.DraftRetentionHours) || 24)
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
    sessionNonce: session.sessionNonce || '',
    candidateName: session.candidateName,
    startedAtMs: session.startedAt.getTime(),
    deadlineAtMs: session.deadlineAt.getTime(),
    serverNowMs: now.getTime(),
    currentNumber: index + 1,
    totalQuestions: questions.length,
    questionShownAtMs: session.currentQuestionShownAt.getTime(),
    snapshotIntervalSeconds,
    maxSnapshotsPerQuestion,
    maxAnswerCharacters: getMaxAnswerCharacters_(settings),
    draftRetentionHours,
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

function getMaxAnswerCharacters_(settings) {
  return Math.min(49000, Math.max(1000, Number(settings.MaxAnswerCharacters) || 40000));
}

function isSafeHttpsImageUrl_(value) {
  const url = String(value || '').trim();
  return Boolean(url) &&
    url.length <= 2048 &&
    /^https:\/\/[^\s<>"']+$/i.test(url) &&
    !/^https:\/\/[^/]*@/i.test(url);
}

function sessionNonceMatches_(session, suppliedNonce) {
  // Sessions created before the nonce migration remain valid.
  if (!session.sessionNonce) return true;
  return secureEquals_(session.sessionNonce, String(suppliedNonce || ''));
}

function secureEquals_(left, right) {
  const first = String(left || '');
  const second = String(right || '');
  let difference = first.length ^ second.length;
  const length = Math.max(first.length, second.length);
  for (let index = 0; index < length; index += 1) {
    difference |= (first.charCodeAt(index) || 0) ^ (second.charCodeAt(index) || 0);
  }
  return difference === 0;
}

function isValidIdentifier_(value, maximumLength) {
  const text = String(value || '');
  return Boolean(text) && text.length <= maximumLength && !/[\u0000-\u001f\u007f]/.test(text);
}

/** Best-effort abuse protection. Cache failure never blocks a legitimate candidate. */
function isRateLimited_(operation, identifier, maximumRequests, windowSeconds) {
  try {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      `${operation}:${identifier}`,
      Utilities.Charset.UTF_8
    );
    const hash = digest
      .map((byte) => ((byte + 256) % 256).toString(16).padStart(2, '0'))
      .join('');
    const cache = CacheService.getScriptCache();
    const key = `rate:${operation}:${hash}`;
    const count = Number(cache.get(key) || 0) + 1;
    cache.put(key, String(count), windowSeconds);
    return count > maximumRequests;
  } catch (error) {
    return false;
  }
}

/** Administrative entry points are editor-only even though their names stay public. */
function assertAdministrator_() {
  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const configuredAdmins = String(
    PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || ''
  )
    .split(',')
    .map((email) => email.trim().toLowerCase())
    .filter(Boolean);
  const isAdministrator = Boolean(activeEmail) && (
    activeEmail === effectiveEmail || configuredAdmins.includes(activeEmail)
  );
  if (!isAdministrator) {
    throw new Error('This administrative function can only be run by an authorised editor.');
  }
}

function flushAndReleaseLock_(lock) {
  try {
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
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
