/**
 * Imports compatible questions from an existing Google Form into the
 * Questions sheet used by the Candidate Assessment Timer.
 *
 * Before running:
 * 1. Run setupAssessment so SourceFormUrl exists in Settings.
 * 2. Paste the full EDIT URL of the Google Form in the Value column.
 * 3. Run importQuestionsFromExistingForm from the Apps Script editor.
 *
 * This updates previously imported GF_ questions by QuestionID and appends new
 * ones. It preserves unrelated Questions rows, prior import logs, the source
 * Google Form, and its existing responses.
 */
function importQuestionsFromExistingForm() {
  assertAdministrator_();
  const settings = getSettings_();
  const sourceFormUrl = String(settings.SourceFormUrl || '').trim();

  if (!/^https:\/\/docs\.google\.com\/forms\//i.test(sourceFormUrl)) {
    throw new Error(
      'Add SourceFormUrl to the Settings sheet and paste the full Google Form edit URL.'
    );
  }

  const form = FormApp.openByUrl(sourceFormUrl);
  const items = form.getItems();
  const rows = [];
  const logRows = [
    ['FormItemIndex', 'FormItemID', 'FormItemType', 'Title', 'ImportResult', 'Notes'],
  ];

  let sectionNumber = 1;
  let groupId = 'Section 1';
  let questionOrder = 0;

  items.forEach((item, itemIndex) => {
    const type = item.getType();
    const typeName = String(type);
    const title = String(item.getTitle() || '').trim();

    if (type === FormApp.ItemType.PAGE_BREAK) {
      sectionNumber += 1;
      const page = item.asPageBreakItem();
      groupId = String(page.getTitle() || `Section ${sectionNumber}`).trim();
      logRows.push([
        itemIndex + 1,
        item.getId(),
        typeName,
        title,
        'Used as GroupID',
        groupId,
      ]);
      return;
    }

    if (type === FormApp.ItemType.SECTION_HEADER) {
      const header = item.asSectionHeaderItem();
      const headerTitle = String(header.getTitle() || '').trim();
      if (headerTitle) groupId = headerTitle;
      logRows.push([
        itemIndex + 1,
        item.getId(),
        typeName,
        title,
        'Used as GroupID',
        groupId,
      ]);
      return;
    }

    let imported = null;
    let note = '';

    switch (type) {
      case FormApp.ItemType.TEXT: {
        const source = item.asTextItem();
        imported = makeImportedQuestion_(source, 'short_text', []);
        break;
      }

      case FormApp.ItemType.PARAGRAPH_TEXT: {
        const source = item.asParagraphTextItem();
        imported = makeImportedQuestion_(source, 'long_text', []);
        break;
      }

      case FormApp.ItemType.MULTIPLE_CHOICE: {
        const source = item.asMultipleChoiceItem();
        const options = source.getChoices().map((choice) => choice.getValue());
        imported = makeImportedQuestion_(source, 'multiple_choice', options);
        if (source.hasOtherOption()) {
          note = 'The original question has an Other option. Add it manually if required.';
        }
        break;
      }

      case FormApp.ItemType.LIST: {
        const source = item.asListItem();
        const options = source.getChoices().map((choice) => choice.getValue());
        imported = makeImportedQuestion_(source, 'multiple_choice', options);
        note = 'Imported as radio-button multiple choice rather than a dropdown.';
        break;
      }

      case FormApp.ItemType.SCALE: {
        const source = item.asScaleItem();
        const options = [];
        for (let value = source.getLowerBound(); value <= source.getUpperBound(); value += 1) {
          options.push(String(value));
        }
        imported = makeImportedQuestion_(source, 'multiple_choice', options);
        const labels = [source.getLeftLabel(), source.getRightLabel()]
          .map((value) => String(value || '').trim())
          .filter(Boolean);
        if (labels.length) {
          imported.text += `\n\nScale labels: ${labels.join(' — ')}`;
        }
        note = 'Imported as numbered multiple choice.';
        break;
      }

      default:
        note = unsupportedTypeNote_(typeName);
    }

    if (!imported) {
      logRows.push([
        itemIndex + 1,
        item.getId(),
        typeName,
        title,
        'Not imported',
        note,
      ]);
      return;
    }

    questionOrder += 1;
    const questionId = `GF_${item.getId()}`;
    rows.push({
      Order: questionOrder,
      QuestionID: questionId,
      GroupID: groupId,
      QuestionText: imported.text,
      AnswerType: imported.answerType,
      Options: imported.options.map(cleanImportedOption_).join('|'),
      Required: imported.required,
      MaxWords: '',
      Active: true,
    });

    logRows.push([
      itemIndex + 1,
      item.getId(),
      typeName,
      title,
      `Imported as ${imported.answerType}`,
      note,
    ]);
  });

  if (!rows.length) {
    throw new Error('No compatible questions were found in the source Google Form.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  let addedCount = 0;
  let updatedCount = 0;
  let archivedCount = 0;
  try {
    const ss = getSpreadsheet_();
    const table = getTable_(APP.SHEETS.QUESTIONS);
    const rowIndexByQuestionId = {};
    let nextOrder = table.rows.reduce((maximum, row) => {
      const order = Number(row[table.map.Order]);
      return Number.isFinite(order) ? Math.max(maximum, order) : maximum;
    }, 0);

    table.rows.forEach((row, index) => {
      const questionId = String(row[table.map.QuestionID] || '').trim();
      if (!questionId) return;
      if (questionId in rowIndexByQuestionId) {
        throw new Error(
          `Questions contains duplicate QuestionID ${questionId}. Correct it before importing.`
        );
      }
      rowIndexByQuestionId[questionId] = index;
    });

    const importedQuestionIds = new Set(rows.map((row) => row.QuestionID));
    rows.forEach((valuesByHeader) => {
      const questionId = valuesByHeader.QuestionID;
      const existingIndex = rowIndexByQuestionId[questionId];
      const updates = {};
      Object.keys(valuesByHeader).forEach((header) => {
        if (!(header in table.map)) {
          throw new Error(`Questions is missing ${header}. Run setupAssessment.`);
        }
        const value = valuesByHeader[header];
        if (existingIndex != null && header === 'MaxWords' && value === '') return;
        if (existingIndex != null && header === 'Order') return;
        updates[header] = typeof value === 'string' ? safeForSheet_(value) : value;
      });

      if (existingIndex == null) {
        nextOrder += 1;
        updates.Order = nextOrder;
        appendMappedRowToSheet_(table.sheet, updates);
        addedCount += 1;
      } else {
        const existingOrder = Number(table.rows[existingIndex][table.map.Order]);
        if (!(Number.isFinite(existingOrder) && existingOrder > 0)) {
          nextOrder += 1;
          updates.Order = nextOrder;
        }
        setMappedValuesInRow_(table, existingIndex + 2, updates);
        updatedCount += 1;
      }
    });

    // Keep historical rows but remove source questions that no longer exist
    // from future assessments. Existing sessions retain their frozen questions.
    table.rows.forEach((row, index) => {
      const questionId = String(row[table.map.QuestionID] || '').trim();
      if (!questionId.startsWith('GF_') || importedQuestionIds.has(questionId)) return;
      if (String(row[table.map.Active] || '').trim().toLowerCase() === 'false') return;
      setMappedValuesInRow_(table, index + 2, { Active: false });
      archivedCount += 1;
    });

    const logSheetName = 'ImportLog';
    const logSheet = ss.getSheetByName(logSheetName) || ss.insertSheet(logSheetName);
    const logStartRow = logSheet.getLastRow() > 0 ? logSheet.getLastRow() + 2 : 1;
    const safeLogRows = logRows.map((row) => row.map((value) =>
      typeof value === 'string' ? safeForSheet_(value) : value
    ));
    logSheet
      .getRange(logStartRow, 1, safeLogRows.length, safeLogRows[0].length)
      .setValues(safeLogRows);
    logSheet.setFrozenRows(1);
    logSheet.autoResizeColumns(1, logRows[0].length);

    if (!String(settings.AssessmentTitle || '').trim()) {
      updateSettingValue_('AssessmentTitle', form.getTitle());
    }
  } finally {
    flushAndReleaseLock_(lock);
  }

  return [
    `Imported ${rows.length} questions from “${form.getTitle()}”.`,
    `Added ${addedCount} questions and updated ${updatedCount} existing imported questions.`,
    `Archived ${archivedCount} previously imported questions that are no longer in the Form.`,
    `Reviewed ${items.length} total form items. Existing unrelated questions and prior ImportLog entries were preserved.`,
    'Open the ImportLog sheet to review skipped or converted items.',
  ].join('\n');
}

function makeImportedQuestion_(source, answerType, options) {
  const title = String(source.getTitle() || '').trim();
  const helpText = String(source.getHelpText() || '').trim();
  return {
    text: helpText ? `${title}\n\n${helpText}` : title,
    answerType,
    options,
    required: source.isRequired(),
  };
}

function cleanImportedOption_(value) {
  // The assessment timer stores options separated by |, so replace literal |
  // characters inside an option to avoid accidentally splitting the choice.
  return String(value == null ? '' : value).replace(/\|/g, '¦').trim();
}

function unsupportedTypeNote_(typeName) {
  const notes = {
    CHECKBOX: 'Checkbox questions require multi-select support in Index.html and Code.js.',
    CHECKBOX_GRID: 'Checkbox grids require a custom grid/multi-select interface.',
    GRID: 'Multiple-choice grids require a custom grid interface.',
    DATE: 'Date validation is not supported by the current timer page.',
    DATETIME: 'Date-and-time validation is not supported by the current timer page.',
    DURATION: 'Duration validation is not supported by the current timer page.',
    TIME: 'Time validation is not supported by the current timer page.',
    IMAGE: 'Layout images are not copied. Add them manually to the web app if needed.',
    VIDEO: 'Videos are not copied.',
    RATING: 'Rating questions require an additional renderer in Index.html.',
  };
  return notes[typeName] || 'This Google Forms item type is not supported by the current timer page.';
}

function updateSettingValue_(key, value) {
  const table = getTable_(APP.SHEETS.SETTINGS);
  const storedValue = typeof value === 'string' ? safeForSheet_(value) : value;
  for (let index = 0; index < table.rows.length; index += 1) {
    if (String(table.rows[index][table.map.Key] || '').trim() === key) {
      table.sheet.getRange(index + 2, table.map.Value + 1).setValue(storedValue);
      return;
    }
  }
  appendMappedRowToSheet_(table.sheet, { Key: key, Value: storedValue });
}
