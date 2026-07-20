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
  try {
    const ss = getSpreadsheet_();
    const table = getTable_(APP.SHEETS.QUESTIONS);
    const mergedRows = table.rows.map((row) => row.slice());
    const rowIndexByQuestionId = {};
    const importOrderBase = mergedRows.reduce((maximum, row) => {
      const order = Number(row[table.map.Order]);
      return Number.isFinite(order) ? Math.max(maximum, order) : maximum;
    }, 0);

    mergedRows.forEach((row, index) => {
      const questionId = String(row[table.map.QuestionID] || '').trim();
      if (!questionId) return;
      if (questionId in rowIndexByQuestionId) {
        throw new Error(
          `Questions contains duplicate QuestionID ${questionId}. Correct it before importing.`
        );
      }
      rowIndexByQuestionId[questionId] = index;
    });

    rows.forEach((valuesByHeader) => {
      valuesByHeader.Order = importOrderBase + valuesByHeader.Order;
      const questionId = valuesByHeader.QuestionID;
      const existingIndex = rowIndexByQuestionId[questionId];
      const row = existingIndex == null
        ? Array(table.headers.length).fill('')
        : mergedRows[existingIndex].slice();

      Object.keys(valuesByHeader).forEach((header) => {
        if (!(header in table.map)) {
          throw new Error(`Questions is missing ${header}. Run setupAssessment.`);
        }
        const value = valuesByHeader[header];
        if (existingIndex != null && header === 'MaxWords' && value === '') return;
        row[table.map[header]] = typeof value === 'string' ? safeForSheet_(value) : value;
      });

      if (existingIndex == null) {
        rowIndexByQuestionId[questionId] = mergedRows.length;
        mergedRows.push(row);
        addedCount += 1;
      } else {
        mergedRows[existingIndex] = row;
        updatedCount += 1;
      }
    });

    table.sheet
      .getRange(2, 1, mergedRows.length, table.headers.length)
      .setValues(mergedRows);
    table.sheet.autoResizeColumns(1, table.headers.length);

    const logSheetName = 'ImportLog';
    const logSheet = ss.getSheetByName(logSheetName) || ss.insertSheet(logSheetName);
    const logStartRow = logSheet.getLastRow() > 0 ? logSheet.getLastRow() + 2 : 1;
    logSheet
      .getRange(logStartRow, 1, logRows.length, logRows[0].length)
      .setValues(logRows);
    logSheet.setFrozenRows(1);
    logSheet.autoResizeColumns(1, logRows[0].length);

    updateSettingValue_('AssessmentTitle', form.getTitle());
  } finally {
    lock.releaseLock();
  }

  return [
    `Imported ${rows.length} questions from “${form.getTitle()}”.`,
    `Added ${addedCount} questions and updated ${updatedCount} existing imported questions.`,
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
  for (let index = 0; index < table.rows.length; index += 1) {
    if (String(table.rows[index][table.map.Key] || '').trim() === key) {
      table.sheet.getRange(index + 2, table.map.Value + 1).setValue(value);
      return;
    }
  }
  table.sheet.appendRow([key, value]);
}
