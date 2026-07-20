/**
 * Imports compatible questions from an existing Google Form into the
 * Questions sheet used by the Candidate Assessment Timer.
 *
 * Before running:
 * 1. Run setupAssessment so SourceFormUrl exists in Settings.
 * 2. Paste the full EDIT URL of the Google Form in the Value column.
 * 3. Run importQuestionsFromExistingForm from the Apps Script editor.
 *
 * This replaces the current rows in the Questions sheet. It does not modify
 * the source Google Form or its existing responses.
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
    rows.push([
      questionOrder,
      questionId,
      groupId,
      imported.text,
      imported.answerType,
      imported.options.map(cleanImportedOption_).join('|'),
      imported.required,
      '',
    ]);

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

  const ss = getSpreadsheet_();
  const questionsSheet = ss.getSheetByName(APP.SHEETS.QUESTIONS);
  if (!questionsSheet) {
    throw new Error('Questions sheet not found. Run setupAssessment first.');
  }

  if (questionsSheet.getLastRow() > 1) {
    questionsSheet
      .getRange(2, 1, questionsSheet.getLastRow() - 1, questionsSheet.getLastColumn())
      .clearContent();
  }

  questionsSheet
    .getRange(2, 1, rows.length, HEADERS.QUESTIONS.length)
    .setValues(rows);
  questionsSheet.autoResizeColumns(1, HEADERS.QUESTIONS.length);

  const logSheetName = 'ImportLog';
  const oldLog = ss.getSheetByName(logSheetName);
  if (oldLog) ss.deleteSheet(oldLog);
  const logSheet = ss.insertSheet(logSheetName);
  logSheet.getRange(1, 1, logRows.length, logRows[0].length).setValues(logRows);
  logSheet.setFrozenRows(1);
  logSheet.autoResizeColumns(1, logRows[0].length);

  updateSettingValue_('AssessmentTitle', form.getTitle());

  return [
    `Imported ${rows.length} questions from “${form.getTitle()}”.`,
    `Reviewed ${items.length} total form items.`,
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
    CHECKBOX: 'Checkbox questions require multi-select support in Index.html and Code.gs.',
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
