/**
 * AI Experience Repository — Google Apps Script backend
 *
 * Bound to a Google Sheet with a tab named "Submissions" and this header row:
 * Timestamp | Name | Email | Role | Department | Tool | Title | Story | Status | Tags | Summary | Category | Type
 *
 * Status lifecycle: NEW -> TAGGED (by the curation agent) -> APPROVED (by you) -> visible on the site
 *
 * Setup (see README.md):
 * 1. Script Properties: add ANTHROPIC_API_KEY
 * 2. Deploy > New deployment > Web app, execute as Me, access: Anyone
 * 3. Triggers: time-driven trigger on curateNewSubmissions, every 15 minutes
 */

var SHEET_NAME = 'Submissions';

var COL = {
  TIMESTAMP: 1, NAME: 2, EMAIL: 3, ROLE: 4, DEPARTMENT: 5, TOOL: 6,
  TITLE: 7, STORY: 8, STATUS: 9, TAGS: 10, SUMMARY: 11, CATEGORY: 12, TYPE: 13
};

var ENTRY_TYPES = ['Prompt template', 'Experience story'];

// Controlled vocabulary the agent tags against. Edit freely — the agent
// will only use tags from this list, which keeps the filter UI clean.
var TAG_VOCABULARY = [
  'teaching', 'course-design', 'assessment', 'research', 'writing',
  'data-analysis', 'literature-review', 'grading-feedback', 'lesson-planning',
  'email-drafting', 'meeting-notes', 'brainstorming', 'summarization',
  'translation', 'coding', 'presentations', 'accessibility', 'advising',
  'administrative', 'policy', 'student-support', 'prompt-engineering'
];

var CATEGORIES = ['Teaching', 'Research', 'Administration', 'Student Support', 'Professional Development'];

// ---------------------------------------------------------------------------
// Web app endpoints
// ---------------------------------------------------------------------------

/** GET — returns approved entries as JSON for the front end. */
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  var entries = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[COL.STATUS - 1]).toUpperCase() !== 'APPROVED') continue;
    entries.push({
      id: i + 1,
      date: r[COL.TIMESTAMP - 1],
      name: r[COL.NAME - 1],
      role: r[COL.ROLE - 1],
      department: r[COL.DEPARTMENT - 1],
      tool: r[COL.TOOL - 1],
      title: r[COL.TITLE - 1],
      story: r[COL.STORY - 1],
      tags: String(r[COL.TAGS - 1]).split(',').map(function (t) { return t.trim(); }).filter(String),
      summary: r[COL.SUMMARY - 1],
      category: r[COL.CATEGORY - 1],
      type: r[COL.TYPE - 1] || ''
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ entries: entries }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST — accepts a new submission from the front end form. */
function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Invalid JSON' });
  }

  if (!data.story || !data.tool) {
    return jsonResponse({ ok: false, error: 'Missing required fields (tool, story)' });
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([
    new Date(),
    data.name || 'Anonymous',
    data.email || '',
    data.role || '',
    data.department || '',
    data.tool,
    data.title || '',
    data.story,
    'NEW', '', '', '',
    ENTRY_TYPES.indexOf(data.type) !== -1 ? data.type : 'Experience story'
  ]);

  return jsonResponse({ ok: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Curation agent — run on a time-driven trigger (e.g. every 15 minutes)
// ---------------------------------------------------------------------------

/** Finds NEW submissions, asks Claude to tag/summarize/categorize them, marks them TAGGED. */
function curateNewSubmissions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[COL.STATUS - 1]).toUpperCase() !== 'NEW') continue;

    try {
      var result = curateWithClaude({
        tool: r[COL.TOOL - 1],
        role: r[COL.ROLE - 1],
        title: r[COL.TITLE - 1],
        story: r[COL.STORY - 1]
      });

      var rowNum = i + 1;
      sheet.getRange(rowNum, COL.TAGS).setValue(result.tags.join(', '));
      sheet.getRange(rowNum, COL.SUMMARY).setValue(result.summary);
      sheet.getRange(rowNum, COL.CATEGORY).setValue(result.category);
      sheet.getRange(rowNum, COL.STATUS).setValue('TAGGED');
    } catch (err) {
      // Leave the row as NEW so the next run retries; log for inspection.
      console.error('Curation failed for row ' + (i + 1) + ': ' + err);
    }
  }
}

/** Calls the Claude API and returns {summary, tags, category}. */
function curateWithClaude(submission) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY is not set in Script Properties');

  var payload = {
    model: 'claude-opus-4-8',
    max_tokens: 1024,
    system: 'You curate a university repository of staff and faculty AI experiences. ' +
      'For each submission, write a one-to-two sentence summary in plain language, ' +
      'choose the most relevant tags ONLY from the provided vocabulary (2 to 5 tags), ' +
      'and pick exactly one category from the provided list.',
    messages: [{
      role: 'user',
      content: 'Tag vocabulary: ' + TAG_VOCABULARY.join(', ') + '\n' +
        'Categories: ' + CATEGORIES.join(', ') + '\n\n' +
        'Submission:\n' +
        'Tool used: ' + submission.tool + '\n' +
        'Submitter role: ' + submission.role + '\n' +
        'Title: ' + submission.title + '\n' +
        'Story: ' + submission.story
    }],
    output_config: {
      format: {
        type: 'json_schema',
        schema: {
          type: 'object',
          properties: {
            summary: { type: 'string' },
            tags: { type: 'array', items: { type: 'string' } },
            category: { type: 'string', enum: CATEGORIES }
          },
          required: ['summary', 'tags', 'category'],
          additionalProperties: false
        }
      }
    }
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Claude API error ' + code + ': ' + response.getContentText());
  }

  var body = JSON.parse(response.getContentText());
  var text = '';
  for (var i = 0; i < body.content.length; i++) {
    if (body.content[i].type === 'text') { text = body.content[i].text; break; }
  }

  var result = JSON.parse(text);
  // Keep only tags that are actually in the vocabulary.
  result.tags = result.tags.filter(function (t) { return TAG_VOCABULARY.indexOf(t) !== -1; });
  return result;
}

/** Run once manually to verify the API key and connection work. */
function testCuration() {
  var result = curateWithClaude({
    tool: 'Claude',
    role: 'Faculty',
    title: 'Faster literature reviews',
    story: 'I upload PDFs of journal articles and ask Claude to extract the methodology and key findings into a comparison table. It cut my lit review prep time roughly in half.'
  });
  console.log(JSON.stringify(result, null, 2));
}
