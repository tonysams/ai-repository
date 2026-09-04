/**
 * AI Experience Repository — Google Apps Script backend
 *
 * Bound to a Google Sheet with a tab named "Submissions" and this header row:
 * Timestamp | Name | Email | Role | Department | Tool | Title | Story | Status | Tags | Summary | Category | Type | Embedding | Related
 *
 * Status lifecycle: NEW -> TAGGED (by the curation agent) -> APPROVED (by you) -> visible on the site
 *
 * Setup (see README.md):
 * 1. Script Properties: add ANTHROPIC_API_KEY (curation) and GEMINI_API_KEY (embeddings)
 * 2. Deploy > New deployment > Web app, execute as Me, access: Anyone
 * 3. Triggers: time-driven trigger on curateNewSubmissions, every 15 minutes
 * 4. (optional) time-driven trigger on refreshRelated, every 15 minutes, to keep the
 *    "Related experiences" links current as you approve entries. Or use the custom
 *    "AI Repository > Rebuild related links" menu button after an approval session.
 *
 * Semantic "Related experiences" (columns N/O):
 *   - N (Embedding): each entry's vector, cached as JSON so it is only computed once.
 *   - O (Related): JSON array of the top matches [{id,title,department,tool,score}].
 *   Anthropic has no embeddings API, so embeddings use Google's text-embedding-004
 *   (Gemini API) — a vendor the University already provides institutionally.
 */

var SHEET_NAME = 'Submissions';

// Where to send "new submission" notifications. Change to your uMail if you prefer.
// Leave as '' to turn notifications off.
var NOTIFY_EMAIL = 'twowheeltony@gmail.com';
var SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Frg1bD2fEYAzJ813ZoehSYMHjiV4srzMnGNzTakumVk/edit';

var COL = {
  TIMESTAMP: 1, NAME: 2, EMAIL: 3, ROLE: 4, DEPARTMENT: 5, TOOL: 6,
  TITLE: 7, STORY: 8, STATUS: 9, TAGS: 10, SUMMARY: 11, CATEGORY: 12, TYPE: 13,
  EMBEDDING: 14, RELATED: 15
};

// "Related experiences" tuning. Embeddings use Google (Gemini API) — Anthropic has
// no embeddings endpoint, and the University already provides Gemini institutionally,
// so it's a vendor the data can flow to without a new data-processing agreement.
var EMBED_MODEL = 'gemini-embedding-001'; // Google Gemini API (current GA embedding model)
var EMBED_TASK = 'SEMANTIC_SIMILARITY';   // symmetric entry-to-entry matching
var RELATED_TOP_K = 3;                     // how many related entries to show per card
var RELATED_MIN_SIM = 0.55;                // ignore weak matches below this cosine similarity (tune on real data)

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
      type: r[COL.TYPE - 1] || '',
      related: parseRelated(r[COL.RELATED - 1])
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

  // Notify — never let a mail failure break the submission.
  try { notifyNewSubmission(data); } catch (err) { console.error('Notify failed: ' + err); }

  return jsonResponse({ ok: true });
}

/** Emails NOTIFY_EMAIL when a new submission arrives. Does not change the review workflow. */
function notifyNewSubmission(data) {
  if (!NOTIFY_EMAIL) return;
  var type = ENTRY_TYPES.indexOf(data.type) !== -1 ? data.type : 'Experience story';
  var story = String(data.story || '');
  var snippet = story.length > 400 ? story.slice(0, 400) + '…' : story;
  var subject = 'AI Repository — new ' + type.toLowerCase() + ': ' + (data.title || '(untitled)');
  var body =
    'A new submission just arrived and is waiting in the queue (status NEW).\n\n' +
    'Title:      ' + (data.title || '(untitled)') + '\n' +
    'Type:       ' + type + '\n' +
    'Tool:       ' + (data.tool || '') + '\n' +
    'From:       ' + (data.name || 'Anonymous') +
      (data.role ? ' · ' + data.role : '') + (data.department ? ' · ' + data.department : '') + '\n' +
    'Email:      ' + (data.email || '(not provided)') + '\n\n' +
    'What they wrote:\n' + snippet + '\n\n' +
    '— It will be auto-tagged within ~15 min, then it is yours to review.\n' +
    'Open the Sheet to approve: ' + SHEET_URL + '\n';
  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run this in the editor to test email sending directly (select testNotify, click Run). */
function testNotify() {
  notifyNewSubmission({
    name: 'Editor Test',
    email: 'test@example.com',
    role: 'Staff',
    department: 'Digital Learning Technologies',
    tool: 'Claude',
    type: 'Experience story',
    title: 'Email test from the editor',
    story: 'Testing whether MailApp.sendEmail delivers to NOTIFY_EMAIL. If this arrives, notifications work.'
  });
  Logger.log('testNotify ran — check ' + NOTIFY_EMAIL + ' and the execution log.');
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

      // Compute + cache the embedding once, so it never needs re-embedding.
      // A failure here must not block tagging — related links are a nice-to-have.
      try {
        var vec = embedText(r[COL.TITLE - 1] + '\n' + result.summary + '\n' + r[COL.STORY - 1]);
        sheet.getRange(rowNum, COL.EMBEDDING).setValue(JSON.stringify(vec));
      } catch (embErr) {
        console.error('Embedding failed for row ' + rowNum + ': ' + embErr);
      }
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
    model: 'claude-opus-5',
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

// ---------------------------------------------------------------------------
// Related experiences — semantic "similar workflows across disciplines"
// ---------------------------------------------------------------------------

/** Embeds text with the Google (Gemini) API and returns the vector (array of floats). */
function embedText(text) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY is not set in Script Properties');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + EMBED_MODEL + ':embedContent';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      model: 'models/' + EMBED_MODEL,
      content: { parts: [{ text: String(text).slice(0, 8000) }] },
      taskType: EMBED_TASK
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) throw new Error('Embedding API error ' + code + ': ' + response.getContentText());
  var body = JSON.parse(response.getContentText());
  if (!body.embedding || !body.embedding.values) throw new Error('Embedding API: unexpected response ' + response.getContentText().slice(0, 200));
  return body.embedding.values;
}

/** Diagnostic: logs which models your GEMINI_API_KEY can use for embeddings.
 *  Run this if embedText() 404s, then set EMBED_MODEL to one of the names it prints. */
function listEmbeddingModels() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) { Logger.log('GEMINI_API_KEY is not set in Script Properties'); return; }
  var response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models', {
    method: 'get', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) { Logger.log('ListModels error: ' + response.getContentText()); return; }
  var models = (JSON.parse(response.getContentText()).models) || [];
  var found = 0;
  models.forEach(function (m) {
    if ((m.supportedGenerationMethods || []).indexOf('embedContent') !== -1) {
      Logger.log('EMBEDS: ' + m.name);   // e.g. "models/gemini-embedding-001"
      found++;
    }
  });
  Logger.log('--- ' + found + ' embedding model(s). Use one above (drop the "models/" prefix) for EMBED_MODEL. ---');
}

/** Cosine similarity between two equal-length vectors. */
function cosineSim(a, b) {
  var dot = 0, na = 0, nb = 0;
  for (var i = 0; i < a.length; i++) { dot += a[i] * b[i]; na += a[i] * a[i]; nb += b[i] * b[i]; }
  return (na && nb) ? dot / (Math.sqrt(na) * Math.sqrt(nb)) : 0;
}

/** Safely parse the Related cell (JSON string) into an array. */
function parseRelated(cell) {
  if (!cell) return [];
  try { var v = JSON.parse(cell); return Array.isArray(v) ? v : []; } catch (e) { return []; }
}

/**
 * Recomputes the top-K related entries for every APPROVED row that has an
 * embedding, and writes the result into column O. Run on a trigger or via the
 * "AI Repository" menu after approving a batch. Cheap: pure math, no API calls.
 */
function refreshRelated() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();

  // Collect approved rows that have a cached embedding.
  var items = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[COL.STATUS - 1]).toUpperCase() !== 'APPROVED') continue;
    var raw = r[COL.EMBEDDING - 1];
    if (!raw) continue;
    var vec;
    try { vec = JSON.parse(raw); } catch (e) { continue; }
    items.push({
      rowNum: i + 1,
      id: i + 1,
      title: r[COL.TITLE - 1],
      department: r[COL.DEPARTMENT - 1],
      tool: r[COL.TOOL - 1],
      vec: vec
    });
  }

  // For each item, score against every other and keep the top K.
  for (var a = 0; a < items.length; a++) {
    var scores = [];
    for (var b = 0; b < items.length; b++) {
      if (a === b) continue;
      var sim = cosineSim(items[a].vec, items[b].vec);
      if (sim >= RELATED_MIN_SIM) {
        scores.push({
          id: items[b].id,
          title: items[b].title,
          department: items[b].department,
          tool: items[b].tool,
          score: Math.round(sim * 100) / 100
        });
      }
    }
    scores.sort(function (x, y) { return y.score - x.score; });
    sheet.getRange(items[a].rowNum, COL.RELATED).setValue(JSON.stringify(scores.slice(0, RELATED_TOP_K)));
  }

  console.log('refreshRelated: updated ' + items.length + ' approved entries.');
}

/** Adds a custom menu so you can rebuild related links from the Sheet UI. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Repository')
    .addItem('Rebuild related links', 'refreshRelated')
    .addItem('Curate new submissions now', 'curateNewSubmissions')
    .addToUi();
}

/** One-off: embed any APPROVED/TAGGED rows that are missing an embedding, then refresh related. */
function backfillEmbeddings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  var done = 0;
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var status = String(r[COL.STATUS - 1]).toUpperCase();
    if (status !== 'APPROVED' && status !== 'TAGGED') continue;
    if (r[COL.EMBEDDING - 1]) continue; // already embedded
    try {
      var vec = embedText(r[COL.TITLE - 1] + '\n' + r[COL.SUMMARY - 1] + '\n' + r[COL.STORY - 1]);
      sheet.getRange(i + 1, COL.EMBEDDING).setValue(JSON.stringify(vec));
      done++;
    } catch (err) {
      console.error('Backfill embedding failed for row ' + (i + 1) + ': ' + err);
    }
  }
  console.log('backfillEmbeddings: embedded ' + done + ' rows.');
  refreshRelated();
}
