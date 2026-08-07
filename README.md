# AI Experience Repository

A searchable, AI-curated repository of staff and faculty AI experiences — prompts, workflows, and lessons learned with ChatGPT Edu, Google Gemini, NotebookLM, Claude, and Microsoft Copilot.

## How it works

```
Submission form (index.html)
        │  POST
        ▼
Google Apps Script web app ──▶ Google Sheet (row status: NEW)
        │
        ▼  every 15 min (trigger)
Curation agent (Claude API) ──▶ adds Tags + Summary + Category (status: TAGGED)
        │
        ▼  you flip status to APPROVED in the Sheet
Front end (index.html) ──▶ fetches APPROVED entries, searchable + filterable
```

The human-in-the-loop step is deliberate: nothing appears on the site until you change a row's Status to `APPROVED`. The AI agent does the tedious part (tagging, summarizing, categorizing); you keep editorial control.

## Setup (about 20 minutes)

### 1. Create the Google Sheet

1. Create a new Google Sheet. Rename the first tab to `Submissions`.
2. Add this header row in row 1:

   `Timestamp | Name | Email | Role | Department | Tool | Title | Story | Status | Tags | Summary | Category`

### 2. Add the Apps Script backend

1. In the Sheet: **Extensions → Apps Script**.
2. Delete the default code and paste in the contents of `backend/Code.gs`.
3. **Project Settings (gear icon) → Script Properties → Add script property**:
   - Property: `ANTHROPIC_API_KEY`
   - Value: your Anthropic API key (from platform.claude.com)
4. Run the `testCuration` function once (select it in the toolbar, click Run). Approve the permissions prompts. Check the execution log — you should see a JSON result with a summary, tags, and category.

### 3. Deploy the web app

1. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
2. Copy the web app URL.

### 4. Set the curation trigger

1. In Apps Script: **Triggers (clock icon) → Add Trigger**
2. Function: `curateNewSubmissions`, Event source: **Time-driven**, every **15 minutes**.

### 5. Connect and publish the front end

1. In `index.html`, replace `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` with your web app URL.
2. Host `index.html` anywhere static — GitHub Pages or Vercel (same flow as The Navigator). It is a single self-contained file with no build step.

## Day-to-day workflow

1. Someone submits an experience → new row appears with Status `NEW`.
2. Within 15 minutes the agent tags it → Status becomes `TAGGED`.
3. You skim the row, edit anything you like, and change Status to `APPROVED`.
4. It appears on the site immediately.

## Customizing

- **Tags**: edit `TAG_VOCABULARY` in `Code.gs`. The agent only uses tags from this list, which keeps the filter bar clean.
- **Categories**: edit `CATEGORIES` in `Code.gs` (the front end displays whatever the backend sends).
- **Tools list**: edit the two `<select>` lists in `index.html` (filter + form).
- **Re-tag a row**: set its Status back to `NEW` and the next trigger run will redo it.

## Notes

- Submissions never appear publicly without your approval — only `APPROVED` rows are returned by the API.
- The Claude API call costs a fraction of a cent per submission (model: `claude-opus-4-8`, ~1K tokens each way).
- If a curation call fails, the row stays `NEW` and is retried on the next trigger run; errors appear in Apps Script's execution log.
