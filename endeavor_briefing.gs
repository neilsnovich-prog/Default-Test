// ============================================================
// ENDEAVOR HEALTH SYSTEM — WEEKLY LITERATURE SCAN
// Google Apps Script  |  Splits research into parallel mini-calls
//                         to minimise tokens per API request
// ============================================================
//
// SETUP (one-time):
//   1. Go to script.google.com → New project → paste this file
//   2. Set the four constants below (API key, email addresses)
//   3. Run setupTrigger() once from the menu to schedule Monday 6 AM CST
//   4. Grant the requested OAuth scopes when prompted
//
// HOW IT WORKS:
//   Each TOPIC is sent to Claude in its own small API call (~1-2k tokens).
//   A final "assembler" call merges all mini-results into one formatted
//   HTML email — no single call ever carries the full briefing context.
// ============================================================

// ── USER CONFIGURATION ──────────────────────────────────────
var ANTHROPIC_API_KEY = 'sk-ant-YOUR_KEY_HERE';   // Anthropic API key
var TO_EMAIL_1        = 'neilsnovich@gmail.com';
var TO_EMAIL_2        = 'nndngrp@ameritech.net';
var FROM_ALIAS        = 'shadowext9@gmail.com';    // must be a Gmail send-as alias
var MODEL             = 'claude-sonnet-4-6';
// ────────────────────────────────────────────────────────────

var HOSPITALS = [
  'Endeavor Health',
  'Evanston Hospital',
  'Glenbrook Hospital',
  'Skokie Hospital',
  'Highland Park Hospital',
  'Lake Forest Hospital'
];

var HOSPITAL_LIST = HOSPITALS.join(', ');

var TOPICS = [
  { key: 'AI',          label: 'AI in Hospitals & Health Systems' },
  { key: 'CYBER',       label: 'Cybersecurity & Data Breaches' },
  { key: 'MA',          label: 'Hospital M&A & Strategic Consolidation' },
  { key: 'FINANCE',     label: 'Hospital Finances, Medicaid & Federal Policy' },
  { key: 'MEDS',        label: 'New Medications, Treatments & GLP-1s' },
  { key: 'CLINICAL',    label: 'Clinical Trials & FDA Regulatory Shifts' },
  { key: 'HEALTHIT',    label: 'Health IT / EHR & Digital Infrastructure' },
  { key: 'SAFETY',      label: 'Hospital Safety, Workforce Violence & Nursing Shortage' },
  { key: 'ADMIN',       label: 'Hospital Administration & Organizational Structure' },
  { key: 'COMMUNITY',   label: 'Community Relations & Health Equity' },
  { key: 'ENDEAVOR',    label: 'Endeavor Health System Deep Dive' }
];

// ── ENTRY POINT (trigger calls this) ────────────────────────
function runWeeklyBriefing() {
  var runDate = Utilities.formatDate(new Date(), 'America/Chicago', 'MMMM d, yyyy');
  var results = {};

  // Step 1 — research each topic in a separate mini-call
  for (var i = 0; i < TOPICS.length; i++) {
    var t = TOPICS[i];
    Logger.log('Researching: ' + t.label);
    results[t.key] = researchTopic(t, runDate);
    Utilities.sleep(1500); // stay within API rate limits
  }

  // Step 2 — assemble the final HTML briefing
  Logger.log('Assembling final briefing...');
  var html = assembleBriefing(results, runDate);

  // Step 3 — send email
  sendBriefing(html, runDate);
  Logger.log('Done — briefing sent for ' + runDate);
}

// ── MINI-CALL: one topic at a time ──────────────────────────
function researchTopic(topic, runDate) {
  var isEndeavor = (topic.key === 'ENDEAVOR');

  var systemPrompt = isEndeavor
    ? 'You are a healthcare intelligence analyst. Search the web for the most recent news ' +
      '(PAST 7 DAYS ONLY — from ' + runDate + ' back 7 days) about: ' + HOSPITAL_LIST + '. ' +
      'Return a JSON object with this exact shape: ' +
      '{"stories":[{"headline":"...","summary":"...two sentences...","source":"pub name","url":"...","hospital":"which hospital"}],' +
      '"deepDive":[{"area":"...","findings":"...3-5 sentences...","relevance":"...2-3 sentences...","source":"...","url":"..."}]}. ' +
      'deepDive areas: System Overview, Evanston Hospital, Glenbrook Hospital, Skokie Hospital, ' +
      'Highland Park Hospital, Lake Forest Hospital, Leadership (Gabrielle Cummings-Johnson), ' +
      'AI & Innovation, Clinical Trials & Research, Community Investment. ' +
      'Include ONLY stories that explicitly mention one of these entities: ' + HOSPITAL_LIST + '. ' +
      'Return ONLY valid JSON, no markdown, no explanation.'
    : 'You are a healthcare intelligence analyst. Search the web for the most recent news ' +
      '(PAST 7 DAYS ONLY — from ' + runDate + ' back 7 days) about the topic: "' + topic.label + '". ' +
      'Also look for any stories that mention ' + HOSPITAL_LIST + ' in relation to this topic. ' +
      'Return a JSON object with this exact shape: ' +
      '{"briefing":{"whatIsHappening":"...3-5 sentences with data points...","whyItMatters":"...2-3 sentences...","trend":"one-line verdict","source":"pub name","url":"..."},' +
      '"stories":[{"headline":"...","summary":"...two sentences...","source":"pub name","url":"...","hospital":"which hospital or null"}]}. ' +
      '"stories" must have AT LEAST 4 items (more is better). ' +
      'Return ONLY valid JSON, no markdown, no explanation.';

  var userMsg = isEndeavor
    ? 'Research all six Endeavor Health System hospitals for the past 7 days ending ' + runDate + '.'
    : 'Research "' + topic.label + '" healthcare news for the past 7 days ending ' + runDate + '.';

  try {
    var raw = callClaude(systemPrompt, userMsg, 1500);
    return JSON.parse(raw);
  } catch(e) {
    Logger.log('Parse error for ' + topic.key + ': ' + e + '\nRaw: ' + raw);
    return { error: 'Research unavailable for this topic.' };
  }
}

// ── ASSEMBLER CALL ───────────────────────────────────────────
function assembleBriefing(results, runDate) {
  // Build a compact summary for the assembler so it stays small
  var summary = JSON.stringify(results, null, 0);

  var systemPrompt =
    'You are a senior healthcare analyst. Convert the supplied JSON research data into a ' +
    'polished HTML executive briefing. Use inline CSS only. ' +
    'Color scheme: navy headers (#1F4E79 white text), alternating white/#EBF0F8 rows, Arial font. ' +
    'Structure EXACTLY as follows — no other sections:\n\n' +

    'HEADER: Title "Endeavor Health System — Healthcare Executive Briefing", subtitle ' +
    '"Weekly Literature Scan", date "' + runDate + '" in large navy bold text.\n\n' +

    'SECTION 1 — TOPIC BRIEFING (table)\n' +
    'Columns: Topic Area | What\'s Going On | Why It Matters | Trend | Key Source\n' +
    'One row per topic from the briefing data (exclude Endeavor deep dive here).\n\n' +

    'SECTION 2 — NEWS STORIES BY TOPIC\n' +
    'For each topic: bold topic heading, then a numbered list of stories.\n' +
    'Story format: [N]. **Headline** — two-sentence summary. Source: linked pub name + date.\n' +
    'Include ALL stories provided (aim for 4-6 per topic).\n' +
    'If no stories, write: "No qualifying stories found for this topic this week."\n\n' +

    'SECTION 3 — ENDEAVOR HEALTH DEEP DIVE (table)\n' +
    'Columns: Area | Key Findings | Relevance to Your Meeting | Source\n' +
    'Use the deepDive data from the ENDEAVOR key.\n\n' +

    'SECTION 4 — INDUSTRY TRENDS WITH ENDEAVOR COVERAGE (at the END)\n' +
    'Identify 6-8 overarching cross-cutting trends from all the research data.\n' +
    'For each trend:\n' +
    '  • Bold trend name as a subheading\n' +
    '  • 2-3 sentence description of the trend\n' +
    '  • Sub-list: "Endeavor & Hospital Articles on This Trend" — list every story from ' +
    '    any topic whose hospital field is not null (i.e., mentions one of the six Endeavor hospitals). ' +
    '    Aim for 3-6 bullet points per trend. Format: hospital name — headline — one sentence — source link.\n' +
    '  • If no hospital-specific articles found for a trend, write "No Endeavor-specific coverage this week."\n\n' +

    'FOOTER: Methodology note — sources, date range (past 7 days), filter rules.\n\n' +
    'Return ONLY the full HTML document starting with <!DOCTYPE html>.';

  var userMsg = 'Here is the research data:\n\n' + summary;

  try {
    return callClaude(systemPrompt, userMsg, 6000);
  } catch(e) {
    return '<html><body><h1>Assembly error</h1><pre>' + e + '</pre></body></html>';
  }
}

// ── CLAUDE API WRAPPER ───────────────────────────────────────
function callClaude(systemPrompt, userMessage, maxTokens) {
  var payload = {
    model: MODEL,
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [{ role: 'user', content: userMessage }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  var code = response.getResponseCode();
  var text = response.getContentText();

  if (code !== 200) {
    throw new Error('API error ' + code + ': ' + text);
  }

  var parsed = JSON.parse(text);
  return parsed.content[0].text;
}

// ── EMAIL SENDER ─────────────────────────────────────────────
function sendBriefing(html, runDate) {
  var subject = 'Endeavor Health Executive Briefing — ' + runDate;
  var recipients = TO_EMAIL_1 + ',' + TO_EMAIL_2;

  GmailApp.sendEmail(recipients, subject, '', {
    htmlBody: html,
    from: FROM_ALIAS,
    name: 'Endeavor Health Briefing'
  });
}

// ── TRIGGER SETUP (run once manually) ───────────────────────
function setupTrigger() {
  // Delete any existing triggers for runWeeklyBriefing
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runWeeklyBriefing') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create: every Monday at 6–7 AM (Apps Script uses local timezone set in project settings)
  // Set project timezone to America/Chicago in Project Settings before running this.
  ScriptApp.newTrigger('runWeeklyBriefing')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  Logger.log('Trigger created: runWeeklyBriefing every Monday at 6 AM (project timezone).');
}

// ── MANUAL TEST (run from Apps Script editor to test now) ────
function testRunNow() {
  runWeeklyBriefing();
}
