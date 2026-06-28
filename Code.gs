/**
 * Work Schedule → Google Calendar (Gmail HTML import)
 * ===================================================
 *
 * SETUP GUIDE
 * -----------
 * 1. Create a new Google Apps Script project: https://script.google.com
 *    Paste this entire file into Code.gs and save.
 *
 * 2. Create two Gmail labels (Gmail → Settings → Labels → Create new label):
 *    - "Work Schedule Import"   — apply to incoming schedule emails (optional filter)
 *    - "Work Schedule Imported" — applied automatically after a successful import
 *
 * 3. Optional Gmail filter:
 *    - Matches: subject contains "Week Starting"
 *    - Action: apply label "Work Schedule Import"
 *
 * 4. Add the Gmail API service (required for schedule image maps):
 *    GmailApp.getBody() strips <map>/<area> tags from schedule emails. The script
 *    reads the raw text/html MIME part via the Advanced Gmail service instead.
 *
 *    In the Apps Script editor:
 *    a. Click "Services" in the left sidebar (or the "+" next to Services).
 *    b. Select "Gmail API" from the list (Google identifier: Gmail).
 *    c. Leave the identifier as "Gmail" and version as "v1", then click Add.
 *
 *    If Google Cloud asks you to enable the API:
 *    a. Click the link to open the linked Cloud project.
 *    b. Enable "Gmail API" for that project.
 *    c. Return to Apps Script and confirm "Gmail" appears under Services.
 *
 *    After setup, a successful import log should show:
 *      HTML source: Gmail API raw MIME | ... | ImageMap1 present=true
 *    If you see "Advanced Gmail service not enabled", repeat this step.
 *
 * 5. Authorise the script:
 *    Run importLatestWorkSchedule() once from the editor (Run ▶).
 *    Accept Calendar + Gmail permissions when prompted.
 *
 * 6. Test parsing without Gmail:
 *    Run testParseWorkScheduleHtml() — sample HTML for week 29/06/2026 is pre-loaded.
 *    To test other weeks, replace TEST_HTML / TEST_SUBJECT or call
 *    testParseWorkScheduleHtml(htmlString, subject). See test-fixture.html in repo.
 *
 * 7. Test safely with real email:
 *    Set DRY_RUN = true (default below), run importLatestWorkSchedule(), check Logs
 *    (View → Executions or Ctrl+Enter after run).
 *
 * 8. Go live:
 *    Set DRY_RUN = false, run importLatestWorkSchedule() again, and verify "Work Rota"
 *    calendar events look correct.
 *
 * 9. Schedule hourly imports (choose one method):
 *
 *    Option A — run once from the editor:
 *    Run createHourlyTrigger() once. It creates an hourly trigger for
 *    importLatestWorkSchedule() and skips if one already exists.
 *
 *    Option B — create the trigger manually (Triggers tab → Add Trigger):
 *      Function:              importLatestWorkSchedule
 *      Deployment:            Head
 *      Event source:          Time-driven
 *      Type:                  Hour timer
 *      Hour interval:         Every hour
 *      Failure notifications: Notify me hourly (or daily, your preference)
 *
 *    Do NOT schedule createHourlyTrigger — that only sets up the trigger once.
 *    The function that imports your schedule every hour is importLatestWorkSchedule.
 *
 *    No "Deploy → New deployment" is required for personal use; Head is correct
 *    for time-driven triggers on your own account.
 *
 * 10. To remove the hourly trigger later:
 *    Triggers (clock icon) → delete the importLatestWorkSchedule trigger.
 *
 * CONFIGURATION
 * -------------
 * Adjust the constants below to match your labels, calendar name, and search query.
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

/** Set to true to parse and log without deleting or creating calendar events. */
const DRY_RUN = false;

/** IANA timezone for all calendar event times. */
const TIMEZONE = 'Europe/London';

/** Dedicated calendar for imported rota events. Created automatically if missing. */
const CALENDAR_NAME = 'Work Rota';

/** Default location on every imported event. */
const EVENT_LOCATION = 'Capita TFL';

/**
 * Gmail search query used to find candidate schedule emails.
 * Adjust sender, label, or age as needed.
 *
 * Sample queries:
 *   subject:"Week Starting" label:"Work Schedule Import" -label:"Work Schedule Imported" newer_than:30d
 *   from:me subject:"Week Starting" -label:"Work Schedule Imported"
 */
const GMAIL_SEARCH_QUERY =
  'subject:"Week Starting" label:"Work Schedule Import" -label:"Work Schedule Imported" newer_than:30d';

/** Label applied to source emails (optional filter target). */
const LABEL_IMPORT = 'Work Schedule Import';

/** Label applied after a successful import; emails with this label are skipped. */
const LABEL_IMPORTED = 'Work Schedule Imported';

/** Map element names in the HTML body (Monday = 1 … Sunday = 7). */
const IMAGE_MAP_NAMES = [
  'ImageMap1',
  'ImageMap2',
  'ImageMap3',
  'ImageMap4',
  'ImageMap5',
  'ImageMap6',
  'ImageMap7',
];

const DAY_NAMES = [
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
  'Sunday',
];

// ---------------------------------------------------------------------------
// Test fixtures — paste your real email HTML here for offline parsing tests
// ---------------------------------------------------------------------------

/** Subject line containing "Week Starting DD/MM/YYYY". */
const TEST_SUBJECT = 'Week Starting 29/06/2026';

/** Optional plain-text body fallback for date extraction (usually leave empty). */
const TEST_PLAIN_BODY = '';

/**
 * Sample schedule HTML for offline tests (maps only — see test-fixture.html in repo).
 * Replace with your full email HTML if needed; base64 images are not required for parsing.
 */
const TEST_HTML = `
<div id="divResults">
<map name="ImageMap1" id="ImageMap1">
  <area shape="rect" coords="410,50,417,86" title="06:15 PM - 06:25 PM PBRK3 " alt="06:15 PM - 06:25 PM PBRK3 ">
  <area shape="rect" coords="319,50,330,86" title="04:00 PM - 04:15 PM PBRK2 " alt="04:00 PM - 04:15 PM PBRK2 ">
  <area shape="rect" coords="218,50,239,86" title="01:30 PM - 02:00 PM LUNCH " alt="01:30 PM - 02:00 PM LUNCH ">
  <area shape="rect" coords="137,50,148,86" title="11:30 AM - 11:45 AM PBRK1 " alt="11:30 AM - 11:45 AM PBRK1 ">
  <area shape="rect" coords="57,41,481,77" title="09:30 AM - 08:00 PM HOLS Holiday day id [7305888] and hol detail id [161825] (awr ID [132714921])" alt="09:30 AM - 08:00 PM HOLS Holiday day id [7305888] and hol detail id [161825] (awr ID [132714921])">
  <area shape="rect" coords="57,31,481,68" title="09:30 AM - 08:00 PM TFLEOPS " alt="09:30 AM - 08:00 PM TFLEOPS ">
  <area shape="rect" coords="57,22,481,59" title="09:30 AM - 08:00 PM SHIFT " alt="09:30 AM - 08:00 PM SHIFT ">
</map>
<map name="ImageMap2" id="ImageMap2">
  <area shape="rect" coords="416,40,429,74" title="06:15 PM - 06:30 PM PBRK2 " alt="06:15 PM - 06:30 PM PBRK2 ">
  <area shape="rect" coords="273,40,299,74" title="03:30 PM - 04:00 PM LUNCH " alt="03:30 PM - 04:00 PM LUNCH ">
  <area shape="rect" coords="156,40,169,74" title="01:15 PM - 01:30 PM PBRK1 " alt="01:15 PM - 01:30 PM PBRK1 ">
  <area shape="rect" coords="117,40,156,74" title="12:30 PM - 01:15 PM COACH TM COACHING" alt="12:30 PM - 01:15 PM COACH TM COACHING">
  <area shape="rect" coords="91,31,507,66" title="12:00 PM - 08:00 PM TFLEOPS " alt="12:00 PM - 08:00 PM TFLEOPS ">
  <area shape="rect" coords="91,22,507,57" title="12:00 PM - 08:00 PM SHIFT " alt="12:00 PM - 08:00 PM SHIFT ">
</map>
<map name="ImageMap3" id="ImageMap3">
  <area shape="rect" coords="442,40,455,74" title="06:45 PM - 07:00 PM PBRK2 " alt="06:45 PM - 07:00 PM PBRK2 ">
  <area shape="rect" coords="286,40,312,74" title="03:45 PM - 04:15 PM LUNCH " alt="03:45 PM - 04:15 PM LUNCH ">
  <area shape="rect" coords="169,40,182,74" title="01:30 PM - 01:45 PM PBRK1 " alt="01:30 PM - 01:45 PM PBRK1 ">
  <area shape="rect" coords="91,31,507,66" title="12:00 PM - 08:00 PM TFLEOPS " alt="12:00 PM - 08:00 PM TFLEOPS ">
  <area shape="rect" coords="91,22,507,57" title="12:00 PM - 08:00 PM SHIFT " alt="12:00 PM - 08:00 PM SHIFT ">
</map>
<map name="ImageMap4" id="ImageMap4">
  <area shape="rect" coords="433,40,446,74" title="06:35 PM - 06:50 PM PBRK2 " alt="06:35 PM - 06:50 PM PBRK2 ">
  <area shape="rect" coords="273,40,299,74" title="03:30 PM - 04:00 PM LUNCH " alt="03:30 PM - 04:00 PM LUNCH ">
  <area shape="rect" coords="208,40,221,74" title="02:15 PM - 02:30 PM MEETT SPARK" alt="02:15 PM - 02:30 PM MEETT SPARK">
  <area shape="rect" coords="156,40,169,74" title="01:15 PM - 01:30 PM PBRK1 " alt="01:15 PM - 01:30 PM PBRK1 ">
  <area shape="rect" coords="91,31,507,66" title="12:00 PM - 08:00 PM TFLEOPS " alt="12:00 PM - 08:00 PM TFLEOPS ">
  <area shape="rect" coords="91,22,507,57" title="12:00 PM - 08:00 PM SHIFT " alt="12:00 PM - 08:00 PM SHIFT ">
</map>
<map name="ImageMap5" id="ImageMap5">
  <area shape="rect" coords="442,40,455,74" title="06:45 PM - 07:00 PM PBRK2 " alt="06:45 PM - 07:00 PM PBRK2 ">
  <area shape="rect" coords="291,40,317,74" title="03:50 PM - 04:20 PM LUNCH " alt="03:50 PM - 04:20 PM LUNCH ">
  <area shape="rect" coords="156,40,169,74" title="01:15 PM - 01:30 PM PBRK1 " alt="01:15 PM - 01:30 PM PBRK1 ">
  <area shape="rect" coords="91,31,507,66" title="12:00 PM - 08:00 PM TFLEOPS " alt="12:00 PM - 08:00 PM TFLEOPS ">
  <area shape="rect" coords="91,22,507,57" title="12:00 PM - 08:00 PM SHIFT " alt="12:00 PM - 08:00 PM SHIFT ">
</map>
<map name="ImageMap6" id="ImageMap6">
  <area shape="rect" coords="91,44,507,88" title="No data to display" alt="No data to display">
</map>
<map name="ImageMap7" id="ImageMap7">
  <area shape="rect" coords="91,44,507,88" title="No data to display" alt="No data to display">
</map>
</div>
`;

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/**
 * Manual entry point — run from the Apps Script editor or an hourly trigger.
 */
function importLatestWorkSchedule() {
  Logger.log('=== importLatestWorkSchedule started (DRY_RUN=%s) ===', DRY_RUN);

  const threads = GmailApp.search(GMAIL_SEARCH_QUERY, 0, 10);
  Logger.log('Gmail search "%s" returned %s thread(s).', GMAIL_SEARCH_QUERY, threads.length);

  if (threads.length === 0) {
    Logger.log('No matching schedule emails found. Nothing to do.');
    return;
  }

  // Process newest thread first; only one week per run.
  const thread = threads[0];
  const messages = thread.getMessages();
  const message = messages[messages.length - 1]; // latest message in thread

  processScheduleMessage_(message);
  Logger.log('=== importLatestWorkSchedule finished ===');
}

/**
 * Creates an hourly time-driven trigger for importLatestWorkSchedule().
 * Run once manually after you are happy with DRY_RUN = false behaviour.
 */
function createHourlyTrigger() {
  const handler = 'importLatestWorkSchedule';
  const existing = ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === handler;
  });

  if (existing.length > 0) {
    Logger.log(
      'Hourly trigger already exists (%s). Delete it from Triggers if you want a fresh one.',
      existing.length
    );
    return;
  }

  ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
  Logger.log('Created hourly trigger for %s().', handler);
}

/**
 * Offline test — parses schedule HTML without Gmail or Calendar.
 *
 * Uses TEST_HTML / TEST_SUBJECT by default. Optionally pass your own HTML:
 *   testParseWorkScheduleHtml(htmlString, 'Week Starting 29/06/2026')
 *
 * @param {string} [html] Optional HTML body; defaults to TEST_HTML
 * @param {string} [subject] Optional subject; defaults to TEST_SUBJECT
 * @param {string} [plainBody] Optional plain body for date fallback
 * @returns {Array<{summary: string, start: Date, end: Date, location: string}>|null}
 */
function testParseWorkScheduleHtml(html, subject, plainBody) {
  html = html || TEST_HTML;
  subject = subject || TEST_SUBJECT;
  plainBody = plainBody !== undefined ? plainBody : TEST_PLAIN_BODY;

  Logger.log('=== testParseWorkScheduleHtml started ===');
  Logger.log('Subject: "%s"', subject);

  if (!html || !html.trim()) {
    Logger.log('No HTML provided. Set TEST_HTML or pass html to testParseWorkScheduleHtml(html, subject).');
    Logger.log('=== testParseWorkScheduleHtml finished ===');
    return null;
  }

  const events = processScheduleContent_(subject, plainBody, html, {
    testMode: true,
  });

  if (events) {
    logEventsByDay_(events, extractWeekStartDate_(subject, plainBody));
  }

  Logger.log('=== testParseWorkScheduleHtml finished (%s event(s)) ===', events ? events.length : 0);
  return events;
}

// ---------------------------------------------------------------------------
// Core processing
// ---------------------------------------------------------------------------

/**
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 */
function processScheduleMessage_(message) {
  const subject = message.getSubject() || '';
  const messageId = message.getId();
  Logger.log('Processing message id=%s subject="%s"', messageId, subject);

  if (hasImportedLabel_(message)) {
    Logger.log('Message already has label "%s". Skipping.', LABEL_IMPORTED);
    return;
  }

  const events = processScheduleContent_(
    subject,
    message.getPlainBody(),
    getMessageHtmlBody_(message),
    { testMode: false, subject: subject, message: message }
  );

  if (!events) {
    return;
  }

  if (DRY_RUN) {
    logWouldImportEvents_(events);
    Logger.log('[DRY RUN] Would mark message with label "%s". Calendar unchanged.', LABEL_IMPORTED);
    return;
  }

  const weekStart = extractWeekStartDate_(subject, message.getPlainBody());
  const calendar = getOrCreateWorkRotaCalendar_();
  const weekEnd = addDays_(weekStart, 7);
  deleteExistingWeekEvents_(calendar, weekStart, weekEnd);
  createCalendarEvents_(calendar, events, subject);
  applyImportedLabel_(message);

  Logger.log(
    'Successfully imported %s event(s) for week starting %s.',
    events.length,
    formatDateOnly_(weekStart)
  );
}

/**
 * Shared parse pipeline for Gmail import and offline HTML tests.
 *
 * @param {string} subject
 * @param {string} plainBody
 * @param {string} htmlBody
 * @param {{testMode: boolean, message?: GoogleAppsScript.Gmail.GmailMessage}} options
 * @returns {Array<{summary: string, start: Date, end: Date, location: string}>|null}
 */
function processScheduleContent_(subject, plainBody, htmlBody, options) {
  options = options || { testMode: false };

  const weekStart = extractWeekStartDate_(subject, plainBody);
  if (!weekStart) {
    Logger.log(
      'Could not extract week start date from subject or body (expected "Week Starting DD/MM" or "Week Starting DD/MM/YYYY"). Skipping.'
    );
    return null;
  }
  Logger.log('Week start date: %s (%s)', formatDateOnly_(weekStart), DAY_NAMES[0]);

  if (!htmlBody || !htmlBody.trim()) {
    Logger.log('HTML body is empty. Skipping.');
    return null;
  }

  const eventsByDay = parseImageMaps_(htmlBody);
  const totalParsed = countParsedEvents_(eventsByDay);

  if (totalParsed === 0) {
    Logger.log('No schedule events found in ImageMap1–ImageMap7. Calendar will not be modified.');
    return null;
  }

  Logger.log('Parsed %s event(s) across %s day map(s).', totalParsed, IMAGE_MAP_NAMES.length);

  const calendarEvents = buildCalendarEvents_(eventsByDay, weekStart);
  const eventsToCreate = calendarEvents.filter(function (event) {
    return !shouldIgnoreTitle_(event.summary);
  });

  Logger.log(
    '%s event(s) after filtering SHIFT entries (%s ignored).',
    eventsToCreate.length,
    calendarEvents.length - eventsToCreate.length
  );

  if (eventsToCreate.length === 0) {
    Logger.log('No events remain after filtering. Calendar will not be modified.');
    return null;
  }

  if (options.testMode) {
    logWouldImportEvents_(eventsToCreate);
    return eventsToCreate;
  }

  return eventsToCreate;
}

/**
 * @param {Array<{summary: string, start: Date, end: Date, location: string}>} events
 */
function logWouldImportEvents_(events) {
  Logger.log('[DRY RUN] Would import into calendar "%s":', CALENDAR_NAME);
  events.forEach(function (event, index) {
    Logger.log(
      '  [%s] %s | %s %s–%s | %s',
      index + 1,
      event.summary,
      formatDateOnly_(event.start),
      Utilities.formatDate(event.start, TIMEZONE, 'HH:mm'),
      Utilities.formatDate(event.end, TIMEZONE, 'HH:mm'),
      event.location
    );
  });
}

// ---------------------------------------------------------------------------
// Gmail helpers
// ---------------------------------------------------------------------------

/**
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @returns {boolean}
 */
function hasImportedLabel_(message) {
  const labels = message.getThread().getLabels();
  return labels.some(function (label) {
    return label.getName() === LABEL_IMPORTED;
  });
}

/**
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 */
function applyImportedLabel_(message) {
  const label = GmailApp.getUserLabelByName(LABEL_IMPORTED) || GmailApp.createLabel(LABEL_IMPORTED);
  message.getThread().addLabel(label);
  Logger.log('Applied Gmail label "%s".', LABEL_IMPORTED);
}

/**
 * Best-effort HTML body for schedule parsing.
 * GmailApp.getBody() often strips <map>/<area>; raw MIME usually preserves them.
 *
 * @param {GoogleAppsScript.Gmail.GmailMessage} message
 * @returns {string}
 */
function getMessageHtmlBody_(message) {
  const messageId = message.getId();
  let html = '';
  let source = 'none';

  const rawHtml = getRawHtmlBodyViaGmailApi_(messageId);
  if (rawHtml && htmlContainsScheduleMaps_(rawHtml)) {
    html = rawHtml;
    source = 'Gmail API raw MIME';
  }

  if (!html) {
    const bodyHtml = message.getBody() || '';
    if (htmlContainsScheduleMaps_(bodyHtml)) {
      html = bodyHtml;
      source = 'GmailApp.getBody()';
    } else if (rawHtml) {
      html = rawHtml;
      source = 'Gmail API raw MIME (no maps detected in getBody)';
    } else {
      html = bodyHtml;
      source = 'GmailApp.getBody()';
    }
  }

  html = prepareHtmlForParsing_(html);
  Logger.log(
    'HTML source: %s | length=%s | ImageMap1 present=%s',
    source,
    html.length,
    html.indexOf('ImageMap1') !== -1
  );

  if (!htmlContainsScheduleMaps_(html)) {
    Logger.log(
      'WARNING: No ImageMap markers in HTML. Schedule tables may have been stripped by Gmail.'
    );
    Logger.log('HTML preview: %s', html.substring(0, 400).replace(/\s+/g, ' '));
  }

  return html;
}

/**
 * @param {string} html
 * @returns {boolean}
 */
function htmlContainsScheduleMaps_(html) {
  return /ImageMap[1-7]/i.test(html || '');
}

/**
 * Decode structural HTML entities so tag patterns can match email source HTML.
 *
 * @param {string} html
 * @returns {string}
 */
function prepareHtmlForParsing_(html) {
  if (!html) {
    return '';
  }

  let normalized = html.replace(/\uFEFF/g, '');

  if (/&lt;(?:map|area|div)\b/i.test(normalized)) {
    normalized = normalized
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/g, "'")
      .replace(/&apos;/gi, "'")
      .replace(/&nbsp;/gi, ' ')
      .replace(/&#160;/g, ' ')
      .replace(/&amp;/gi, '&');
  }

  return normalized;
}

/**
 * Read text/html from the raw Gmail MIME payload (requires Advanced Gmail service).
 *
 * @param {string} messageId
 * @returns {string}
 */
function getRawHtmlBodyViaGmailApi_(messageId) {
  if (typeof Gmail === 'undefined' || !Gmail.Users) {
    Logger.log('Advanced Gmail service not enabled — using GmailApp.getBody() only.');
    return '';
  }

  try {
    const message = Gmail.Users.Messages.get('me', messageId, { format: 'full' });
    return extractHtmlFromMimePart_(message.payload) || '';
  } catch (err) {
    Logger.log('Failed to read raw MIME HTML: %s', err);
    return '';
  }
}

/**
 * @param {GoogleAppsScript.Gmail.GmailV1.Schema.MessagePart|null|undefined} part
 * @returns {string}
 */
function extractHtmlFromMimePart_(part) {
  if (!part) {
    return '';
  }

  if (part.mimeType === 'text/html' && part.body && part.body.data) {
    return decodeGmailBase64_(part.body.data);
  }

  if (part.parts) {
    for (var i = 0; i < part.parts.length; i++) {
      const html = extractHtmlFromMimePart_(part.parts[i]);
      if (html) {
        return html;
      }
    }
  }

  return '';
}

/**
 * @param {string} encoded
 * @returns {string}
 */
function decodeGmailBase64_(encoded) {
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(encoded)).getDataAsString('UTF-8');
}

// ---------------------------------------------------------------------------
// Date / subject parsing
// ---------------------------------------------------------------------------

/**
 * Extract week start (Monday) from subject or plain body.
 * Accepts "Week Starting DD/MM/YYYY" or "Week Starting DD/MM" (year inferred).
 *
 * @param {string} subject
 * @param {string} plainBody
 * @returns {Date|null} Date at 00:00 Europe/London on the Monday, or null
 */
function extractWeekStartDate_(subject, plainBody) {
  const sources = [subject || '', plainBody || ''];
  const datePattern = /Week\s+Starting\s+(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?/i;
  let match = null;

  for (var i = 0; i < sources.length; i++) {
    match = sources[i].match(datePattern);
    if (match) {
      break;
    }
  }

  if (!match) {
    return null;
  }

  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const explicitYear = match[3] ? parseInt(match[3], 10) : null;

  if (explicitYear) {
    return buildWeekStartDate_(day, month, explicitYear);
  }

  return inferWeekStartDateFromDayMonth_(day, month);
}

/**
 * @param {number} day
 * @param {number} month
 * @param {number} year
 * @returns {Date|null}
 */
function buildWeekStartDate_(day, month, year) {
  if (!isValidCalendarDate_(year, month, day)) {
    Logger.log('Invalid calendar date: %s/%s/%s', day, month, year);
    return null;
  }

  return londonDate_(year, month, day, 0, 0);
}

/**
 * Infer DD/MM year from today in Europe/London — pick closest date among
 * previous, current, and next calendar year (checking backwards and forwards).
 *
 * @param {number} day
 * @param {number} month
 * @returns {Date|null}
 */
function inferWeekStartDateFromDayMonth_(day, month) {
  const today = getTodayLondon_();
  const todayYear = parseInt(Utilities.formatDate(today, TIMEZONE, 'yyyy'), 10);
  let bestDate = null;
  let bestDiff = null;
  let bestYear = null;

  for (var year = todayYear - 1; year <= todayYear + 1; year++) {
    if (!isValidCalendarDate_(year, month, day)) {
      continue;
    }

    const candidate = londonDate_(year, month, day, 0, 0);
    const diff = Math.abs(candidate.getTime() - today.getTime());

    if (
      bestDiff === null ||
      diff < bestDiff ||
      (diff === bestDiff && year > bestYear)
    ) {
      bestDiff = diff;
      bestYear = year;
      bestDate = candidate;
    }
  }

  if (!bestDate) {
    Logger.log('Could not infer a valid year for %s/%s.', day, month);
    return null;
  }

  Logger.log(
    'Inferred week start %s for %s/%s (closest to today %s).',
    formatDateOnly_(bestDate),
    pad2_(day),
    pad2_(month),
    formatDateOnly_(today)
  );

  return bestDate;
}

/**
 * Today at 00:00 in Europe/London.
 *
 * @returns {Date}
 */
function getTodayLondon_() {
  const parts = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd').split('-');
  return londonDate_(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10),
    parseInt(parts[2], 10),
    0,
    0
  );
}

/**
 * @param {number} year
 * @param {number} month 1–12
 * @param {number} day 1–31
 * @returns {boolean}
 */
function isValidCalendarDate_(year, month, day) {
  const test = new Date(Date.UTC(year, month - 1, day));
  return (
    test.getUTCFullYear() === year &&
    test.getUTCMonth() === month - 1 &&
    test.getUTCDate() === day
  );
}

/**
 * Build a Date for a wall-clock time on a calendar day in Europe/London.
 *
 * @param {number} year
 * @param {number} month 1–12
 * @param {number} day 1–31
 * @param {number} hour 0–23
 * @param {number} minute 0–59
 * @returns {Date}
 */
function londonDate_(year, month, day, hour, minute) {
  hour = hour || 0;
  minute = minute || 0;

  const localString =
    String(year) +
    '-' +
    pad2_(month) +
    '-' +
    pad2_(day) +
    'T' +
    pad2_(hour) +
    ':' +
    pad2_(minute) +
    ':00';

  return new Date(
    Utilities.parseDate(localString, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss").getTime()
  );
}

/**
 * Add calendar days in Europe/London, preserving the local time of day.
 *
 * @param {Date} anchorDate
 * @param {number} days
 * @returns {Date}
 */
function addDays_(anchorDate, days) {
  const datePart = Utilities.formatDate(anchorDate, TIMEZONE, 'yyyy-MM-dd');
  const timePart = Utilities.formatDate(anchorDate, TIMEZONE, 'HH:mm:ss');
  const parts = datePart.split('-').map(function (part) {
    return parseInt(part, 10);
  });

  const shifted = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + days));
  const newDatePart =
    shifted.getUTCFullYear() +
    '-' +
    pad2_(shifted.getUTCMonth() + 1) +
    '-' +
    pad2_(shifted.getUTCDate());

  return new Date(
    Utilities.parseDate(newDatePart + 'T' + timePart, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss").getTime()
  );
}

/**
 * @param {Date} date
 * @returns {string}
 */
function formatDateOnly_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

/**
 * @param {number} n
 * @returns {string}
 */
function pad2_(n) {
  return n < 10 ? '0' + n : String(n);
}

// ---------------------------------------------------------------------------
// HTML parsing
// ---------------------------------------------------------------------------

/**
 * Decode common HTML entities in attribute values.
 *
 * @param {string} text
 * @returns {string}
 */
function decodeHtmlEntities_(text) {
  if (!text) {
    return '';
  }

  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&#160;/g, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/gi, "'")
    .replace(/&#(\d+);/g, function (_match, code) {
      return String.fromCharCode(parseInt(code, 10));
    })
    .replace(/&#x([0-9a-f]+);/gi, function (_match, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Parse ImageMap1–ImageMap7 and extract area title attributes.
 *
 * @param {string} html
 * @returns {string[][]} titles per day index 0–6
 */
function parseImageMaps_(html) {
  html = prepareHtmlForParsing_(html);

  const results = IMAGE_MAP_NAMES.map(function () {
    return [];
  });

  IMAGE_MAP_NAMES.forEach(function (mapName, dayIndex) {
    const mapInner = extractMapInnerHtml_(html, mapName);

    if (!mapInner) {
      Logger.log('Map %s (%s): not found in HTML.', mapName, DAY_NAMES[dayIndex]);
      return;
    }

    const titles = extractTitlesFromMapInner_(mapInner);
    results[dayIndex] = titles;

    Logger.log('Map %s (%s): found %s area title(s).', mapName, DAY_NAMES[dayIndex], titles.length);
  });

  return results;
}

/**
 * Extract inner HTML of a named schedule map (macro-compatible patterns first).
 *
 * @param {string} html
 * @param {string} mapName
 * @returns {string|null}
 */
function extractMapInnerHtml_(html, mapName) {
  const patterns = [
    new RegExp('(?<=<map name="' + mapName + '" id="' + mapName + '">)[\\s\\S]+?(?=</map>)', 'i'),
    new RegExp('(?<=<map id="' + mapName + '" name="' + mapName + '">)[\\s\\S]+?(?=</map>)', 'i'),
    new RegExp('<map\\b[^>]*\\bname=["\']?' + mapName + '["\']?[^>]*>([\\s\\S]*?)<\\/map>', 'i'),
    new RegExp('<map\\b[^>]*\\bid=["\']?' + mapName + '["\']?[^>]*>([\\s\\S]*?)<\\/map>', 'i'),
  ];

  for (var i = 0; i < patterns.length; i++) {
    const match = html.match(patterns[i]);
    if (match) {
      return match[1] || match[0];
    }
  }

  return null;
}

/**
 * Extract title strings from inside a map block (area tags, then macro-style fallback).
 *
 * @param {string} mapInner
 * @returns {string[]}
 */
function extractTitlesFromMapInner_(mapInner) {
  const titles = [];
  const seen = {};

  const areaRegex = /<area\b[^>]*\btitle\s*=\s*["']([^"']*)["'][^>]*>/gi;
  let areaMatch;

  while ((areaMatch = areaRegex.exec(mapInner)) !== null) {
    const title = decodeHtmlEntities_(areaMatch[1]);
    if (title && !seen[title]) {
      seen[title] = true;
      titles.push(title);
    }
  }

  if (titles.length > 0) {
    return titles;
  }

  const macroTitleRegex = /(?<=title=")(.+?)(?=")/gi;
  let macroMatch;

  while ((macroMatch = macroTitleRegex.exec(mapInner)) !== null) {
    const title = decodeHtmlEntities_(macroMatch[1]);
    if (title && !seen[title]) {
      seen[title] = true;
      titles.push(title);
    }
  }

  return titles;
}

/**
 * @param {string[][]} eventsByDay
 * @returns {number}
 */
function countParsedEvents_(eventsByDay) {
  return eventsByDay.reduce(function (sum, dayEvents) {
    return sum + dayEvents.length;
  }, 0);
}

// ---------------------------------------------------------------------------
// Title / time parsing
// ---------------------------------------------------------------------------

/**
 * Parse "08:00 AM - 04:00 PM TFLBOPS " style titles.
 *
 * @param {string} titleAttr
 * @returns {{startTime: string, endTime: string, summary: string}|null}
 */
function parseTitleAttribute_(titleAttr) {
  const cleaned = decodeHtmlEntities_(titleAttr);
  const match = cleaned.match(
    /^(\d{1,2}:\d{2}\s*[AP]M)\s*-\s*(\d{1,2}:\d{2}\s*[AP]M)\s+(.+)$/i
  );

  if (!match) {
    if (!isIgnorableAreaTitle_(cleaned)) {
      Logger.log('Could not parse title attribute: "%s"', cleaned);
    }
    return null;
  }

  return {
    startTime: normalizeTimeToken_(match[1]),
    endTime: normalizeTimeToken_(match[2]),
    summary: normalizeEventTitle_(match[3]),
  };
}

/**
 * @param {string} token e.g. "08:00 AM"
 * @returns {string}
 */
function normalizeTimeToken_(token) {
  return token.replace(/\s+/g, ' ').toUpperCase();
}

/**
 * Apply rename rules and trim trailing spaces from codes.
 *
 * @param {string} rawTitle
 * @returns {string}
 */
function normalizeEventTitle_(rawTitle) {
  let title = rawTitle.trim();

  if (/PBRK/i.test(title)) {
    return title.replace(/PBRK/gi, 'BREAK ').trim();
  }
  if (/LUNCH/i.test(title)) {
    return 'LUNCH BREAK';
  }

  return title.trim();
}

/**
 * @param {string} summary
 * @returns {boolean}
 */
function shouldIgnoreTitle_(summary) {
  return /^SHIFT\s*$/i.test(summary.trim());
}

/**
 * Area titles that are expected on empty days — skip without warning.
 *
 * @param {string} title
 * @returns {boolean}
 */
function isIgnorableAreaTitle_(title) {
  return /^no data to display$/i.test(title.trim());
}

/**
 * Log parsed events grouped by day (test helper).
 *
 * @param {Array<{summary: string, start: Date, end: Date, location: string}>} events
 * @param {Date|null} weekStart
 */
function logEventsByDay_(events, weekStart) {
  if (!weekStart) {
    return;
  }

  IMAGE_MAP_NAMES.forEach(function (_mapName, dayIndex) {
    const dayStart = addDays_(weekStart, dayIndex);
    const dayEnd = addDays_(weekStart, dayIndex + 1);
    const dayEvents = events.filter(function (event) {
      return event.start.getTime() >= dayStart.getTime() && event.start.getTime() < dayEnd.getTime();
    });

    Logger.log(
      '%s (%s): %s event(s)',
      DAY_NAMES[dayIndex],
      formatDateOnly_(dayStart),
      dayEvents.length
    );
    dayEvents.forEach(function (event) {
      Logger.log(
        '    %s %s–%s',
        event.summary,
        Utilities.formatDate(event.start, TIMEZONE, 'HH:mm'),
        Utilities.formatDate(event.end, TIMEZONE, 'HH:mm')
      );
    });
  });
}

/**
 * @param {string[][]} eventsByDay
 * @param {Date} weekStart Monday 00:00 London
 * @returns {Array<{summary: string, start: Date, end: Date, location: string}>}
 */
function buildCalendarEvents_(eventsByDay, weekStart) {
  const output = [];

  eventsByDay.forEach(function (dayTitles, dayIndex) {
    const dayDate = addDays_(weekStart, dayIndex);

    dayTitles.forEach(function (titleAttr) {
      const parsed = parseTitleAttribute_(titleAttr);
      if (!parsed) {
        return;
      }

      const start = combineDateAndTimeLondon_(dayDate, parsed.startTime);
      let end = combineDateAndTimeLondon_(dayDate, parsed.endTime);

      // Overnight shift segment (rare): end before start → roll end to next day.
      if (end.getTime() <= start.getTime()) {
        end = addDays_(end, 1);
      }

      output.push({
        summary: parsed.summary,
        start: start,
        end: end,
        location: EVENT_LOCATION,
      });
    });
  });

  return output;
}

/**
 * Combine a calendar date with "HH:MM AM/PM" in Europe/London.
 *
 * @param {Date} dayDate anchor day (local London date)
 * @param {string} timeToken e.g. "08:00 AM"
 * @returns {Date}
 */
function combineDateAndTimeLondon_(dayDate, timeToken) {
  const datePart = Utilities.formatDate(dayDate, TIMEZONE, 'yyyy-MM-dd');
  const match = timeToken.match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/i);
  if (!match) {
    throw new Error('Invalid time token: ' + timeToken);
  }

  let hour = parseInt(match[1], 10);
  const minute = parseInt(match[2], 10);
  const meridiem = match[3].toUpperCase();

  if (meridiem === 'AM') {
    if (hour === 12) {
      hour = 0;
    }
  } else if (hour !== 12) {
    hour += 12;
  }

  const localString =
    datePart + 'T' + pad2_(hour) + ':' + pad2_(minute) + ':00';

  return new Date(
    Utilities.parseDate(localString, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss").getTime()
  );
}

// ---------------------------------------------------------------------------
// Google Calendar
// ---------------------------------------------------------------------------

/**
 * @returns {GoogleAppsScript.Calendar.Calendar}
 */
function getOrCreateWorkRotaCalendar_() {
  const calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (calendars.length > 0) {
    Logger.log('Using existing calendar "%s" (id=%s).', CALENDAR_NAME, calendars[0].getId());
    return calendars[0];
  }

  const created = CalendarApp.createCalendar(CALENDAR_NAME, {
    timeZone: TIMEZONE,
  });
  Logger.log('Created calendar "%s" (id=%s).', CALENDAR_NAME, created.getId());
  return created;
}

/**
 * Delete events in [weekStart, weekEnd) — Monday 00:00 through next Monday 00:00.
 *
 * @param {GoogleAppsScript.Calendar.Calendar} calendar
 * @param {Date} weekStart
 * @param {Date} weekEnd
 */
function deleteExistingWeekEvents_(calendar, weekStart, weekEnd) {
  const events = calendar.getEvents(weekStart, weekEnd);
  Logger.log(
    'Deleting %s existing event(s) in "%s" for %s → %s.',
    events.length,
    CALENDAR_NAME,
    formatDateOnly_(weekStart),
    formatDateOnly_(weekEnd)
  );

  events.forEach(function (event) {
    try {
      event.deleteEvent();
    } catch (err) {
      Logger.log('Failed to delete event "%s": %s', event.getTitle(), err);
    }
  });
}

/**
 * @param {GoogleAppsScript.Calendar.Calendar} calendar
 * @param {Array<{summary: string, start: Date, end: Date, location: string}>} events
 * @param {string} emailSubject
 */
function createCalendarEvents_(calendar, events, emailSubject) {
  const description =
    'Imported from work schedule email\n\nGmail subject: ' + emailSubject;

  events.forEach(function (eventData) {
    calendar.createEvent(eventData.summary, eventData.start, eventData.end, {
      location: eventData.location,
      description: description,
    });

    Logger.log(
      'Created: %s | %s %s–%s',
      eventData.summary,
      formatDateOnly_(eventData.start),
      Utilities.formatDate(eventData.start, TIMEZONE, 'HH:mm'),
      Utilities.formatDate(eventData.end, TIMEZONE, 'HH:mm')
    );
  });
}
