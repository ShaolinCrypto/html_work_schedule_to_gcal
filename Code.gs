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
 * 4. Add the Gmail API service (required for schedule HTML):
 *    GmailApp.getBody() often strips schedule markup (image maps or Kendo scheduler
 *    accessibility spans). The script reads the raw text/html MIME part via the
 *    Advanced Gmail service instead.
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
 *      HTML source: Gmail API raw MIME | ... | scheduler=true
 *    If you see "Advanced Gmail service not enabled", repeat this step.
 *
 * 5. Authorise the script:
 *    Run importLatestWorkSchedule() once from the editor (Run ▶).
 *    Accept Calendar + Gmail permissions when prompted.
 *
 * 6. Test parsing without Gmail:
 *    Run testParseWorkScheduleHtml() — Kendo scheduler sample for week 29/06/2026 is
 *    pre-loaded. For a full email body use test-fixture-scheduler.html in repo, or call
 *    testParseWorkScheduleHtml(htmlString, subject). Legacy ImageMap HTML still works
 *    (see test-fixture.html).
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

/** Subject line containing "Week Starting DD/MM" or "Week Starting DD/MM/YYYY". */
const TEST_SUBJECT = 'Week Starting 29/06';

/** Optional plain-text body fallback for date extraction (usually leave empty). */
const TEST_PLAIN_BODY = '';

/**
 * Sample Kendo scheduler HTML for offline tests (see test-fixture-scheduler.html for full email).
 * Legacy ImageMap sample is in test-fixture.html.
 */
const TEST_HTML = `
<div id="seScheduler">
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Holiday All Day. Memo. Holiday day id [7305889] and hol detail id [161826] (awr ID [132714923]). </span>
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Transport for London Enforceme From 9:30 AM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Shift (container) From 9:30 AM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Paid Break 1 From 11:30 AM To 11:45 AM. </span>
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Lunch From 1:30 PM To 2:00 PM. </span>
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Paid Break 2 From 4:00 PM To 4:15 PM. </span>
<span class="accessibility-screen-reader"> Monday, June 29, 2026 Segment Paid Break 3 From 6:15 PM To 6:25 PM. </span>
<span class="accessibility-screen-reader"> Tuesday, June 30, 2026 Segment Transport for London Enforceme From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Tuesday, June 30, 2026 Segment Shift (container) From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Tuesday, June 30, 2026 Segment Coaching Session From 12:30 PM To 1:15 PM. Memo. TM COACHING. </span>
<span class="accessibility-screen-reader"> Tuesday, June 30, 2026 Segment Paid Break 1 From 2:00 PM To 2:15 PM. </span>
<span class="accessibility-screen-reader"> Tuesday, June 30, 2026 Segment Lunch From 3:30 PM To 4:00 PM. </span>
<span class="accessibility-screen-reader"> Tuesday, June 30, 2026 Segment Paid Break 2 From 5:30 PM To 5:45 PM. </span>
<span class="accessibility-screen-reader"> Wednesday, July 1, 2026 Segment Transport for London Enforceme From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Wednesday, July 1, 2026 Segment Shift (container) From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Wednesday, July 1, 2026 Segment Paid Break 1 From 1:30 PM To 1:45 PM. </span>
<span class="accessibility-screen-reader"> Wednesday, July 1, 2026 Segment Lunch From 3:45 PM To 4:15 PM. </span>
<span class="accessibility-screen-reader"> Wednesday, July 1, 2026 Segment Paid Break 2 From 6:45 PM To 7:00 PM. </span>
<span class="accessibility-screen-reader"> Thursday, July 2, 2026 Segment Transport for London Enforceme From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Thursday, July 2, 2026 Segment Shift (container) From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Thursday, July 2, 2026 Segment Paid Break 1 From 1:30 PM To 1:45 PM. </span>
<span class="accessibility-screen-reader"> Thursday, July 2, 2026 Segment Lunch From 3:30 PM To 4:00 PM. </span>
<span class="accessibility-screen-reader"> Thursday, July 2, 2026 Segment Paid Break 2 From 6:45 PM To 7:00 PM. </span>
<span class="accessibility-screen-reader"> Friday, July 3, 2026 Segment Transport for London Enforceme From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Friday, July 3, 2026 Segment Shift (container) From 12:00 PM To 8:00 PM. </span>
<span class="accessibility-screen-reader"> Friday, July 3, 2026 Segment Paid Break 1 From 1:15 PM To 1:30 PM. </span>
<span class="accessibility-screen-reader"> Friday, July 3, 2026 Segment Lunch From 3:30 PM To 4:00 PM. </span>
<span class="accessibility-screen-reader"> Friday, July 3, 2026 Segment Paid Break 2 From 6:10 PM To 6:25 PM. </span>
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

  const result = processScheduleContent_(subject, plainBody, html, {
    testMode: true,
  });

  if (result) {
    logEventsByDay_(result.events, result.weekStart);
  }

  Logger.log(
    '=== testParseWorkScheduleHtml finished (%s event(s)) ===',
    result ? result.events.length : 0
  );
  return result ? result.events : null;
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

  const result = processScheduleContent_(
    subject,
    message.getPlainBody(),
    getMessageHtmlBody_(message),
    { testMode: false, subject: subject, message: message }
  );

  if (!result) {
    return;
  }

  const events = result.events;
  const weekStart = result.weekStart;

  if (DRY_RUN) {
    logWouldImportEvents_(events);
    Logger.log('[DRY RUN] Would mark message with label "%s". Calendar unchanged.', LABEL_IMPORTED);
    return;
  }

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
 * @returns {{events: Array<{summary: string, start: Date, end: Date, location: string}>, weekStart: Date}|null}
 */
function processScheduleContent_(subject, plainBody, htmlBody, options) {
  options = options || { testMode: false };

  if (!htmlBody || !htmlBody.trim()) {
    Logger.log('HTML body is empty. Skipping.');
    return null;
  }

  const isScheduler = htmlContainsSchedulerSegments_(htmlBody);
  let weekStart = resolveWeekStartForSchedule_(subject, plainBody, htmlBody, null);

  if (!isScheduler && !weekStart) {
    Logger.log(
      'Could not extract week start for legacy ImageMap HTML (expected "Week Starting DD/MM" in subject or scheduler week range in HTML). Skipping.'
    );
    return null;
  }

  if (weekStart) {
    Logger.log('Week start date: %s (%s)', formatDateOnly_(weekStart), DAY_NAMES[0]);
  } else {
    Logger.log(
      'No week start in subject or HTML header; will infer from parsed segment dates (Kendo scheduler).'
    );
  }

  const parseResult = parseScheduleHtml_(htmlBody, weekStart);
  const calendarEvents = parseResult.events;

  if (calendarEvents.length === 0) {
    Logger.log(
      'No schedule events found in HTML (%s parser). Calendar will not be modified.',
      parseResult.format
    );
    return null;
  }

  if (!weekStart) {
    weekStart = resolveWeekStartForSchedule_(subject, plainBody, htmlBody, calendarEvents);
  }
  if (!weekStart) {
    Logger.log('Could not determine week start date for calendar import. Skipping.');
    return null;
  }
  Logger.log('Week start for import: %s (%s)', formatDateOnly_(weekStart), DAY_NAMES[0]);

  Logger.log(
    'Parser: %s | %s raw segment(s)/title(s) → %s event(s) before filtering.',
    parseResult.format,
    parseResult.rawCount,
    calendarEvents.length
  );

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
    return { events: eventsToCreate, weekStart: weekStart };
  }

  return { events: eventsToCreate, weekStart: weekStart };
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
  if (rawHtml && htmlContainsScheduleData_(rawHtml)) {
    html = rawHtml;
    source = 'Gmail API raw MIME';
  }

  if (!html) {
    const bodyHtml = message.getBody() || '';
    if (htmlContainsScheduleData_(bodyHtml)) {
      html = bodyHtml;
      source = 'GmailApp.getBody()';
    } else if (rawHtml) {
      html = rawHtml;
      source = 'Gmail API raw MIME (no schedule markers in getBody)';
    } else {
      html = bodyHtml;
      source = 'GmailApp.getBody()';
    }
  }

  html = prepareHtmlForParsing_(html);
  Logger.log(
    'HTML source: %s | length=%s | scheduler=%s | ImageMap1=%s',
    source,
    html.length,
    htmlContainsSchedulerSegments_(html),
    html.indexOf('ImageMap1') !== -1
  );

  if (!htmlContainsScheduleData_(html)) {
    Logger.log(
      'WARNING: No schedule markers in HTML. Schedule content may have been stripped by Gmail.'
    );
    Logger.log('HTML preview: %s', html.substring(0, 400).replace(/\s+/g, ' '));
  }

  return html;
}

/**
 * @param {string} html
 * @returns {boolean}
 */
function htmlContainsScheduleData_(html) {
  return htmlContainsSchedulerSegments_(html) || htmlContainsScheduleMaps_(html);
}

/**
 * @param {string} html
 * @returns {boolean}
 */
function htmlContainsSchedulerSegments_(html) {
  return (
    /accessibility-screen-reader/i.test(html || '') && /\bSegment\b/i.test(html || '')
  );
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
 * Extract Monday week start from Kendo scheduler date range in HTML.
 *
 * @param {string} html
 * @returns {Date|null}
 */
function extractWeekStartFromSchedulerHtml_(html) {
  if (!html) {
    return null;
  }

  const match = html.match(
    /(Monday,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+-\s+Sunday,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4}/i
  );
  if (!match) {
    return null;
  }

  const weekStart = parseSegmentDate_(match[1]);
  if (weekStart) {
    Logger.log('Week start from scheduler HTML: %s', formatDateOnly_(weekStart));
  }
  return weekStart;
}

/**
 * Resolve Monday week start for import/delete window.
 * Order: subject/body → scheduler header in HTML → earliest parsed event date.
 *
 * @param {string} subject
 * @param {string} plainBody
 * @param {string} htmlBody
 * @param {Array<{start: Date}>|null} parsedEvents
 * @returns {Date|null}
 */
function resolveWeekStartForSchedule_(subject, plainBody, htmlBody, parsedEvents) {
  let weekStart = extractWeekStartDate_(subject, plainBody);
  if (weekStart) {
    return weekStart;
  }

  if (htmlBody) {
    weekStart = extractWeekStartFromSchedulerHtml_(htmlBody);
    if (weekStart) {
      return weekStart;
    }
  }

  if (parsedEvents && parsedEvents.length > 0) {
    weekStart = inferWeekStartFromParsedEvents_(parsedEvents);
    if (weekStart) {
      Logger.log(
        'Inferred week start %s from earliest parsed event (%s).',
        formatDateOnly_(weekStart),
        formatDateOnly_(parsedEvents.reduce(function (earliest, event) {
          return event.start.getTime() < earliest.start.getTime() ? event : earliest;
        }, parsedEvents[0]).start)
      );
    }
  }

  return weekStart;
}

/**
 * Monday 00:00 London for the week containing the earliest parsed event.
 *
 * @param {Array<{start: Date}>} events
 * @returns {Date|null}
 */
function inferWeekStartFromParsedEvents_(events) {
  if (!events || events.length === 0) {
    return null;
  }

  let earliest = events[0].start;
  events.forEach(function (event) {
    if (event.start.getTime() < earliest.getTime()) {
      earliest = event.start;
    }
  });

  return getMondayOfWeekLondon_(earliest);
}

/**
 * @param {Date} date any instant in Europe/London
 * @returns {Date|null} that week's Monday at 00:00 London
 */
function getMondayOfWeekLondon_(date) {
  const dayName = Utilities.formatDate(date, TIMEZONE, 'EEEE');
  const dayIndex = DAY_NAMES.indexOf(dayName);
  if (dayIndex < 0) {
    return null;
  }

  const midnight = londonDateFromInstant_(date);
  return addDays_(midnight, -dayIndex);
}

/**
 * @param {Date} date
 * @returns {Date} same calendar day at 00:00 Europe/London
 */
function londonDateFromInstant_(date) {
  const parts = Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd').split('-');
  return londonDate_(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10),
    parseInt(parts[2], 10),
    0,
    0
  );
}

/**
 * Parse "Monday, June 29, 2026" in Europe/London.
 *
 * @param {string} dateStr
 * @returns {Date|null}
 */
function parseSegmentDate_(dateStr) {
  try {
    return Utilities.parseDate(dateStr.trim(), TIMEZONE, 'EEEE, MMMM d, yyyy');
  } catch (err) {
    Logger.log('Could not parse segment date "%s": %s', dateStr, err);
    return null;
  }
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

/** Regex for each accessibility span segment line (non-capturing All Day branch). */
const SEGMENT_LINE_REGEX_ =
  /^([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+Segment\s+(.+?)\s+(?:(?:From\s+(\d{1,2}:\d{2}\s+[AP]M)\s+To\s+(\d{1,2}:\d{2}\s+[AP]M))|(?:All Day))\.?(?:\s+Memo\.\s*(.*?))?$/i;

/**
 * Route HTML to Kendo scheduler or legacy ImageMap parser.
 *
 * @param {string} html
 * @param {Date} weekStart Monday 00:00 London (ImageMap path only)
 * @returns {{format: string, rawCount: number, events: Array<{summary: string, start: Date, end: Date, location: string}>}}
 */
function parseScheduleHtml_(html, weekStart) {
  html = prepareHtmlForParsing_(html);

  if (htmlContainsSchedulerSegments_(html)) {
    const lines = extractAccessibilitySegmentLines_(html);
    const segments = lines
      .map(function (line) {
        return parseSegmentLine_(line);
      })
      .filter(function (segment) {
        return segment !== null;
      });

    Logger.log(
      'Kendo scheduler: %s accessibility line(s), %s parsed segment(s).',
      lines.length,
      segments.length
    );

    return {
      format: 'kendo-scheduler',
      rawCount: lines.length,
      events: buildCalendarEventsFromSegments_(segments),
    };
  }

  const eventsByDay = parseImageMaps_(html);
  const totalParsed = countParsedEvents_(eventsByDay);

  return {
    format: 'imagemap',
    rawCount: totalParsed,
    events: buildCalendarEvents_(eventsByDay, weekStart),
  };
}

/**
 * Extract segment description lines from accessibility spans (and aria-label fallback).
 *
 * @param {string} html
 * @returns {string[]}
 */
function extractAccessibilitySegmentLines_(html) {
  const lines = [];
  const seen = {};

  try {
    const lookbehind =
      /(?<=<span class="accessibility-screen-reader">[| \u00a0]).+?(?=[| \u00a0]<\/span>)/gi;
    let match;
    while ((match = lookbehind.exec(html)) !== null) {
      addUniqueSegmentLine_(lines, seen, decodeHtmlEntities_(match[0]));
    }
  } catch (err) {
    Logger.log('Lookbehind span extraction unavailable (%s); using fallback pattern.', err);
  }

  const fallback =
    /<span class="accessibility-screen-reader">\s*[|\u00a0]?\s*(.+?)\s*[|\u00a0]?\s*<\/span>/gi;
  let fbMatch;
  while ((fbMatch = fallback.exec(html)) !== null) {
    addUniqueSegmentLine_(lines, seen, decodeHtmlEntities_(fbMatch[1]));
  }

  const ariaRegex = /aria-label="([^"]*\bSegment\b[^"]*)"/gi;
  let ariaMatch;
  while ((ariaMatch = ariaRegex.exec(html)) !== null) {
    addUniqueSegmentLine_(lines, seen, decodeHtmlEntities_(ariaMatch[1]));
  }

  return lines.filter(function (line) {
    return /\bSegment\b/i.test(line);
  });
}

/**
 * @param {string[]} lines
 * @param {Object<string, boolean>} seen
 * @param {string} line
 */
function addUniqueSegmentLine_(lines, seen, line) {
  const normalized = (line || '').replace(/\s+/g, ' ').trim();
  if (!normalized || seen[normalized]) {
    return;
  }
  seen[normalized] = true;
  lines.push(normalized);
}

/**
 * @param {string} line
 * @returns {{date: Date, segmentName: string, startTime: string|null, endTime: string|null, isAllDay: boolean, memo: string|null}|null}
 */
function parseSegmentLine_(line) {
  const cleaned = decodeHtmlEntities_(line);
  const match = cleaned.match(SEGMENT_LINE_REGEX_);
  if (!match) {
    Logger.log('Could not parse segment line: "%s"', cleaned.substring(0, 120));
    return null;
  }

  const date = parseSegmentDate_(match[1]);
  if (!date) {
    return null;
  }

  return {
    date: date,
    segmentName: match[2].trim(),
    startTime: match[3] ? normalizeTimeToken_(match[3]) : null,
    endTime: match[4] ? normalizeTimeToken_(match[4]) : null,
    isAllDay: !match[3],
    memo: match[5] ? match[5].replace(/\.$/, '').trim() : null,
  };
}

/**
 * Build calendar events from parsed Kendo scheduler segments.
 *
 * @param {Array<{date: Date, segmentName: string, startTime: string|null, endTime: string|null, isAllDay: boolean, memo: string|null}>} segments
 * @returns {Array<{summary: string, start: Date, end: Date, location: string}>}
 */
function buildCalendarEventsFromSegments_(segments) {
  const shiftByDate = {};

  segments.forEach(function (segment) {
    if (isShiftContainerSegment_(segment.segmentName) && segment.startTime && segment.endTime) {
      shiftByDate[formatDateOnly_(segment.date)] = {
        startTime: segment.startTime,
        endTime: segment.endTime,
      };
    }
  });

  const output = [];

  segments.forEach(function (segment) {
    if (isShiftContainerSegment_(segment.segmentName)) {
      return;
    }

    let startTime = segment.startTime;
    let endTime = segment.endTime;

    if (segment.isAllDay) {
      const shift = shiftByDate[formatDateOnly_(segment.date)];
      if (!shift) {
        Logger.log(
          'All Day segment "%s" on %s: no Shift (container) times; skipping.',
          segment.segmentName,
          formatDateOnly_(segment.date)
        );
        return;
      }
      startTime = shift.startTime;
      endTime = shift.endTime;
    }

    if (!startTime || !endTime) {
      return;
    }

    const start = combineDateAndTimeLondon_(segment.date, startTime);
    let end = combineDateAndTimeLondon_(segment.date, endTime);
    if (end.getTime() <= start.getTime()) {
      end = addDays_(end, 1);
    }

    output.push({
      summary: normalizeSegmentTitle_(segment.segmentName, segment.memo),
      start: start,
      end: end,
      location: EVENT_LOCATION,
    });
  });

  return output;
}

/**
 * @param {string} segmentName
 * @returns {boolean}
 */
function isShiftContainerSegment_(segmentName) {
  return /^shift\s*\(container\)$/i.test((segmentName || '').trim());
}

/**
 * Apply Kendo segment title rules (no "Segment" prefix; Paid Break → Break, etc.).
 *
 * @param {string} segmentName
 * @param {string|null} memo
 * @returns {string}
 */
function normalizeSegmentTitle_(segmentName, memo) {
  let title = (segmentName || '').trim();

  if (/^paid break\s+/i.test(title)) {
    title = title.replace(/^paid break\s+/i, 'Break ');
  } else if (/^lunch$/i.test(title)) {
    title = 'LUNCH BREAK';
  } else if (/transport for london enforceme/i.test(title)) {
    title = 'TFL EOPs';
  } else if (/^coaching session$/i.test(title) && memo) {
    title = 'COACH - ' + memo;
  }

  return title.trim();
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
