const OUTLOOK_BASE_URL = "https://outlook.office.com/calendar/0/deeplink/compose";
const URL_GLOBAL_REGEX = /(https?:\/\/[^\s<>"']+)/g;
const SCHEDULE_WINDOW_MS = 7000;
const PENDING_SCHEDULE_TTL_MS = 20000;
const DEDUP_STORAGE_KEY = "capturedJournalSysIds";
const LOG_BODY_PREVIEW_LEN = 220;
const LOG_PREFIX = "[SN MEETING]";
const INSTANCE_ID = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
const HEARTBEAT_MS = 30000;
const INIT_GUARD_KEY = "__SN_MEETING_MONITOR_INITIALIZED__";

function getFrameKind() {
  try {
    return window.top === window ? "top" : "iframe";
  } catch (_) {
    return "iframe";
  }
}

const FRAME_KIND = getFrameKind();

console.info(`${LOG_PREFIX} content.js evaluated`, {
  instanceId: INSTANCE_ID,
  href: window.location.href,
  frameKind: FRAME_KIND
});

let pendingSchedule = null;
let seenJournalSysIds = new Set();

function logDebug(message, extra) {
  const scopedPrefix = `${LOG_PREFIX} [${INSTANCE_ID}] [${FRAME_KIND}]`;
  if (extra !== undefined) {
    console.log(`${scopedPrefix} ${message}`, extra);
    return;
  }
  console.log(`${scopedPrefix} ${message}`);
}

function startHeartbeat() {
  window.setInterval(() => {
    logDebug("Heartbeat: extension monitor still running", {
      href: window.location.href
    });
  }, HEARTBEAT_MS);
}

function injectRunningBadge() {
  if (FRAME_KIND !== "top") {
    return;
  }

  if (!document.body) {
    return;
  }

  const badgeId = "sn-meeting-extension-running-badge";
  if (document.getElementById(badgeId)) {
    return;
  }

  const badge = document.createElement("div");
  badge.id = badgeId;
  badge.textContent = `${LOG_PREFIX} running`;
  badge.style.position = "fixed";
  badge.style.right = "12px";
  badge.style.bottom = "12px";
  badge.style.zIndex = "2147483647";
  badge.style.padding = "6px 10px";
  badge.style.fontSize = "12px";
  badge.style.fontFamily = "ui-monospace, SFMono-Regular, Menlo, monospace";
  badge.style.color = "#ffffff";
  badge.style.background = "#0078d4";
  badge.style.borderRadius = "8px";
  badge.style.boxShadow = "0 2px 6px rgba(0,0,0,0.25)";
  badge.style.pointerEvents = "none";
  badge.style.opacity = "0.95";
  document.body.appendChild(badge);
}

function shortText(value, maxLen = LOG_BODY_PREVIEW_LEN) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (text.length <= maxLen) {
    return text;
  }
  return `${text.slice(0, maxLen)}...`;
}

function normalizeUrl(url) {
  try {
    return new URL(url, window.location.href).toString();
  } catch (_) {
    return String(url || "");
  }
}

function normalizeDateString(dateValue) {
  if (!dateValue) {
    return "";
  }

  const raw = String(dateValue).trim();
  if (!raw) {
    return "";
  }

  if (/^\d{8}T\d{6}Z?$/.test(raw)) {
    const isUtc = raw.endsWith("Z");
    const clean = raw.replace(/Z$/, "");
    const yyyy = clean.slice(0, 4);
    const mm = clean.slice(4, 6);
    const dd = clean.slice(6, 8);
    const hh = clean.slice(9, 11);
    const min = clean.slice(11, 13);
    const ss = clean.slice(13, 15);
    return `${yyyy}-${mm}-${dd}T${hh}:${min}:${ss}${isUtc ? "Z" : ""}`;
  }

  if (/^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$/i.test(raw)) {
    const normalized = raw.replace(" ", "T");
    return normalized.includes("Z") || /[+-]\d{2}:?\d{2}$/.test(normalized)
      ? normalized
      : `${normalized}${normalized.length === 16 ? ":00" : ""}`;
  }

  const parsed = Date.parse(raw);
  if (Number.isNaN(parsed)) {
    return "";
  }

  return new Date(parsed).toISOString().replace(/\.\d{3}Z$/, "Z");
}

function parseServiceNowDateToMs(value) {
  if (!value) {
    return NaN;
  }

  const raw = String(value).trim();
  if (!raw) {
    return NaN;
  }

  const direct = Date.parse(raw);
  if (!Number.isNaN(direct)) {
    return direct;
  }

  const match = raw.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})$/);
  if (match) {
    const localGuess = Date.parse(`${match[1]}T${match[2]}`);
    if (!Number.isNaN(localGuess)) {
      return localGuess;
    }
  }

  logDebug("Unable to parse ServiceNow date", { value: raw });
  return NaN;
}

function extractCaseSysIdFromPage() {
  const url = new URL(window.location.href);
  const direct = url.searchParams.get("sys_id");
  if (direct) {
    return direct;
  }

  const uri = url.searchParams.get("uri") || "";
  const embeddedMatch = uri.match(/[?&]sys_id=([a-f0-9]{32})/i);
  return embeddedMatch ? embeddedMatch[1] : "";
}

function extractScheduleFieldsFromBody(body) {
  const result = {
    topic: "",
    startdt: "",
    enddt: "",
    caseSysId: ""
  };

  const assignIfMatches = (key, value) => {
    if (value == null) {
      return;
    }

    const normalizedValue = String(value).trim();
    if (!normalizedValue) {
      return;
    }

    if (!result.topic && /(topic|subject|title|meeting_name|meetingTopic)/i.test(key)) {
      result.topic = normalizedValue;
    }

    if (!result.startdt && /(start|begin|from|start_time|startTime|meeting_start|meetingStart)/i.test(key)) {
      result.startdt = normalizeDateString(normalizedValue);
    }

    if (!result.enddt && /(end|to|finish|end_time|endTime|meeting_end|meetingEnd)/i.test(key)) {
      result.enddt = normalizeDateString(normalizedValue);
    }

    if (!result.caseSysId && /(u_case_zoom_schedule\.u_case|u_case|case_sys_id|case)/i.test(key)) {
      const sysIdMatch = normalizedValue.match(/\b[a-f0-9]{32}\b/i);
      if (sysIdMatch) {
        result.caseSysId = sysIdMatch[0];
      }
    }
  };

  if (!body) {
    logDebug("Schedule POST body missing");
    return result;
  }

  if (typeof body === "string") {
    const text = body.trim();
    if (!text) {
      return result;
    }

    try {
      const parsedJson = JSON.parse(text);
      const stack = [parsedJson];
      while (stack.length > 0) {
        const current = stack.pop();
        if (!current || typeof current !== "object") {
          continue;
        }

        for (const [key, value] of Object.entries(current)) {
          if (value && typeof value === "object") {
            stack.push(value);
          } else {
            assignIfMatches(key, value);
          }
        }
      }
      logDebug("Extracted schedule fields from JSON body", result);
      return result;
    } catch (_) {
      const params = new URLSearchParams(text);
      for (const [key, value] of params.entries()) {
        assignIfMatches(key, value);
      }
      logDebug("Extracted schedule fields from URL-encoded body", result);
      return result;
    }
  }

  if (typeof FormData !== "undefined" && body instanceof FormData) {
    for (const [key, value] of body.entries()) {
      assignIfMatches(key, value);
    }
  }

  if (typeof URLSearchParams !== "undefined" && body instanceof URLSearchParams) {
    for (const [key, value] of body.entries()) {
      assignIfMatches(key, value);
    }
  }

  logDebug("Extracted schedule fields from non-string body", result);

  return result;
}

function isMeetingSchedulePost(url, method, bodyText) {
  if (method !== "POST") {
    return false;
  }

  const normalized = normalizeUrl(url);
  const loweredUrl = normalized.toLowerCase();
  const loweredBody = (bodyText || "").toLowerCase();

  const isZoomScheduleEndpoint = /\/u_case_zoom_schedule\.do(?:\?|$)/.test(loweredUrl);
  const hasRequiredBodySignals =
    loweredBody.includes("sys_target=u_case_zoom_schedule") ||
    (loweredBody.includes("u_case_zoom_schedule.u_start_date_time") && loweredBody.includes("u_case_zoom_schedule.u_end_date_time"));

  const urlSignal = isZoomScheduleEndpoint;
  const bodySignal = hasRequiredBodySignals;

  const matched = urlSignal || bodySignal;
  logDebug("Evaluated meeting POST candidate", {
    method,
    url: normalized,
    matched,
    urlSignal,
    bodySignal,
    bodyPreview: shortText(bodyText)
  });
  return matched;
}

function isListHistoryRequest(url) {
  return /\/list_history\.do(\?|$)/i.test(url);
}

function getMeetingUrl(text) {
  const urls = text.match(URL_GLOBAL_REGEX) || [];
  if (urls.length === 0) {
    return "";
  }

  const preferred = urls.find((url) => /zoom|teams|meet\.google|webex/i.test(url));
  return preferred || urls[0];
}

function getPassword(text) {
  const match = text.match(/^Password:\s*(.+)$/mi);
  return match ? match[1].trim() : "";
}

function getTopic(text) {
  const match = text.match(/^Topic:\s*(.+)$/mi);
  return match ? match[1].trim() : "";
}

function getSubjectFromJournal(text) {
  return getTopic(text) || pendingSchedule?.topic || "ServiceNow Meeting";
}

function buildOutlookDeeplink(details) {
  const bodyParts = [];
  if (details.journalText) {
    bodyParts.push(details.journalText);
  }
  if (details.meetingLink && !details.journalText.includes(details.meetingLink)) {
    bodyParts.push(`Meeting Link: ${details.meetingLink}`);
  }
  if (details.password && !details.journalText.includes(`Password: ${details.password}`)) {
    bodyParts.push(`Password: ${details.password}`);
  }

  const params = new URLSearchParams({
    subject: details.subject,
    startdt: details.startdt,
    enddt: details.enddt,
    body: bodyParts.join("\n\n"),
    location: "Online",
    online: "1"
  });

  return `${OUTLOOK_BASE_URL}?${params.toString()}`;
}

function sendOutlookOpen(url) {
  logDebug("Sending deeplink open request to background", {
    urlPreview: shortText(url, 300)
  });
  chrome.runtime.sendMessage({
    type: "OPEN_OUTLOOK_DEEPLINK",
    url
  });
}

function flattenObjects(root) {
  const items = [];
  const stack = [root];

  while (stack.length > 0) {
    const current = stack.pop();
    if (!current || typeof current !== "object") {
      continue;
    }

    items.push(current);

    if (Array.isArray(current)) {
      for (const item of current) {
        if (item && typeof item === "object") {
          stack.push(item);
        }
      }
    } else {
      for (const value of Object.values(current)) {
        if (value && typeof value === "object") {
          stack.push(value);
        }
      }
    }
  }

  return items;
}

function getJournalTextFromNode(node) {
  const keys = [
    "value",
    "new_value",
    "display_value",
    "comments",
    "work_notes",
    "text",
    "entry",
    "message",
    "journal"
  ];

  for (const key of keys) {
    if (typeof node[key] === "string" && node[key].trim()) {
      return node[key].trim();
    }
  }

  return "";
}

function isCommentsOrWorkNotesNode(node) {
  const keys = ["element", "field", "name", "type", "journal_type", "field_name", "field_label"];
  for (const key of keys) {
    if (typeof node[key] !== "string") {
      continue;
    }

    const value = node[key].toLowerCase();
    if (value.includes("comments") || value.includes("work_notes") || value.includes("work notes")) {
      return true;
    }
  }

  return false;
}

function extractJournalEntries(payload) {
  const allNodes = flattenObjects(payload);
  const entries = [];

  for (const node of allNodes) {
    if (!node || typeof node !== "object") {
      continue;
    }

    const sysId = typeof node.sys_id === "string" ? node.sys_id : "";
    const created = typeof node.sys_created_on_adjusted === "string" ? node.sys_created_on_adjusted : "";
    const text = getJournalTextFromNode(node);

    if (!sysId || !created || !text) {
      continue;
    }

    if (!isCommentsOrWorkNotesNode(node)) {
      continue;
    }

    if (!/scheduled\s+zoom\s+meeting|join\s+from\s+pc|topic:/i.test(text)) {
      continue;
    }

    entries.push({
      sysId,
      created,
      text
    });
  }

  logDebug("Extracted journal entries from list_history payload", {
    count: entries.length
  });
  return entries;
}

function isWithinPendingWindow(createdMs) {
  if (!pendingSchedule) {
    return false;
  }

  const deltaMs = Math.abs(createdMs - pendingSchedule.timestamp);
  const within = deltaMs <= SCHEDULE_WINDOW_MS;
  logDebug("Timestamp gate evaluated", {
    createdMs,
    pendingTimestamp: pendingSchedule.timestamp,
    deltaMs,
    windowMs: SCHEDULE_WINDOW_MS,
    within
  });
  return within;
}

function isPendingScheduleFresh() {
  if (!pendingSchedule) {
    return false;
  }

  const ageMs = Date.now() - pendingSchedule.timestamp;
  const fresh = ageMs <= PENDING_SCHEDULE_TTL_MS;
  if (!fresh) {
    logDebug("Pending schedule expired", {
      ageMs,
      ttlMs: PENDING_SCHEDULE_TTL_MS,
      pendingSchedule
    });
  }
  return fresh;
}

async function persistSeenSysIds() {
  const ids = Array.from(seenJournalSysIds);
  await chrome.storage.local.set({ [DEDUP_STORAGE_KEY]: ids.slice(-500) });
  logDebug("Persisted dedup sys_id cache", { size: Math.min(ids.length, 500) });
}

async function markSysIdSeen(sysId) {
  seenJournalSysIds.add(sysId);
  await persistSeenSysIds();
}

function parseJsonSafely(text) {
  try {
    return JSON.parse(text);
  } catch (_) {
    logDebug("Failed to parse list_history response as JSON", {
      preview: shortText(text)
    });
    return null;
  }
}

async function processHistoryResponse(responseText, sourceUrl) {
  if (!pendingSchedule) {
    logDebug("Ignoring list_history because no pending schedule exists", { sourceUrl });
    return;
  }

  if (!isPendingScheduleFresh()) {
    pendingSchedule = null;
    return;
  }

  logDebug("Processing list_history response", {
    sourceUrl,
    responseLength: responseText.length,
    pendingSchedule
  });

  const payload = parseJsonSafely(responseText);
  if (!payload) {
    logDebug("Skipping list_history response due to invalid JSON", { sourceUrl });
    return;
  }

  const entries = extractJournalEntries(payload);
  if (entries.length === 0) {
    logDebug("list_history parsed but no journal entries found", sourceUrl);
    return;
  }

  for (const entry of entries) {
    if (seenJournalSysIds.has(entry.sysId)) {
      logDebug("Skipping journal entry: sys_id already processed", { sysId: entry.sysId });
      continue;
    }

    const createdMs = parseServiceNowDateToMs(entry.created);
    if (Number.isNaN(createdMs) || !isWithinPendingWindow(createdMs)) {
      logDebug("Skipping journal entry: outside timestamp window", {
        sysId: entry.sysId,
        created: entry.created
      });
      continue;
    }

    const subject = getSubjectFromJournal(entry.text);
    const meetingLink = getMeetingUrl(entry.text);
    const password = getPassword(entry.text);

    const startdt = pendingSchedule.startdt || "";
    const enddt = pendingSchedule.enddt || "";

    if (!startdt || !enddt) {
      logDebug("Skipping matched journal because schedule POST did not expose start/end", pendingSchedule);
      continue;
    }

    const deeplink = buildOutlookDeeplink({
      subject,
      startdt,
      enddt,
      journalText: entry.text,
      meetingLink,
      password
    });

    await markSysIdSeen(entry.sysId);
    sendOutlookOpen(deeplink);
    logDebug("Opened Outlook from matched journal entry", {
      sysId: entry.sysId,
      created: entry.created
    });

    pendingSchedule = null;
    logDebug("Cleared pending schedule after successful Outlook open");
    break;
  }
}

function getRequestBodyText(body) {
  if (typeof body === "string") {
    return body;
  }

  if (body instanceof URLSearchParams) {
    return body.toString();
  }

  return "";
}

function startPendingSchedule(method, url, body) {
  const bodyText = getRequestBodyText(body);
  if (!isMeetingSchedulePost(url, method, bodyText)) {
    if (method === "POST") {
      logDebug("POST ignored (did not match meeting pattern)", {
        url,
        bodyPreview: shortText(bodyText)
      });
    }
    return;
  }

  const fields = extractScheduleFieldsFromBody(body);
  if (!fields.startdt || !fields.enddt) {
    logDebug("Meeting-like POST found but missing start/end fields", {
      url,
      fields,
      bodyPreview: shortText(bodyText)
    });
    return;
  }

  pendingSchedule = {
    timestamp: Date.now(),
    requestUrl: url,
    caseSysId: fields.caseSysId || extractCaseSysIdFromPage(),
    topic: fields.topic,
    startdt: fields.startdt,
    enddt: fields.enddt
  };

  logDebug("Captured pending meeting schedule POST", pendingSchedule);
}

function installFetchMonitor() {
  const originalFetch = window.fetch.bind(window);
  logDebug("Installing fetch monitor");

  window.fetch = async (...args) => {
    const input = args[0];
    const init = args[1] || {};

    const requestUrl = normalizeUrl(typeof input === "string" ? input : input?.url);
    const method = String(init.method || input?.method || "GET").toUpperCase();

    if (method === "POST" || isListHistoryRequest(requestUrl)) {
      logDebug("Fetch observed", {
        method,
        requestUrl,
        bodyPreview: shortText(getRequestBodyText(init.body || input?.body))
      });
    }

    startPendingSchedule(method, requestUrl, init.body || input?.body);

    const response = await originalFetch(...args);

    if (isListHistoryRequest(requestUrl)) {
      logDebug("Fetch list_history response received", {
        requestUrl,
        status: response.status
      });
      response.clone().text()
        .then((text) => processHistoryResponse(text, requestUrl))
        .catch((error) => {
          logDebug("Failed reading fetch list_history response body", { requestUrl, error: String(error) });
        });
    }

    return response;
  };
}

function installXhrMonitor() {
  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSend = XMLHttpRequest.prototype.send;
  logDebug("Installing XHR monitor");

  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    this.__snMeetingMethod = String(method || "GET").toUpperCase();
    this.__snMeetingUrl = normalizeUrl(url);
    return originalOpen.call(this, method, url, ...rest);
  };

  XMLHttpRequest.prototype.send = function (body) {
    const method = this.__snMeetingMethod || "GET";
    const url = this.__snMeetingUrl || "";

    if (method === "POST" || isListHistoryRequest(url)) {
      logDebug("XHR observed", {
        method,
        url,
        bodyPreview: shortText(getRequestBodyText(body))
      });
    }

    startPendingSchedule(method, url, body);

    this.addEventListener("load", () => {
      if (!isListHistoryRequest(url)) {
        return;
      }

      logDebug("XHR list_history response received", {
        url,
        status: this.status
      });

      if (typeof this.responseText !== "string" || !this.responseText.trim()) {
        logDebug("XHR list_history empty responseText", { url });
        return;
      }

      processHistoryResponse(this.responseText, url).catch((error) => {
        logDebug("Failed processing XHR list_history response", { url, error: String(error) });
      });
    });

    return originalSend.call(this, body);
  };
}

async function loadSeenSysIds() {
  const stored = await chrome.storage.local.get(DEDUP_STORAGE_KEY);
  const values = Array.isArray(stored[DEDUP_STORAGE_KEY]) ? stored[DEDUP_STORAGE_KEY] : [];
  seenJournalSysIds = new Set(values.filter((item) => typeof item === "string" && item));
  logDebug("Loaded journal sys_id dedup cache", { size: seenJournalSysIds.size });
}

async function init() {
  console.info(`${LOG_PREFIX} Extension content script loaded`, {
    instanceId: INSTANCE_ID,
    frameKind: FRAME_KIND,
    href: window.location.href,
    host: window.location.hostname
  });

  if (window[INIT_GUARD_KEY]) {
    logDebug("Init skipped: monitor already initialized in this frame");
    return;
  }
  window[INIT_GUARD_KEY] = true;
  logDebug("Init guard set for frame");

  logDebug("Content script init started", { href: window.location.href });

  if (document.readyState === "complete" || document.readyState === "interactive") {
    injectRunningBadge();
  } else {
    document.addEventListener("DOMContentLoaded", injectRunningBadge, { once: true });
  }

  await loadSeenSysIds();
  installFetchMonitor();
  installXhrMonitor();
  startHeartbeat();
  logDebug("Meeting monitor initialized", { href: window.location.href });
}

init().catch((error) => {
  console.error("ServiceNow meeting monitor failed to initialize", error);
});
