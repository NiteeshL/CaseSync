const PROCESSED_ICS_URLS = new Set();
const SEEN_ICS_URLS_AT_LOAD = new Set();

const CASE_REGEX = /\b(?:INC|REQ|RITM|TASK|CHG|PRB|SCTASK|SCREQ)\d{4,}\b/i;
const URL_REGEX = /(https?:\/\/[^\s<>"']+)/i;

function normalizeUrl(url) {
  try {
    const parsed = new URL(url, window.location.href);
    parsed.hash = "";
    return parsed.toString();
  } catch (_) {
    return String(url || "");
  }
}

function isIcsAnchor(anchor) {
  if (!(anchor instanceof HTMLAnchorElement)) {
    return false;
  }

  const href = anchor.getAttribute("href") || "";
  const text = (anchor.textContent || "").toLowerCase();
  const fullUrl = normalizeUrl(href).toLowerCase();

  if (fullUrl.includes(".ics") || href.toLowerCase().includes(".ics")) {
    return true;
  }

  if (text.includes(".ics") || text.includes("text/calendar")) {
    return true;
  }

  if (fullUrl.includes("sys_attachment.do") && /filename=.*\.ics|file_name=.*\.ics/.test(fullUrl)) {
    return true;
  }

  return false;
}

function unfoldIcsLines(raw) {
  const normalized = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  const unfolded = [];

  for (const line of lines) {
    if (/^[ \t]/.test(line) && unfolded.length > 0) {
      unfolded[unfolded.length - 1] += line.slice(1);
    } else {
      unfolded.push(line);
    }
  }

  return unfolded;
}

function parseEventFromIcs(rawIcs) {
  const lines = unfoldIcsLines(rawIcs);
  let summary = "";
  let description = "";
  let dtStartRaw = "";
  let dtEndRaw = "";

  for (const line of lines) {
    if (!summary && line.startsWith("SUMMARY")) {
      summary = line.split(":").slice(1).join(":").trim();
    } else if (!description && line.startsWith("DESCRIPTION")) {
      description = line.split(":").slice(1).join(":").trim();
    } else if (!dtStartRaw && line.startsWith("DTSTART")) {
      dtStartRaw = line.split(":").slice(1).join(":").trim();
    } else if (!dtEndRaw && line.startsWith("DTEND")) {
      dtEndRaw = line.split(":").slice(1).join(":").trim();
    }
  }

  return {
    summary,
    description: description.replace(/\\n/g, "\n"),
    dtStartRaw,
    dtEndRaw
  };
}

function formatAsIsoDateTime(icsDateValue) {
  if (!icsDateValue) {
    return "";
  }

  const cleaned = icsDateValue.replace(/Z$/, "").trim();
  const compact = cleaned.replace(/[^0-9T]/g, "");

  if (/^\d{8}$/.test(compact)) {
    const yyyy = compact.slice(0, 4);
    const mm = compact.slice(4, 6);
    const dd = compact.slice(6, 8);
    return `${yyyy}-${mm}-${dd}T00:00:00`;
  }

  if (/^\d{8}T\d{6}$/.test(compact)) {
    const yyyy = compact.slice(0, 4);
    const mm = compact.slice(4, 6);
    const dd = compact.slice(6, 8);
    const hh = compact.slice(9, 11);
    const min = compact.slice(11, 13);
    const ss = compact.slice(13, 15);
    return `${yyyy}-${mm}-${dd}T${hh}:${min}:${ss}`;
  }

  return "";
}

function extractCaseNumber(summary, description) {
  const summaryMatch = summary.match(CASE_REGEX);
  if (summaryMatch) {
    return summaryMatch[0].toUpperCase();
  }

  const descriptionMatch = description.match(CASE_REGEX);
  return descriptionMatch ? descriptionMatch[0].toUpperCase() : "";
}

function extractMeetingLink(description) {
  const match = description.match(URL_REGEX);
  return match ? match[1] : "";
}

function buildOutlookDeeplink(eventData) {
  const { summary, description, startdt, enddt, caseNumber, meetingLink } = eventData;
  const subject = caseNumber && !summary.includes(caseNumber)
    ? `${summary} [${caseNumber}]`
    : summary;

  const bodyParts = [];
  if (description) {
    bodyParts.push(description);
  }
  if (meetingLink && !description.includes(meetingLink)) {
    bodyParts.push(`Meeting Link: ${meetingLink}`);
  }

  const params = new URLSearchParams({
    subject: subject || "ServiceNow Meeting",
    startdt,
    enddt,
    body: bodyParts.join("\n\n"),
    location: "Online",
    online: "1"
  });

  return `https://outlook.office.com/calendar/0/deeplink/compose?${params.toString()}`;
}

async function fetchIcsText(icsUrl) {
  const response = await fetch(icsUrl, {
    credentials: "include"
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch ICS file: ${response.status}`);
  }

  return response.text();
}

async function processIcsUrl(rawUrl) {
  const icsUrl = normalizeUrl(rawUrl);
  if (!icsUrl || PROCESSED_ICS_URLS.has(icsUrl) || SEEN_ICS_URLS_AT_LOAD.has(icsUrl)) {
    return;
  }

  PROCESSED_ICS_URLS.add(icsUrl);

  try {
    const icsText = await fetchIcsText(icsUrl);
    const parsed = parseEventFromIcs(icsText);

    const startdt = formatAsIsoDateTime(parsed.dtStartRaw);
    const enddt = formatAsIsoDateTime(parsed.dtEndRaw);

    if (!startdt || !enddt) {
      throw new Error("Invalid DTSTART/DTEND format in ICS.");
    }

    const caseNumber = extractCaseNumber(parsed.summary, parsed.description);
    const meetingLink = extractMeetingLink(parsed.description);

    const deeplink = buildOutlookDeeplink({
      summary: parsed.summary,
      description: parsed.description,
      startdt,
      enddt,
      caseNumber,
      meetingLink
    });

    chrome.runtime.sendMessage({
      type: "OPEN_OUTLOOK_DEEPLINK",
      url: deeplink
    });
  } catch (error) {
    console.error("ServiceNow ICS to Outlook: failed to process ICS", error);
  }
}

function getIcsAnchorsFromNode(node) {
  if (!(node instanceof Element)) {
    return [];
  }

  const anchors = [];

  if (node instanceof HTMLAnchorElement) {
    anchors.push(node);
  }

  anchors.push(...node.querySelectorAll("a[href]"));
  return anchors.filter(isIcsAnchor);
}

function markExistingIcsLinksAsSeen() {
  const existingAnchors = document.querySelectorAll("a[href]");
  for (const anchor of existingAnchors) {
    if (isIcsAnchor(anchor)) {
      SEEN_ICS_URLS_AT_LOAD.add(normalizeUrl(anchor.href));
    }
  }
}

function setupObserver() {
  const observer = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      for (const addedNode of mutation.addedNodes) {
        const anchors = getIcsAnchorsFromNode(addedNode);
        for (const anchor of anchors) {
          processIcsUrl(anchor.href);
        }
      }
    }
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}

function init() {
  markExistingIcsLinksAsSeen();
  setupObserver();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init, { once: true });
} else {
  init();
}
