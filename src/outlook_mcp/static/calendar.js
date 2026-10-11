"use strict";

(function () {
  var BOOT = window.__OUTLOOK_CALENDAR_BOOTSTRAP__ || {};
  if (!BOOT.apiBase || !window.FullCalendar) {
    return;
  }
  if (!BOOT.calendarTimezone) {
    throw new Error("calendarTimezone is required");
  }
  BOOT.capabilities = BOOT.capabilities || {};

  var refs = {
    calendar: document.getElementById("calendar"),
    loading: document.getElementById("cal-loading"),
    calendarStatus: document.getElementById("calendar-status"),
    cacheStatus: document.getElementById("calendar-cache-status"),
    calendarTab: document.getElementById("tab-calendar"),
    draftsTab: document.getElementById("tab-drafts"),
    calendarView: document.getElementById("calendar-view"),
    draftsView: document.getElementById("drafts-view"),
    draftsList: document.getElementById("drafts-list"),
    draftsEmpty: document.getElementById("drafts-empty"),
    draftsStatus: document.getElementById("drafts-status"),
    refreshDraftsButton: document.getElementById("btn-refresh-drafts"),
    retentionMode: document.getElementById("draft-retention-mode"),
    retentionDays: document.getElementById("draft-retention-days"),
    saveRetentionButton: document.getElementById("btn-save-retention"),
    panel: document.getElementById("editor-panel"),
    backdrop: document.getElementById("panel-backdrop"),
    panelEmpty: document.getElementById("panel-empty"),
    panelDayAgenda: document.getElementById("panel-day-agenda"),
    dayAgendaTitle: document.getElementById("day-agenda-title"),
    dayAgendaSummary: document.getElementById("day-agenda-summary"),
    dayAgendaList: document.getElementById("day-agenda-list"),
    dayAgendaEmpty: document.getElementById("day-agenda-empty"),
    dayNewButton: document.getElementById("btn-day-new-event"),
    closeDayAgendaButton: document.getElementById("btn-close-day-agenda"),
    panelLoading: document.getElementById("panel-loading"),
    panelLoadingTitle: document.getElementById("panel-loading-title"),
    panelLoadingMessage: document.getElementById("panel-loading-message"),
    form: document.getElementById("event-form"),
    formBody: document.getElementById("event-form-body"),
    panelTitle: document.getElementById("panel-title"),
    panelError: document.getElementById("panel-error"),
    remoteUpdateBanner: document.getElementById("remote-update-banner"),
    remoteUpdateMessage: document.getElementById("remote-update-message"),
    applyRemoteDetailButton: document.getElementById("btn-apply-remote-detail"),
    timezone: document.getElementById("calendar-timezone"),
    refreshButton: document.getElementById("btn-refresh-calendar"),
    newButton: document.getElementById("btn-new-event"),
    emptyNewButton: document.getElementById("btn-empty-new-event"),
    emptyTodayButton: document.getElementById("btn-empty-today"),
    upcomingList: document.getElementById("upcoming-list"),
    upcomingEmpty: document.getElementById("upcoming-empty"),
    closeButton: document.getElementById("btn-close-panel"),
    closeLoadingButton: document.getElementById("btn-close-loading"),
    cancelButton: document.getElementById("btn-cancel"),
    saveButton: document.getElementById("btn-save"),
    saveDraftButton: document.getElementById("btn-save-draft"),
    deleteButton: document.getElementById("btn-delete"),
    deleteDraftButton: document.getElementById("btn-delete-draft"),
    unlockDraftButton: document.getElementById("btn-unlock-draft"),
    descriptionEditor: document.getElementById("description-editor"),
    bodyInput: document.getElementById("description-textarea"),
    addAttendeeButton: document.getElementById("add-attendee"),
    attendeesList: document.getElementById("attendees-list"),
    attendeesEmpty: document.getElementById("attendees-empty"),
    reminderNote: document.getElementById("reminder-note"),
    meetingNote: document.getElementById("meeting-note"),
    meetingSummary: document.getElementById("meeting-summary"),
    meetingSettings: document.getElementById("meeting-settings"),
    toast: document.getElementById("toast-container"),
    itemId: document.getElementById("event-item-id"),
    allDay: document.getElementById("is-all-day"),
    sendInvites: document.getElementById("send-invites"),
  };

  var state = {
    calendar: null,
    mode: "empty",
    dirty: false,
    formOperation: null,
    formOperationSeq: 0,
    activeTab: "calendar",
    currentDraftId: "",
    currentDraftVersion: null,
    currentDraftStatus: "",
    currentDraftType: "",
    sourceEventId: "",
    sourceEventSnapshot: null,
    draftItems: [],
    calendarTimezone: BOOT.calendarTimezone,
    selectedDate: "",
    selectedEventId: "",
    visibleEvents: [],
    hasLoadedSuccessfully: false,
    lastLoadFailed: false,
    lastSuccessfulEventInputs: [],
    descriptionEditor: null,
    originalSnapshot: null,
    suppressDescriptionChange: false,
    forceNextNetwork: false,
    detailLoadSeq: 0,
    draftListLoadSeq: 0,
    draftSettingsLoadSeq: 0,
    draftEventLoadSeq: 0,
    calendarLoadSeq: 0,
    retentionOperationSeq: 0,
    retentionSaving: false,
    detailCache: {},
    pendingRemoteDraft: null,
    pendingRemoteItemId: "",
    suppressNextSelect: false,
    nowIndicatorSyncFrame: 0,
    focusMode: "monthGrid",
    keyboardFocusDate: "",
    keyboardFocusEventId: "",
    focusDateAfterRender: "",
    focusReturnTarget: null,
    restoreFocusAfterRender: false,
  };

  var CACHE_VERSION = 3;
  var CACHE_STALE_AFTER_MS = 15 * 60 * 1000;
  var CACHE_MAX_RANGES = 12;
  var DETAIL_CACHE_TTL_MS = 60 * 1000;
  var PAGE_LIMIT = 100;
  var MAX_VIEW_ITEMS = 1000;
  var API_TIMEOUT_MS = 60 * 1000;
  var HALF_HOUR_MINUTES = 30;
  var fieldErrors = {};
  var descriptionEditorPromise = null;
  refs.timezone.textContent = "显示时区：" + state.calendarTimezone;

  function ApiError(code, message, data) {
    this.name = "ApiError";
    this.code = code;
    this.message = message;
    this.data = data || {};
  }
  ApiError.prototype = Object.create(Error.prototype);
  ApiError.prototype.constructor = ApiError;

  function apiBearerToken() {
    var parts = String(BOOT.authPath || "").split("/").filter(Boolean);
    return parts.length ? decodeURIComponent(parts[parts.length - 1]) : "";
  }

  function pad(num) {
    return String(num).padStart(2, "0");
  }

  function wallDateValue(value) {
    return value.getUTCFullYear() + "-" + pad(value.getUTCMonth() + 1) + "-" + pad(value.getUTCDate());
  }

  function wallDateTimeValue(value) {
    return wallDateValue(value) + "T" + pad(value.getUTCHours()) + ":" + pad(value.getUTCMinutes());
  }

  function projectDateTimeParts(value) {
    var result = {};
    new Intl.DateTimeFormat("en-CA", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
      timeZone: state.calendarTimezone,
    }).formatToParts(new Date(value)).forEach(function (part) {
      if (part.type !== "literal") {
        result[part.type] = part.value;
      }
    });
    return result;
  }

  function projectDateValue(value) {
    var parts = projectDateTimeParts(value);
    return parts.year + "-" + parts.month + "-" + parts.day;
  }

  function projectDateTimeValue(value) {
    var parts = projectDateTimeParts(value);
    return parts.year + "-" + parts.month + "-" + parts.day
      + "T" + parts.hour + ":" + parts.minute + ":" + parts.second;
  }

  function roundUpToHalfHour(value) {
    var rounded = new Date(value.getTime());
    var remainder = rounded.getUTCMinutes() % HALF_HOUR_MINUTES;
    if (remainder !== 0) {
      rounded.setUTCMinutes(rounded.getUTCMinutes() + HALF_HOUR_MINUTES - remainder);
    }
    rounded.setUTCSeconds(0, 0);
    return rounded;
  }

  function suggestedEndValue(startValue) {
    var end = new Date(timedApiValue(startValue) + "Z");
    end.setUTCMinutes(end.getUTCMinutes() + HALF_HOUR_MINUTES);
    return wallDateTimeValue(roundUpToHalfHour(end));
  }

  function parseCalendarIsoParts(value) {
    var match = String(value || "").trim().match(/^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2})(?::\d{2}(?:\.\d+)?)?)?/);
    if (!match) {
      return null;
    }
    return { date: match[1], time: match[2] || "" };
  }

  function shiftIsoDateValue(value, days) {
    var parsed = parseCalendarIsoParts(value);
    if (!parsed) {
      return "";
    }
    var shifted = new Date(parsed.date + "T00:00:00Z");
    shifted.setUTCDate(shifted.getUTCDate() + days);
    return wallDateValue(shifted);
  }

  function calendarInputValue(value, allDay) {
    var parsed = parseCalendarIsoParts(value);
    if (!parsed) {
      return "";
    }
    return allDay ? parsed.date : parsed.date + "T" + parsed.time;
  }

  function calendarEndInputValue(value, allDay) {
    if (!allDay) {
      return calendarInputValue(value, false);
    }
    return shiftIsoDateValue(value, -1);
  }

  function shiftDate(value, days) {
    var copy = new Date(value.getTime());
    copy.setUTCDate(copy.getUTCDate() + days);
    return copy;
  }

  function timedApiValue(value) {
    return value.length === 16 ? value + ":00" : value;
  }

  function allDayEndValue(value) {
    return shiftIsoDateValue(value, 1);
  }

  function syncDateInputMode() {
    var nextType = refs.allDay.checked ? "date" : "datetime-local";
    if (refs.form.start.type === nextType) {
      return;
    }
    if (refs.allDay.checked) {
      refs.form.start.value = refs.form.start.value ? refs.form.start.value.slice(0, 10) : "";
      refs.form.end.value = refs.form.end.value ? refs.form.end.value.slice(0, 10) : "";
    } else {
      refs.form.start.value = refs.form.start.value ? refs.form.start.value.slice(0, 10) + "T09:00" : "";
      refs.form.end.value = refs.form.end.value ? refs.form.end.value.slice(0, 10) + "T10:00" : "";
    }
    refs.form.start.type = nextType;
    refs.form.end.type = nextType;
  }

  function syncEndAfterStartChange() {
    if (
      refs.allDay.checked ||
      !refs.form.start.value ||
      !refs.form.end.value ||
      refs.form.start.value < refs.form.end.value
    ) {
      return;
    }
    refs.form.end.value = suggestedEndValue(refs.form.start.value);
  }

  function showNativePicker(event) {
    var input = event.currentTarget;
    if (event.detail === 0 || typeof input.showPicker !== "function") {
      return;
    }
    input.showPicker();
  }

  function eventEnd(event) {
    if (event.end) {
      return event.end;
    }
    return shiftDate(event.start, event.allDay ? 1 : 0);
  }

  function dateFromValue(dateValue) {
    return new Date(String(dateValue || "") + "T00:00:00Z");
  }

  function formatDayAgendaTitle(dateValue) {
    return new Intl.DateTimeFormat("zh-CN", {
      year: "numeric",
      month: "long",
      day: "numeric",
      weekday: "long",
      timeZone: "UTC",
    }).format(dateFromValue(dateValue));
  }

  function eventBoundaryDate(event, boundary) {
    var rawValue = boundary === "start" ? event.startStr : event.endStr;
    var parsed = parseCalendarIsoParts(rawValue);
    if (parsed) {
      return parsed.date;
    }
    var value = boundary === "start" ? event.start : eventEnd(event);
    return value ? wallDateValue(value) : "";
  }

  function eventBoundaryTime(event, boundary) {
    var rawValue = boundary === "start" ? event.startStr : event.endStr;
    var parsed = parseCalendarIsoParts(rawValue);
    if (parsed && parsed.time) {
      return parsed.time;
    }
    var value = boundary === "start" ? event.start : eventEnd(event);
    return value ? wallDateTimeValue(value).slice(11, 16) : "";
  }

  function eventOverlapsDate(event, dateValue) {
    var startDate = eventBoundaryDate(event, "start");
    var endDate = eventBoundaryDate(event, "end") || startDate;
    if (event.allDay) {
      return startDate <= dateValue && dateValue < endDate;
    }
    if (startDate === endDate) {
      return startDate === dateValue;
    }
    if (eventBoundaryTime(event, "end") === "00:00") {
      return startDate <= dateValue && dateValue < endDate;
    }
    return startDate <= dateValue && dateValue <= endDate;
  }

  function formatDayAgendaMeta(event, dateValue) {
    if (event.allDay) {
      return "全天";
    }
    var startDate = eventBoundaryDate(event, "start");
    var endDate = eventBoundaryDate(event, "end") || startDate;
    var startTime = eventBoundaryTime(event, "start");
    var endTime = eventBoundaryTime(event, "end");
    if (startDate === dateValue && endDate === dateValue) {
      return startTime + " - " + endTime;
    }
    if (startDate !== dateValue && endDate === dateValue) {
      return "跨日 · " + endTime + " 结束";
    }
    if (startDate === dateValue && endDate !== dateValue) {
      return startTime + " 开始 · 跨日";
    }
    return "跨日进行中";
  }

  function cloneCalendarItem(item) {
    return Object.assign({}, item || {}, {
      required_attendees: (item && item.required_attendees || []).slice(),
      optional_attendees: (item && item.optional_attendees || []).slice(),
    });
  }

  function cacheDetailItem(item) {
    if (!item || !item.id) {
      return;
    }
    state.detailCache[item.id] = {
      fetchedAt: Date.now(),
      item: cloneCalendarItem(item),
    };
  }

  function cachedDetailItem(itemId) {
    var entry = state.detailCache[itemId];
    if (!entry) {
      return null;
    }
    if (Date.now() - entry.fetchedAt > DETAIL_CACHE_TTL_MS) {
      return null;
    }
    return cloneCalendarItem(entry.item);
  }

  function staleDetailItem(itemId) {
    var entry = state.detailCache[itemId];
    return entry ? cloneCalendarItem(entry.item) : null;
  }

  function clearDetailCache(itemId) {
    if (itemId) {
      delete state.detailCache[itemId];
      return;
    }
    state.detailCache = {};
  }

  function showToast(message, type) {
    var item = document.createElement("div");
    item.className = "toast toast-" + (type || "info");
    item.textContent = message;
    refs.toast.appendChild(item);
    window.setTimeout(function () {
      item.remove();
    }, 3200);
  }

  function showPanelError(message, actionHref, actionLabel) {
    refs.panelError.hidden = false;
    refs.panelError.textContent = "";
    var text = document.createElement("span");
    text.textContent = message;
    refs.panelError.appendChild(text);
    if (actionHref && actionLabel) {
      var link = document.createElement("a");
      link.href = actionHref;
      link.textContent = actionLabel;
      refs.panelError.appendChild(link);
    }
  }

  function clearPanelError() {
    refs.panelError.hidden = true;
    refs.panelError.textContent = "";
  }

  function clearRemoteUpdateNotice() {
    state.pendingRemoteDraft = null;
    state.pendingRemoteItemId = "";
    refs.remoteUpdateBanner.hidden = true;
    refs.remoteUpdateMessage.textContent = "";
    refs.applyRemoteDetailButton.hidden = true;
    delete refs.remoteUpdateBanner.dataset.tone;
  }

  function showRemoteSyncingNotice() {
    state.pendingRemoteDraft = null;
    state.pendingRemoteItemId = "";
    refs.remoteUpdateBanner.hidden = false;
    refs.remoteUpdateMessage.textContent = "正在同步最新详情...";
    refs.applyRemoteDetailButton.hidden = true;
    delete refs.remoteUpdateBanner.dataset.tone;
  }

  function showRemoteUpdatedNotice(itemId, draft) {
    state.pendingRemoteDraft = draft;
    state.pendingRemoteItemId = itemId;
    refs.remoteUpdateBanner.hidden = false;
    refs.remoteUpdateMessage.textContent = "远端详情已更新。";
    refs.applyRemoteDetailButton.hidden = false;
    refs.remoteUpdateBanner.dataset.tone = "updated";
  }

  function setFieldError(name, message) {
    fieldErrors[name] = message;
    var target = document.querySelector('[data-field-error="' + name + '"]');
    if (target) {
      target.textContent = message || "";
    }
  }

  function clearFieldErrors() {
    fieldErrors = {};
    refs.form.querySelectorAll("[data-field-error]").forEach(function (node) {
      node.textContent = "";
    });
  }

  function clearChildren(node) {
    while (node.firstChild) {
      node.removeChild(node.firstChild);
    }
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function plainTextToEditorHtml(value) {
    var normalized = String(value || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
    if (!normalized) {
      return "";
    }
    return normalized.split(/\n{2,}/).map(function (block) {
      return "<p>" + block.split("\n").map(escapeHtml).join("<br>") + "</p>";
    }).join("");
  }

  var HTML_DOCUMENT_MARKER_RE = /<\s*(?:!\s*doctype|\/?html|head|body)\b/i;
  var ESCAPED_HTML_DOCUMENT_MARKER_RE = /&lt;\s*(?:!\s*doctype|\/?html|head|body)\b/i;
  var NON_TEXT_HTML_RE = /<\s*(?:img|table|hr|iframe|video|audio|ul|ol|blockquote|pre|object|embed|svg|canvas)\b/i;

  function decodeHtmlEntities(value) {
    var textarea = document.createElement("textarea");
    textarea.innerHTML = String(value || "");
    return textarea.value;
  }

  function looksLikeHtmlDocument(value) {
    var raw = String(value || "");
    if (HTML_DOCUMENT_MARKER_RE.test(raw)) {
      return true;
    }
    if (!ESCAPED_HTML_DOCUMENT_MARKER_RE.test(raw)) {
      return false;
    }
    return HTML_DOCUMENT_MARKER_RE.test(decodeHtmlEntities(raw));
  }

  function isEmptyHtmlFragment(value) {
    var html = String(value || "");
    if (!html.trim()) {
      return true;
    }
    if (NON_TEXT_HTML_RE.test(html)) {
      return false;
    }
    var parsed = new DOMParser().parseFromString(html, "text/html");
    return !String(parsed.body ? parsed.body.textContent : "").replace(/\u00a0/g, " ").trim();
  }

  function extractHtmlBodyFragment(value) {
    var raw = String(value || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
    if (!raw) {
      return "";
    }
    if (!HTML_DOCUMENT_MARKER_RE.test(raw) && ESCAPED_HTML_DOCUMENT_MARKER_RE.test(raw)) {
      var decoded = decodeHtmlEntities(raw);
      if (HTML_DOCUMENT_MARKER_RE.test(decoded)) {
        raw = decoded;
      }
    }
    var parsed = new DOMParser().parseFromString(raw, "text/html");
    if (parsed.body) {
      parsed.body.querySelectorAll("script, style, meta, link, title").forEach(function (node) {
        node.remove();
      });
      var fragment = normalizeDescriptionHtml(parsed.body.innerHTML);
      return isEmptyHtmlFragment(fragment) ? "" : fragment;
    }
    return normalizeDescriptionHtml(raw);
  }

  function normalizeEditorHtml(value, contentType) {
    var raw = String(value || "");
    if (!raw.trim()) {
      return "";
    }
    if (contentType === "html" || looksLikeHtmlDocument(raw)) {
      return extractHtmlBodyFragment(raw);
    }
    return plainTextToEditorHtml(raw);
  }

  function bodyToEditorHtml(value, contentType) {
    return normalizeEditorHtml(value, contentType);
  }

  function getDescriptionEditor() {
    return state.descriptionEditor || (window.tinymce && window.tinymce.get("description-textarea")) || null;
  }

  function currentDescriptionHtml() {
    var editor = getDescriptionEditor();
    if (editor) {
      var html = normalizeDescriptionHtml(editor.getContent({ format: "html" }));
      if (!editor.getContent({ format: "text" }).trim() && !/<(?:img|table|hr|iframe|video|audio|ul|ol|blockquote|pre)\b/i.test(html)) {
        return "";
      }
      return html;
    }
    return normalizeDescriptionHtml(refs.bodyInput.value);
  }

  function ensureDescriptionEditor() {
    if (descriptionEditorPromise) {
      return descriptionEditorPromise;
    }
    if (!window.tinymce) {
      return Promise.reject(new Error("TinyMCE 未加载。"));
    }
    descriptionEditorPromise = window.tinymce.init({
      target: refs.bodyInput,
      license_key: "gpl",
      promotion: false,
      branding: false,
      menubar: false,
      statusbar: false,
      resize: false,
      height: 190,
      plugins: "autolink advlist lists link table code",
      toolbar: "undo redo | bold italic underline | bullist numlist | link table | removeformat | code",
      toolbar_mode: "wrap",
      contextmenu: false,
      content_style: [
        "body {",
        "  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;",
        "  font-size: 14px;",
        "  line-height: 1.55;",
        "  margin: 12px;",
        "}",
        "p, ul, ol, pre, blockquote, table { margin: 0 0 8px; }",
      ].join("")
    }).then(function (editors) {
      var editor = editors && editors[0] ? editors[0] : (window.tinymce && window.tinymce.get("description-textarea"));
      if (!editor) {
        throw new Error("TinyMCE 初始化失败。");
      }
      state.descriptionEditor = editor;
      editor.on("keydown", handleDescriptionEditorKeydown);
      editor.on("input change undo redo keyup SetContent", function () {
        if (state.suppressDescriptionChange) {
          return;
        }
        editor.save();
        if (state.mode !== "empty") {
          markDirty();
          clearPanelError();
        }
      });
      editor.save();
      editor.show();
      syncDescriptionEditorOperationState();
      return editor;
    }).catch(function (error) {
      descriptionEditorPromise = null;
      throw error;
    });
    return descriptionEditorPromise;
  }

  function setDescriptionHtml(value) {
    var html = normalizeEditorHtml(value, "html");
    refs.bodyInput.value = html;
    return ensureDescriptionEditor().then(function (editor) {
      state.suppressDescriptionChange = true;
      try {
        editor.setContent(html);
        editor.save();
      } finally {
        state.suppressDescriptionChange = false;
      }
      return editor;
    }).catch(function () {
      return null;
    });
  }

  function normalizeDescriptionHtml(value) {
    return String(value || "").trim();
  }

  function arraysEqual(left, right) {
    if (left.length !== right.length) {
      return false;
    }
    for (var index = 0; index < left.length; index += 1) {
      if (left[index] !== right[index]) {
        return false;
      }
    }
    return true;
  }

  function currentFormSnapshot() {
    var attendees = collectAttendees();
    return {
      subject: refs.form.subject.value.trim(),
      location: refs.form.location.value.trim(),
      body: normalizeDescriptionHtml(currentDescriptionHtml()),
      start: refs.allDay.checked ? refs.form.start.value : timedApiValue(refs.form.start.value),
      end: refs.allDay.checked ? allDayEndValue(refs.form.end.value) : timedApiValue(refs.form.end.value),
      isAllDay: !!refs.allDay.checked,
      sendInvites: !!refs.sendInvites.checked,
      reminder: refs.form.reminder_minutes_before_start.value,
      requiredAttendees: attendees.required,
      optionalAttendees: attendees.optional,
    };
  }

  function snapshotsEqual(left, right) {
    if (!left || !right) {
      return false;
    }
    return left.subject === right.subject
      && left.location === right.location
      && left.body === right.body
      && left.start === right.start
      && left.end === right.end
      && left.isAllDay === right.isAllDay
      && left.sendInvites === right.sendInvites
      && left.reminder === right.reminder
      && arraysEqual(left.requiredAttendees, right.requiredAttendees)
      && arraysEqual(left.optionalAttendees, right.optionalAttendees);
  }

  function buildSnapshotFromDraft(draft) {
    var requiredAttendees = [];
    var optionalAttendees = [];
    (draft.rows || []).forEach(function (row) {
      if (!row || !row.email) {
        return;
      }
      if (row.type === "optional") {
        optionalAttendees.push(row.email);
      } else {
        requiredAttendees.push(row.email);
      }
    });
    return {
      subject: String(draft.subject || "").trim(),
      location: String(draft.location || "").trim(),
      body: normalizeDescriptionHtml(draft.body),
      start: draft.isAllDay ? draft.start : timedApiValue(draft.start),
      end: draft.isAllDay ? allDayEndValue(draft.end) : timedApiValue(draft.end),
      isAllDay: !!draft.isAllDay,
      sendInvites: !!draft.sendInvites,
      reminder: String(draft.reminder || ""),
      requiredAttendees: requiredAttendees,
      optionalAttendees: optionalAttendees,
    };
  }

  function syncDirtyState() {
    if (state.mode !== "edit" && state.mode !== "create") {
      state.dirty = false;
      return;
    }
    if (!state.originalSnapshot) {
      state.dirty = true;
      return;
    }
    state.dirty = !snapshotsEqual(currentFormSnapshot(), state.originalSnapshot);
  }

  function setCalendarStatus(message, tone, actionHref, actionLabel) {
    refs.calendarStatus.hidden = false;
    refs.calendarStatus.dataset.tone = tone || "info";
    refs.calendarStatus.textContent = "";
    var text = document.createElement("span");
    text.textContent = message;
    refs.calendarStatus.appendChild(text);
    if (actionHref && actionLabel) {
      var link = document.createElement("a");
      link.href = actionHref;
      link.textContent = actionLabel;
      refs.calendarStatus.appendChild(link);
    }
  }

  function clearCalendarStatus() {
    refs.calendarStatus.hidden = true;
    refs.calendarStatus.textContent = "";
    delete refs.calendarStatus.dataset.tone;
  }

  function setCacheStatus(message, tone) {
    refs.cacheStatus.textContent = message;
    if (tone) {
      refs.cacheStatus.dataset.tone = tone;
      return;
    }
    delete refs.cacheStatus.dataset.tone;
  }

  function formatInstantTime(value) {
    return new Intl.DateTimeFormat("zh-CN", {
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
      timeZone: state.calendarTimezone,
    }).format(new Date(value));
  }

  function formatCacheTimestamp(value) {
    var date = new Date(value);
    var now = new Date();
    var sameDay = projectDateValue(date) === projectDateValue(now);
    return new Intl.DateTimeFormat("zh-CN", {
      month: sameDay ? undefined : "numeric",
      day: sameDay ? undefined : "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
      timeZone: state.calendarTimezone,
    }).format(date);
  }

  function storagePrefix() {
    return [
      "outlook-mcp",
      "calendar",
      "v" + CACHE_VERSION,
      BOOT.account || "unknown",
      state.calendarTimezone,
    ].join(":");
  }

  function storageIndexPrefix() {
    return storagePrefix() + ":";
  }

  function cacheKeyForRange(rangeStart, rangeEnd) {
    return storageIndexPrefix() + rangeStart + ":" + rangeEnd;
  }

  function readStorage(key) {
    try {
      return window.localStorage.getItem(key);
    } catch (error) {
      return null;
    }
  }

  function removeStorage(key) {
    try {
      window.localStorage.removeItem(key);
    } catch (error) {
      return;
    }
  }

  function parseCachedPayload(rawValue) {
    if (!rawValue) {
      return null;
    }
    try {
      var parsed = JSON.parse(rawValue);
      if (
        !parsed
        || typeof parsed !== "object"
        || typeof parsed.range_start !== "string"
        || typeof parsed.range_end !== "string"
        || typeof parsed.fetched_at !== "number"
        || typeof parsed.complete !== "boolean"
        || !Array.isArray(parsed.items)
      ) {
        return null;
      }
      return parsed;
    } catch (error) {
      return null;
    }
  }

  function loadCachedRange(rangeStart, rangeEnd) {
    var key = cacheKeyForRange(rangeStart, rangeEnd);
    var parsed = parseCachedPayload(readStorage(key));
    if (!parsed || parsed.range_start !== rangeStart || parsed.range_end !== rangeEnd) {
      removeStorage(key);
      return null;
    }
    return {
      key: key,
      fetchedAt: parsed.fetched_at,
      items: parsed.items.map(cloneEventInput),
      complete: parsed.complete,
      isExpired: Date.now() - parsed.fetched_at > CACHE_STALE_AFTER_MS,
    };
  }

  function pruneCachedRanges() {
    var keys = [];
    var prefix = storageIndexPrefix();
    try {
      for (var index = 0; index < window.localStorage.length; index += 1) {
        var key = window.localStorage.key(index);
        if (key && key.indexOf(prefix) === 0) {
          var parsed = parseCachedPayload(window.localStorage.getItem(key));
          if (!parsed) {
            removeStorage(key);
          } else {
            keys.push({ key: key, fetchedAt: parsed.fetched_at });
          }
        }
      }
    } catch (error) {
      return;
    }
    keys.sort(function (left, right) {
      return right.fetchedAt - left.fetchedAt;
    });
    keys.slice(CACHE_MAX_RANGES).forEach(function (item) {
      removeStorage(item.key);
    });
  }

  function saveCachedRange(rangeStart, rangeEnd, eventInputs, complete) {
    var key = cacheKeyForRange(rangeStart, rangeEnd);
    try {
      window.localStorage.setItem(key, JSON.stringify({
        range_start: rangeStart,
        range_end: rangeEnd,
        fetched_at: Date.now(),
        complete: complete,
        items: eventInputs.map(cloneEventInput),
      }));
      pruneCachedRanges();
      return true;
    } catch (error) {
      return false;
    }
  }

  function cachedRangeEntries() {
    var ranges = [];
    var prefix = storageIndexPrefix();
    try {
      for (var index = 0; index < window.localStorage.length; index += 1) {
        var key = window.localStorage.key(index);
        if (!key || key.indexOf(prefix) !== 0) {
          continue;
        }
        var parsed = parseCachedPayload(window.localStorage.getItem(key));
        if (parsed) {
          ranges.push({ key: key, payload: parsed });
        }
      }
    } catch (error) {
      return null;
    }
    return ranges;
  }

  function invalidateCachedRanges() {
    var ranges = cachedRangeEntries();
    if (ranges === null) {
      return false;
    }
    var removed = true;
    ranges.forEach(function (range) {
      try {
        window.localStorage.removeItem(range.key);
      } catch (error) {
        removed = false;
      }
    });
    return removed;
  }

  function eventInputIsProjectable(eventInput) {
    if (!eventInput || typeof eventInput.id !== "string" || !eventInput.id.trim()) {
      return false;
    }
    if (eventInput.allDay) {
      return /^\d{4}-\d{2}-\d{2}$/.test(eventInput.start)
        && /^\d{4}-\d{2}-\d{2}$/.test(eventInput.end);
    }
    return Number.isFinite(Date.parse(eventInput.start))
      && Number.isFinite(Date.parse(eventInput.end));
  }

  function eventInputOverlapsRange(eventInput, rangeStart, rangeEnd) {
    if (eventInput.allDay) {
      return eventInput.start < rangeEnd.slice(0, 10)
        && eventInput.end > rangeStart.slice(0, 10);
    }
    return Date.parse(eventInput.start) < Date.parse(rangeEnd)
      && Date.parse(eventInput.end) > Date.parse(rangeStart);
  }

  function projectCalendarMutationIntoCachedRanges(operation, data, previousItemId) {
    if (operation !== "upsert" && operation !== "delete") {
      return false;
    }

    var eventInput = null;
    var affectedIds = [];
    if (operation === "upsert") {
      eventInput = toEvent(data || {});
      if (!eventInputIsProjectable(eventInput)) {
        return false;
      }
      affectedIds.push(String(eventInput.id));
    }
    if (previousItemId) {
      affectedIds.push(String(previousItemId));
    }
    if (operation === "delete" && affectedIds.length === 0) {
      return false;
    }

    var ranges = cachedRangeEntries();
    if (ranges === null) {
      return false;
    }

    var failed = false;
    ranges.forEach(function (range) {
      if (!range.payload.complete) {
        try {
          window.localStorage.removeItem(range.key);
        } catch (error) {
          failed = true;
        }
        return;
      }

      var nextItems = range.payload.items.filter(function (item) {
        return affectedIds.indexOf(String(item.id)) === -1;
      });
      if (
        eventInput
        && eventInputOverlapsRange(
          eventInput,
          range.payload.range_start,
          range.payload.range_end
        )
      ) {
        nextItems.push(cloneEventInput(eventInput));
      }

      try {
        window.localStorage.setItem(range.key, JSON.stringify({
          range_start: range.payload.range_start,
          range_end: range.payload.range_end,
          fetched_at: range.payload.fetched_at,
          complete: true,
          items: nextItems,
        }));
      } catch (error) {
        failed = true;
      }
    });

    if (failed) {
      invalidateCachedRanges();
      return false;
    }
    return true;
  }

  function syncCalendarMutation(operation, data, previousItemId) {
    clearDetailCache(previousItemId || (data && data.id) || null);
    var cacheUpdated = projectCalendarMutationIntoCachedRanges(operation, data, previousItemId);
    if (!cacheUpdated) {
      setCacheStatus("日程操作已成功，正在从服务器同步当前视图...", "loading");
    }
    return cacheUpdated;
  }

  function formatUpcomingMeta(event) {
    var startDate = eventBoundaryDate(event, "start");
    var endDate = eventBoundaryDate(event, "end") || startDate;
    var startTime = eventBoundaryTime(event, "start");
    var endTime = eventBoundaryTime(event, "end");
    if (event.allDay) {
      return startDate + " · 全天";
    }
    var sameDay = startDate === endDate;
    var datePart = new Intl.DateTimeFormat("zh-CN", {
      month: "numeric",
      day: "numeric",
      weekday: "short",
      timeZone: "UTC",
    }).format(dateFromValue(startDate));
    if (sameDay) {
      return datePart + " · " + startTime + " - " + endTime;
    }
    return datePart + " · " + startTime + " - " + endDate;
  }

  function isUpcomingEvent(event, todayDate, nowMillis) {
    var props = event.extendedProps || {};
    if (event.allDay) {
      return props.sourceEnd ? props.sourceEnd > todayDate : props.sourceStart >= todayDate;
    }
    return Date.parse(props.sourceEnd || props.sourceStart) >= nowMillis;
  }

  function compareUpcomingEvents(left, right) {
    return String(left.startStr).localeCompare(String(right.startStr));
  }

  function renderUpcomingEvents(events) {
    clearChildren(refs.upcomingList);
    var now = new Date();
    var todayDate = projectDateValue(now);
    var upcomingEvents = events.filter(function (event) {
      return isUpcomingEvent(event, todayDate, now.getTime());
    }).sort(compareUpcomingEvents).slice(0, 6);

    if (!upcomingEvents.length) {
      refs.upcomingEmpty.hidden = false;
      refs.upcomingEmpty.textContent = state.lastLoadFailed
        ? "当前视图加载失败。"
        : "当前视图没有接下来的日程。";
      return;
    }

    refs.upcomingEmpty.hidden = true;
    upcomingEvents.forEach(function (event) {
      var item = document.createElement("button");
      item.type = "button";
      item.className = "cal-upcoming-item";
      item.addEventListener("click", function () {
        openCalendarEntry(event);
      });

      var title = document.createElement("span");
      title.className = "cal-upcoming-item__title";
      title.textContent = event.title || "(无标题)";
      item.appendChild(title);

      var meta = document.createElement("span");
      meta.className = "cal-upcoming-item__meta";
      meta.textContent = formatUpcomingMeta(event);
      item.appendChild(meta);

      refs.upcomingList.appendChild(item);
    });
  }

  function renderDayAgenda() {
    clearChildren(refs.dayAgendaList);
    if (!state.selectedDate) {
      refs.dayAgendaTitle.textContent = "当天日程";
      refs.dayAgendaSummary.textContent = "";
      refs.dayAgendaEmpty.hidden = true;
      return;
    }

    refs.dayAgendaTitle.textContent = formatDayAgendaTitle(state.selectedDate);
    var dayEvents = state.visibleEvents.filter(function (event) {
      return eventOverlapsDate(event, state.selectedDate);
    }).sort(function (left, right) {
      if (left.allDay !== right.allDay) {
        return left.allDay ? -1 : 1;
      }
      return left.start.getTime() - right.start.getTime();
    });

    refs.dayAgendaSummary.textContent = dayEvents.length
      ? dayEvents.length + " 项日程"
      : "当天没有日程";
    refs.dayAgendaEmpty.hidden = dayEvents.length > 0;

    dayEvents.forEach(function (event, index) {
      var item = document.createElement("button");
      item.type = "button";
      item.className = "cal-day-agenda-item";
      item.id = "cal-day-agenda-item-" + index;
      item.dataset.eventId = event.id;
      if (event.extendedProps && event.extendedProps.entityType === "draft") {
        item.classList.add("cal-event-draft");
      }
      if (event.extendedProps && event.extendedProps.isCancelled) {
        item.classList.add("cal-event-cancelled");
      }
      item.addEventListener("click", function () {
        state.focusReturnTarget = { mode: "dayAgenda", eventId: event.id };
        openCalendarEntry(event);
        if (state.selectedEventId !== event.id) {
          state.focusReturnTarget = null;
        }
      });

      var meta = document.createElement("span");
      meta.className = "cal-day-agenda-item__meta";
      meta.textContent = formatDayAgendaMeta(event, state.selectedDate);
      item.appendChild(meta);

      var title = document.createElement("span");
      title.className = "cal-day-agenda-item__title";
      title.textContent = event.title || "(无标题)";
      item.appendChild(title);

      if (event.extendedProps && event.extendedProps.location) {
        var location = document.createElement("span");
        location.className = "cal-day-agenda-item__location";
        location.textContent = event.extendedProps.location;
        item.appendChild(location);
      }
      refs.dayAgendaList.appendChild(item);
    });
    syncDayAgendaAccessibility();
  }

  function applyRemoteDetailDraft(draft) {
    if (!draft || !draft.id) {
      return;
    }
    if (state.currentDraftId) {
      if (!window.confirm("覆盖为远端详情会删除当前草稿。确认继续？")) {
        return;
      }
      var draftId = state.currentDraftId;
      var draftVersion = state.currentDraftVersion;
      var operationToken = beginFormOperation("remote_replace");
      if (operationToken === null) {
        return;
      }
      apiCall("calendar/drafts/delete", {
        confirm: true,
        draft_id: draftId,
        version: draftVersion,
      }).then(function () {
        if (!isFormOperationActive(operationToken) || state.currentDraftId !== draftId) {
          return;
        }
        openPanel("edit", draft, { force: true });
        showToast("已删除草稿并覆盖为远端详情。", "info");
        refreshDraftViews();
      }).catch(function (error) {
        if (isFormOperationActive(operationToken)) {
          handleApiError(error);
        }
      }).finally(function () {
        endFormOperation(operationToken);
      });
      return;
    }
    openPanel("edit", draft, { force: true });
    showToast("已覆盖为远端详情。", "info");
  }

  function openCalendarEntry(event) {
    var props = event && event.extendedProps || {};
    if (props.entityType === "draft") {
      openPersistedDraft(props.draftRecord);
      return;
    }
    if (event && event.id) {
      openEventDetail(event.id);
    }
  }

  function openEventDetail(itemId) {
    if (!itemId) {
      return;
    }
    var freshItem = cachedDetailItem(itemId);
    if (freshItem) {
      openPanel("edit", buildEditDraftFromItem(freshItem));
      return;
    }

    var event = state.calendar ? state.calendar.getEventById(itemId) : null;
    var staleItem = staleDetailItem(itemId);
    if (staleItem) {
      if (!openPanel("edit", buildEditDraftFromItem(staleItem))) {
        return;
      }
    } else if (event) {
      if (!openPanel("edit", buildEditDraftFromEvent(event))) {
        return;
      }
    } else {
      if (!canDiscard()) {
        return;
      }
      showDetailLoading(itemId);
    }

    state.detailLoadSeq += 1;
    var loadSeq = state.detailLoadSeq;
    if (state.mode === "edit" && state.selectedEventId === itemId) {
      showRemoteSyncingNotice();
    }
    apiCall("calendar/get", {
      item_id: itemId,
      timezone: state.calendarTimezone,
    }).then(function (data) {
      if (loadSeq !== state.detailLoadSeq) {
        return;
      }
      var item = data.item || data;
      cacheDetailItem(item);
      var remoteDraft = buildEditDraftFromItem(item);
      if (state.mode === "loading" && state.selectedEventId === itemId) {
        openPanel("edit", remoteDraft, { force: true });
        return;
      }
      if (state.mode !== "edit" || state.selectedEventId !== itemId) {
        return;
      }
      if (state.dirty || isFormBusy() || state.currentDraftId) {
        showRemoteUpdatedNotice(itemId, remoteDraft);
        showToast("远端详情已更新，可点击覆盖为远端。", "info");
        return;
      }
      openPanel("edit", remoteDraft, { force: true });
    }).catch(function (error) {
      if (loadSeq !== state.detailLoadSeq) {
        return;
      }
      if (state.mode === "edit" && state.selectedEventId === itemId) {
        clearRemoteUpdateNotice();
        showToast("同步最新详情失败：" + error.message, "error");
        return;
      }
      showDetailLoadingError("加载日程详情失败：" + error.message);
      showToast("加载日程详情失败：" + error.message, "error");
    });
  }

  function apiCall(path, payload) {
    var controller = new AbortController();
    var timeoutId = window.setTimeout(function () {
      controller.abort();
    }, API_TIMEOUT_MS);
    return fetch(BOOT.apiBase + "/" + path, {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + apiBearerToken(),
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload || {}),
      signal: controller.signal,
    }).then(function (response) {
      return response.text().then(function (text) {
        var body = {};
        if (text) {
          try {
            body = JSON.parse(text);
          } catch (error) {
            throw new ApiError("network_error", "服务返回了无法解析的响应。");
          }
        }
        if (!body.ok) {
          throw new ApiError(body.error || "unknown", body.message || "请求失败。", body.data || {});
        }
        return body.data || {};
      });
    }).catch(function (error) {
      if (error instanceof ApiError) {
        throw error;
      }
      if (error && error.name === "AbortError") {
        throw new ApiError("network_timeout", "请求超过 60 秒未返回，请稍后重试。");
      }
      throw new ApiError("network_error", error.message || "网络请求失败。");
    }).finally(function () {
      window.clearTimeout(timeoutId);
    });
  }

  function attendeeRows(requiredAttendees, optionalAttendees) {
    var rows = [];
    (requiredAttendees || []).forEach(function (email) {
      rows.push({ email: email, type: "required" });
    });
    (optionalAttendees || []).forEach(function (email) {
      rows.push({ email: email, type: "optional" });
    });
    return rows;
  }

  function calendarItemSnapshot(item) {
    return {
      id: String(item.id || ""),
      changekey: String(item.changekey || ""),
      subject: item.subject || "",
      body: item.body || "",
      start: item.start_local || item.start || "",
      end: item.end_local || item.end || "",
      location: item.location || "",
      is_all_day: !!item.is_all_day,
      is_cancelled: !!item.is_cancelled,
      required_attendees: (item.required_attendees || []).slice(),
      optional_attendees: (item.optional_attendees || []).slice(),
    };
  }

  function buildCreateDraft(start, end, allDay) {
    return {
      id: "",
      subject: "",
      location: "",
      body: "",
      reminder: "",
      isAllDay: !!allDay,
      sendInvites: true,
      rows: [],
      start: allDay ? wallDateValue(start) : wallDateTimeValue(start),
      end: allDay ? wallDateValue(shiftDate(end, -1)) : wallDateTimeValue(end),
      isMeeting: false,
      meetingSettingsOpen: false,
      sourceSnapshot: null,
    };
  }

  function buildCreateDraftFromCalendarRange(startValue, endValue, allDay) {
    return {
      id: "",
      subject: "",
      location: "",
      body: "",
      reminder: "",
      isAllDay: !!allDay,
      sendInvites: true,
      rows: [],
      start: calendarInputValue(startValue, !!allDay),
      end: calendarEndInputValue(endValue, !!allDay),
      isMeeting: false,
      meetingSettingsOpen: false,
      sourceSnapshot: null,
    };
  }

  function eventApiStartIso(event) {
    if (event.startStr) {
      return event.startStr;
    }
    return event.allDay ? wallDateValue(event.start) : timedApiValue(wallDateTimeValue(event.start));
  }

  function eventApiEndIso(event) {
    if (event.endStr) {
      return event.endStr;
    }
    var end = eventEnd(event);
    return event.allDay ? wallDateValue(end) : timedApiValue(wallDateTimeValue(end));
  }

  function buildEditDraftFromItem(item) {
    var startValue = item.start_local || item.start;
    var endValue = item.end_local || item.end;
    var bodyContentType = item.body_content_type || "text";
    var bodySource = item.body;
    if (bodySource == null || bodySource === "") {
      bodySource = item.body_plain || "";
    }
    var bodyHtml = bodyToEditorHtml(bodySource, bodyContentType);
    return {
      id: item.id,
      subject: item.subject || "",
      location: item.location || "",
      body: bodyHtml,
      reminder: "",
      isAllDay: !!item.is_all_day,
      sendInvites: true,
      rows: attendeeRows(item.required_attendees, item.optional_attendees),
      start: calendarInputValue(startValue, !!item.is_all_day),
      end: calendarEndInputValue(endValue, !!item.is_all_day),
      isMeeting: !!item.is_meeting,
      meetingSettingsOpen: false,
      sourceSnapshot: calendarItemSnapshot(item),
    };
  }

  function buildEditDraftFromEvent(event) {
    var props = event.extendedProps || {};
    var draft = buildEditDraftFromItem({
      id: event.id,
      changekey: props.changeKey || "",
      subject: event.title || "",
      body: props.body || "",
      body_plain: props.bodyPlain || "",
      body_content_type: props.bodyContentType || "text",
      start: eventApiStartIso(event),
      end: eventApiEndIso(event),
      location: props.location || "",
      is_all_day: !!event.allDay,
      is_meeting: !!props.isMeeting,
      required_attendees: props.requiredAttendees || [],
      optional_attendees: props.optionalAttendees || [],
    });
    draft.sourceSnapshot = props.sourceSnapshot || draft.sourceSnapshot;
    return draft;
  }

  function buildPanelDraftFromRecord(record) {
    var payload = record.payload || {};
    return {
      id: record.draft_type === "update" ? (record.source_event_id || "") : "",
      subject: payload.subject || "",
      location: payload.location || "",
      body: payload.body || "",
      reminder: payload.reminder_minutes_before_start == null ? "" : String(payload.reminder_minutes_before_start),
      isAllDay: !!payload.is_all_day,
      sendInvites: payload.send_invites !== false,
      rows: attendeeRows(payload.required_attendees, payload.optional_attendees),
      start: payload.start || "",
      end: payload.end || "",
      isMeeting: (payload.required_attendees || []).length + (payload.optional_attendees || []).length > 0,
      meetingSettingsOpen: (payload.required_attendees || []).length + (payload.optional_attendees || []).length > 0,
      sourceSnapshot: record.base_snapshot || null,
      draftId: record.id,
      draftVersion: record.version,
      draftStatus: record.status,
      draftType: record.draft_type,
    };
  }

  function openPersistedDraft(record) {
    if (!record) {
      return;
    }
    if (!record.payload) {
      state.detailLoadSeq += 1;
      var loadSeq = state.detailLoadSeq;
      apiCall("calendar/drafts/get", { draft_id: record.id }).then(function (data) {
        if (loadSeq !== state.detailLoadSeq) {
          return;
        }
        openPersistedDraft(data.draft);
      }).catch(function (error) {
        if (loadSeq === state.detailLoadSeq) {
          showToast("加载草稿失败：" + error.message, "error");
        }
      });
      return;
    }
    var mode = record.draft_type === "update" ? "edit" : "create";
    openPanel(mode, buildPanelDraftFromRecord(record));
    if (record.status === "outcome_unknown") {
      showPanelError("上次提交结果未知。请先在 Outlook 中核对，当前禁止再次提交。", null, null);
    } else if (record.last_error_message) {
      showPanelError(record.last_error_message);
    }
  }

  function renderAttendeeRow(entry) {
    var row = document.createElement("div");
    row.className = "attendee-row";

    var email = document.createElement("input");
    email.type = "email";
    email.className = "attendee-email";
    email.placeholder = "name@example.com";
    email.value = entry.email || "";
    row.appendChild(email);

    var select = document.createElement("select");
    select.className = "attendee-type";
    [["required", "必需"], ["optional", "可选"]].forEach(function (item) {
      var option = document.createElement("option");
      option.value = item[0];
      option.textContent = item[1];
      option.selected = item[0] === entry.type;
      select.appendChild(option);
    });
    row.appendChild(select);

    var remove = document.createElement("button");
    remove.type = "button";
    remove.className = "cal-icon-button";
    remove.textContent = "×";
    remove.setAttribute("aria-label", "移除参会人");
    remove.addEventListener("click", function () {
      row.remove();
      syncFormState();
      markDirty();
    });
    row.appendChild(remove);
    refs.attendeesList.appendChild(row);
  }

  function renderAttendees(rows) {
    refs.attendeesList.textContent = "";
    rows.forEach(renderAttendeeRow);
    refs.attendeesEmpty.hidden = rows.length > 0;
  }

  function collectAttendees() {
    var required = [];
    var optional = [];
    refs.attendeesList.querySelectorAll(".attendee-row").forEach(function (row) {
      var email = row.querySelector(".attendee-email").value.trim();
      var type = row.querySelector(".attendee-type").value;
      if (!email) {
        return;
      }
      if (type === "optional") {
        optional.push(email);
      } else {
        required.push(email);
      }
    });
    return { required: required, optional: optional };
  }

  function setPanelOpen(open) {
    refs.panel.dataset.open = open ? "1" : "0";
    refs.backdrop.hidden = !open;
  }

  function syncSelectedEventClass() {
    refs.calendar.querySelectorAll(".fc-event[data-event-id]").forEach(function (node) {
      node.classList.toggle("cal-event-selected", node.dataset.eventId === state.selectedEventId);
    });
  }

  function syncSelectedDateClass() {
    refs.calendar.querySelectorAll(".fc-daygrid-day[data-date]").forEach(function (node) {
      node.classList.toggle("cal-day-selected", node.dataset.date === state.selectedDate);
    });
  }

  function isMonthView() {
    return !!(state.calendar && state.calendar.view.type === "dayGridMonth");
  }

  function monthDateCellFor(dateValue) {
    var cells = refs.calendar.querySelectorAll(".fc-daygrid-day[data-date]");
    for (var index = 0; index < cells.length; index += 1) {
      if (cells[index].dataset.date === dateValue) {
        return cells[index];
      }
    }
    return null;
  }

  function formatKeyboardDateLabel(dateValue) {
    var label = new Intl.DateTimeFormat("zh-CN", {
      year: "numeric",
      month: "long",
      day: "numeric",
      weekday: "long",
      timeZone: "UTC",
    }).format(dateFromValue(dateValue));
    return dateValue === projectDateValue(new Date()) ? label + "，今天" : label;
  }

  function syncMonthGridAccessibility() {
    if (!isMonthView()) {
      return;
    }
    var cells = Array.prototype.slice.call(refs.calendar.querySelectorAll(".fc-daygrid-day[data-date]"));
    if (!cells.length) {
      return;
    }

    var activeDate = state.keyboardFocusDate || wallDateValue(state.calendar.getDate());
    if (!monthDateCellFor(activeDate)) {
      activeDate = cells[0].dataset.date;
      state.keyboardFocusDate = activeDate;
    }
    var grid = refs.calendar.querySelector(".fc-daygrid-body table");
    if (grid) {
      grid.setAttribute("aria-label", "月视图日期");
    }

    cells.forEach(function (cell) {
      var isActive = cell.dataset.date === activeDate;
      var isSelected = cell.dataset.date === state.selectedDate;
      cell.tabIndex = isActive ? 0 : -1;
      cell.setAttribute("aria-label", formatKeyboardDateLabel(cell.dataset.date));
      cell.setAttribute("aria-selected", isSelected ? "true" : "false");
      if (cell.dataset.date === projectDateValue(new Date())) {
        cell.setAttribute("aria-current", "date");
      } else {
        cell.removeAttribute("aria-current");
      }
      cell.classList.toggle("cal-day-keyboard-active", isActive);
      cell.querySelectorAll(".fc-daygrid-event, .fc-daygrid-more-link").forEach(function (node) {
        node.tabIndex = -1;
      });
    });
  }

  function syncDayAgendaAccessibility() {
    var items = Array.prototype.slice.call(refs.dayAgendaList.querySelectorAll(".cal-day-agenda-item"));
    refs.dayAgendaList.setAttribute("role", "list");
    if (!items.length) {
      state.keyboardFocusEventId = "";
      return;
    }

    var activeItem = items.find(function (item) {
      return item.dataset.eventId === state.keyboardFocusEventId;
    }) || items[0];
    state.keyboardFocusEventId = activeItem.dataset.eventId;
    items.forEach(function (item) {
      var isActive = item === activeItem;
      item.tabIndex = isActive ? 0 : -1;
      item.classList.toggle("cal-agenda-item-active", isActive);
    });
  }

  function focusDayAgenda(preferredEventId) {
    state.focusMode = "dayAgenda";
    var items = Array.prototype.slice.call(refs.dayAgendaList.querySelectorAll(".cal-day-agenda-item"));
    var target = items.find(function (item) {
      return item.dataset.eventId === (preferredEventId || state.keyboardFocusEventId);
    }) || items[0];
    if (!target) {
      state.keyboardFocusEventId = "";
      syncDayAgendaAccessibility();
      refs.dayNewButton.focus();
      return;
    }
    state.keyboardFocusEventId = target.dataset.eventId;
    syncDayAgendaAccessibility();
    target.focus();
  }

  function focusMonthDate(dateValue, navigate) {
    if (!dateValue || !isMonthView()) {
      return;
    }
    state.focusMode = "monthGrid";
    state.keyboardFocusDate = dateValue;
    var cell = monthDateCellFor(dateValue);
    if (!cell && navigate) {
      state.focusDateAfterRender = dateValue;
      state.calendar.gotoDate(dateFromValue(dateValue));
      return;
    }
    syncMonthGridAccessibility();
    cell = monthDateCellFor(dateValue);
    if (cell) {
      cell.focus();
    }
  }

  function shiftMonthDateValue(dateValue, months) {
    var source = dateFromValue(dateValue);
    var day = source.getUTCDate();
    var target = new Date(Date.UTC(source.getUTCFullYear(), source.getUTCMonth() + months, 1));
    var lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
    target.setUTCDate(Math.min(day, lastDay));
    return wallDateValue(target);
  }

  function captureFocusReturnTarget() {
    if (state.focusMode === "dayAgenda" && state.selectedDate) {
      return { mode: "dayAgenda", eventId: state.keyboardFocusEventId };
    }
    return {
      mode: "monthGrid",
      date: state.keyboardFocusDate || (state.calendar ? wallDateValue(state.calendar.getDate()) : ""),
    };
  }

  function restoreFocusTarget(target) {
    if (!target || target.focus === false) {
      return;
    }
    if (target.mode === "dayAgenda" && state.selectedDate) {
      state.focusMode = "dayAgenda";
      state.keyboardFocusEventId = target.eventId || "";
      focusDayAgenda(target.eventId);
      state.restoreFocusAfterRender = true;
      return;
    }
    if (target.mode === "monthGrid") {
      focusMonthDate(target.date || state.keyboardFocusDate, false);
    }
  }

  function clearWeeklyNowIndicatorStyles() {
    refs.calendar.querySelectorAll(".cal-timegrid-now-indicator-expanded").forEach(function (container) {
      container.classList.remove("cal-timegrid-now-indicator-expanded");
    });
    refs.calendar.querySelectorAll(".fc-timegrid-now-indicator-line").forEach(function (line) {
      line.style.removeProperty("left");
      line.style.removeProperty("right");
    });
  }

  function syncWeeklyNowIndicator() {
    if (!state.calendar || state.calendar.view.type !== "timeGridWeek") {
      clearWeeklyNowIndicatorStyles();
      return;
    }

    var line = refs.calendar.querySelector(".fc-timegrid-now-indicator-line");
    var columns = Array.prototype.slice.call(refs.calendar.querySelectorAll(".fc-timegrid-col[data-date]"));
    var indicatorContainer = line && line.closest(".fc-timegrid-now-indicator-container");
    var positioningParent = line && (line.offsetParent || line.parentElement);
    if (!line || !indicatorContainer || !positioningParent || !columns.length) {
      clearWeeklyNowIndicatorStyles();
      return;
    }

    var firstColumnRect = columns[0].getBoundingClientRect();
    var lastColumnRect = columns[columns.length - 1].getBoundingClientRect();
    var parentRect = positioningParent.getBoundingClientRect();
    indicatorContainer.classList.add("cal-timegrid-now-indicator-expanded");
    line.style.left = (firstColumnRect.left - parentRect.left) + "px";
    line.style.right = (parentRect.right - lastColumnRect.right) + "px";
  }

  function scheduleWeeklyNowIndicatorSync() {
    if (state.nowIndicatorSyncFrame) {
      window.cancelAnimationFrame(state.nowIndicatorSyncFrame);
    }
    state.nowIndicatorSyncFrame = window.requestAnimationFrame(function () {
      state.nowIndicatorSyncFrame = 0;
      syncWeeklyNowIndicator();
    });
  }

  function setSelectedEventId(itemId) {
    state.selectedEventId = itemId || "";
    syncSelectedEventClass();
  }

  function renderDayAgendaPanel(open) {
    state.mode = "dayAgenda";
    state.focusMode = "dayAgenda";
    state.dirty = false;
    state.originalSnapshot = null;
    clearFieldErrors();
    clearPanelError();
    clearRemoteUpdateNotice();
    refs.panelEmpty.hidden = true;
    refs.panelDayAgenda.hidden = false;
    refs.panelLoading.hidden = true;
    refs.form.hidden = true;
    setSelectedEventId("");
    renderDayAgenda();
    setPanelOpen(open);
  }

  function selectCalendarDate(dateValue, options) {
    if (!dateValue || !canDiscard()) {
      return;
    }
    if (state.selectedDate === dateValue && state.mode === "dayAgenda" && refs.panel.dataset.open === "1") {
      state.selectedDate = "";
      syncSelectedDateClass();
      resetPanel();
      return;
    }
    state.detailLoadSeq += 1;
    state.selectedDate = dateValue;
    state.keyboardFocusDate = dateValue;
    state.keyboardFocusEventId = "";
    state.focusMode = "dayAgenda";
    state.focusReturnTarget = null;
    syncSelectedDateClass();
    renderDayAgendaPanel(true);
    if (options && options.focusAgenda) {
      focusDayAgenda();
    }
  }

  function clearSelectedDate(options) {
    if (!canDiscard()) {
      return;
    }
    var returnDate = state.keyboardFocusDate || state.selectedDate;
    state.selectedDate = "";
    syncSelectedDateClass();
    state.focusReturnTarget = {
      mode: "monthGrid",
      date: returnDate,
      focus: !!(options && options.focus),
    };
    resetPanel();
  }

  function showDetailLoading(itemId) {
    state.mode = "loading";
    state.dirty = false;
    state.originalSnapshot = null;
    clearFieldErrors();
    clearPanelError();
    setSelectedEventId(itemId);
    refs.panelEmpty.hidden = true;
    refs.form.hidden = true;
    refs.panelLoading.hidden = false;
    refs.panelLoadingTitle.textContent = "加载日程...";
    refs.panelLoadingMessage.textContent = "正在读取日程详情...";
    setPanelOpen(true);
  }

  function showDetailLoadingError(message) {
    refs.panelLoading.hidden = false;
    refs.panelLoadingTitle.textContent = "加载失败";
    refs.panelLoadingMessage.textContent = message;
  }

  function fillForm(draft) {
    refs.itemId.value = draft.id || "";
    refs.form.subject.value = draft.subject;
    refs.form.location.value = draft.location;
    refs.bodyInput.value = draft.body;
    refs.form.reminder_minutes_before_start.value = draft.reminder;
    refs.allDay.checked = draft.isAllDay;
    refs.form.start.type = draft.isAllDay ? "date" : "datetime-local";
    refs.form.end.type = draft.isAllDay ? "date" : "datetime-local";
    refs.form.start.value = draft.start;
    refs.form.end.value = draft.end;
    refs.sendInvites.checked = draft.sendInvites;
    refs.meetingSettings.open = !!draft.meetingSettingsOpen;
    renderAttendees(draft.rows);
    state.originalSnapshot = buildSnapshotFromDraft(draft);
    setDescriptionHtml(draft.body).then(function () {
      state.originalSnapshot = currentFormSnapshot();
      syncDirtyState();
    });
    syncFormState();
  }

  function openPanel(mode, draft, options) {
    var force = !!(options && options.force);
    var keepRemoteNotice = !!(options && options.keepRemoteNotice);
    if (!force && !canDiscard()) {
      return false;
    }
    if (!state.focusReturnTarget) {
      state.focusReturnTarget = captureFocusReturnTarget();
    }
    state.detailLoadSeq += 1;
    state.mode = mode;
    state.focusMode = "eventEditor";
    state.dirty = false;
    state.currentDraftId = draft.draftId || "";
    state.currentDraftVersion = draft.draftVersion == null ? null : draft.draftVersion;
    state.currentDraftStatus = draft.draftStatus || "";
    state.currentDraftType = draft.draftType || (mode === "edit" ? "update" : "create");
    state.sourceEventId = mode === "edit" ? String(draft.id || "") : "";
    state.sourceEventSnapshot = draft.sourceSnapshot || null;
    setSelectedEventId(mode === "edit" ? draft.id : "");
    clearFieldErrors();
    clearPanelError();
    if (!keepRemoteNotice) {
      clearRemoteUpdateNotice();
    }
    refs.panelLoading.hidden = true;
    refs.panelEmpty.hidden = true;
    refs.panelDayAgenda.hidden = true;
    refs.form.hidden = false;
    refs.panelTitle.textContent = state.currentDraftId
      ? (mode === "edit" ? "编辑日程草稿" : "新建日程草稿")
      : (mode === "edit" ? "编辑日程" : "新建日程");
    refs.deleteButton.hidden = mode !== "edit" || !!state.currentDraftId;
    refs.deleteDraftButton.hidden = !state.currentDraftId;
    refs.unlockDraftButton.hidden = state.currentDraftStatus !== "outcome_unknown";
    state.originalSnapshot = null;
    fillForm(draft);
    setPanelOpen(true);
    refs.formBody.scrollTop = 0;
    refs.form.subject.focus();
    return true;
  }

  function resetPanel() {
    var returnTarget = state.focusReturnTarget;
    state.focusReturnTarget = null;
    state.detailLoadSeq += 1;
    refs.form.reset();
    refs.panelEmpty.hidden = false;
    refs.panelDayAgenda.hidden = true;
    refs.panelLoading.hidden = true;
    refs.form.hidden = true;
    refs.attendeesList.textContent = "";
    refs.attendeesEmpty.hidden = false;
    refs.deleteButton.hidden = true;
    refs.deleteDraftButton.hidden = true;
    refs.unlockDraftButton.hidden = true;
    clearFieldErrors();
    clearPanelError();
    clearRemoteUpdateNotice();
    state.mode = "empty";
    state.dirty = false;
    setSelectedEventId("");
    state.originalSnapshot = null;
    state.currentDraftId = "";
    state.currentDraftVersion = null;
    state.currentDraftStatus = "";
    state.currentDraftType = "";
    state.sourceEventId = "";
    state.sourceEventSnapshot = null;
    setDescriptionHtml("");
    if (state.selectedDate) {
      renderDayAgendaPanel(false);
      restoreFocusTarget(returnTarget);
      return;
    }
    setPanelOpen(false);
    restoreFocusTarget(returnTarget);
  }

  function canDiscard() {
    return !isFormBusy() && (!state.dirty || window.confirm("放弃未保存的更改？"));
  }

  function isTextEntryTarget(target) {
    return !!(target && typeof target.closest === "function" && target.closest("input, textarea, select, [contenteditable=\"true\"]"));
  }

  function isDraftShortcut(event) {
    return (event.ctrlKey || event.metaKey) && !event.altKey && String(event.key || "").toLowerCase() === "s";
  }

  function isSubmitShortcut(event) {
    return (event.ctrlKey || event.metaKey) && !event.altKey && event.key === "Enter";
  }

  function submitFormFromKeyboard() {
    if (isFormBusy()) {
      return;
    }
    if (typeof refs.form.requestSubmit === "function") {
      refs.form.requestSubmit();
    } else {
      refs.saveButton.click();
    }
  }

  function saveDraftFromKeyboard() {
    if (isFormBusy()) {
      return;
    }
    if (state.currentDraftStatus === "outcome_unknown" || state.currentDraftStatus === "submitting") {
      showPanelError("该草稿的提交结果未知，核对 Outlook 前不能继续暂存或提交。");
      return;
    }
    if (state.mode === "edit" && !state.currentDraftId && !state.dirty) {
      showToast("没有变更需要暂存。", "info");
      return;
    }
    var operationToken = beginFormOperation("draft_save");
    if (operationToken === null) {
      return;
    }
    persistCurrentDraft().catch(function (error) {
      if (error.code === "draft_not_found" && state.currentDraftId && window.confirm(
        "该草稿已过期或被另一设备删除。是否将当前内容另存为新草稿？"
      )) {
        state.currentDraftId = "";
        state.currentDraftVersion = null;
        state.currentDraftStatus = "";
        return persistCurrentDraft();
      }
      handleApiError(error);
      return null;
    }).finally(function () {
      endFormOperation(operationToken);
    });
  }

  function handleEditShortcut(event) {
    if (event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || String(event.key || "").toLowerCase() !== "e") {
      return false;
    }
    var draftRow = event.target && event.target.closest ? event.target.closest(".cal-draft-row__open") : null;
    if (draftRow) {
      event.preventDefault();
      draftRow.click();
      return true;
    }
    var agendaItem = event.target && event.target.closest ? event.target.closest(".cal-day-agenda-item") : null;
    if (agendaItem) {
      event.preventDefault();
      agendaItem.click();
      return true;
    }
    var calendarEvent = event.target && event.target.closest ? event.target.closest(".fc-event[data-event-id]") : null;
    if (calendarEvent && state.calendar) {
      var entry = state.calendar.getEventById(calendarEvent.dataset.eventId);
      if (entry) {
        event.preventDefault();
        openCalendarEntry(entry);
        return true;
      }
    }
    return false;
  }

  function handlePanelEscape(event) {
    if (state.mode === "edit" || state.mode === "create") {
      if (!canDiscard()) {
        return true;
      }
      event.preventDefault();
      resetPanel();
      return true;
    }
    if (state.mode === "dayAgenda") {
      event.preventDefault();
      clearSelectedDate({ focus: true });
      return true;
    }
    if (state.mode === "loading") {
      event.preventDefault();
      resetPanel();
      return true;
    }
    return false;
  }

  function handleMonthGridKeydown(event, cell) {
    var dateValue = cell.dataset.date;
    var nextDate = "";
    var current = dateFromValue(dateValue);
    switch (event.key) {
      case "ArrowLeft":
        current.setUTCDate(current.getUTCDate() - 1);
        nextDate = wallDateValue(current);
        break;
      case "ArrowRight":
        current.setUTCDate(current.getUTCDate() + 1);
        nextDate = wallDateValue(current);
        break;
      case "ArrowUp":
        current.setUTCDate(current.getUTCDate() - 7);
        nextDate = wallDateValue(current);
        break;
      case "ArrowDown":
        current.setUTCDate(current.getUTCDate() + 7);
        nextDate = wallDateValue(current);
        break;
      case "Home":
        current.setUTCDate(current.getUTCDate() - ((current.getUTCDay() + 6) % 7));
        nextDate = wallDateValue(current);
        break;
      case "End":
        current.setUTCDate(current.getUTCDate() + (6 - ((current.getUTCDay() + 6) % 7)));
        nextDate = wallDateValue(current);
        break;
      case "PageUp":
        nextDate = shiftMonthDateValue(dateValue, event.shiftKey ? -12 : -1);
        break;
      case "PageDown":
        nextDate = shiftMonthDateValue(dateValue, event.shiftKey ? 12 : 1);
        break;
      case "Enter":
      case " ":
        event.preventDefault();
        selectCalendarDate(dateValue, { focusAgenda: true });
        return;
      default:
        return;
    }
    event.preventDefault();
    focusMonthDate(nextDate, true);
  }

  function moveDayAgendaFocus(direction) {
    var items = Array.prototype.slice.call(refs.dayAgendaList.querySelectorAll(".cal-day-agenda-item"));
    if (!items.length) {
      return;
    }
    var currentIndex = -1;
    for (var index = 0; index < items.length; index += 1) {
      if (items[index] === document.activeElement || items[index].dataset.eventId === state.keyboardFocusEventId) {
        currentIndex = index;
        break;
      }
    }
    if (currentIndex < 0) {
      focusDayAgenda();
      return;
    }
    var nextIndex = currentIndex + direction;
    if (nextIndex < 0 || nextIndex >= items.length) {
      return;
    }
    focusDayAgenda(items[nextIndex].dataset.eventId);
  }

  function handleDayAgendaKeydown(event) {
    var item = event.target && event.target.closest ? event.target.closest(".cal-day-agenda-item") : null;
    if (!item) {
      return;
    }
    switch (event.key) {
      case "ArrowUp":
        event.preventDefault();
        moveDayAgendaFocus(-1);
        break;
      case "ArrowDown":
        event.preventDefault();
        moveDayAgendaFocus(1);
        break;
      case "Home":
        event.preventDefault();
        focusDayAgenda();
        break;
      case "End":
        event.preventDefault();
        var items = refs.dayAgendaList.querySelectorAll(".cal-day-agenda-item");
        if (items.length) {
          focusDayAgenda(items[items.length - 1].dataset.eventId);
        }
        break;
      default:
        break;
    }
  }

  function handleCalendarKeydown(event) {
    if (isDraftShortcut(event) && (state.mode === "edit" || state.mode === "create")) {
      event.preventDefault();
      saveDraftFromKeyboard();
      return;
    }
    if (isSubmitShortcut(event) && (state.mode === "edit" || state.mode === "create")) {
      event.preventDefault();
      submitFormFromKeyboard();
      return;
    }
    if (event.key === "Escape" && state.mode !== "empty") {
      handlePanelEscape(event);
      return;
    }
    if (isTextEntryTarget(event.target)) {
      return;
    }
    if (handleEditShortcut(event)) {
      return;
    }
    var cell = event.target && event.target.closest ? event.target.closest(".fc-daygrid-day[data-date]") : null;
    if (state.mode !== "edit" && state.mode !== "create" && state.mode !== "loading" && isMonthView() && cell && !(event.target.closest && event.target.closest(".fc-event, .fc-daygrid-more-link"))) {
      handleMonthGridKeydown(event, cell);
      return;
    }
    if (state.mode === "dayAgenda") {
      handleDayAgendaKeydown(event);
    }
  }

  function handleDescriptionEditorKeydown(event) {
    if (isDraftShortcut(event) && (state.mode === "edit" || state.mode === "create")) {
      event.preventDefault();
      saveDraftFromKeyboard();
      return;
    }
    if (isSubmitShortcut(event) && (state.mode === "edit" || state.mode === "create")) {
      event.preventDefault();
      submitFormFromKeyboard();
      return;
    }
    if (event.key === "Escape" && state.mode !== "empty") {
      handlePanelEscape(event);
    }
  }

  function markDirty() {
    syncDirtyState();
  }

  function syncFormState() {
    var attendees = collectAttendees();
    var rowCount = refs.attendeesList.querySelectorAll(".attendee-row").length;
    var attendeeCount = attendees.required.length + attendees.optional.length;
    refs.attendeesEmpty.hidden = rowCount > 0;
    refs.sendInvites.disabled = state.mode === "edit" || attendeeCount > 0;
    refs.sendInvites.checked = state.mode === "edit" || attendeeCount > 0 ? true : refs.sendInvites.checked;
    if (attendeeCount > 0 && (state.mode !== "edit" || refs.meetingSettings.open)) {
      refs.meetingSettings.open = true;
    }
    refs.meetingSummary.textContent = attendeeCount > 0
      ? "已添加 " + attendeeCount + " 位参会人，Graph 会自动发送邀请。"
      : state.mode === "edit"
        ? "需要时可展开，继续管理提醒和邀请设置。"
        : "私人日程可保持收起。";
    refs.meetingNote.textContent = attendeeCount > 0
      ? "只要存在参会人，Graph 就会自动发送邀请。"
      : state.mode === "edit"
        ? "Graph 更新日程时不支持 send_invites=false。"
        : "没有参会人时，你可以关闭邀请，保存为私人日程。";
    refs.form.reminder_minutes_before_start.disabled = state.mode === "edit";
    refs.reminderNote.hidden = state.mode !== "edit";
    var blockedDraft = state.currentDraftStatus === "outcome_unknown" || state.currentDraftStatus === "submitting";
    refs.saveButton.disabled = isFormBusy() || blockedDraft;
    refs.saveDraftButton.disabled = isFormBusy() || blockedDraft;
    markDirty();
  }

  function validateForm() {
    var valid = true;
    clearFieldErrors();
    clearPanelError();
    if (!refs.form.subject.value.trim()) {
      setFieldError("subject", "请填写主题。");
      valid = false;
    }
    if (!refs.form.start.value) {
      setFieldError("start", "请填写开始时间。");
      valid = false;
    }
    if (!refs.form.end.value) {
      setFieldError("end", "请填写结束时间。");
      valid = false;
    }
    if (refs.form.start.value && refs.form.end.value) {
      if (refs.allDay.checked ? refs.form.end.value < refs.form.start.value : refs.form.end.value <= refs.form.start.value) {
        setFieldError("end", refs.allDay.checked ? "结束日期不能早于开始日期。" : "结束时间必须晚于开始时间。");
        valid = false;
      }
    }
    refs.attendeesList.querySelectorAll(".attendee-email").forEach(function (input) {
      if (input.value.trim() && !input.checkValidity()) {
        setFieldError("attendees", "每位参会人都需要填写合法的邮箱地址。");
        valid = false;
      }
    });
    return valid;
  }

  function draftPayloadForForm() {
    var attendees = collectAttendees();
    var reminder = refs.form.reminder_minutes_before_start.value.trim();
    return {
      subject: refs.form.subject.value,
      body: currentDescriptionHtml(),
      location: refs.form.location.value,
      start: refs.form.start.value,
      end: refs.form.end.value,
      is_all_day: !!refs.allDay.checked,
      timezone: state.calendarTimezone,
      send_invites: !!refs.sendInvites.checked,
      required_attendees: attendees.required,
      optional_attendees: attendees.optional,
      reminder_minutes_before_start: reminder === "" ? null : reminder,
    };
  }

  function applySavedDraftContext(record) {
    state.currentDraftId = record.id;
    state.currentDraftVersion = record.version;
    state.currentDraftStatus = record.status;
    state.currentDraftType = record.draft_type;
    refs.panelTitle.textContent = state.mode === "edit" ? "编辑日程草稿" : "新建日程草稿";
    refs.deleteButton.hidden = true;
    refs.deleteDraftButton.hidden = false;
    refs.unlockDraftButton.hidden = record.status !== "outcome_unknown";
    state.originalSnapshot = currentFormSnapshot();
    state.dirty = false;
    syncFormState();
  }

  function persistCurrentDraft(options) {
    if (state.mode !== "edit" && state.mode !== "create") {
      return Promise.reject(new ApiError("validation_error", "当前没有可暂存的日程。"));
    }
    var params = {
      confirm: true,
      draft_type: state.mode === "edit" ? "update" : "create",
      source_event_id: state.mode === "edit" ? (state.sourceEventId || refs.itemId.value) : null,
      base_snapshot: state.mode === "edit" ? state.sourceEventSnapshot : null,
      payload: draftPayloadForForm(),
    };
    if (state.currentDraftId) {
      params.draft_id = state.currentDraftId;
      params.version = state.currentDraftVersion;
    }
    return apiCall("calendar/drafts/save", params).then(function (data) {
      var record = data.draft;
      applySavedDraftContext(record);
      if (!options || !options.silent) {
        showToast("草稿已跨设备暂存。", "success");
      }
      refreshDraftViews();
      return record;
    });
  }

  function applyDraftErrorContext(error) {
    var record = error && error.data ? error.data.draft : null;
    if (!record || record.id !== state.currentDraftId) {
      return;
    }
    state.currentDraftVersion = record.version;
    state.currentDraftStatus = record.status;
    if (error.code === "remote_conflict" && error.data.remote && state.sourceEventId) {
      showRemoteUpdatedNotice(state.sourceEventId, buildEditDraftFromItem(error.data.remote));
    }
    syncFormState();
  }

  function convertCurrentDraftToCreate() {
    return apiCall("calendar/drafts/convert-to-create", {
      confirm: true,
      draft_id: state.currentDraftId,
      version: state.currentDraftVersion,
    }).then(function (data) {
      state.mode = "create";
      state.sourceEventId = "";
      state.sourceEventSnapshot = null;
      refs.itemId.value = "";
      setSelectedEventId("");
      applySavedDraftContext(data.draft);
      showToast("已转为新建日程草稿。", "success");
      refreshDraftViews();
    });
  }

  function submitCurrentForm() {
    if (isFormBusy()) {
      return;
    }
    if (state.currentDraftStatus === "outcome_unknown" || state.currentDraftStatus === "submitting") {
      showPanelError("该草稿的提交结果未知，核对 Outlook 前不能再次提交。");
      return;
    }
    if (state.mode === "edit" && !state.currentDraftId && !state.dirty) {
      showToast("没有变更需要保存。", "info");
      return;
    }
    if (!validateForm()) {
      return;
    }
    var wasEditing = state.mode === "edit";
    var previousItemId = state.sourceEventId || refs.itemId.value || null;
    var operationToken = beginFormOperation("calendar_submit");
    if (operationToken === null) {
      return;
    }
    persistCurrentDraft({ silent: true }).then(function (record) {
      return apiCall("calendar/drafts/submit", {
        confirm: true,
        draft_id: record.id,
        version: record.version,
      });
    }).then(function (data) {
      var calendarData = data.calendar || {};
      var cacheUpdated = calendarData.id
        ? syncCalendarMutation("upsert", calendarData, previousItemId)
        : false;
      showToast(wasEditing ? "已更新日程。" : "已创建日程。", "success");
      resetPanel();
      refreshDraftViews();
      refreshCalendar({ forceNetwork: !cacheUpdated });
    }).catch(function (error) {
      applyDraftErrorContext(error);
      if (error.code === "draft_not_found" && state.currentDraftId && window.confirm(
        "该草稿已过期或被另一设备删除。是否先将当前内容另存为新草稿？"
      )) {
        state.currentDraftId = "";
        state.currentDraftVersion = null;
        state.currentDraftStatus = "";
        return persistCurrentDraft().then(function () {
          showPanelError("已另存为新草稿，请核对后再次保存日程。");
        }).catch(handleApiError);
      }
      if (error.code === "source_missing" && state.currentDraftId && window.confirm(
        "原日程已不存在。是否保留当前内容并转为新建日程草稿？"
      )) {
        return convertCurrentDraftToCreate().catch(handleApiError);
      }
      handleApiError(error);
    }).finally(function () {
      endFormOperation(operationToken);
    });
  }

  function isFormBusy() {
    return state.formOperation !== null;
  }

  function isFormOperationActive(token) {
    return !!(state.formOperation && state.formOperation.token === token);
  }

  function beginFormOperation(name) {
    if (isFormBusy()) {
      return null;
    }
    state.formOperationSeq += 1;
    state.formOperation = { name: name, token: state.formOperationSeq };
    syncFormOperationState();
    return state.formOperation.token;
  }

  function endFormOperation(token) {
    if (!isFormOperationActive(token)) {
      return;
    }
    state.formOperation = null;
    syncFormOperationState();
    syncFormState();
  }

  function syncDescriptionEditorOperationState() {
    var editor = getDescriptionEditor();
    if (editor && editor.mode && typeof editor.mode.set === "function") {
      editor.mode.set(isFormBusy() ? "readonly" : "design");
    }
  }

  function syncFormOperationState() {
    var busy = isFormBusy();
    refs.form.querySelectorAll("input, textarea, select, button").forEach(function (node) {
      if (busy) {
        if (!node.disabled) {
          node.dataset.operationDisabled = "1";
          node.disabled = true;
        }
        return;
      }
      if (node.dataset.operationDisabled === "1") {
        node.disabled = false;
        delete node.dataset.operationDisabled;
      }
    });
    syncDescriptionEditorOperationState();
    refs.saveButton.textContent = "保存";
    refs.saveDraftButton.textContent = "暂存";
    refs.deleteButton.textContent = "删除";
    refs.deleteDraftButton.textContent = "删除草稿";
    refs.unlockDraftButton.textContent = "已核对未提交";
    refs.applyRemoteDetailButton.textContent = "覆盖为远端";
    if (!state.formOperation) {
      return;
    }
    switch (state.formOperation.name) {
      case "draft_save":
        refs.saveDraftButton.textContent = "暂存中...";
        break;
      case "calendar_submit":
        refs.saveButton.textContent = "保存中...";
        break;
      case "event_delete":
        refs.deleteButton.textContent = "删除中...";
        break;
      case "draft_delete":
        refs.deleteDraftButton.textContent = "删除中...";
        break;
      case "outcome_resolve":
        refs.unlockDraftButton.textContent = "解锁中...";
        break;
      case "remote_replace":
        refs.applyRemoteDetailButton.textContent = "覆盖中...";
        break;
      default:
        throw new Error("unknown form operation: " + state.formOperation.name);
    }
  }

  function toEvent(item) {
    var props = {
      sourceStart: item.start,
      sourceEnd: item.end,
      body: item.body || "",
      bodyPlain: item.body_plain || item.body || "",
      bodyContentType: item.body_content_type || "text",
      location: item.location || "",
      organizer: item.organizer || "",
      requiredAttendees: item.required_attendees || [],
      optionalAttendees: item.optional_attendees || [],
      isMeeting: !!item.is_meeting,
      isCancelled: !!item.is_cancelled,
      entityType: "event",
      changeKey: item.changekey || "",
      sourceSnapshot: calendarItemSnapshot(item),
    };
    var locked = props.isCancelled || (props.isMeeting && !!item.is_all_day && BOOT.capabilities.disable_drag_resize_for_all_day_meetings);
    var allDay = !!item.is_all_day;
    var startValue = item.start_local || item.start;
    var endValue = item.end_local || item.end;
    return {
      id: item.id,
      title: item.subject || "(无标题)",
      start: allDay ? calendarInputValue(startValue, true) : startValue,
      end: allDay ? calendarInputValue(endValue, true) : endValue,
      allDay: allDay,
      editable: !props.isCancelled,
      startEditable: !locked,
      durationEditable: !locked,
      classNames: props.isCancelled ? ["cal-event-cancelled"] : [],
      extendedProps: props,
    };
  }

  function toDraftEvent(record) {
    var start = record.is_all_day
      ? record.start_date
      : (record.start_at == null ? "" : new Date(record.start_at * 1000).toISOString());
    var end = record.is_all_day
      ? record.end_date
      : (record.end_at == null ? "" : new Date(record.end_at * 1000).toISOString());
    if (!start || !end) {
      return null;
    }
    return {
      id: "draft:" + record.id,
      title: "[草稿] " + (record.subject || "(无标题)"),
      start: start,
      end: end,
      allDay: !!record.is_all_day,
      editable: false,
      startEditable: false,
      durationEditable: false,
      classNames: ["cal-event-draft"],
      extendedProps: {
        entityType: "draft",
        draftId: record.id,
        sourceEventId: record.source_event_id || "",
        sourceStart: start,
        sourceEnd: end,
        location: "",
        draftRecord: record,
      },
    };
  }

  function loadDraftEventInputs(info) {
    state.draftEventLoadSeq += 1;
    var loadSeq = state.draftEventLoadSeq;
    var collected = [];
    function nextPage(offset) {
      return apiCall("calendar/drafts/list", {
        start: info.startStr,
        end: info.endStr,
        limit: PAGE_LIMIT,
        offset: offset,
      }).then(function (data) {
        var items = data.items || [];
        collected = collected.concat(items);
        if (data.has_more) {
          return nextPage(offset + items.length);
        }
        return collected.map(toDraftEvent).filter(Boolean);
      });
    }
    return nextPage(0).catch(function (error) {
      if (loadSeq === state.draftEventLoadSeq) {
        showToast("加载草稿占位失败：" + error.message, "error");
      }
      return [];
    });
  }

  function formatDraftMeta(record) {
    var typeLabel = record.draft_type === "update" ? "编辑草稿" : "新建草稿";
    var timeLabel = "时间未设置";
    if (record.is_all_day && record.start_date && record.end_date) {
      timeLabel = record.start_date + " · 全天";
    } else if (record.start_at != null && record.end_at != null) {
      var timeFormatter = new Intl.DateTimeFormat("zh-CN", {
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        timeZone: state.calendarTimezone,
      });
      timeLabel = timeFormatter.format(new Date(record.start_at * 1000))
        + " - " + timeFormatter.format(new Date(record.end_at * 1000));
    }
    var updatedLabel = new Intl.DateTimeFormat("zh-CN", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    }).format(new Date(record.updated_at * 1000));
    return typeLabel + " · " + timeLabel + " · 最后暂存 " + updatedLabel;
  }

  function renderDraftList(items) {
    clearChildren(refs.draftsList);
    state.draftItems = items.slice();
    refs.draftsEmpty.hidden = items.length > 0;
    items.forEach(function (record) {
      var row = document.createElement("article");
      row.className = "cal-draft-row";
      row.setAttribute("role", "listitem");
      row.dataset.status = record.status;
      row.dataset.draftId = record.id;

      var open = document.createElement("button");
      open.type = "button";
      open.className = "cal-draft-row__open";
      open.dataset.draftId = record.id;
      var title = document.createElement("span");
      title.className = "cal-draft-row__title";
      title.textContent = "[草稿] " + (record.subject || "(无标题)");
      var meta = document.createElement("span");
      meta.className = "cal-draft-row__meta";
      meta.textContent = formatDraftMeta(record);
      open.appendChild(title);
      open.appendChild(meta);
      if (record.last_error_message) {
        var error = document.createElement("span");
        error.className = "cal-draft-row__error";
        error.textContent = record.last_error_message;
        open.appendChild(error);
      }
      open.addEventListener("click", function () {
        openPersistedDraft(record);
      });
      row.appendChild(open);

      var remove = document.createElement("button");
      remove.type = "button";
      remove.className = "cal-btn cal-btn-danger cal-btn-inline";
      remove.textContent = "删除";
      remove.addEventListener("click", function () {
        deleteDraftRecord(record);
      });
      row.appendChild(remove);
      refs.draftsList.appendChild(row);
    });
  }

  function loadDraftList() {
    state.draftListLoadSeq += 1;
    var loadSeq = state.draftListLoadSeq;
    refs.draftsStatus.hidden = false;
    refs.draftsStatus.dataset.tone = "";
    refs.draftsStatus.textContent = "正在加载草稿...";
    var collected = [];
    function nextPage(offset) {
      return apiCall("calendar/drafts/list", { limit: PAGE_LIMIT, offset: offset }).then(function (data) {
        var items = data.items || [];
        collected = collected.concat(items);
        return data.has_more ? nextPage(offset + items.length) : collected;
      });
    }
    return nextPage(0).then(function (items) {
      if (loadSeq !== state.draftListLoadSeq) {
        return items;
      }
      renderDraftList(items);
      refs.draftsStatus.hidden = true;
      return items;
    }).catch(function (error) {
      if (loadSeq !== state.draftListLoadSeq) {
        return [];
      }
      refs.draftsStatus.hidden = false;
      refs.draftsStatus.dataset.tone = "error";
      refs.draftsStatus.textContent = "加载草稿失败：" + error.message;
      return [];
    });
  }

  function deleteDraftRecord(record) {
    if (!window.confirm("确认删除这个草稿？")) {
      return;
    }
    apiCall("calendar/drafts/delete", {
      confirm: true,
      draft_id: record.id,
      version: record.version,
    }).then(function () {
      if (state.currentDraftId === record.id) {
        resetPanel();
      }
      showToast("已删除草稿。", "success");
      refreshDraftViews();
    }).catch(function (error) {
      showToast("删除草稿失败：" + error.message, "error");
    });
  }

  function loadDraftSettings() {
    state.draftSettingsLoadSeq += 1;
    var loadSeq = state.draftSettingsLoadSeq;
    return apiCall("calendar/drafts/settings/get", {}).then(function (data) {
      if (loadSeq !== state.draftSettingsLoadSeq || state.retentionSaving) {
        return;
      }
      var days = data.retention_days;
      refs.retentionMode.value = days == null ? "forever" : "days";
      refs.retentionDays.hidden = days == null;
      if (days != null) {
        refs.retentionDays.value = String(days);
      }
    }).catch(function (error) {
      if (loadSeq === state.draftSettingsLoadSeq && !state.retentionSaving) {
        showToast("加载草稿保留设置失败：" + error.message, "error");
      }
    });
  }

  function beginRetentionOperation() {
    if (state.retentionSaving) {
      return null;
    }
    state.retentionOperationSeq += 1;
    state.retentionSaving = true;
    state.draftSettingsLoadSeq += 1;
    refs.retentionMode.disabled = true;
    refs.retentionDays.disabled = true;
    refs.saveRetentionButton.disabled = true;
    refs.saveRetentionButton.textContent = "保存中...";
    return state.retentionOperationSeq;
  }

  function endRetentionOperation(token) {
    if (!state.retentionSaving || token !== state.retentionOperationSeq) {
      return;
    }
    state.retentionSaving = false;
    refs.retentionMode.disabled = false;
    refs.retentionDays.disabled = false;
    refs.saveRetentionButton.disabled = false;
    refs.saveRetentionButton.textContent = "保存设置";
  }

  function refreshDraftViews() {
    if (state.calendar) {
      var source = state.calendar.getEventSourceById("drafts");
      if (source) {
        source.refetch();
      }
    }
    if (state.activeTab === "drafts") {
      loadDraftList();
    }
  }

  function showStageTab(tabName) {
    if (tabName === state.activeTab) {
      return;
    }
    if (!canDiscard()) {
      return;
    }
    if (state.mode !== "empty") {
      state.selectedDate = "";
      syncSelectedDateClass();
      resetPanel();
    }
    state.activeTab = tabName;
    var showingDrafts = tabName === "drafts";
    refs.calendarTab.setAttribute("aria-selected", showingDrafts ? "false" : "true");
    refs.draftsTab.setAttribute("aria-selected", showingDrafts ? "true" : "false");
    refs.calendarView.hidden = showingDrafts;
    refs.draftsView.hidden = !showingDrafts;
    if (showingDrafts) {
      loadDraftList();
      loadDraftSettings();
    } else {
      window.requestAnimationFrame(function () {
        state.calendar.updateSize();
      });
    }
  }

  function refreshCalendar(options) {
    if (state.calendar) {
      state.forceNextNetwork = !!(options && options.forceNetwork);
      var source = state.calendar.getEventSourceById("remote");
      if (source) {
        source.refetch();
      } else {
        state.calendar.refetchEvents();
      }
    }
  }

  function cloneEventInput(eventInput) {
    return {
      id: eventInput.id,
      title: eventInput.title,
      start: eventInput.start,
      end: eventInput.end,
      allDay: eventInput.allDay,
      editable: eventInput.editable,
      startEditable: eventInput.startEditable,
      durationEditable: eventInput.durationEditable,
      classNames: (eventInput.classNames || []).slice(),
      extendedProps: Object.assign({}, eventInput.extendedProps || {}),
    };
  }

  function loadEventInputs(info) {
    state.calendarLoadSeq += 1;
    var loadSeq = state.calendarLoadSeq;
    var rangeStart = info.startStr;
    var rangeEnd = info.endStr;
    var forceNetwork = state.forceNextNetwork;
    state.forceNextNetwork = false;
    if (!forceNetwork) {
      var cached = loadCachedRange(rangeStart, rangeEnd);
      if (cached) {
        state.hasLoadedSuccessfully = true;
        state.lastLoadFailed = false;
        state.lastSuccessfulEventInputs = cached.items.map(cloneEventInput);
        setCacheStatus(
          cached.isExpired
            ? "显示缓存内容 · 最近更新 " + formatCacheTimestamp(cached.fetchedAt) + " · 缓存较旧，正在后台刷新最新日程"
            : "显示缓存内容 · 最近更新 " + formatCacheTimestamp(cached.fetchedAt) + " · 点击刷新获取最新日程",
          cached.isExpired ? "expired" : "stale"
        );
        if (cached.isExpired) {
          window.setTimeout(function () {
            refreshCalendar({ forceNetwork: true });
          }, 0);
        }
        return Promise.resolve(cached.items.map(cloneEventInput));
      }
      setCacheStatus("正在加载最新日程...", "loading");
    } else {
      setCacheStatus("正在刷新最新日程...", "loading");
    }

    var collectedItems = [];
    var offset = 0;

    function nextPage() {
      return apiCall("calendar/list", {
        start: rangeStart,
        end: rangeEnd,
        limit: PAGE_LIMIT,
        offset: offset,
        timezone: state.calendarTimezone,
      }).then(function (data) {
        var items = data.items || [];
        collectedItems = collectedItems.concat(items);
        offset += items.length;
        if (items.length === PAGE_LIMIT && collectedItems.length < MAX_VIEW_ITEMS) {
          return nextPage();
        }

        var truncated = collectedItems.length >= MAX_VIEW_ITEMS && items.length === PAGE_LIMIT;
        var eventInputs = collectedItems.slice(0, MAX_VIEW_ITEMS).map(toEvent);
        if (loadSeq === state.calendarLoadSeq) {
          clearDetailCache();
          state.hasLoadedSuccessfully = true;
          state.lastLoadFailed = false;
          state.lastSuccessfulEventInputs = eventInputs.map(cloneEventInput);
          saveCachedRange(rangeStart, rangeEnd, eventInputs, !truncated);
          setCacheStatus("已更新 · " + formatInstantTime(Date.now()), "");

          if (truncated) {
            setCalendarStatus("当前视图事件过多，暂时只显示前 1000 条。", "warning");
          } else {
            clearCalendarStatus();
          }
        }
        return eventInputs;
      });
    }

    return nextPage().catch(function (error) {
      if (loadSeq !== state.calendarLoadSeq) {
        return [];
      }
      state.lastLoadFailed = true;
      if (error.code === "auth_required") {
        setCalendarStatus(error.message, "error", BOOT.authPath, "重新授权");
      } else {
        setCalendarStatus("加载日程失败：" + error.message, "error");
      }
      if (state.hasLoadedSuccessfully) {
        setCacheStatus("显示缓存内容 · 刷新失败", "stale");
        showToast("刷新失败，已保留当前日历内容。", "error");
        return state.lastSuccessfulEventInputs.map(cloneEventInput);
      }
      setCacheStatus("加载失败，暂无缓存可用。", "stale");
      showToast("加载日程失败：" + error.message, "error");
      return [];
    });
  }

  function handleApiError(error) {
    if (error.code === "auth_required") {
      showPanelError(error.message, BOOT.authPath, "重新授权");
      return;
    }
    showPanelError(error.message);
  }

  function openDefaultCreatePanel() {
    var current = projectDateTimeValue(new Date());
    var startDate = state.selectedDate || current.slice(0, 10);
    var startValue = startDate + current.slice(10, 13) + ":00";
    openPanel("create", buildCreateDraftFromCalendarRange(
      startValue,
      suggestedEndValue(startValue),
      false
    ));
  }

  function openSelectedDateCreatePanel() {
    if (!state.selectedDate) {
      return;
    }
    openDefaultCreatePanel();
  }

  function stageCalendarMutation(info, message) {
    if (isFormBusy()) {
      info.revert();
      return;
    }
    var draft = buildEditDraftFromEvent(info.event);
    info.revert();
    if (!openPanel("edit", draft)) {
      return;
    }
    var operationToken = beginFormOperation("draft_save");
    if (operationToken === null) {
      return;
    }
    persistCurrentDraft().then(function () {
      showToast(message, "success");
    }).catch(handleApiError).finally(function () {
      endFormOperation(operationToken);
    });
  }

  refs.calendarTab.addEventListener("click", function () { showStageTab("calendar"); });
  refs.draftsTab.addEventListener("click", function () { showStageTab("drafts"); });
  refs.newButton.addEventListener("click", openDefaultCreatePanel);
  refs.refreshButton.addEventListener("click", function () {
    setCacheStatus("正在刷新最新日程...", "loading");
    refreshCalendar({ forceNetwork: true });
    refreshDraftViews();
  });
  refs.refreshDraftsButton.addEventListener("click", loadDraftList);
  refs.retentionMode.addEventListener("change", function () {
    state.draftSettingsLoadSeq += 1;
    refs.retentionDays.hidden = refs.retentionMode.value !== "days";
  });
  refs.retentionDays.addEventListener("input", function () {
    state.draftSettingsLoadSeq += 1;
  });
  refs.saveRetentionButton.addEventListener("click", function () {
    if (state.retentionSaving) {
      return;
    }
    var retentionDays = refs.retentionMode.value === "forever" ? null : Number(refs.retentionDays.value);
    if (retentionDays !== null && (!Number.isInteger(retentionDays) || retentionDays < 1 || retentionDays > 3650)) {
      showToast("保留天数必须是 1 到 3650 之间的整数。", "error");
      return;
    }
    var operationToken = beginRetentionOperation();
    if (operationToken === null) {
      return;
    }
    apiCall("calendar/drafts/settings/preview", {
      retention_days: retentionDays,
    }).then(function (preview) {
      if (preview.expired_count > 0 && !window.confirm(
        "设置后会立即删除 " + preview.expired_count + " 条超过期限的草稿。确认继续？"
      )) {
        return null;
      }
      return apiCall("calendar/drafts/settings/update", {
        confirm: true,
        retention_days: retentionDays,
      });
    }).then(function (data) {
      if (!data || operationToken !== state.retentionOperationSeq) {
        return;
      }
      var suffix = data.expired_count ? "，已清理 " + data.expired_count + " 条过期草稿" : "";
      showToast("草稿保留设置已保存" + suffix + "。", "success");
      refreshDraftViews();
    }).catch(function (error) {
      if (operationToken === state.retentionOperationSeq) {
        showToast("保存保留设置失败：" + error.message, "error");
      }
    }).finally(function () {
      endRetentionOperation(operationToken);
    });
  });
  refs.emptyNewButton.addEventListener("click", openDefaultCreatePanel);
  refs.emptyTodayButton.addEventListener("click", function () {
    clearSelectedDate();
    state.calendar.today();
  });
  refs.dayNewButton.addEventListener("click", openSelectedDateCreatePanel);
  refs.closeDayAgendaButton.addEventListener("click", clearSelectedDate);
  refs.closeButton.addEventListener("click", function () { if (canDiscard()) resetPanel(); });
  refs.closeLoadingButton.addEventListener("click", function () { resetPanel(); });
  refs.cancelButton.addEventListener("click", function () { if (canDiscard()) resetPanel(); });
  refs.saveDraftButton.addEventListener("click", saveDraftFromKeyboard);
  refs.backdrop.addEventListener("click", function () { if (canDiscard()) resetPanel(); });
  refs.applyRemoteDetailButton.addEventListener("click", function () {
    if (state.pendingRemoteDraft && state.pendingRemoteItemId === state.selectedEventId) {
      applyRemoteDetailDraft(state.pendingRemoteDraft);
    }
  });
  refs.addAttendeeButton.addEventListener("click", function () {
    refs.meetingSettings.open = true;
    renderAttendeeRow({ email: "", type: "required" });
    syncFormState();
    markDirty();
  });
  refs.allDay.addEventListener("change", function () { syncDateInputMode(); markDirty(); syncFormState(); });
  refs.form.start.addEventListener("click", showNativePicker);
  refs.form.end.addEventListener("click", showNativePicker);
  refs.form.start.addEventListener("change", syncEndAfterStartChange);
  refs.form.addEventListener("input", function () { markDirty(); syncFormState(); clearPanelError(); });
  refs.form.addEventListener("change", function () { markDirty(); syncFormState(); });
  refs.calendar.addEventListener("focusin", function (event) {
    if (!isMonthView() || !event.target || !event.target.closest) {
      return;
    }
    var cell = event.target.closest(".fc-daygrid-day[data-date]");
    if (!cell || (event.target.closest(".fc-event, .fc-daygrid-more-link"))) {
      return;
    }
    state.focusMode = "monthGrid";
    state.keyboardFocusDate = cell.dataset.date;
    syncMonthGridAccessibility();
  });
  refs.dayAgendaList.addEventListener("focusin", function (event) {
    var item = event.target && event.target.closest ? event.target.closest(".cal-day-agenda-item") : null;
    if (!item) {
      return;
    }
    state.focusMode = "dayAgenda";
    state.keyboardFocusEventId = item.dataset.eventId;
    syncDayAgendaAccessibility();
  });
  document.addEventListener("keydown", handleCalendarKeydown);

  refs.deleteButton.addEventListener("click", function () {
    if (!refs.itemId.value || !window.confirm("确认删除这个日程？")) {
      return;
    }
    var deletedItemId = refs.itemId.value;
    var operationToken = beginFormOperation("event_delete");
    if (operationToken === null) {
      return;
    }
    apiCall("calendar/delete", { confirm: true, item_id: refs.itemId.value, send_cancellations: true }).then(function (data) {
      var cacheUpdated = syncCalendarMutation("delete", data, deletedItemId);
      showToast("已删除日程。", "success");
      resetPanel();
      refreshCalendar({ forceNetwork: !cacheUpdated });
    }).catch(handleApiError).finally(function () { endFormOperation(operationToken); });
  });

  refs.deleteDraftButton.addEventListener("click", function () {
    if (!state.currentDraftId || !window.confirm("确认删除这个草稿？")) {
      return;
    }
    var operationToken = beginFormOperation("draft_delete");
    if (operationToken === null) {
      return;
    }
    apiCall("calendar/drafts/delete", {
      confirm: true,
      draft_id: state.currentDraftId,
      version: state.currentDraftVersion,
    }).then(function () {
      showToast("已删除草稿。", "success");
      resetPanel();
      refreshDraftViews();
    }).catch(handleApiError).finally(function () {
      endFormOperation(operationToken);
    });
  });

  refs.unlockDraftButton.addEventListener("click", function () {
    if (!state.currentDraftId || !window.confirm(
      "仅当你已在 Outlook 中确认上次操作没有生效时才能解锁。确认继续？"
    )) {
      return;
    }
    var operationToken = beginFormOperation("outcome_resolve");
    if (operationToken === null) {
      return;
    }
    apiCall("calendar/drafts/resolve-outcome", {
      confirm: true,
      draft_id: state.currentDraftId,
      version: state.currentDraftVersion,
    }).then(function (data) {
      applySavedDraftContext(data.draft);
      clearPanelError();
      showToast("草稿已解锁，可以再次提交。", "success");
      refreshDraftViews();
    }).catch(handleApiError).finally(function () {
      endFormOperation(operationToken);
    });
  });

  refs.form.addEventListener("submit", function (event) {
    event.preventDefault();
    submitCurrentForm();
  });

  state.calendar = new FullCalendar.Calendar(refs.calendar, {
    initialView: "dayGridMonth",
    locale: "zh-cn",
    timeZone: state.calendarTimezone,
    now: function () {
      return projectDateTimeValue(new Date());
    },
    firstDay: 1,
    selectable: true,
    editable: true,
    eventResizableFromStart: true,
    slotEventOverlap: false,
    dayMaxEvents: true,
    expandRows: true,
    height: "100%",
    nowIndicator: true,
    titleFormat: { year: "numeric", month: "long" },
    dayHeaderFormat: { weekday: "short" },
    eventTimeFormat: { hour: "2-digit", minute: "2-digit", hour12: false },
    headerToolbar: { left: "prev,next today", center: "title", right: "dayGridMonth,timeGridWeek,timeGridDay,listWeek" },
    loading: function (loading) {
      refs.loading.hidden = !loading;
      refs.refreshButton.disabled = loading;
    },
    eventSources: [
      { id: "remote", events: loadEventInputs },
      { id: "drafts", events: loadDraftEventInputs },
    ],
    eventsSet: function (events) {
      state.visibleEvents = events.slice();
      renderUpcomingEvents(events);
      renderDayAgenda();
      syncSelectedEventClass();
      syncSelectedDateClass();
      syncMonthGridAccessibility();
      syncDayAgendaAccessibility();
      if (state.restoreFocusAfterRender) {
        state.restoreFocusAfterRender = false;
        window.requestAnimationFrame(function () {
          if (state.focusMode === "dayAgenda") {
            focusDayAgenda(state.keyboardFocusEventId);
          }
        });
      }
      scheduleWeeklyNowIndicatorSync();
    },
    datesSet: function () {
      var focusDateAfterRender = state.focusDateAfterRender;
      state.focusDateAfterRender = "";
      if (state.selectedDate) {
        state.selectedDate = "";
        if (state.mode === "dayAgenda") {
          state.focusReturnTarget = {
            mode: "monthGrid",
            date: state.keyboardFocusDate,
            focus: false,
          };
          resetPanel();
        }
      }
      if (!state.keyboardFocusDate && state.calendar) {
        state.keyboardFocusDate = wallDateValue(state.calendar.getDate());
      }
      syncSelectedDateClass();
      syncMonthGridAccessibility();
      if (focusDateAfterRender) {
        window.requestAnimationFrame(function () {
          focusMonthDate(focusDateAfterRender, false);
        });
      }
      scheduleWeeklyNowIndicatorSync();
    },
    select: function (info) {
      state.calendar.unselect();
      window.setTimeout(function () {
        if (state.suppressNextSelect) {
          state.suppressNextSelect = false;
          return;
        }
        if (info.startStr && info.endStr) {
          openPanel("create", buildCreateDraftFromCalendarRange(info.startStr, info.endStr, info.allDay));
        } else {
          openPanel("create", buildCreateDraft(info.start, info.end, info.allDay));
        }
      }, 0);
    },
    dateClick: function (info) {
      if (info.view.type === "dayGridMonth") {
        state.suppressNextSelect = true;
        window.setTimeout(function () {
          state.suppressNextSelect = false;
        }, 50);
        focusMonthDate(info.dateStr, false);
        selectCalendarDate(info.dateStr, { focusAgenda: false });
      }
    },
    eventClick: function (info) {
      state.keyboardFocusDate = wallDateValue(info.event.start);
      openCalendarEntry(info.event);
      if (state.selectedEventId === info.event.id) {
        state.selectedDate = wallDateValue(info.event.start);
        state.focusReturnTarget = { mode: "dayAgenda", eventId: info.event.id };
        state.focusMode = "eventEditor";
        syncSelectedDateClass();
      }
    },
    eventDrop: function (info) {
      stageCalendarMutation(info, "移动结果已暂存为编辑草稿。");
    },
    eventResize: function (info) {
      stageCalendarMutation(info, "时间调整已暂存为编辑草稿。");
    },
    eventDidMount: function (info) {
      info.el.dataset.eventId = info.event.id;
      info.el.tabIndex = info.view.type === "dayGridMonth" ? -1 : 0;
      info.el.setAttribute("aria-label", info.event.title || "日程");
      info.el.classList.toggle("cal-event-selected", info.event.id === state.selectedEventId);
      var details = [];
      if (info.event.extendedProps.location) { details.push(info.event.extendedProps.location); }
      if (info.event.extendedProps.isCancelled) { details.push("已取消"); }
      if (details.length) { info.el.title = details.join(" · "); }
    },
  });

  window.addEventListener("resize", scheduleWeeklyNowIndicatorSync);
  state.calendar.render();
  ensureDescriptionEditor().catch(function (error) {
    console.error(error);
    showToast("说明编辑器加载失败，已回退到原始文本框。", "error");
  });
})();
