from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest


CALENDAR_SCRIPT = Path(__file__).resolve().parents[1] / "src/outlook_mcp/static/calendar.js"

pytestmark = pytest.mark.nodejs


@pytest.mark.parametrize(
    ("scenario", "expected"),
    [
        (
            "draft_save",
            {"save": "保存", "draft": "暂存中...", "deleteEvent": "删除", "deleteDraft": "删除草稿"},
        ),
        (
            "calendar_submit",
            {"save": "保存中...", "draft": "暂存", "deleteEvent": "删除", "deleteDraft": "删除草稿"},
        ),
        (
            "event_delete",
            {"save": "保存", "draft": "暂存", "deleteEvent": "删除中...", "deleteDraft": "删除草稿"},
        ),
        (
            "draft_delete",
            {"save": "保存", "draft": "暂存", "deleteEvent": "删除", "deleteDraft": "删除中..."},
        ),
        ("draft_list_race", {"draftTitle": "[草稿] 新响应"}),
        ("draft_settings_race", {"retentionMode": "days", "retentionDays": "45"}),
        ("calendar_load_race", {"cachedTitle": "新响应", "cacheWrites": 1}),
        (
            "retention_save",
            {"text": "保存中...", "modeDisabled": True, "daysDisabled": True, "buttonDisabled": True},
        ),
    ],
)
def test_calendar_ui_runtime_state(
    scenario: str,
    expected: dict[str, str],
) -> None:
    node = shutil.which("node")
    if node is None:
        pytest.skip("Node.js is required for the calendar UI state test")

    node_source = r"""
const fs = require("fs");
const vm = require("vm");
const scenario = process.argv[2];
const nodes = new Map();
const documentListeners = {};
const revealTexts = [];
const pendingDraftLists = [];
const pendingDraftSettings = [];
const pendingCalendarLists = [];
const localStorageValues = new Map();
const localStorageWrites = [];

const initialText = {
  "btn-save": "保存",
  "btn-save-draft": "暂存",
  "btn-delete": "删除",
  "btn-delete-draft": "删除草稿",
  "btn-unlock-draft": "已核对未提交",
};
const initiallyHidden = new Set([
  "btn-delete", "btn-delete-draft", "btn-unlock-draft", "event-form",
  "panel-day-agenda", "panel-loading", "drafts-view", "drafts-status",
]);

function makeNode(id) {
  const listeners = {};
  const target = {
    id: id || "",
    children: [],
    dataset: {},
    style: { removeProperty() {} },
    classList: { add() {}, remove() {}, toggle() {} },
    value: "",
    checked: false,
    hidden: initiallyHidden.has(id),
    disabled: false,
    type: "",
    textContent: initialText[id] || "",
    appendChild(child) { this.children.push(child); return child; },
    removeChild(child) { this.children.splice(this.children.indexOf(child), 1); },
    addEventListener(type, listener) { (listeners[type] ||= []).push(listener); },
    dispatch(type, event = {}) { (listeners[type] || []).forEach((listener) => listener(event)); },
    click() { this.dispatch("click", { preventDefault() {} }); },
    querySelectorAll() { return []; },
    querySelector() { return null; },
    setAttribute() {},
    removeAttribute() {},
    focus() {},
    remove() {},
    closest(selector) {
      if (selector.includes("input") && this.id.includes("subject")) return this;
      return null;
    },
    reset() {},
  };
  Object.defineProperty(target, "firstChild", {
    get() { return this.children[0] || null; },
  });
  return new Proxy(target, {
    get(object, property) {
      if (property in object) return object[property];
      if (typeof property === "string") {
        object[property] = makeNode(id + "." + property);
        return object[property];
      }
    },
    set(object, property, value) {
      object[property] = value;
      if (id === "btn-delete-draft" && property === "hidden" && value === false) {
        revealTexts.push(object.textContent);
      }
      return true;
    },
  });
}

const document = {
  activeElement: null,
  getElementById(id) {
    if (!nodes.has(id)) nodes.set(id, makeNode(id));
    return nodes.get(id);
  },
  createElement(tag) { return makeNode(tag); },
  querySelector() { return null; },
  querySelectorAll() { return []; },
  addEventListener(type, listener) { (documentListeners[type] ||= []).push(listener); },
};

const form = document.getElementById("event-form");
const formControls = [
  "btn-save", "btn-save-draft", "btn-delete", "btn-delete-draft",
  "btn-unlock-draft", "btn-cancel", "btn-close-panel", "event-form.subject",
].map((id) => document.getElementById(id));
form.querySelectorAll = function (selector) {
  if (selector === "input, textarea, select, button") return formControls;
  if (selector === "[data-field-error]") return [];
  return [];
};
form.subject = document.getElementById("event-form.subject");
form.location = document.getElementById("event-form.location");
form.start = document.getElementById("event-form.start");
form.end = document.getElementById("event-form.end");
form.reminder_minutes_before_start = document.getElementById("event-form.reminder");
form.requestSubmit = function () {
  form.dispatch("submit", { preventDefault() {} });
};

let calendarOptions;
let calendarInstance;
class Calendar {
  constructor(element, options) {
    calendarOptions = options;
    calendarInstance = this;
    this.view = { type: "dayGridMonth" };
    this.events = new Map();
  }
  render() {}
  getDate() { return new Date("2026-07-16T01:00:00Z"); }
  getEventById(id) { return this.events.get(id) || null; }
  getEventSourceById() { return { refetch() {} }; }
  updateSize() {}
}
const FullCalendar = { Calendar };
const editor = {
  mode: { current: "design", set(value) { this.current = value; } },
  on() {}, save() {}, show() {}, setContent() {},
  getContent(options) { return options && options.format === "text" ? "" : ""; },
};
const window = {
  __OUTLOOK_CALENDAR_BOOTSTRAP__: {
    apiBase: "/api",
    authPath: "/graph/person/secret",
    calendarTimezone: "Asia/Shanghai",
    capabilities: {},
  },
  FullCalendar,
  tinymce: {
    init() { return Promise.resolve([editor]); },
    get() { return editor; },
  },
  addEventListener() {},
  requestAnimationFrame(callback) { callback(); return 1; },
  cancelAnimationFrame() {},
  setTimeout() { return 1; },
  clearTimeout() {},
  confirm() { return true; },
  localStorage: {
    getItem(key) { return localStorageValues.get(key) || null; },
    setItem(key, value) { localStorageValues.set(key, value); localStorageWrites.push({ key, value }); },
    removeItem(key) { localStorageValues.delete(key); },
    key(index) { return Array.from(localStorageValues.keys())[index] || null; },
    get length() { return localStorageValues.size; },
  },
};
global.window = window;
global.document = document;
global.FullCalendar = FullCalendar;
global.DOMParser = class {};
function responseFor(body) {
  return { text() { return Promise.resolve(JSON.stringify(body)); } };
}
global.fetch = function (url) {
  if (scenario === "draft_list_race" && url.endsWith("calendar/drafts/list")) {
    return new Promise((resolve) => pendingDraftLists.push((body) => resolve(responseFor(body))));
  }
  if (scenario === "draft_settings_race" && url.endsWith("calendar/drafts/settings/get")) {
    return new Promise((resolve) => pendingDraftSettings.push((body) => resolve(responseFor(body))));
  }
  if (scenario === "calendar_load_race" && url.endsWith("calendar/list")) {
    return new Promise((resolve) => pendingCalendarLists.push((body) => resolve(responseFor(body))));
  }
  const body = url.endsWith("calendar/drafts/save")
    ? { ok: true, data: { draft: {
        id: "draft-1", version: 1, status: "active", draft_type: "create",
      } } }
    : { ok: true, data: { items: [], has_more: false } };
  return Promise.resolve(responseFor(body));
};

vm.runInThisContext(fs.readFileSync(process.argv[1], "utf8"));
let shortcutEvent = null;

function captureDuring() {
  return {
    save: nodes.get("btn-save").textContent,
    draft: nodes.get("btn-save-draft").textContent,
    deleteEvent: nodes.get("btn-delete").textContent,
    deleteDraft: nodes.get("btn-delete-draft").textContent,
    editorMode: editor.mode.current,
  };
}

function writeResult(during) {
  process.stdout.write(JSON.stringify({
    during,
    revealTexts,
    prevented: shortcutEvent ? shortcutEvent.prevented : null,
  }));
}

function pressShortcut(key) {
  shortcutEvent = {
    key, ctrlKey: true, metaKey: false, altKey: false, shiftKey: false,
    target: form.subject, preventDefault() { this.prevented = true; },
  };
  documentListeners.keydown[0](shortcutEvent);
}

if (scenario === "draft_save") {
  nodes.get("btn-new-event").click();
  pressShortcut("s");
  const during = captureDuring();
  setImmediate(() => setImmediate(() => writeResult(during)));
} else if (scenario === "calendar_submit") {
  nodes.get("btn-new-event").click();
  form.subject.value = "测试日程";
  pressShortcut("Enter");
  const during = captureDuring();
  setImmediate(() => setImmediate(() => writeResult(during)));
} else if (scenario === "event_delete") {
  const event = {
    id: "event-1", title: "已有日程", start: new Date("2026-07-16T01:00:00Z"),
    end: new Date("2026-07-16T02:00:00Z"), allDay: false, extendedProps: {},
  };
  calendarInstance.events.set(event.id, event);
  calendarOptions.eventClick({ event });
  nodes.get("btn-delete").click();
  const during = captureDuring();
  setImmediate(() => setImmediate(() => writeResult(during)));
} else if (scenario === "draft_delete") {
  nodes.get("btn-new-event").click();
  pressShortcut("s");
  setImmediate(() => setImmediate(() => {
    shortcutEvent = null;
    nodes.get("btn-delete-draft").click();
    const during = captureDuring();
    setImmediate(() => setImmediate(() => writeResult(during)));
  }));
} else if (scenario === "draft_list_race") {
  nodes.get("tab-drafts").click();
  nodes.get("btn-refresh-drafts").click();
  const record = (id, subject) => ({
    id, subject, draft_type: "create", status: "active", version: 1,
    updated_at: 1784160000, is_all_day: false, start_at: null, end_at: null,
  });
  pendingDraftLists[1]({ ok: true, data: { items: [record("new", "新响应")], has_more: false } });
  setImmediate(() => setImmediate(() => {
    pendingDraftLists[0]({ ok: true, data: { items: [record("old", "旧响应")], has_more: false } });
    setImmediate(() => setImmediate(() => {
      const list = nodes.get("drafts-list");
      const draftTitle = list.children[0].children[0].children[0].textContent;
      process.stdout.write(JSON.stringify({ draftTitle }));
    }));
  }));
} else if (scenario === "draft_settings_race") {
  nodes.get("tab-drafts").click();
  const mode = nodes.get("draft-retention-mode");
  const days = nodes.get("draft-retention-days");
  mode.value = "days";
  mode.dispatch("change");
  days.value = "45";
  days.dispatch("input");
  pendingDraftSettings[0]({ ok: true, data: { retention_days: null } });
  setImmediate(() => setImmediate(() => {
    process.stdout.write(JSON.stringify({ retentionMode: mode.value, retentionDays: days.value }));
  }));
} else if (scenario === "calendar_load_race") {
  const loadEvents = calendarOptions.eventSources[0].events;
  const info = { startStr: "2026-07-01T00:00:00Z", endStr: "2026-08-01T00:00:00Z" };
  const item = (id, subject) => ({
    id, subject, start: "2026-07-16T09:00:00+08:00", end: "2026-07-16T10:00:00+08:00",
    is_all_day: false,
  });
  const first = loadEvents(info);
  const second = loadEvents(info);
  pendingCalendarLists[1]({ ok: true, data: { items: [item("new", "新响应")] } });
  setImmediate(() => setImmediate(() => {
    pendingCalendarLists[0]({ ok: true, data: { items: [item("old", "旧响应")] } });
    Promise.all([first, second]).then(() => {
      const payload = JSON.parse(localStorageWrites[localStorageWrites.length - 1].value);
      process.stdout.write(JSON.stringify({
        cachedTitle: payload.items[0].title,
        cacheWrites: localStorageWrites.length,
      }));
    });
  }));
} else if (scenario === "retention_save") {
  nodes.get("tab-drafts").click();
  const mode = nodes.get("draft-retention-mode");
  const days = nodes.get("draft-retention-days");
  const button = nodes.get("btn-save-retention");
  mode.value = "forever";
  button.click();
  const result = {
    text: button.textContent,
    modeDisabled: mode.disabled,
    daysDisabled: days.disabled,
    buttonDisabled: button.disabled,
  };
  setImmediate(() => setImmediate(() => process.stdout.write(JSON.stringify(result))));
} else {
  throw new Error("unknown scenario: " + scenario);
}
"""
    completed = subprocess.run(
        [node, "-e", node_source, str(CALENDAR_SCRIPT), scenario],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    result = json.loads(completed.stdout)
    if scenario in {"draft_save", "calendar_submit"}:
        assert result["prevented"] is True
    if scenario in {"draft_list_race", "draft_settings_race", "calendar_load_race", "retention_save"}:
        assert result == expected
    else:
        assert result["during"].pop("editorMode") == "readonly"
        assert result["during"] == expected
        if scenario == "draft_save":
            assert result["revealTexts"] == ["删除草稿"]
