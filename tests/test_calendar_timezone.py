from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
FULLCALENDAR_BUNDLE = PROJECT_ROOT / "src/outlook_mcp/static/fullcalendar-index.global.min.js"
CALENDAR_SCRIPT = PROJECT_ROOT / "src/outlook_mcp/static/calendar.js"

pytestmark = pytest.mark.nodejs


def test_fullcalendar_preserves_named_timezone_as_wall_time() -> None:
    node = shutil.which("node")
    if node is None:
        pytest.skip("Node.js is required for the FullCalendar runtime contract test")

    node_source = r"""
const fs = require("fs");
const vm = require("vm");
vm.runInThisContext(fs.readFileSync(process.argv[1], "utf8"));

const dateEnv = new FullCalendar.Internal.DateEnv({
  timeZone: "Asia/Shanghai",
  namedTimeZoneImpl: null,
  calendarSystem: "gregory",
  locale: { week: { dow: 1, doy: 4 }, options: { weekText: "W" } },
  weekNumberCalculation: "ISO",
  defaultSeparator: " - ",
});
const start = dateEnv.createMarkerMeta("2026-01-01T17:00:00+08:00");
const end = dateEnv.createMarkerMeta("2026-01-01T18:00:00+08:00");
const event = new FullCalendar.Internal.EventImpl(
  { dateEnv },
  { allDay: false, hasEnd: true, extendedProps: {} },
  {
    range: { start: start.marker, end: end.marker },
    forcedStartTzo: start.forcedTzo,
    forcedEndTzo: end.forcedTzo,
  }
);
process.stdout.write(JSON.stringify({
  start: event.start.toISOString(),
  startStr: event.startStr,
  end: event.end.toISOString(),
  endStr: event.endStr,
}));
"""
    completed = subprocess.run(
        [node, "-e", node_source, str(FULLCALENDAR_BUNDLE)],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    result = json.loads(completed.stdout)
    assert result == {
        "start": "2026-01-01T17:00:00.000Z",
        "startStr": "2026-01-01T17:00:00+08:00",
        "end": "2026-01-01T18:00:00.000Z",
        "endStr": "2026-01-01T18:00:00+08:00",
    }


def test_calendar_script_uses_project_timezone_and_wall_time() -> None:
    script = CALENDAR_SCRIPT.read_text(encoding="utf-8")

    assert "calendarTimezone: BOOT.calendarTimezone" in script
    assert "resolvedOptions().timeZone" not in script
    assert "timeZone: state.calendarTimezone" in script
    assert 'timeZone: "UTC"' in script
    assert "sourceStart: item.start" in script
    assert "sourceEnd: item.end" in script
    assert 'var startDate = eventBoundaryDate(event, "start");' in script
    assert 'var startTime = eventBoundaryTime(event, "start");' in script
    assert "Date.parse(props.sourceEnd || props.sourceStart)" in script
    assert "startOfDay" not in script


@pytest.mark.parametrize("browser_timezone", ["UTC", "America/New_York"])
def test_upcoming_renders_project_wall_time(browser_timezone: str) -> None:
    node = shutil.which("node")
    if node is None:
        pytest.skip("Node.js is required for the calendar WebUI runtime test")

    node_source = r"""
const fs = require("fs");
const vm = require("vm");
const nodes = new Map();

function makeNode(id) {
  const target = {
    id: id || "",
    children: [],
    dataset: {},
    style: { removeProperty() {} },
    classList: { add() {}, remove() {}, toggle() {} },
    value: "",
    checked: false,
    hidden: false,
    disabled: false,
    type: "",
    textContent: "",
    appendChild(child) { this.children.push(child); return child; },
    removeChild(child) { this.children.splice(this.children.indexOf(child), 1); },
    addEventListener() {},
    querySelectorAll() { return []; },
    querySelector() { return null; },
    setAttribute() {},
    removeAttribute() {},
    focus() {},
    remove() {},
    closest() { return null; },
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
  });
}

const document = {
  getElementById(id) {
    if (!nodes.has(id)) nodes.set(id, makeNode(id));
    return nodes.get(id);
  },
  createElement(tag) { return makeNode(tag); },
  querySelector() { return null; },
  querySelectorAll() { return []; },
  addEventListener() {},
};
let calendarOptions;
class Calendar {
  constructor(element, options) {
    calendarOptions = options;
    this.view = { type: "dayGridMonth" };
  }
  render() {}
  getDate() { return new Date("2099-01-01T00:00:00Z"); }
  getEventById() { return null; }
}
const FullCalendar = { Calendar };
const editor = { on() {}, save() {}, show() {}, getContent() { return ""; } };
const window = {
  __OUTLOOK_CALENDAR_BOOTSTRAP__: {
    apiBase: "/api",
    calendarTimezone: "Asia/Shanghai",
    capabilities: {},
  },
  FullCalendar,
  tinymce: {
    init() { return Promise.resolve([editor]); },
    get() { return editor; },
  },
  addEventListener() {},
  requestAnimationFrame() { return 1; },
  cancelAnimationFrame() {},
  setTimeout() { return 1; },
  clearTimeout() {},
  localStorage: {
    getItem() { return null; },
    setItem() {},
    removeItem() {},
    key() { return null; },
    length: 0,
  },
};
global.window = window;
global.document = document;
global.FullCalendar = FullCalendar;
global.DOMParser = class {};
vm.runInThisContext(fs.readFileSync(process.argv[1], "utf8"));

function event(id, sourceStart, sourceEnd) {
  return {
    id,
    title: id,
    allDay: false,
    start: new Date("2099-01-01T17:00:00Z"),
    end: new Date("2099-01-01T18:00:00Z"),
    startStr: "2099-01-01T17:00:00+08:00",
    endStr: "2099-01-01T18:00:00+08:00",
    extendedProps: { sourceStart, sourceEnd },
  };
}
calendarOptions.eventsSet([
  event("expired", "2000-01-01T09:00:00+00:00", "2000-01-01T10:00:00+00:00"),
  event("future", "2099-01-01T09:00:00+00:00", "2099-01-01T10:00:00+00:00"),
]);
const upcoming = nodes.get("upcoming-list");
process.stdout.write(JSON.stringify({
  calendarTimezone: calendarOptions.timeZone,
  itemCount: upcoming.children.length,
  title: upcoming.children[0].children[0].textContent,
  meta: upcoming.children[0].children[1].textContent,
}));
"""
    completed = subprocess.run(
        [node, "-e", node_source, str(CALENDAR_SCRIPT)],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env={**os.environ, "TZ": browser_timezone},
    )

    result = json.loads(completed.stdout)
    assert result == {
        "calendarTimezone": "Asia/Shanghai",
        "itemCount": 1,
        "title": "future",
        "meta": "1/1周四 · 17:00 - 18:00",
    }
