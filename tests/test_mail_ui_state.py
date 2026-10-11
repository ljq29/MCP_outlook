from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest


MAIL_SCRIPT = Path(__file__).resolve().parents[1] / "src/outlook_mcp/static/mail.js"

pytestmark = pytest.mark.nodejs


def test_mail_ui_runtime_helpers_and_send_payload() -> None:
    node = shutil.which("node")
    if node is None:
        pytest.skip("Node.js is required for the mail UI state test")

    node_source = r"""
const fs = require("fs");
const vm = require("vm");
const nodes = new Map();
const calls = [];
const historyCalls = [];
const windowListeners = {};
const documentListeners = {};
const initiallyHidden = new Set([
  "mail-status", "folders-loading", "list-loading", "list-empty", "detail-content",
  "body-truncated", "attachments-section", "compose-backdrop", "compose-panel",
  "btn-clear-search", "btn-load-images", "folder-backdrop",
]);

function makeNode(id) {
  const listeners = {};
  const classNames = new Set();
  const node = {
    id: id || "",
    children: [],
    dataset: {},
    style: { removeProperty() {} },
    className: "",
    value: "",
    hidden: initiallyHidden.has(id),
    disabled: false,
    inert: false,
    textContent: "",
    selectedIndex: 0,
    classList: {
      add(...names) { names.forEach((name) => classNames.add(name)); node.className = Array.from(classNames).join(" "); },
      remove(...names) { names.forEach((name) => classNames.delete(name)); node.className = Array.from(classNames).join(" "); },
      toggle(name, force) {
        const enabled = force === undefined ? !classNames.has(name) : Boolean(force);
        if (enabled) { classNames.add(name); } else { classNames.delete(name); }
        node.className = Array.from(classNames).join(" ");
        return enabled;
      },
      contains(name) { return classNames.has(name); },
    },
    appendChild(child) { this.children.push(child); return child; },
    replaceChildren(...children) { this.children = children; },
    cloneNode() {
      const clone = makeNode(this.id);
      clone.className = this.className;
      clone.title = this.title;
      clone.sandbox = this.sandbox;
      clone.referrerpolicy = this.referrerpolicy;
      return clone;
    },
    replaceWith(replacement) { nodes.set(this.id, replacement); },
    setAttribute(name, value) {
      if (name === "inert") { this.inert = true; return; }
      this[name] = String(value);
      if (name === "class") { this.className = String(value); }
    },
    removeAttribute(name) {
      if (name === "inert") { this.inert = false; return; }
      delete this[name];
    },
    addEventListener(name, handler) {
      if (!listeners[name]) { listeners[name] = []; }
      listeners[name].push(handler);
    },
    dispatch(name, event) {
      let result;
      (listeners[name] || []).forEach((handler) => { result = handler(event || {}); });
      return result;
    },
    click() { return this.dispatch("click", { target: this, preventDefault() {} }); },
    focus() { document.activeElement = this; },
    querySelector() { return makeNode("nested-heading"); },
    querySelectorAll() {
      if (this.id !== "compose-panel") { return []; }
      return ["btn-compose-close", "compose-to", "compose-cc", "compose-bcc", "compose-subject", "compose-body", "btn-send"]
        .map((nodeId) => document.getElementById(nodeId));
    },
  };
  Object.defineProperty(node, "options", { get() { return this.children; } });
  return node;
}

global.document = {
  activeElement: null,
  body: null,
  getElementById(id) {
    if (!nodes.has(id)) { nodes.set(id, makeNode(id)); }
    return nodes.get(id);
  },
  createElement(tag) { return makeNode(tag); },
  querySelectorAll(selector) {
    if (!selector.includes(".mail-detail__actions")) { return []; }
    return ["btn-reply", "btn-forward", "btn-toggle-read", "move-folder", "btn-move", "btn-delete"]
      .map((nodeId) => document.getElementById(nodeId));
  },
  addEventListener(name, handler) {
    if (!documentListeners[name]) { documentListeners[name] = []; }
    documentListeners[name].push(handler);
  },
  dispatch(name, event) {
    (documentListeners[name] || []).forEach((handler) => handler(event || {}));
  },
};
document.body = makeNode("body");
document.activeElement = document.body;
global.window = global;
window.location = {
  origin: "https://mail.example.test",
  href: "https://mail.example.test/graph/person@outlook.com/test-secret-at-least-24-chars/mail",
};
window.__OUTLOOK_MAIL_TESTING__ = true;
window.__OUTLOOK_MAIL_BOOTSTRAP__ = {
  apiBase: "/graph/person@outlook.com/api",
  authPath: "/graph/person@outlook.com/test-secret-at-least-24-chars",
  pageSize: 20,
  bodyMaxChars: 200000,
};
window.addEventListener = function (name, handler) {
  if (!windowListeners[name]) { windowListeners[name] = []; }
  windowListeners[name].push(handler);
};
window.matchMedia = function (query) {
  return {
    media: query,
    matches: false,
    listeners: [],
    addEventListener(name, handler) { if (name === "change") { this.listeners.push(handler); } },
    addListener(handler) { this.listeners.push(handler); },
  };
};
window.history = {
  state: null,
  pushState(state, title, url) { this.state = state; historyCalls.push({ state, url }); },
  back() {
    this.state = null;
    (windowListeners.popstate || []).forEach((handler) => handler({}));
  },
};
window.confirm = function () { return true; };

function response(data) {
  return Promise.resolve({
    text() { return Promise.resolve(JSON.stringify({ ok: true, data })); },
  });
}

global.fetch = function (url, options) {
  const body = options && options.body ? JSON.parse(options.body) : {};
  calls.push({ url, body, headers: options.headers });
  if (url.endsWith("/mail/folders")) {
    return response({
      items: [
        { id: "folder-inbox", name: "Inbox", unread_count: 2 },
        { id: "folder-sent", name: "Sent Items", unread_count: 0 },
      ],
    });
  }
  if (url.endsWith("/mail/list")) {
    if (body.folder === "folder-sent") {
      return response({ items: [{ id: "sent-1", subject: "Sent", author: "person@outlook.com", is_read: true }], has_more: false });
    }
    if (body.offset === 20) { return response({ items: [], has_more: false }); }
    return response({
      items: [
        {
          id: "mail-1", subject: "Quarterly report", author: "sender@example.com",
          datetime_received: "2026-07-15T10:00:00Z", is_read: false, has_attachments: true,
        },
        {
          id: "mail-2", subject: "Second message", author: "second@example.com",
          datetime_received: "2026-07-15T09:00:00Z", is_read: true, has_attachments: false,
        },
        {
          id: "mail-3", subject: "Empty message", author: "empty@example.com",
          datetime_received: "2026-07-15T08:00:00Z", is_read: true, has_attachments: false,
        },
      ],
      has_more: true,
    });
  }
  if (url.endsWith("/mail/search")) {
    return response({ items: [{ id: "search-1", subject: "Quarterly result", author: "sender@example.com" }], has_more: false });
  }
  if (url.endsWith("/mail/get")) {
    if (body.item_id === "mail-2") {
      return response({
        id: "mail-2", subject: "Second message", author: "second@example.com",
        display_to: "person@outlook.com", display_cc: "", datetime_received: "2026-07-15T09:00:00Z",
        is_read: true, has_attachments: false, body: "Second detailed body", body_content_type: "text", body_truncated: false,
      });
    }
    if (body.item_id === "mail-3") {
      return response({
        id: "mail-3", subject: "Empty message", author: "empty@example.com",
        display_to: "person@outlook.com", display_cc: "", datetime_received: "2026-07-15T08:00:00Z",
        is_read: true, has_attachments: false, body: "", body_content_type: "text", body_truncated: false,
      });
    }
    return response({
      id: "mail-1", subject: "Quarterly report", author: "sender@example.com",
      display_to: "person@outlook.com", display_cc: "", datetime_received: "2026-07-15T10:00:00Z",
      is_read: false, has_attachments: true, body: "Detailed body", body_content_type: "text", body_truncated: false,
    });
  }
  if (url.endsWith("/mail/attachments")) {
    return response({ items: [{ name: "report.pdf", content_type: "application/pdf", size: 2048 }], total: 1 });
  }
  if (url.endsWith("/mail/send")) { return response({ sent: true }); }
  if (url.endsWith("/mail/reply")) { return response({ replied: true, id: "mail-1" }); }
  if (url.endsWith("/mail/forward")) { return response({ forwarded: true }); }
  if (url.endsWith("/mail/mark_read")) { return response({ marked: "read", id: "mail-1" }); }
  if (url.endsWith("/mail/mark_unread")) { return response({ marked: "unread", id: "mail-1" }); }
  if (url.endsWith("/mail/move")) { return response({ moved: true, id: "mail-1", target_folder: "folder-sent" }); }
  if (url.endsWith("/mail/delete")) { return response({ deleted: true, mode: "move_to_trash" }); }
  throw new Error("unexpected fetch: " + url);
};

vm.runInThisContext(fs.readFileSync(process.argv[1], "utf8"), { filename: process.argv[1] });

async function settle() {
  await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setImmediate(resolve));
  await new Promise((resolve) => setImmediate(resolve));
}

(async function () {
  await settle();
  const hooks = window.__OUTLOOK_MAIL_TEST_HOOKS__;
  const initialFolder = hooks.stateSnapshot().folder;
  const initialListCount = nodes.get("message-list").children.length;
  nodes.get("message-list").children[0].dispatch("click");
  await settle();
  const firstBodyFrame = nodes.get("mail-body-frame");
  const detailBodyLoaded = firstBodyFrame.srcdoc.includes("Detailed body");
  const attachmentCount = nodes.get("attachments-list").children.length;
  nodes.get("message-list").children[1].dispatch("click");
  await settle();
  const secondBodyFrame = nodes.get("mail-body-frame");
  const secondBodyLoaded = secondBodyFrame.srcdoc.includes("Second detailed body") &&
    !secondBodyFrame.srcdoc.includes("Detailed body</pre>");
  nodes.get("message-list").children[2].dispatch("click");
  await settle();
  const emptyBodyFrame = nodes.get("mail-body-frame");
  const emptyBodyExplained = emptyBodyFrame.srcdoc.includes("此邮件没有可显示的正文");
  nodes.get("btn-next").dispatch("click");
  await settle();
  nodes.get("search-query").value = "Quarterly";
  nodes.get("search-form").dispatch("submit", { preventDefault() {} });
  await settle();
  nodes.get("folder-list").children[1].dispatch("click");
  await settle();

  hooks.openCompose("new");
  nodes.get("compose-to").value = "alice@example.com; bob@example.com";
  nodes.get("compose-cc").value = "copy@example.com";
  nodes.get("compose-subject").value = "HTML message";
  nodes.get("compose-body").value = "<p>Hello</p>";
  await hooks.submitCompose({ preventDefault() {} });
  await settle();

  const detail = {
    id: "mail-1", subject: "Status", author: "sender@example.com", display_to: "person@outlook.com",
    display_cc: "", datetime_received: "2026-07-15T10:00:00Z", is_read: false,
    has_attachments: false, body: "Plain body", body_content_type: "text", body_truncated: false,
  };
  hooks.setDetail(detail);
  hooks.openCompose("reply");
  nodes.get("compose-body").value = "<p>Reply</p>";
  await hooks.submitCompose({ preventDefault() {} });
  hooks.setDetail(detail);
  hooks.openCompose("forward");
  nodes.get("compose-to").value = "forward@example.com";
  nodes.get("compose-body").value = "<p>Forward note</p>";
  await hooks.submitCompose({ preventDefault() {} });
  hooks.setDetail(detail);
  await hooks.toggleRead();
  await hooks.toggleRead();
  hooks.setDetail(detail);
  nodes.get("move-folder").value = "folder-inbox";
  await hooks.moveSelected();
  hooks.setDetail(detail);
  await hooks.deleteSelected();
  await settle();

  hooks.setViewport(true, true);
  hooks.setDetail(detail);
  hooks.returnToList(false);
  nodes.get("message-list").children[0].dispatch("click");
  await settle();
  const mobileDetailState = hooks.stateSnapshot();
  const mobileDetailVisibility = {
    listHidden: nodes.get("list-pane").hidden,
    detailHidden: nodes.get("detail-pane").hidden,
  };
  hooks.returnToList(true);
  await settle();
  const mobileListState = hooks.stateSnapshot();
  const mobileListVisibility = {
    listHidden: nodes.get("list-pane").hidden,
    detailHidden: nodes.get("detail-pane").hidden,
  };

  hooks.openFolders();
  const drawerOpen = {
    state: hooks.stateSnapshot().folderDrawerOpen,
    dataOpen: nodes.get("folder-pane").dataset.open,
    backdropHidden: nodes.get("folder-backdrop").hidden,
    listInert: nodes.get("list-pane").inert,
  };
  hooks.closeFolders(false);
  const drawerClosedListInert = nodes.get("list-pane").inert;

  nodes.get("btn-compose").focus();
  hooks.openCompose("new");
  await settle();
  const composeFocused = document.activeElement.id;
  const composeWorkspaceInert = nodes.get("mail-workspace").inert;
  nodes.get("compose-body").dispatch("input");
  let escapePrevented = false;
  document.dispatch("keydown", { key: "Escape", preventDefault() { escapePrevented = true; } });
  await settle();
  const restoredFocus = document.activeElement.id;
  const restoredWorkspaceInert = nodes.get("mail-workspace").inert;

  const sendCall = calls.find((call) => call.url.endsWith("/mail/send"));
  const getCall = calls.find((call) => call.url.endsWith("/mail/get"));
  const attachmentsCall = calls.find((call) => call.url.endsWith("/mail/attachments"));
  const nextListCall = calls.find((call) => call.url.endsWith("/mail/list") && call.body.offset === 20);
  const searchCall = calls.find((call) => call.url.endsWith("/mail/search"));
  const sentListCall = calls.find((call) => call.url.endsWith("/mail/list") && call.body.folder === "folder-sent");
  const writeCalls = Object.fromEntries(calls.filter((call) => /\/mail\/(reply|forward|mark_read|mark_unread|move|delete)$/.test(call.url))
    .map((call) => [call.url.split("/").pop(), call]));
  process.stdout.write(JSON.stringify({
    addresses: hooks.parseAddresses("a@example.com; b@example.com, c@example.com"),
    unsafeUrl: hooks.safeUrl("javascript:alert(1)", true),
    safeUrl: hooks.safeUrl("https://example.com/path", false),
    unsafeStyle: hooks.safeStyle("background:url(https://tracker.example/pixel)"),
    remoteImage: hooks.hasRemoteImages('<img src="https://tracker.example/pixel">'),
    replySubject: hooks.replySubject("Status"),
    initialFolder,
    finalFolder: hooks.stateSnapshot().folder,
    initialCalls: calls.slice(0, 2).map((call) => call.url),
    initialListCount,
    detailBodyLoaded,
    bodyFrameReplaced: firstBodyFrame !== secondBodyFrame && secondBodyFrame !== emptyBodyFrame,
    secondBodyLoaded,
    emptyBodyExplained,
    attachmentCount,
    getCall,
    attachmentsCall,
    nextListCall,
    searchCall,
    sentListCall,
    send: sendCall,
    writes: writeCalls,
    composeHidden: nodes.get("compose-panel").hidden,
    mobileDetailState,
    mobileDetailVisibility,
    mobileListState,
    mobileListVisibility,
    drawerOpen,
    drawerClosedListInert,
    historyCalls,
    composeFocused,
    composeWorkspaceInert,
    escapePrevented,
    restoredFocus,
    restoredWorkspaceInert,
  }));
})().catch((error) => {
  process.stderr.write(error.stack || String(error));
  process.exitCode = 1;
});
"""
    completed = subprocess.run(
        [node, "-e", node_source, str(MAIL_SCRIPT)],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    result = json.loads(completed.stdout)
    assert result["addresses"] == ["a@example.com", "b@example.com", "c@example.com"]
    assert result["unsafeUrl"] == ""
    assert result["safeUrl"] == "https://example.com/path"
    assert result["unsafeStyle"] == ""
    assert result["remoteImage"] is True
    assert result["replySubject"] == "Re: Status"
    assert result["initialFolder"] == {"id": "folder-inbox", "name": "Inbox"}
    assert result["finalFolder"] == {"id": "folder-sent", "name": "Sent Items"}
    assert result["initialCalls"] == [
        "/graph/person@outlook.com/api/mail/folders",
        "/graph/person@outlook.com/api/mail/list",
    ]
    assert result["initialListCount"] == 3
    assert result["detailBodyLoaded"] is True
    assert result["bodyFrameReplaced"] is True
    assert result["secondBodyLoaded"] is True
    assert result["emptyBodyExplained"] is True
    assert result["attachmentCount"] == 1
    assert result["getCall"]["body"] == {
        "item_id": "mail-1",
        "folder": "folder-inbox",
        "include_body": True,
        "body_max_chars": 200000,
    }
    assert result["attachmentsCall"]["body"] == {"item_id": "mail-1", "folder": "folder-inbox"}
    assert result["nextListCall"]["body"]["offset"] == 20
    assert result["nextListCall"]["body"]["limit"] == 20
    assert result["searchCall"]["body"] == {
        "folder": "folder-inbox",
        "limit": 20,
        "offset": 0,
        "query": "Quarterly",
    }
    assert result["sentListCall"]["body"]["folder"] == "folder-sent"
    assert result["sentListCall"]["body"]["offset"] == 0
    assert result["send"]["headers"]["Authorization"] == "Bearer test-secret-at-least-24-chars"
    assert result["send"]["body"] == {
        "to": ["alice@example.com", "bob@example.com"],
        "cc": ["copy@example.com"],
        "bcc": None,
        "subject": "HTML message",
        "body": "<p>Hello</p>",
        "html": True,
        "confirm": True,
    }
    assert result["writes"]["reply"]["body"]["html"] is True
    assert result["writes"]["reply"]["body"]["confirm"] is True
    assert result["writes"]["reply"]["body"]["item_id"] == "mail-1"
    assert result["writes"]["forward"]["body"]["to"] == ["forward@example.com"]
    assert result["writes"]["forward"]["body"]["html"] is True
    assert result["writes"]["mark_read"]["body"]["confirm"] is True
    assert result["writes"]["mark_unread"]["body"]["confirm"] is True
    assert result["writes"]["move"]["body"]["target_folder"] == "folder-inbox"
    assert result["writes"]["move"]["body"]["confirm"] is True
    assert result["writes"]["delete"]["body"]["mode"] == "move_to_trash"
    assert result["writes"]["delete"]["body"]["confirm"] is True
    assert result["composeHidden"] is True
    assert result["mobileDetailState"]["mobileView"] == "detail"
    assert result["mobileDetailVisibility"] == {"listHidden": True, "detailHidden": False}
    assert result["mobileListState"]["mobileView"] == "list"
    assert result["mobileListVisibility"] == {"listHidden": False, "detailHidden": True}
    assert result["drawerOpen"] == {
        "state": True,
        "dataOpen": "true",
        "backdropHidden": False,
        "listInert": True,
    }
    assert result["drawerClosedListInert"] is False
    assert len(result["historyCalls"]) == 1
    assert result["historyCalls"][0]["state"] == {"outlookMailView": "detail"}
    assert result["composeFocused"] == "compose-to"
    assert result["composeWorkspaceInert"] is True
    assert result["escapePrevented"] is True
    assert result["restoredFocus"] == "btn-compose"
    assert result["restoredWorkspaceInert"] is False
