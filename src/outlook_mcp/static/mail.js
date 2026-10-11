(function () {
  "use strict";

  var BOOT = window.__OUTLOOK_MAIL_BOOTSTRAP__ || {};
  var PAGE_SIZE = Number(BOOT.pageSize || 20);
  var BODY_MAX_CHARS = Number(BOOT.bodyMaxChars || 200000);
  var API_TIMEOUT_MS = 60000;
  var BLOCKED_ELEMENTS = "script,iframe,object,embed,svg,math,form,input,button,textarea,select,meta,link,base";

  function mediaQuery(query) {
    if (window.matchMedia) { return window.matchMedia(query); }
    return { matches: false, addEventListener: function () {}, addListener: function () {} };
  }

  var compactMedia = mediaQuery("(max-width: 1199px)");
  var mobileMedia = mediaQuery("(max-width: 767px)");

  var refs = {
    workspace: document.getElementById("mail-workspace"),
    status: document.getElementById("mail-status"),
    folderPane: document.getElementById("folder-pane"),
    folderBackdrop: document.getElementById("folder-backdrop"),
    foldersOpen: document.getElementById("btn-folders-open"),
    foldersClose: document.getElementById("btn-folders-close"),
    foldersLoading: document.getElementById("folders-loading"),
    folderList: document.getElementById("folder-list"),
    listPane: document.getElementById("list-pane"),
    searchForm: document.getElementById("search-form"),
    searchQuery: document.getElementById("search-query"),
    clearSearch: document.getElementById("btn-clear-search"),
    listTitle: document.getElementById("message-list-title"),
    listMode: document.getElementById("list-mode"),
    listLoading: document.getElementById("list-loading"),
    messageList: document.getElementById("message-list"),
    listEmpty: document.getElementById("list-empty"),
    prev: document.getElementById("btn-prev"),
    next: document.getElementById("btn-next"),
    pageLabel: document.getElementById("page-label"),
    detailPane: document.getElementById("detail-pane"),
    detailBack: document.getElementById("btn-detail-back"),
    detailEmpty: document.getElementById("detail-empty"),
    detailContent: document.getElementById("detail-content"),
    detailFolderLabel: document.getElementById("detail-folder-label"),
    detailTitle: document.getElementById("detail-title"),
    detailAuthor: document.getElementById("detail-author"),
    detailTo: document.getElementById("detail-to"),
    detailCc: document.getElementById("detail-cc"),
    detailCcRow: document.getElementById("detail-cc-row"),
    detailTime: document.getElementById("detail-time"),
    detailReadState: document.getElementById("detail-read-state"),
    toggleRead: document.getElementById("btn-toggle-read"),
    moveFolder: document.getElementById("move-folder"),
    bodyFrame: document.getElementById("mail-body-frame"),
    loadImages: document.getElementById("btn-load-images"),
    bodyTruncated: document.getElementById("body-truncated"),
    attachmentsSection: document.getElementById("attachments-section"),
    attachmentsList: document.getElementById("attachments-list"),
    composeBackdrop: document.getElementById("compose-backdrop"),
    composePanel: document.getElementById("compose-panel"),
    composeForm: document.getElementById("compose-form"),
    composeModeLabel: document.getElementById("compose-mode-label"),
    composeTitle: document.getElementById("compose-title"),
    composeTo: document.getElementById("compose-to"),
    composeCc: document.getElementById("compose-cc"),
    composeBcc: document.getElementById("compose-bcc"),
    composeSubject: document.getElementById("compose-subject"),
    composeBody: document.getElementById("compose-body"),
    composeStatus: document.getElementById("compose-status"),
    send: document.getElementById("btn-send"),
  };

  var state = {
    folders: [],
    folder: { id: "inbox", name: "收件箱" },
    query: "",
    offset: 0,
    hasMore: false,
    items: [],
    selectedId: null,
    detail: null,
    attachments: [],
    listSeq: 0,
    detailSeq: 0,
    operation: null,
    composeMode: "new",
    composeItemId: null,
    composeDirty: false,
    allowRemoteImages: false,
    mobileView: "list",
    folderDrawerOpen: false,
    composeReturnFocus: null,
  };

  function ApiError(code, message, data) {
    this.name = "ApiError";
    this.code = code;
    this.message = message;
    this.data = data || {};
  }
  ApiError.prototype = Object.create(Error.prototype);

  function apiBearerToken() {
    var parts = String(BOOT.authPath || "").split("/").filter(Boolean);
    return decodeURIComponent(parts[parts.length - 1] || "");
  }

  function apiCall(path, payload) {
    var controller = new AbortController();
    var timeoutId = window.setTimeout(function () { controller.abort(); }, API_TIMEOUT_MS);
    return fetch(String(BOOT.apiBase || "") + "/" + path, {
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
          try { body = JSON.parse(text); }
          catch (error) { throw new ApiError("network_error", "服务返回了无法解析的响应。"); }
        }
        if (!body.ok) {
          throw new ApiError(body.error || "unknown", body.message || "请求失败。", body.data || {});
        }
        return body.data || {};
      });
    }).catch(function (error) {
      if (error instanceof ApiError) { throw error; }
      if (error && error.name === "AbortError") {
        throw new ApiError("network_timeout", "请求超过 60 秒未返回，请先在 Outlook 中核对状态。");
      }
      throw new ApiError("network_error", error && error.message ? error.message : "网络请求失败。");
    }).finally(function () { window.clearTimeout(timeoutId); });
  }

  function showStatus(message, kind) {
    refs.status.textContent = message;
    refs.status.dataset.kind = kind || "error";
    refs.status.hidden = false;
  }

  function clearStatus() {
    refs.status.hidden = true;
    refs.status.textContent = "";
    refs.status.removeAttribute("data-kind");
  }

  function setNodeInert(node, inert) {
    if (!node) { return; }
    node.inert = Boolean(inert);
    if (inert) { node.setAttribute("inert", ""); }
    else { node.removeAttribute("inert"); }
  }

  function syncBodyOverlay() {
    if (!document.body || !document.body.classList) { return; }
    document.body.classList.toggle(
      "mail-overlay-open",
      state.folderDrawerOpen || !refs.composePanel.hidden
    );
  }

  function renderLayout() {
    var compact = Boolean(compactMedia.matches);
    var mobile = Boolean(mobileMedia.matches);
    var folderOpen = compact && state.folderDrawerOpen;
    var detailOpen = mobile && state.mobileView === "detail";
    var listInactive = detailOpen || folderOpen;
    var detailInactive = (mobile && !detailOpen) || folderOpen;

    refs.workspace.dataset.mobileView = state.mobileView;
    refs.folderPane.dataset.open = folderOpen ? "true" : "false";
    refs.folderPane.setAttribute("aria-hidden", compact && !folderOpen ? "true" : "false");
    refs.foldersOpen.setAttribute("aria-expanded", folderOpen ? "true" : "false");
    refs.folderBackdrop.hidden = !folderOpen;
    setNodeInert(refs.folderPane, compact && !folderOpen);

    refs.listPane.hidden = detailOpen;
    refs.listPane.setAttribute("aria-hidden", listInactive ? "true" : "false");
    setNodeInert(refs.listPane, listInactive);

    refs.detailPane.hidden = mobile && !detailOpen;
    refs.detailPane.setAttribute("aria-hidden", detailInactive ? "true" : "false");
    setNodeInert(refs.detailPane, detailInactive);
    syncBodyOverlay();
  }

  function setMobileView(view, pushHistory) {
    var previous = state.mobileView;
    state.mobileView = view === "detail" ? "detail" : "list";
    if (pushHistory && mobileMedia.matches && previous !== "detail" && window.history && window.history.pushState) {
      window.history.pushState({ outlookMailView: "detail" }, "", window.location.href);
    }
    renderLayout();
  }

  function focusLater(node) {
    if (!node || !node.focus) { return; }
    window.setTimeout(function () { node.focus(); }, 0);
  }

  function returnToList(useHistory) {
    if (useHistory && window.history && window.history.state &&
        window.history.state.outlookMailView === "detail" && window.history.back) {
      window.history.back();
      return;
    }
    setMobileView("list", false);
    focusLater(refs.listTitle);
  }

  function openFolders() {
    if (!compactMedia.matches) { return; }
    state.folderDrawerOpen = true;
    renderLayout();
    focusLater(refs.folderList.children[0] || refs.foldersClose);
  }

  function closeFolders(restoreFocus) {
    if (!state.folderDrawerOpen) { return; }
    state.folderDrawerOpen = false;
    renderLayout();
    if (restoreFocus) { focusLater(refs.foldersOpen); }
  }

  function handleMediaChange() {
    if (!compactMedia.matches) { state.folderDrawerOpen = false; }
    renderLayout();
  }

  function listenForMediaChange(query, handler) {
    if (query.addEventListener) { query.addEventListener("change", handler); }
    else if (query.addListener) { query.addListener(handler); }
  }

  function formatDate(value) {
    if (!value) { return ""; }
    var date = new Date(value);
    return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString("zh-CN");
  }

  function formatListDate(value) {
    if (!value) { return ""; }
    var date = new Date(value);
    if (Number.isNaN(date.getTime())) { return String(value); }
    var now = new Date();
    if (date.toDateString() === now.toDateString()) {
      return date.toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" });
    }
    return date.toLocaleDateString("zh-CN", {
      year: date.getFullYear() === now.getFullYear() ? undefined : "numeric",
      month: "2-digit",
      day: "2-digit",
    });
  }

  function senderInitials(value) {
    var name = String(value || "?").replace(/<[^>]+>/g, "").trim().split("@")[0];
    var parts = name.split(/[\s._-]+/).filter(Boolean);
    if (parts.length > 1) {
      return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
    }
    return Array.from(name || "?").slice(0, 2).join("").toUpperCase();
  }

  function formatSize(value) {
    var size = Number(value || 0);
    if (size < 1024) { return size + " B"; }
    if (size < 1024 * 1024) { return (size / 1024).toFixed(1) + " KiB"; }
    return (size / (1024 * 1024)).toFixed(1) + " MiB";
  }

  function parseAddresses(value) {
    return String(value || "").split(/[;,\n]/).map(function (item) {
      return item.trim();
    }).filter(Boolean);
  }

  function safeUrl(value, allowMailto) {
    var raw = String(value || "").trim();
    if (!raw) { return ""; }
    if (!/^(https?:|mailto:)/i.test(raw)) { return ""; }
    try {
      var parsed = new URL(raw, window.location.origin);
      if (parsed.protocol === "http:" || parsed.protocol === "https:") { return parsed.href; }
      if (allowMailto && parsed.protocol === "mailto:") { return raw; }
    } catch (error) {
      return "";
    }
    return "";
  }

  function safeStyle(value) {
    var style = String(value || "");
    return /url\s*\(|expression\s*\(|behavior\s*:|-moz-binding|@import/i.test(style) ? "" : style;
  }

  function escapeHtml(value) {
    return String(value || "").replace(/[&<>"']/g, function (char) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[char];
    });
  }

  function hasRemoteImages(value) {
    return /<img\b[^>]*\bsrc\s*=\s*["']?https?:/i.test(String(value || ""));
  }

  function sanitizeMailHtml(source, allowRemoteImages) {
    var parsed = new DOMParser().parseFromString(String(source || ""), "text/html");
    parsed.querySelectorAll(BLOCKED_ELEMENTS).forEach(function (node) { node.remove(); });
    parsed.querySelectorAll("*").forEach(function (node) {
      Array.from(node.attributes || []).forEach(function (attribute) {
        var name = attribute.name.toLowerCase();
        if (/^on/.test(name) || name === "srcdoc" || name === "action" || name === "formaction") {
          node.removeAttribute(attribute.name);
          return;
        }
        if (["srcset", "ping", "usemap", "longdesc"].indexOf(name) >= 0) {
          node.removeAttribute(attribute.name);
          return;
        }
        if (name === "style") {
          var style = safeStyle(attribute.value);
          if (style) { node.setAttribute("style", style); }
          else { node.removeAttribute("style"); }
        }
        if (["src", "href", "xlink:href", "poster", "background"].indexOf(name) >= 0 &&
            node.tagName !== "A" && node.tagName !== "IMG") {
          node.removeAttribute(attribute.name);
        }
      });
      if (node.tagName === "STYLE") {
        node.textContent = String(node.textContent || "")
          .replace(/@import[^;]+;?/gi, "")
          .replace(/url\s*\([^)]*\)/gi, "none");
      }
      if (node.tagName === "A") {
        var href = safeUrl(node.getAttribute("href"), true);
        if (href) { node.setAttribute("href", href); }
        else { node.removeAttribute("href"); }
        node.setAttribute("target", "_blank");
        node.setAttribute("rel", "noopener noreferrer");
      }
      if (node.tagName === "IMG") {
        var src = String(node.getAttribute("src") || "").trim();
        var remote = safeUrl(src, false);
        var inlineData = /^data:image\/(png|jpeg|gif|webp);base64,/i.test(src);
        if (inlineData || (allowRemoteImages && remote)) {
          node.setAttribute("src", inlineData ? src : remote);
        } else {
          node.removeAttribute("src");
          node.setAttribute("alt", src.toLowerCase().startsWith("cid:") ? "[内嵌图片暂不可用]" : "[外部图片已阻止]");
        }
      }
    });
    return parsed.body.innerHTML;
  }

  function bodyDocument(detail, allowRemoteImages) {
    var body = String(detail.body || "");
    var content = detail.body_content_type === "html"
      ? sanitizeMailHtml(body, allowRemoteImages)
      : "<pre class=\"plain-text\">" + escapeHtml(body) + "</pre>";
    if (!body.trim() || !String(content).trim()) {
      content = "<p class=\"empty-body\">此邮件没有可显示的正文。</p>";
    }
    var imagePolicy = allowRemoteImages ? "data: http: https:" : "data:";
    return "<!doctype html><html><head><meta charset=\"utf-8\"><meta name=\"referrer\" content=\"no-referrer\">" +
      "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src " + imagePolicy +
      "; style-src 'unsafe-inline';\"><base target=\"_blank\"><style>body{font:14px/1.55 system-ui,sans-serif;color:#172033;margin:16px;overflow-wrap:anywhere}.plain-text{white-space:pre-wrap;font:inherit}img{max-width:100%;height:auto}table{max-width:100%;border-collapse:collapse}td,th{padding:4px}</style></head><body>" +
      content + "</body></html>";
  }

  function replaceBodyFrame(documentHtml) {
    var previous = refs.bodyFrame;
    var replacement = previous.cloneNode(false);
    replacement.srcdoc = documentHtml;
    previous.replaceWith(replacement);
    refs.bodyFrame = replacement;
  }

  function createText(tagName, className, text) {
    var node = document.createElement(tagName);
    if (className) { node.className = className; }
    node.textContent = text == null ? "" : String(text);
    return node;
  }

  function currentFolderName() {
    return state.folder && state.folder.name ? state.folder.name : "邮件";
  }

  function renderFolders() {
    refs.folderList.replaceChildren();
    refs.moveFolder.replaceChildren();
    state.folders.forEach(function (folder) {
      var button = document.createElement("button");
      button.type = "button";
      button.className = "folder-button";
      button.dataset.folderId = folder.id;
      button.setAttribute("aria-current", folder.id === state.folder.id ? "page" : "false");
      button.appendChild(createText("span", "", folder.name || folder.id));
      button.appendChild(createText("span", "", String(folder.unread_count || 0)));
      button.addEventListener("click", function () { selectFolder(folder); });
      refs.folderList.appendChild(button);

      if (folder.id !== state.folder.id) {
        var option = document.createElement("option");
        option.value = folder.id;
        option.textContent = folder.name || folder.id;
        refs.moveFolder.appendChild(option);
      }
    });
  }

  function renderList() {
    refs.messageList.replaceChildren();
    state.items.forEach(function (item) {
      var sender = item.author || item.sender || "未知发件人";
      var subject = item.subject || "（无主题）";
      var receivedAt = item.datetime_received || item.datetime_sent;
      var button = document.createElement("button");
      button.type = "button";
      button.className = "message-card";
      button.dataset.itemId = item.id;
      button.dataset.unread = item.is_read ? "false" : "true";
      button.setAttribute("aria-current", item.id === state.selectedId ? "true" : "false");
      button.setAttribute("aria-label", sender + "，" + subject + "，" + formatDate(receivedAt));
      button.title = subject;

      button.appendChild(createText("span", "message-card__avatar", senderInitials(sender)));

      var content = document.createElement("span");
      content.className = "message-card__content";

      var top = document.createElement("div");
      top.className = "message-card__top";
      top.appendChild(createText("span", "message-card__sender", sender));
      top.appendChild(createText("time", "", formatListDate(receivedAt)));
      content.appendChild(top);
      content.appendChild(createText("p", "message-card__subject", subject));

      var meta = document.createElement("div");
      meta.className = "message-card__meta";
      meta.appendChild(createText("span", "", item.has_attachments ? "含附件" : ""));
      meta.appendChild(createText("span", "", item.importance === "high" ? "高优先级" : ""));
      content.appendChild(meta);
      button.appendChild(content);
      button.addEventListener("click", function () { selectMessage(item.id); });
      refs.messageList.appendChild(button);
    });
    refs.listEmpty.hidden = state.items.length !== 0;
    refs.prev.disabled = state.offset <= 0;
    refs.next.disabled = !state.hasMore;
    refs.pageLabel.textContent = "第 " + (Math.floor(state.offset / PAGE_SIZE) + 1) + " 页";
    refs.listTitle.textContent = currentFolderName();
    refs.listMode.textContent = state.query ? "主题搜索：" + state.query : "最近邮件";
    refs.clearSearch.hidden = !state.query;
  }

  function clearDetail() {
    state.selectedId = null;
    state.detail = null;
    state.attachments = [];
    state.allowRemoteImages = false;
    state.mobileView = "list";
    refs.detailContent.hidden = true;
    refs.detailEmpty.hidden = false;
    refs.bodyFrame.srcdoc = "";
    renderLayout();
    renderList();
  }

  function renderAttachments() {
    refs.attachmentsList.replaceChildren();
    state.attachments.forEach(function (attachment) {
      var text = (attachment.name || "未命名附件") + " · " +
        (attachment.content_type || "未知类型") + " · " + formatSize(attachment.size);
      refs.attachmentsList.appendChild(createText("li", "", text));
    });
    refs.attachmentsSection.hidden = state.attachments.length === 0;
  }

  function renderDetail() {
    var detail = state.detail;
    if (!detail) { clearDetail(); return; }
    refs.detailEmpty.hidden = true;
    refs.detailContent.hidden = false;
    refs.detailFolderLabel.textContent = currentFolderName();
    refs.detailTitle.textContent = detail.subject || "（无主题）";
    refs.detailAuthor.textContent = detail.author || detail.sender || "";
    refs.detailTo.textContent = detail.display_to || "";
    refs.detailCc.textContent = detail.display_cc || "";
    refs.detailCcRow.hidden = !detail.display_cc;
    refs.detailTime.textContent = formatDate(detail.datetime_received || detail.datetime_sent);
    refs.detailReadState.textContent = detail.is_read ? "已读" : "未读";
    refs.toggleRead.textContent = detail.is_read ? "标记未读" : "标记已读";
    refs.bodyTruncated.hidden = !detail.body_truncated;
    refs.loadImages.hidden = detail.body_content_type !== "html" || !hasRemoteImages(detail.body) || state.allowRemoteImages;
    replaceBodyFrame(bodyDocument(detail, state.allowRemoteImages));
    renderAttachments();
    renderFolders();
    renderList();
  }

  function selectFolder(folder) {
    if (!folder) { return; }
    if (folder.id === state.folder.id) {
      closeFolders(true);
      return;
    }
    state.folder = folder;
    state.query = "";
    state.offset = 0;
    refs.searchQuery.value = "";
    closeFolders(false);
    clearDetail();
    renderFolders();
    loadMessages();
  }

  function loadFolders() {
    refs.foldersLoading.hidden = false;
    return apiCall("mail/folders", {}).then(function (data) {
      state.folders = Array.isArray(data.items) ? data.items : [];
      var inbox = state.folders.find(function (folder) {
        var name = String(folder.name || "").toLowerCase();
        return name === "inbox" || name === "收件箱" || String(folder.id || "").toLowerCase() === "inbox";
      });
      var selected = state.folders.find(function (folder) { return folder.id === state.folder.id; });
      if (selected) { state.folder = selected; }
      else if (inbox) { state.folder = inbox; }
      renderFolders();
    }).catch(function (error) {
      state.folders = [state.folder];
      renderFolders();
      showStatus("加载文件夹失败：" + error.message);
    }).finally(function () { refs.foldersLoading.hidden = true; });
  }

  function loadMessages() {
    var seq = ++state.listSeq;
    refs.listLoading.hidden = false;
    var path = state.query ? "mail/search" : "mail/list";
    var payload = {
      folder: state.folder.id,
      limit: PAGE_SIZE,
      offset: state.offset,
    };
    if (state.query) { payload.query = state.query; }
    return apiCall(path, payload).then(function (data) {
      if (seq !== state.listSeq) { return; }
      state.items = Array.isArray(data.items) ? data.items : [];
      state.hasMore = Boolean(data.has_more);
      renderList();
    }).catch(function (error) {
      if (seq !== state.listSeq) { return; }
      state.items = [];
      state.hasMore = false;
      renderList();
      showStatus("加载邮件失败：" + error.message);
    }).finally(function () {
      if (seq === state.listSeq) { refs.listLoading.hidden = true; }
    });
  }

  function selectMessage(itemId) {
    state.selectedId = itemId;
    state.allowRemoteImages = false;
    state.attachments = [];
    setMobileView("detail", true);
    renderList();
    var seq = ++state.detailSeq;
    refs.detailEmpty.hidden = false;
    refs.detailEmpty.querySelector("h2").textContent = "加载邮件中…";
    refs.detailContent.hidden = true;
    return apiCall("mail/get", {
      item_id: itemId,
      folder: state.folder.id,
      include_body: true,
      body_max_chars: BODY_MAX_CHARS,
    }).then(function (detail) {
      if (seq !== state.detailSeq || itemId !== state.selectedId) { return; }
      state.detail = detail;
      renderDetail();
      if (mobileMedia.matches) { focusLater(refs.detailTitle); }
      if (detail.has_attachments) { loadAttachments(itemId, seq); }
    }).catch(function (error) {
      if (seq !== state.detailSeq) { return; }
      state.detail = null;
      refs.detailEmpty.hidden = false;
      refs.detailEmpty.querySelector("h2").textContent = "加载详情失败";
      refs.detailContent.hidden = true;
      showStatus("加载邮件详情失败：" + error.message);
    });
  }

  function loadAttachments(itemId, detailSeq) {
    return apiCall("mail/attachments", { item_id: itemId, folder: state.folder.id }).then(function (data) {
      if (detailSeq !== state.detailSeq || itemId !== state.selectedId) { return; }
      state.attachments = Array.isArray(data.items) ? data.items : [];
      renderAttachments();
    }).catch(function (error) {
      if (detailSeq === state.detailSeq) { showStatus("加载附件信息失败：" + error.message); }
    });
  }

  function editor() {
    return window.tinymce && window.tinymce.get ? window.tinymce.get("compose-body") : null;
  }

  function editorContent() {
    var instance = editor();
    return instance ? instance.getContent() : refs.composeBody.value;
  }

  function setEditorContent(value) {
    var instance = editor();
    if (instance) { instance.setContent(String(value || "")); }
    else { refs.composeBody.value = String(value || ""); }
  }

  function initEditor() {
    if (!window.tinymce || !window.tinymce.init) { return; }
    window.tinymce.init({
      selector: "#compose-body",
      license_key: "gpl",
      menubar: false,
      height: 360,
      plugins: "autolink link lists table code",
      toolbar: "undo redo | bold italic underline | bullist numlist | link table | code",
      toolbar_mode: "wrap",
      branding: false,
      promotion: false,
      setup: function (instance) {
        instance.on("change input undo redo", function () { state.composeDirty = true; });
      },
    });
  }

  function composeLabels(mode) {
    if (mode === "reply") { return { eyebrow: "回复邮件", title: "回复" }; }
    if (mode === "forward") { return { eyebrow: "转发邮件", title: "转发" }; }
    return { eyebrow: "新邮件", title: "写信" };
  }

  function openCompose(mode) {
    var detail = state.detail || {};
    var labels = composeLabels(mode);
    state.composeReturnFocus = document.activeElement || null;
    state.composeMode = mode;
    state.composeItemId = mode === "new" ? null : state.selectedId;
    state.composeDirty = false;
    refs.composeModeLabel.textContent = labels.eyebrow;
    refs.composeTitle.textContent = labels.title;
    refs.composeTo.value = mode === "reply" ? (detail.author || detail.sender || "") : "";
    refs.composeCc.value = "";
    refs.composeBcc.value = "";
    refs.composeCc.disabled = mode === "forward";
    refs.composeBcc.disabled = mode === "forward";
    refs.composeSubject.value = mode === "reply"
      ? replySubject(detail.subject)
      : mode === "forward" ? forwardSubject(detail.subject) : "";
    refs.composeStatus.textContent = mode === "forward" ? "转发接口首版仅支持收件人，不支持抄送和密送。" : "";
    setEditorContent("");
    refs.composeBackdrop.hidden = false;
    refs.composePanel.hidden = false;
    refs.composePanel.setAttribute("aria-hidden", "false");
    refs.workspace.setAttribute("aria-hidden", "true");
    setNodeInert(refs.workspace, true);
    syncBodyOverlay();
    window.setTimeout(function () { refs.composeTo.focus(); }, 0);
  }

  function replySubject(subject) {
    var value = String(subject || "（无主题）");
    return /^re:/i.test(value) ? value : "Re: " + value;
  }

  function forwardSubject(subject) {
    var value = String(subject || "（无主题）");
    return /^fwd:/i.test(value) ? value : "Fwd: " + value;
  }

  function closeCompose(force) {
    if (!force && state.composeDirty && !window.confirm("当前内容尚未发送，确认关闭？")) { return; }
    var returnFocus = state.composeReturnFocus;
    refs.composeBackdrop.hidden = true;
    refs.composePanel.hidden = true;
    refs.composePanel.setAttribute("aria-hidden", "true");
    refs.workspace.setAttribute("aria-hidden", "false");
    setNodeInert(refs.workspace, false);
    refs.composeStatus.textContent = "";
    state.composeDirty = false;
    state.composeItemId = null;
    state.composeReturnFocus = null;
    syncBodyOverlay();
    focusLater(returnFocus);
  }

  function trapComposeFocus(event) {
    if (event.key !== "Tab" || refs.composePanel.hidden || !refs.composePanel.querySelectorAll) { return; }
    var focusable = Array.from(refs.composePanel.querySelectorAll(
      "button:not([disabled]), input:not([disabled]), textarea:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex='-1'])"
    )).filter(function (node) { return !node.hidden; });
    if (!focusable.length) { return; }
    var first = focusable[0];
    var last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }

  function setOperationBusy(busy, label) {
    refs.send.disabled = busy;
    refs.send.textContent = busy ? (label || "处理中…") : "发送";
    document.querySelectorAll(".mail-detail__actions button, .mail-detail__actions select").forEach(function (node) {
      node.disabled = busy;
    });
  }

  function operationErrorMessage(error) {
    if (error.code === "outcome_unknown" || error.code === "network_timeout") {
      return error.message + " 请先在 Outlook 中核对结果，不要立即重试。";
    }
    if (error.code === "auth_required") {
      return "授权已失效，请返回 endpoint 授权页面重新登录。";
    }
    return error.message;
  }

  function runOperation(name, path, payload, onSuccess) {
    if (state.operation) { return Promise.resolve(); }
    clearStatus();
    state.operation = name;
    setOperationBusy(true, name);
    return apiCall(path, payload).then(function (data) {
      onSuccess(data || {});
    }).catch(function (error) {
      showStatus(operationErrorMessage(error));
    }).finally(function () {
      state.operation = null;
      setOperationBusy(false);
    });
  }

  function removeSelectedFromList() {
    var selectedId = state.selectedId;
    state.items = state.items.filter(function (item) { return item.id !== selectedId; });
    clearDetail();
    renderList();
  }

  function updateReadState(isRead) {
    if (state.detail) { state.detail.is_read = isRead; }
    state.items.forEach(function (item) {
      if (item.id === state.selectedId) { item.is_read = isRead; }
    });
    renderDetail();
  }

  function submitCompose(event) {
    event.preventDefault();
    var to = parseAddresses(refs.composeTo.value);
    var cc = parseAddresses(refs.composeCc.value);
    var bcc = parseAddresses(refs.composeBcc.value);
    var subject = refs.composeSubject.value.trim();
    var body = editorContent();
    var mode = state.composeMode;
    if ((mode === "new" || mode === "forward") && to.length === 0) {
      refs.composeStatus.textContent = "至少填写一个收件人。";
      return;
    }
    if (mode === "new" && !subject) {
      refs.composeStatus.textContent = "主题不能为空。";
      return;
    }
    if (!window.confirm("确认" + composeLabels(mode).title + "并立即发送？")) { return; }

    var path;
    var payload;
    if (mode === "reply") {
      path = "mail/reply";
      payload = {
        item_id: state.composeItemId,
        folder: state.folder.id,
        body: body,
        subject: subject || null,
        to: to.length ? to : null,
        cc: cc.length ? cc : null,
        bcc: bcc.length ? bcc : null,
        html: true,
        confirm: true,
      };
    } else if (mode === "forward") {
      path = "mail/forward";
      payload = {
        item_id: state.composeItemId,
        folder: state.folder.id,
        to: to,
        body: body,
        html: true,
        confirm: true,
      };
    } else {
      path = "mail/send";
      payload = {
        to: to,
        cc: cc.length ? cc : null,
        bcc: bcc.length ? bcc : null,
        subject: subject,
        body: body,
        html: true,
        confirm: true,
      };
    }
    refs.composeStatus.textContent = "发送中…";
    return runOperation("发送中…", path, payload, function () {
      closeCompose(true);
      showStatus("邮件已提交发送。", "success");
    });
  }

  function moveSelected() {
    if (!state.detail || !refs.moveFolder.value) { return; }
    var targetName = refs.moveFolder.options[refs.moveFolder.selectedIndex]
      ? refs.moveFolder.options[refs.moveFolder.selectedIndex].textContent : refs.moveFolder.value;
    if (!window.confirm("确认将邮件移动到“" + targetName + "”？")) { return; }
    return runOperation("移动中…", "mail/move", {
      item_id: state.selectedId,
      folder: state.folder.id,
      target_folder: refs.moveFolder.value,
      confirm: true,
    }, function () {
      removeSelectedFromList();
      showStatus("邮件已移动。", "success");
    });
  }

  function deleteSelected() {
    if (!state.detail || !window.confirm("确认将这封邮件移入已删除邮件？")) { return; }
    return runOperation("删除中…", "mail/delete", {
      item_id: state.selectedId,
      folder: state.folder.id,
      mode: "move_to_trash",
      confirm: true,
    }, function () {
      removeSelectedFromList();
      showStatus("邮件已移入已删除邮件。", "success");
    });
  }

  function toggleRead() {
    if (!state.detail) { return; }
    var markRead = !state.detail.is_read;
    if (!window.confirm("确认将邮件标记为" + (markRead ? "已读" : "未读") + "？")) { return; }
    return runOperation("更新中…", markRead ? "mail/mark_read" : "mail/mark_unread", {
      item_id: state.selectedId,
      folder: state.folder.id,
      confirm: true,
    }, function () {
      updateReadState(markRead);
      showStatus("邮件已标记为" + (markRead ? "已读。" : "未读。"), "success");
    });
  }

  function handleGlobalKeydown(event) {
    if (event.key === "Tab" && !refs.composePanel.hidden) {
      trapComposeFocus(event);
      return;
    }
    if (event.key !== "Escape") { return; }
    if (!refs.composePanel.hidden) {
      event.preventDefault();
      closeCompose(false);
      return;
    }
    if (state.folderDrawerOpen) {
      event.preventDefault();
      closeFolders(true);
      return;
    }
    if (mobileMedia.matches && state.mobileView === "detail") {
      event.preventDefault();
      returnToList(true);
    }
  }

  function handlePopState() {
    if (!mobileMedia.matches || state.mobileView !== "detail") { return; }
    setMobileView("list", false);
    focusLater(refs.listTitle);
  }

  function bindEvents() {
    document.getElementById("btn-refresh").addEventListener("click", function () {
      clearStatus();
      Promise.all([loadFolders(), loadMessages()]);
      if (state.selectedId) { selectMessage(state.selectedId); }
    });
    document.getElementById("btn-compose").addEventListener("click", function () { openCompose("new"); });
    refs.foldersOpen.addEventListener("click", openFolders);
    refs.foldersClose.addEventListener("click", function () { closeFolders(true); });
    refs.folderBackdrop.addEventListener("click", function () { closeFolders(true); });
    refs.detailBack.addEventListener("click", function () { returnToList(true); });
    document.getElementById("btn-compose-close").addEventListener("click", function () { closeCompose(false); });
    refs.composeBackdrop.addEventListener("click", function () { closeCompose(false); });
    refs.composeForm.addEventListener("submit", submitCompose);
    [refs.composeTo, refs.composeCc, refs.composeBcc, refs.composeSubject, refs.composeBody].forEach(function (node) {
      node.addEventListener("input", function () { state.composeDirty = true; });
    });
    refs.searchForm.addEventListener("submit", function (event) {
      event.preventDefault();
      state.query = refs.searchQuery.value.trim();
      state.offset = 0;
      clearDetail();
      loadMessages();
    });
    refs.clearSearch.addEventListener("click", function () {
      refs.searchQuery.value = "";
      state.query = "";
      state.offset = 0;
      clearDetail();
      loadMessages();
    });
    refs.prev.addEventListener("click", function () {
      if (state.offset <= 0) { return; }
      state.offset = Math.max(0, state.offset - PAGE_SIZE);
      clearDetail();
      loadMessages();
    });
    refs.next.addEventListener("click", function () {
      if (!state.hasMore) { return; }
      state.offset += PAGE_SIZE;
      clearDetail();
      loadMessages();
    });
    document.getElementById("btn-reply").addEventListener("click", function () {
      if (state.detail) { openCompose("reply"); }
    });
    document.getElementById("btn-forward").addEventListener("click", function () {
      if (state.detail) { openCompose("forward"); }
    });
    document.getElementById("btn-move").addEventListener("click", moveSelected);
    document.getElementById("btn-delete").addEventListener("click", deleteSelected);
    refs.toggleRead.addEventListener("click", toggleRead);
    refs.loadImages.addEventListener("click", function () {
      if (!state.detail) { return; }
      state.allowRemoteImages = true;
      renderDetail();
    });
    document.addEventListener("keydown", handleGlobalKeydown);
    window.addEventListener("popstate", handlePopState);
    window.addEventListener("beforeunload", function (event) {
      if (!state.composeDirty || refs.composePanel.hidden) { return; }
      event.preventDefault();
      event.returnValue = "";
    });
    listenForMediaChange(compactMedia, handleMediaChange);
    listenForMediaChange(mobileMedia, handleMediaChange);
  }

  function stateSnapshot() {
    return {
      folder: { id: state.folder.id, name: state.folder.name },
      query: state.query,
      offset: state.offset,
      hasMore: state.hasMore,
      selectedId: state.selectedId,
      composeMode: state.composeMode,
      operation: state.operation,
      mobileView: state.mobileView,
      folderDrawerOpen: state.folderDrawerOpen,
    };
  }

  function setDetailForTest(detail) {
    state.selectedId = detail.id;
    state.detail = detail;
    state.items = [detail];
    state.attachments = [];
    state.allowRemoteImages = false;
    setMobileView("detail", false);
    renderDetail();
  }

  function setViewportForTest(compact, mobile) {
    compactMedia.matches = Boolean(compact);
    mobileMedia.matches = Boolean(mobile);
    handleMediaChange();
  }

  if (window.__OUTLOOK_MAIL_TESTING__) {
    window.__OUTLOOK_MAIL_TEST_HOOKS__ = {
      parseAddresses: parseAddresses,
      safeUrl: safeUrl,
      safeStyle: safeStyle,
      hasRemoteImages: hasRemoteImages,
      replySubject: replySubject,
      forwardSubject: forwardSubject,
      stateSnapshot: stateSnapshot,
      setDetail: setDetailForTest,
      openCompose: openCompose,
      submitCompose: submitCompose,
      moveSelected: moveSelected,
      deleteSelected: deleteSelected,
      toggleRead: toggleRead,
      returnToList: returnToList,
      openFolders: openFolders,
      closeFolders: closeFolders,
      setViewport: setViewportForTest,
    };
  }

  renderLayout();
  bindEvents();
  initEditor();
  Promise.all([loadFolders(), loadMessages()]);
})();
