var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
var DocumentTimelineWebPart = /** @class */ (function (_super) {
    __extends(DocumentTimelineWebPart, _super);
    function DocumentTimelineWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(DocumentTimelineWebPart.prototype, "siteUrl", {
        get: function () { return this.properties.siteUrl || this.context.pageContext.web.absoluteUrl; },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(DocumentTimelineWebPart.prototype, "listName", {
        get: function () { return this.properties.listName || 'Document Log Tracking'; },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(DocumentTimelineWebPart.prototype, "statusLogListName", {
        get: function () { return this.properties.statusLogListName || 'Document Status Log'; },
        enumerable: false,
        configurable: true
    });
    // ── Render ──────────────────────────────────────────────────────────────
    DocumentTimelineWebPart.prototype.render = function () {
        var user = this.context.pageContext.user;
        this.domElement.innerHTML = "\n      <style>\n        .dtl-wrap { font-family: 'Segoe UI', sans-serif; max-width: 700px; margin: 0 auto; padding: 24px 16px; color: #1c1a16; }\n        .dtl-header { text-align: center; margin-bottom: 28px; padding-bottom: 20px; border-bottom: 2px solid #1c1a16; }\n        .dtl-header h2 { font-size: 22px; font-weight: 700; margin: 0 0 6px; }\n        .dtl-header p { font-size: 13px; color: #7a7368; margin: 0; }\n\n        .dtl-card { background: #fff; border: 1px solid #d8d2c8; padding: 28px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }\n        .dtl-card-title { font-size: 15px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid #d8d2c8; display: flex; align-items: center; gap: 8px; }\n        .dtl-card-title .bar { width: 3px; height: 14px; background: #c9a84c; flex-shrink: 0; }\n\n        .dtl-field { margin-bottom: 18px; }\n        .dtl-label { display: block; font-size: 11px; font-weight: 600; color: #7a7368; margin-bottom: 5px; letter-spacing: 0.07em; text-transform: uppercase; }\n        .dtl-label .req { color: #8b4513; }\n        .dtl-input { width: 100%; padding: 9px 12px; border: 1px solid #d8d2c8; background: #f7f5f0; font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #1c1a16; outline: none; box-sizing: border-box; transition: border-color 0.15s; }\n        .dtl-input:focus { border-color: #8b4513; background: #fff; }\n\n        .dtl-notice { padding: 11px 14px; font-size: 13px; display: flex; gap: 8px; align-items: flex-start; margin-bottom: 18px; line-height: 1.5; }\n        .dtl-notice.info { background: #f5ece4; border-left: 3px solid #c8773a; color: #8b4513; }\n        .dtl-notice.success { background: #eaf3ed; border-left: 3px solid #2d6a4f; color: #2d6a4f; }\n        .dtl-notice.error { background: #fdf5f5; border-left: 3px solid #c0392b; color: #c0392b; }\n\n        .dtl-btn { padding: 10px 22px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.15s; letter-spacing: 0.03em; }\n        .dtl-btn-primary { background: #8b4513; color: #fff; }\n        .dtl-btn-primary:hover { background: #6d3410; }\n        .dtl-btn-primary:disabled { background: #d8d2c8; color: #7a7368; cursor: not-allowed; }\n        .dtl-btn-outline { background: transparent; color: #1c1a16; border: 1px solid #b0a898; }\n        .dtl-btn-outline:hover { border-color: #1c1a16; }\n        .dtl-btn-row { display: flex; justify-content: space-between; align-items: center; margin-top: 24px; gap: 10px; flex-wrap: wrap; }\n\n        .dtl-rtable { width: 100%; border-collapse: collapse; }\n        .dtl-rtable tr { border-bottom: 1px solid #d8d2c8; }\n        .dtl-rtable tr:last-child { border-bottom: none; }\n        .dtl-rtable td { padding: 9px 0; font-size: 13px; vertical-align: top; }\n        .dtl-rtable td:first-child { color: #7a7368; font-size: 11px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; width: 36%; padding-right: 14px; padding-top: 11px; }\n\n        .dtl-search-row { display: flex; gap: 10px; align-items: flex-start; }\n        .dtl-search-row .dtl-input { flex: 1; }\n\n        .dtl-spinner { width: 16px; height: 16px; border: 2px solid #d8d2c8; border-top-color: #8b4513; border-radius: 50%; animation: dtlspin 0.75s linear infinite; flex-shrink: 0; display: inline-block; vertical-align: middle; }\n        @keyframes dtlspin { to { transform: rotate(360deg); } }\n\n        .dtl-screen { display: none; }\n        .dtl-screen.active { display: block; }\n\n        .dtl-code { display: inline-block; background: #1c1a16; color: #f7f3e8; font-family: 'Courier New', monospace; font-size: 22px; font-weight: 700; letter-spacing: 0.18em; padding: 12px 24px; border: 2px solid #c9a84c; }\n\n        /* Timeline */\n        .dtl-timeline { position: relative; padding: 0; margin: 0; list-style: none; }\n        .dtl-timeline::before { content: ''; position: absolute; left: 15px; top: 0; bottom: 0; width: 2px; background: #d8d2c8; }\n\n        .dtl-tl-node { position: relative; padding: 0 0 28px 44px; }\n        .dtl-tl-node:last-child { padding-bottom: 0; }\n        .dtl-tl-dot { position: absolute; left: 8px; top: 2px; width: 16px; height: 16px; border-radius: 50%; border: 2px solid #2d6a4f; background: #2d6a4f; z-index: 1; }\n        .dtl-tl-node.current .dtl-tl-dot { border-color: #8b4513; background: #8b4513; animation: dtlpulse 2s infinite; }\n        .dtl-tl-node.origin .dtl-tl-dot { border-color: #7a7368; background: #7a7368; }\n        @keyframes dtlpulse { 0%,100%{box-shadow:0 0 0 0 rgba(139,69,19,0.4)} 50%{box-shadow:0 0 0 8px rgba(139,69,19,0)} }\n\n        .dtl-tl-content { background: #f7f5f0; border: 1px solid #d8d2c8; padding: 14px 16px; }\n        .dtl-tl-node.current .dtl-tl-content { border-color: #8b4513; background: #fdf8f3; }\n        .dtl-tl-node.origin .dtl-tl-content { border-color: #b0a898; background: #f7f5f0; }\n\n        .dtl-tl-transition { font-size: 14px; font-weight: 700; margin-bottom: 6px; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }\n        .dtl-tl-arrow { color: #8b4513; font-weight: 400; }\n        .dtl-tl-meta { font-size: 11px; color: #7a7368; line-height: 1.6; }\n        .dtl-tl-meta span { margin-right: 14px; }\n\n        .dtl-tl-duration { display: inline-block; font-family: 'Courier New', monospace; font-size: 11px; font-weight: 700; background: #eaf3ed; color: #2d6a4f; padding: 2px 8px; letter-spacing: 0.04em; }\n        .dtl-tl-node.current .dtl-tl-duration { background: #f5ece4; color: #8b4513; }\n\n        .dtl-total-bar { display: flex; align-items: center; justify-content: center; gap: 12px; padding: 16px; background: #1c1a16; color: #f7f3e8; margin-top: 14px; }\n        .dtl-total-label { font-size: 10px; font-weight: 600; letter-spacing: 0.14em; text-transform: uppercase; color: #b0a898; }\n        .dtl-total-value { font-family: 'Courier New', monospace; font-size: 22px; font-weight: 700; letter-spacing: 0.08em; color: #c9a84c; }\n\n        .dtl-status-badge { display: inline-block; padding: 3px 10px; font-size: 11px; font-weight: 600; letter-spacing: 0.05em; text-transform: uppercase; }\n        .dtl-status-badge.st-received { background: #f5ece4; color: #8b4513; }\n        .dtl-status-badge.st-review { background: #e8f0fe; color: #1a56db; }\n        .dtl-status-badge.st-dca { background: #fef3cd; color: #856404; }\n        .dtl-status-badge.st-released { background: #eaf3ed; color: #2d6a4f; }\n        .dtl-status-badge.st-filed { background: #e2e2e2; color: #555; }\n\n        .dtl-empty-tl { text-align: center; padding: 32px 16px; color: #7a7368; font-size: 13px; }\n        .dtl-empty-tl .icon { font-size: 28px; margin-bottom: 8px; }\n      </style>\n\n      <div class=\"dtl-wrap\">\n        <div class=\"dtl-header\">\n          <h2>Document Timeline</h2>\n          <p>Logged in as <strong>".concat(this._esc(user.displayName), "</strong> &nbsp;&middot;&nbsp; ").concat(this._esc(user.email), "</p>\n        </div>\n\n        <!-- Screen 1: Search -->\n        <div class=\"dtl-screen active\" id=\"dtl-s1\">\n          <div class=\"dtl-card\">\n            <div class=\"dtl-card-title\"><span class=\"bar\"></span>Search Document</div>\n            <div class=\"dtl-field\">\n              <label class=\"dtl-label\">Reference Code <span class=\"req\">*</span></label>\n              <div class=\"dtl-search-row\">\n                <input id=\"dtl-search\" type=\"text\" class=\"dtl-input\" placeholder=\"e.g. RCM-SC-0009\" />\n                <button id=\"dtl-btn-search\" class=\"dtl-btn dtl-btn-primary\">Search</button>\n              </div>\n            </div>\n            <div id=\"dtl-search-status\" style=\"margin-top:12px;\"></div>\n          </div>\n        </div>\n\n        <!-- Screen 2: Timeline -->\n        <div class=\"dtl-screen\" id=\"dtl-s2\">\n          <div class=\"dtl-card\" id=\"dtl-summary-card\"></div>\n          <div class=\"dtl-card\">\n            <div class=\"dtl-card-title\"><span class=\"bar\"></span>Status Timeline</div>\n            <div id=\"dtl-timeline-body\"></div>\n          </div>\n          <div id=\"dtl-total-bar\"></div>\n          <div class=\"dtl-btn-row\" style=\"margin-top:14px;\">\n            <button id=\"dtl-btn-back\" class=\"dtl-btn dtl-btn-outline\">&larr; Search Again</button>\n          </div>\n        </div>\n      </div>\n    ");
        this._bindEvents();
    };
    // ── Events ──────────────────────────────────────────────────────────────
    DocumentTimelineWebPart.prototype._bindEvents = function () {
        var _this = this;
        var searchInput = this.domElement.querySelector('#dtl-search');
        searchInput.addEventListener('keydown', function (e) {
            if (e.key === 'Enter')
                _this._doSearch();
        });
        this.domElement.querySelector('#dtl-btn-search').addEventListener('click', function () { return _this._doSearch(); });
        this.domElement.querySelector('#dtl-btn-back').addEventListener('click', function () {
            _this._showScreen(1);
            _this.domElement.querySelector('#dtl-search').value = '';
            _this._el('dtl-search-status').innerHTML = '';
        });
    };
    // ── Search ──────────────────────────────────────────────────────────────
    DocumentTimelineWebPart.prototype._doSearch = function () {
        return __awaiter(this, void 0, void 0, function () {
            var input, code, statusEl, searchBtn, filterCode, _a, docResponse, logResponse, docData, logData, items, doc, logs, err_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        input = this.domElement.querySelector('#dtl-search');
                        code = input.value.trim();
                        statusEl = this._el('dtl-search-status');
                        searchBtn = this.domElement.querySelector('#dtl-btn-search');
                        if (!code) {
                            statusEl.innerHTML = '<div class="dtl-notice error">&cross; &nbsp;Please enter a reference code.</div>';
                            return [2 /*return*/];
                        }
                        searchBtn.disabled = true;
                        statusEl.innerHTML = '<div style="display:flex;align-items:center;gap:10px;font-size:13px;color:#7a7368;"><div class="dtl-spinner"></div>Searching...</div>';
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 5, , 6]);
                        filterCode = encodeURIComponent(code);
                        return [4 /*yield*/, Promise.all([
                                this.context.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items?$filter=ReferenceCode eq '").concat(filterCode, "'"), SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } }),
                                this.context.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.statusLogListName, "')/items?$filter=Title eq '").concat(filterCode, "'&$orderby=Created asc"), SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } })
                            ])];
                    case 2:
                        _a = _b.sent(), docResponse = _a[0], logResponse = _a[1];
                        if (!docResponse.ok) {
                            throw new Error("Document list returned ".concat(docResponse.status));
                        }
                        if (!logResponse.ok) {
                            throw new Error("Status log list returned ".concat(logResponse.status));
                        }
                        return [4 /*yield*/, docResponse.json()];
                    case 3:
                        docData = _b.sent();
                        return [4 /*yield*/, logResponse.json()];
                    case 4:
                        logData = _b.sent();
                        items = docData.value;
                        if (!items || items.length === 0) {
                            statusEl.innerHTML = "<div class=\"dtl-notice error\">&cross; &nbsp;No document found with reference code <strong>".concat(this._esc(code), "</strong>.</div>");
                            searchBtn.disabled = false;
                            return [2 /*return*/];
                        }
                        doc = items[0];
                        logs = logData.value || [];
                        this._renderTimeline(doc, logs);
                        searchBtn.disabled = false;
                        return [3 /*break*/, 6];
                    case 5:
                        err_1 = _b.sent();
                        console.error('DocumentTimeline search error:', err_1);
                        statusEl.innerHTML = '<div class="dtl-notice error">&cross; &nbsp;Could not search SharePoint. Please check your connection and try again.</div>';
                        searchBtn.disabled = false;
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    // ── Render timeline ─────────────────────────────────────────────────────
    DocumentTimelineWebPart.prototype._renderTimeline = function (doc, logs) {
        // Summary card
        var createdDate = new Date(doc.Created);
        var formatLabel = doc.Document_x0020_Format === 'HC' ? 'Physical (HC)' : 'Digital (SC)';
        this._el('dtl-summary-card').innerHTML = "\n      <div style=\"text-align:center;margin-bottom:18px;\">\n        <div style=\"font-family:'Courier New',monospace;font-size:10px;letter-spacing:0.16em;text-transform:uppercase;color:#7a7368;margin-bottom:8px;\">Reference Code</div>\n        <div class=\"dtl-code\">".concat(this._esc(doc.ReferenceCode), "</div>\n      </div>\n      <table class=\"dtl-rtable\">\n        <tr><td>Document Title</td><td>").concat(this._esc(doc.Title), "</td></tr>\n        <tr><td>Document Type</td><td>").concat(this._esc(doc.Document_x0020_Type), "</td></tr>\n        <tr><td>Document Format</td><td>").concat(formatLabel, "</td></tr>\n        <tr><td>From</td><td>").concat(this._esc(doc.From), "</td></tr>\n        <tr><td>Current Status</td><td>").concat(this._statusBadge(doc.Status), "</td></tr>\n        <tr><td>Date Logged</td><td>").concat(this._formatDate(createdDate), "</td></tr>\n      </table>\n    ");
        // Timeline
        var timelineBody = this._el('dtl-timeline-body');
        if (logs.length === 0) {
            timelineBody.innerHTML = "\n        <div class=\"dtl-empty-tl\">\n          <div class=\"icon\">&#8986;</div>\n          <div>No status transitions recorded yet.</div>\n          <div style=\"margin-top:4px;font-size:12px;\">The document is currently at <strong>".concat(this._esc(doc.Status), "</strong>.</div>\n        </div>\n      ");
            this._el('dtl-total-bar').innerHTML = '';
            this._showScreen(2);
            return;
        }
        var totalDuration = 0;
        var html = '<ul class="dtl-timeline">';
        // Origin node: the initial status (FromStatus of the first log entry)
        var firstLog = logs[0];
        var originDate = new Date(doc.Created);
        html += "\n      <li class=\"dtl-tl-node origin\">\n        <div class=\"dtl-tl-dot\"></div>\n        <div class=\"dtl-tl-content\">\n          <div class=\"dtl-tl-transition\">".concat(this._statusBadge(firstLog.FromStatus), "</div>\n          <div class=\"dtl-tl-meta\">\n            <span>Initial status</span>\n            <span>").concat(this._formatDateTime(originDate), "</span>\n          </div>\n        </div>\n      </li>\n    ");
        // Transition nodes
        for (var i = 0; i < logs.length; i++) {
            var log = logs[i];
            var isLast = i === logs.length - 1;
            var nodeClass = isLast ? 'current' : '';
            var logDate = new Date(log.Created);
            var durationMinutes = log.DurationHours || 0;
            totalDuration += durationMinutes;
            html += "\n        <li class=\"dtl-tl-node ".concat(nodeClass, "\">\n          <div class=\"dtl-tl-dot\"></div>\n          <div class=\"dtl-tl-content\">\n            <div class=\"dtl-tl-transition\">\n              ").concat(this._statusBadge(log.FromStatus), "\n              <span class=\"dtl-tl-arrow\">&rarr;</span>\n              ").concat(this._statusBadge(log.ToStatus), "\n            </div>\n            <div class=\"dtl-tl-meta\">\n              <span class=\"dtl-tl-duration\">").concat(this._formatDuration(durationMinutes), "</span>\n              <span>by ").concat(this._esc(log.ChangedBy || 'System'), "</span>\n              <span>").concat(this._formatDateTime(logDate), "</span>\n            </div>\n          </div>\n        </li>\n      ");
        }
        html += '</ul>';
        timelineBody.innerHTML = html;
        // Total bar
        this._el('dtl-total-bar').innerHTML = "\n      <div class=\"dtl-total-bar\">\n        <span class=\"dtl-total-label\">Total Time Across Transitions</span>\n        <span class=\"dtl-total-value\">".concat(this._formatDuration(totalDuration), "</span>\n      </div>\n    ");
        this._showScreen(2);
    };
    // ── Helpers ─────────────────────────────────────────────────────────────
    DocumentTimelineWebPart.prototype._el = function (id) {
        return this.domElement.querySelector("#".concat(id));
    };
    DocumentTimelineWebPart.prototype._showScreen = function (n) {
        this.domElement.querySelectorAll('.dtl-screen').forEach(function (s) { return s.classList.remove('active'); });
        this._el("dtl-s".concat(n)).classList.add('active');
    };
    DocumentTimelineWebPart.prototype._esc = function (text) {
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    };
    DocumentTimelineWebPart.prototype._statusBadge = function (status) {
        var cls = '';
        switch (status) {
            case 'Received':
                cls = 'st-received';
                break;
            case 'In Progress':
                cls = 'st-review';
                break;
            case 'For DCA Approval and Signature':
                cls = 'st-dca';
                break;
            case 'Released':
                cls = 'st-released';
                break;
            case 'Filed':
                cls = 'st-filed';
                break;
        }
        return "<span class=\"dtl-status-badge ".concat(cls, "\">").concat(this._esc(status), "</span>");
    };
    DocumentTimelineWebPart.prototype._formatDate = function (d) {
        return d.toLocaleDateString('en-PH', { year: 'numeric', month: 'long', day: 'numeric' });
    };
    DocumentTimelineWebPart.prototype._formatDateTime = function (d) {
        return d.toLocaleDateString('en-PH', { year: 'numeric', month: 'short', day: 'numeric' }) +
            ' ' + d.toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit' });
    };
    DocumentTimelineWebPart.prototype._formatDuration = function (decimalMinutes) {
        var totalSeconds = Math.round(decimalMinutes * 60);
        var h = Math.floor(totalSeconds / 3600);
        var m = Math.floor((totalSeconds % 3600) / 60);
        var s = totalSeconds % 60;
        if (h > 0) {
            return s > 0 ? "".concat(h, "h ").concat(m, "m ").concat(s, "s") : m > 0 ? "".concat(h, "h ").concat(m, "m") : "".concat(h, "h");
        }
        if (m > 0) {
            return s > 0 ? "".concat(m, "m ").concat(s, "s") : "".concat(m, "m");
        }
        return "".concat(s, "s");
    };
    Object.defineProperty(DocumentTimelineWebPart.prototype, "dataVersion", {
        // ── SPFx lifecycle ──────────────────────────────────────────────────────
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    DocumentTimelineWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: 'Document Timeline Settings' },
                    groups: [
                        {
                            groupName: 'Configuration',
                            groupFields: [
                                PropertyPaneTextField('siteUrl', {
                                    label: 'SharePoint Site URL',
                                    description: 'Leave blank to use the current site',
                                    placeholder: 'https://tenant.sharepoint.com/sites/your-site'
                                }),
                                PropertyPaneTextField('listName', {
                                    label: 'Document Log Tracking List Name',
                                    placeholder: 'Document Log Tracking'
                                }),
                                PropertyPaneTextField('statusLogListName', {
                                    label: 'Status Log List Name',
                                    placeholder: 'Document Status Log'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return DocumentTimelineWebPart;
}(BaseClientSideWebPart));
export default DocumentTimelineWebPart;
//# sourceMappingURL=DocumentTimelineWebPart.js.map