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
var DocumentLogWebPart = /** @class */ (function (_super) {
    __extends(DocumentLogWebPart, _super);
    function DocumentLogWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.pollInterval = null;
        // ── Timers ──────────────────────────────────────────────────────────────
        _this.timers = {};
        _this.tick = null;
        return _this;
    }
    Object.defineProperty(DocumentLogWebPart.prototype, "siteUrl", {
        get: function () { return this.properties.siteUrl || this.context.pageContext.web.absoluteUrl; },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(DocumentLogWebPart.prototype, "listName", {
        get: function () { return this.properties.listName || 'Document Log Tracking'; },
        enumerable: false,
        configurable: true
    });
    DocumentLogWebPart.prototype.dispose = function () {
        if (this.tick)
            clearInterval(this.tick);
        if (this.pollInterval)
            clearInterval(this.pollInterval);
        this.tick = null;
        this.pollInterval = null;
        _super.prototype.dispose.call(this);
    };
    DocumentLogWebPart.prototype.fmt = function (ms) {
        if (!ms && ms !== 0)
            return '—';
        var s = Math.round(ms / 1000);
        return s < 60 ? "".concat(s, "s") : "".concat(Math.floor(s / 60), "m ").concat(s % 60, "s");
    };
    DocumentLogWebPart.prototype.startTick = function (elId, t0) {
        var _this = this;
        if (this.tick)
            clearInterval(this.tick);
        this.tick = setInterval(function () {
            var el = _this.domElement.querySelector("#".concat(elId));
            if (el)
                el.textContent = _this.fmt(Date.now() - t0);
        }, 1000);
    };
    DocumentLogWebPart.prototype.stopTick = function () {
        if (this.tick)
            clearInterval(this.tick);
    };
    // ── Render ──────────────────────────────────────────────────────────────
    DocumentLogWebPart.prototype.render = function () {
        var user = this.context.pageContext.user;
        this.domElement.innerHTML = "\n      <style>\n        .dl-wrap { font-family: 'Segoe UI', sans-serif; max-width: 700px; margin: 0 auto; padding: 24px 16px; color: #1c1a16; }\n        .dl-header { text-align: center; margin-bottom: 28px; padding-bottom: 20px; border-bottom: 2px solid #1c1a16; }\n        .dl-header h2 { font-size: 22px; font-weight: 700; margin: 0 0 6px; }\n        .dl-header p { font-size: 13px; color: #7a7368; margin: 0; }\n\n        .dl-progress { display: flex; align-items: center; margin-bottom: 20px; }\n        .dl-ps { display: flex; align-items: center; gap: 6px; flex: 1; }\n        .dl-ps:not(:last-child)::after { content: ''; flex: 1; height: 1px; background: #d8d2c8; margin: 0 6px; }\n        .dl-dot { width: 26px; height: 26px; border-radius: 50%; border: 2px solid #d8d2c8; display: flex; align-items: center; justify-content: center; font-size: 11px; color: #7a7368; flex-shrink: 0; background: #fff; transition: all 0.2s; }\n        .dl-ps-label { font-size: 11px; color: #7a7368; white-space: nowrap; }\n        .dl-ps.active .dl-dot { border-color: #8b4513; background: #8b4513; color: #fff; }\n        .dl-ps.active .dl-ps-label { color: #8b4513; font-weight: 600; }\n        .dl-ps.done .dl-dot { border-color: #2d6a4f; background: #2d6a4f; color: #fff; }\n        .dl-ps.done .dl-ps-label { color: #2d6a4f; }\n\n        .dl-timers { display: flex; gap: 8px; margin-bottom: 20px; flex-wrap: wrap; }\n        .dl-tc { font-family: 'Courier New', monospace; font-size: 10px; padding: 4px 10px; background: #fff; border: 1px solid #d8d2c8; color: #7a7368; display: flex; align-items: center; gap: 5px; }\n        .dl-tc .dot { width: 5px; height: 5px; border-radius: 50%; background: #d8d2c8; }\n        .dl-tc.running .dot { background: #c8773a; animation: dlblink 1s infinite; }\n        .dl-tc.done .dot { background: #2d6a4f; }\n        @keyframes dlblink { 0%,100%{opacity:1} 50%{opacity:0.2} }\n\n        .dl-card { background: #fff; border: 1px solid #d8d2c8; padding: 28px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }\n        .dl-card-title { font-size: 15px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid #d8d2c8; display: flex; align-items: center; gap: 8px; }\n        .dl-card-title .bar { width: 3px; height: 14px; background: #c9a84c; flex-shrink: 0; }\n\n        .dl-field { margin-bottom: 18px; }\n        .dl-field:last-of-type { margin-bottom: 0; }\n        .dl-label { display: block; font-size: 11px; font-weight: 600; color: #7a7368; margin-bottom: 5px; letter-spacing: 0.07em; text-transform: uppercase; }\n        .dl-label .req { color: #8b4513; }\n        .dl-input, .dl-select, .dl-textarea { width: 100%; padding: 9px 12px; border: 1px solid #d8d2c8; background: #f7f5f0; font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #1c1a16; outline: none; box-sizing: border-box; transition: border-color 0.15s; }\n        .dl-input:focus, .dl-select:focus, .dl-textarea:focus { border-color: #8b4513; background: #fff; }\n        .dl-input.err, .dl-select.err, .dl-textarea.err { border-color: #c0392b; background: #fdf5f5; }\n        .dl-textarea { resize: vertical; min-height: 80px; line-height: 1.5; }\n        .dl-row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }\n        .dl-val { font-size: 11px; color: #c0392b; margin-top: 4px; display: none; }\n        .dl-val.show { display: block; }\n\n        .dl-notice { padding: 11px 14px; font-size: 13px; display: flex; gap: 8px; align-items: flex-start; margin-bottom: 18px; line-height: 1.5; }\n        .dl-notice.info { background: #f5ece4; border-left: 3px solid #c8773a; color: #8b4513; }\n        .dl-notice.success { background: #eaf3ed; border-left: 3px solid #2d6a4f; color: #2d6a4f; }\n        .dl-notice.error { background: #fdf5f5; border-left: 3px solid #c0392b; color: #c0392b; }\n\n        .dl-btn { padding: 10px 22px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.15s; letter-spacing: 0.03em; }\n        .dl-btn-primary { background: #8b4513; color: #fff; }\n        .dl-btn-primary:hover { background: #6d3410; }\n        .dl-btn-primary:disabled { background: #d8d2c8; color: #7a7368; cursor: not-allowed; }\n        .dl-btn-outline { background: transparent; color: #1c1a16; border: 1px solid #b0a898; }\n        .dl-btn-outline:hover { border-color: #1c1a16; }\n        .dl-btn-success { background: #2d6a4f; color: #fff; }\n        .dl-btn-success:hover { background: #1e4d38; }\n        .dl-btn-row { display: flex; justify-content: space-between; align-items: center; margin-top: 24px; gap: 10px; flex-wrap: wrap; }\n\n        .dl-rtable { width: 100%; border-collapse: collapse; }\n        .dl-rtable tr { border-bottom: 1px solid #d8d2c8; }\n        .dl-rtable tr:last-child { border-bottom: none; }\n        .dl-rtable td { padding: 9px 0; font-size: 13px; vertical-align: top; }\n        .dl-rtable td:first-child { color: #7a7368; font-size: 11px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; width: 36%; padding-right: 14px; padding-top: 11px; }\n\n        .dl-code-wrap { text-align: center; margin: 18px 0; }\n        .dl-code-eye { font-family: 'Courier New', monospace; font-size: 10px; letter-spacing: 0.16em; text-transform: uppercase; color: #7a7368; margin-bottom: 8px; }\n        .dl-code { display: inline-block; background: #1c1a16; color: #f7f3e8; font-family: 'Courier New', monospace; font-size: 30px; font-weight: 700; letter-spacing: 0.2em; padding: 18px 32px; border: 3px solid #c9a84c; }\n\n        .dl-time-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 18px; }\n        .dl-tbox { border: 1px solid #d8d2c8; padding: 12px; text-align: center; background: #f7f5f0; }\n        .dl-tbox .tbl { font-family: 'Courier New', monospace; font-size: 9px; letter-spacing: 0.1em; text-transform: uppercase; color: #7a7368; margin-bottom: 5px; }\n        .dl-tbox .tbv { font-family: 'Courier New', monospace; font-size: 17px; font-weight: 700; color: #2d6a4f; }\n\n        .dl-spinner { width: 16px; height: 16px; border: 2px solid #d8d2c8; border-top-color: #8b4513; border-radius: 50%; animation: dlspin 0.75s linear infinite; flex-shrink: 0; }\n        @keyframes dlspin { to { transform: rotate(360deg); } }\n        .dl-status-row { display: flex; align-items: center; gap: 12px; font-family: 'Courier New', monospace; font-size: 12px; color: #7a7368; padding: 10px 0; }\n\n        .dl-screen { display: none; }\n        .dl-screen.active { display: block; }\n\n        .dl-attach-area { border: 1px dashed #b0a898; padding: 14px; text-align: center; background: #f7f5f0; cursor: pointer; }\n        .dl-attach-area:hover { border-color: #8b4513; }\n        .dl-attach-label { font-size: 13px; color: #7a7368; }\n        .dl-attach-label span { color: #8b4513; font-weight: 600; text-decoration: underline; }\n        .dl-attach-list { margin-top: 8px; }\n        .dl-attach-item { font-size: 12px; font-family: 'Courier New', monospace; padding: 4px 0; border-bottom: 1px solid #d8d2c8; display: flex; justify-content: space-between; align-items: center; }\n        .dl-attach-item button { background: none; border: none; cursor: pointer; color: #7a7368; font-size: 11px; }\n        .dl-attach-item button:hover { color: #c0392b; }\n      </style>\n\n      <div class=\"dl-wrap\">\n        <div class=\"dl-header\">\n          <h2>Document Log</h2>\n          <p>Logged in as <strong>".concat(user.displayName, "</strong> &nbsp;\u00B7&nbsp; ").concat(user.email, "</p>\n        </div>\n\n        <div class=\"dl-progress\" id=\"dl-progress\">\n          <div class=\"dl-ps active\" id=\"dl-ps-1\"><div class=\"dl-dot\">1</div><span class=\"dl-ps-label\">Details</span></div>\n          <div class=\"dl-ps\" id=\"dl-ps-2\"><div class=\"dl-dot\">2</div><span class=\"dl-ps-label\">Review</span></div>\n          <div class=\"dl-ps\" id=\"dl-ps-3\"><div class=\"dl-dot\">3</div><span class=\"dl-ps-label\">Log Code</span></div>\n        </div>\n\n        <div class=\"dl-timers\">\n          <div class=\"dl-tc running\" id=\"dl-tc1\"><span class=\"dot\"></span>Filling \u2014 <span id=\"dl-t1\">0s</span></div>\n          <div class=\"dl-tc\" id=\"dl-tc2\"><span class=\"dot\"></span>Review \u2014 <span id=\"dl-t2\">\u2014</span></div>\n          <div class=\"dl-tc\" id=\"dl-tc3\"><span class=\"dot\"></span>Processing \u2014 <span id=\"dl-t3\">\u2014</span></div>\n        </div>\n\n        <!-- Screen 1: Form -->\n        <div class=\"dl-screen active\" id=\"dl-s1\">\n          <div class=\"dl-card\">\n            <div class=\"dl-card-title\"><span class=\"bar\"></span>Document Information</div>\n            <div class=\"dl-field\">\n              <label class=\"dl-label\">Document Title <span class=\"req\">*</span></label>\n              <input id=\"dl-title\" type=\"text\" class=\"dl-input\" placeholder=\"e.g. Re: Administrative Matter No. 25-01-001\" />\n              <div class=\"dl-val\" id=\"dl-v-title\">Document title is required.</div>\n            </div>\n            <div class=\"dl-row2\">\n              <div class=\"dl-field\">\n                <label class=\"dl-label\">Document Type <span class=\"req\">*</span></label>\n                <select id=\"dl-type\" class=\"dl-select\">\n                  <option value=\"\">\u2014 Select \u2014</option>\n                  <option value=\"RCM\">RCM</option>\n                  <option value=\"OTHERS\">OTHERS</option>\n                </select>\n                <div class=\"dl-val\" id=\"dl-v-type\">Please select a document type.</div>\n              </div>\n              <div class=\"dl-field\">\n                <label class=\"dl-label\">Document Format <span class=\"req\">*</span></label>\n                <select id=\"dl-format\" class=\"dl-select\">\n                  <option value=\"\">\u2014 Select \u2014</option>\n                  <option value=\"HC\">Physical (HC)</option>\n                  <option value=\"SC\">Digital (SC)</option>\n                </select>\n                <div class=\"dl-val\" id=\"dl-v-format\">Please select a document format.</div>\n              </div>\n            </div>\n            <div class=\"dl-row2\">\n              <div class=\"dl-field\">\n                <label class=\"dl-label\">From <span class=\"req\">*</span></label>\n                <input id=\"dl-from\" type=\"text\" class=\"dl-input\" placeholder=\"Sender name or office\" />\n                <div class=\"dl-val\" id=\"dl-v-from\">Please indicate the sender.</div>\n              </div>\n              <div class=\"dl-field\">\n                <label class=\"dl-label\">Status <span class=\"req\">*</span></label>\n                <select id=\"dl-status\" class=\"dl-select\">\n                  <option value=\"\">\u2014 Select \u2014</option>\n                  <option>Received</option>\n                  <option>In Progress</option>\n                  <option>For DCA Approval and Signature</option>\n                  <option>Released</option>\n                  <option>Filed</option>\n                </select>\n                <div class=\"dl-val\" id=\"dl-v-status\">Please select a status.</div>\n              </div>\n            </div>\n            <div class=\"dl-field\">\n              <label class=\"dl-label\">Other Remarks</label>\n              <textarea id=\"dl-remarks\" class=\"dl-textarea\" placeholder=\"Optional notes...\"></textarea>\n            </div>\n            <div class=\"dl-field\">\n              <label class=\"dl-label\">Attachments <span style=\"color:#7a7368;font-weight:400;text-transform:none;letter-spacing:0;\">(optional)</span></label>\n              <div class=\"dl-attach-area\" id=\"dl-attach-area\">\n                <div class=\"dl-attach-label\">Drop files here or <span>browse</span></div>\n              </div>\n              <div class=\"dl-attach-list\" id=\"dl-attach-list\"></div>\n            </div>\n            <div class=\"dl-btn-row\">\n              <span style=\"font-size:11px;color:#7a7368;font-family:'Courier New',monospace;\">Fields marked <span style=\"color:#8b4513\">*</span> are required</span>\n              <button id=\"dl-btn-review\" class=\"dl-btn dl-btn-primary\">Review Entry \u2192</button>\n            </div>\n          </div>\n        </div>\n\n        <!-- Screen 2: Review -->\n        <div class=\"dl-screen\" id=\"dl-s2\">\n          <div class=\"dl-card\">\n            <div class=\"dl-card-title\"><span class=\"bar\"></span>Review Before Submitting</div>\n            <div class=\"dl-notice info\">\u2139\uFE0F &nbsp;Please verify all details before submitting.</div>\n            <table class=\"dl-rtable\" id=\"dl-review-table\"></table>\n            <div class=\"dl-btn-row\">\n              <button id=\"dl-btn-back\" class=\"dl-btn dl-btn-outline\">\u2190 Edit Details</button>\n              <button id=\"dl-btn-submit\" class=\"dl-btn dl-btn-primary\">Submit & Generate Code \u2192</button>\n            </div>\n          </div>\n        </div>\n\n        <!-- Screen 3: Result -->\n        <div class=\"dl-screen\" id=\"dl-s3\">\n          <div class=\"dl-card\" id=\"dl-loading\">\n            <div class=\"dl-card-title\"><span class=\"bar\"></span>Processing</div>\n            <div class=\"dl-notice info\">Submitting to SharePoint and waiting for the log code...</div>\n            <div class=\"dl-status-row\">\n              <div class=\"dl-spinner\"></div>\n              <span id=\"dl-status-msg\">Writing entry to document list...</span>\n            </div>\n          </div>\n          <div class=\"dl-card\" id=\"dl-success\" style=\"display:none;\">\n            <div class=\"dl-notice success\">\u2713 &nbsp;Document logged successfully. Reference code generated.</div>\n            <div class=\"dl-code-wrap\">\n              <div class=\"dl-code-eye\">Document Reference Code</div>\n              <div class=\"dl-code\" id=\"dl-code\">\u2014</div>\n            </div>\n            <div style=\"margin-top:24px;\">\n              <div class=\"dl-card-title\"><span class=\"bar\"></span>Log Summary</div>\n              <table class=\"dl-rtable\" id=\"dl-final-table\"></table>\n            </div>\n            <div class=\"dl-time-grid\">\n              <div class=\"dl-tbox\"><div class=\"tbl\">Filling</div><div class=\"tbv\" id=\"dl-ts1\">\u2014</div></div>\n              <div class=\"dl-tbox\"><div class=\"tbl\">Review</div><div class=\"tbv\" id=\"dl-ts2\">\u2014</div></div>\n              <div class=\"dl-tbox\"><div class=\"tbl\">Processing</div><div class=\"tbv\" id=\"dl-ts3\">\u2014</div></div>\n            </div>\n            <div class=\"dl-btn-row\" style=\"margin-top:20px;\">\n              <button id=\"dl-btn-new\" class=\"dl-btn dl-btn-outline\">+ New Entry</button>\n              <div style=\"display:flex;gap:10px;\">\n                <button id=\"dl-btn-copy\" class=\"dl-btn dl-btn-outline\">Copy Code</button>\n                <button id=\"dl-btn-print\" class=\"dl-btn dl-btn-success\">\uD83D\uDDA8 Print Routing Slip</button>\n              </div>\n            </div>\n          </div>\n          <div class=\"dl-card\" id=\"dl-error\" style=\"display:none;\">\n            <div class=\"dl-notice error\">\u2715 &nbsp;Something went wrong. The entry was saved but the code could not be retrieved. Please check your SharePoint list directly.</div>\n            <div class=\"dl-btn-row\">\n              <button id=\"dl-btn-retry-new\" class=\"dl-btn dl-btn-outline\">Start New Entry</button>\n            </div>\n          </div>\n        </div>\n\n      </div>\n    ");
        this._bindEvents();
        this.timers = { s1Start: Date.now() };
        this.startTick('dl-t1', this.timers.s1Start);
    };
    // ── Bind all events ──────────────────────────────────────────────────────
    DocumentLogWebPart.prototype._bindEvents = function () {
        var _this = this;
        var _a;
        var attachments = [];
        var attachArea = this.domElement.querySelector('#dl-attach-area');
        var attachList = this.domElement.querySelector('#dl-attach-list');
        var fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.multiple = true;
        fileInput.style.display = 'none';
        this.domElement.appendChild(fileInput);
        attachArea.addEventListener('click', function (e) {
            e.preventDefault();
            e.stopPropagation();
            fileInput.click();
        });
        fileInput.addEventListener('change', function () {
            Array.from(fileInput.files || []).forEach(function (f) {
                if (!attachments.find(function (a) { return a.name === f.name; }))
                    attachments.push(f);
            });
            fileInput.value = '';
            renderAttachments();
        });
        attachArea.addEventListener('dragover', function (e) {
            e.preventDefault();
            attachArea.style.borderColor = '#8b4513';
        });
        attachArea.addEventListener('dragleave', function () {
            attachArea.style.borderColor = '#b0a898';
        });
        attachArea.addEventListener('drop', function (e) {
            e.preventDefault();
            attachArea.style.borderColor = '#b0a898';
            var files = e.dataTransfer ? e.dataTransfer.files : null;
            if (files) {
                Array.from(files).forEach(function (f) {
                    if (!attachments.find(function (a) { return a.name === f.name; }))
                        attachments.push(f);
                });
                renderAttachments();
            }
        });
        var renderAttachments = function () {
            if (attachments.length === 0) {
                attachArea.innerHTML = '<div class="dl-attach-label">Drop files here or <span>browse</span></div>';
                attachList.innerHTML = '';
                return;
            }
            attachArea.innerHTML = "<div class=\"dl-attach-label\"><span>".concat(attachments.length, " file").concat(attachments.length > 1 ? 's' : '', " attached</span> \u2014 click to add more</div>");
            attachList.innerHTML = attachments.map(function (f, i) {
                return "<div class=\"dl-attach-item\">\n          <span>".concat(i + 1, ". ").concat(f.name, "</span>\n          <button data-name=\"").concat(f.name, "\">\u2715 remove</button>\n        </div>");
            }).join('');
            attachList.querySelectorAll('button').forEach(function (btn) {
                btn.addEventListener('click', function (e) {
                    e.stopPropagation();
                    attachments = attachments.filter(function (a) { return a.name !== btn.getAttribute('data-name'); });
                    renderAttachments();
                });
            });
        };
        this.domElement.querySelector('#dl-btn-review').addEventListener('click', function () {
            if (!_this._validate())
                return;
            var fd = _this._getFormData();
            _this.timers.s1End = Date.now();
            _this.stopTick();
            _this._el('dl-t1').textContent = _this.fmt(_this.timers.s1End - _this.timers.s1Start);
            _this._tc('dl-tc1', 'done');
            _this.timers.s2Start = Date.now();
            _this.startTick('dl-t2', _this.timers.s2Start);
            _this._tc('dl-tc2', 'running');
            var rows = [
                ['Logged By', _this.context.pageContext.user.displayName],
                ['Document Title', fd.title],
                ['Document Type', fd.type],
                ['Document Format', fd.format === 'HC' ? 'Physical (HC)' : 'Digital (SC)'],
                ['From', fd.from],
                ['Status', fd.status],
                ['Remarks', fd.remarks || '—'],
                ['Attachments', attachments.length ? attachments.map(function (a) { return a.name; }).join(', ') : 'None'],
            ];
            _this._el('dl-review-table').innerHTML =
                rows.map(function (_a) {
                    var k = _a[0], v = _a[1];
                    return "<tr><td>".concat(k, "</td><td>").concat(v, "</td></tr>");
                }).join('');
            _this._setProgress(2);
            _this._showScreen(2);
        });
        this.domElement.querySelector('#dl-btn-back').addEventListener('click', function () {
            _this.timers.s1End = undefined;
            _this.startTick('dl-t1', _this.timers.s1Start);
            _this._tc('dl-tc1', 'running');
            _this._tc('dl-tc2', '');
            _this._el('dl-t2').textContent = '—';
            _this.timers.s2Start = undefined;
            _this._setProgress(1);
            _this._showScreen(1);
        });
        this.domElement.querySelector('#dl-btn-submit').addEventListener('click', function () {
            var fd = _this._getFormData();
            _this.timers.s2End = Date.now();
            _this.stopTick();
            _this._el('dl-t2').textContent = _this.fmt(_this.timers.s2End - _this.timers.s2Start);
            _this._tc('dl-tc2', 'done');
            _this.timers.s3Start = Date.now();
            _this.startTick('dl-t3', _this.timers.s3Start);
            _this._tc('dl-tc3', 'running');
            _this._setProgress(3);
            _this._showScreen(3);
            _this._submitToSharePoint(fd, attachments);
        });
        this.domElement.querySelector('#dl-btn-new').addEventListener('click', function () { return _this._reset(); });
        (_a = this.domElement.querySelector('#dl-btn-retry-new')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', function () { return _this._reset(); });
        this.domElement.querySelector('#dl-btn-copy').addEventListener('click', function () {
            var code = _this._el('dl-code').textContent || '';
            navigator.clipboard.writeText(code)
                .then(function () { return alert('Copied: ' + code); })
                .catch(function () {
                window.prompt('Could not copy automatically. Copy the code below:', code);
            });
        });
        this.domElement.querySelector('#dl-btn-print').addEventListener('click', function () {
            var code = _this._el('dl-code').textContent || '';
            var fd = _this._getFormData();
            _this._printSlip(code, fd);
        });
    };
    // ── Print routing slip ───────────────────────────────────────────────────
    DocumentLogWebPart.prototype._printSlip = function (code, fd) {
        var date = new Date().toLocaleDateString('en-PH', { year: 'numeric', month: 'long', day: 'numeric' });
        var loggedBy = this.context.pageContext.user.displayName;
        var html = "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<title>Routing Slip - ".concat(code, "</title>\n<style>\n  * { box-sizing: border-box; margin: 0; padding: 0; }\n  body { font-family: Arial, sans-serif; font-size: 11px; color: #000; padding: 20px 28px; max-width: 800px; margin: 0 auto; }\n  .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px; }\n  .header-left h1 { font-size: 22px; font-weight: 700; }\n  .header-left .sub { font-size: 10px; color: #666; margin-top: 2px; }\n  .header-right { text-align: right; }\n  .header-right .label { font-size: 10px; color: #555; }\n  .ref { font-family: 'Courier New', monospace; font-size: 20px; font-weight: 700; letter-spacing: 0.1em; border: 2px solid #c9a84c; background: #1c1a16; color: #f7f3e8; padding: 6px 14px; display: inline-block; margin-top: 3px; }\n  hr { border: none; border-top: 2px solid #000; margin: 8px 0; }\n  .meta { width: 100%; border-collapse: collapse; margin-bottom: 10px; }\n  .meta td { padding: 3px 6px 3px 0; vertical-align: top; }\n  .meta td:first-child { font-weight: 700; white-space: nowrap; width: 130px; }\n  table.wf { width: 100%; border-collapse: collapse; font-size: 10px; }\n  table.wf th { background: #d9d9d9; border: 1px solid #000; padding: 5px 6px; text-align: left; font-size: 10px; }\n  table.wf td { border: 1px solid #000; padding: 5px 6px; vertical-align: middle; }\n  table.wf td.chk { text-align: center; width: 22px; font-size: 13px; }\n  table.wf tr.sec td { background: #d9d9d9; font-weight: 700; }\n  .remarks-box { border: 1px solid #000; padding: 8px; min-height: 48px; margin-top: 4px; }\n  .footer { margin-top: 10px; font-size: 10px; color: #555; display: flex; justify-content: space-between; }\n  @media print {\n    body { padding: 10px 16px; }\n    @page { margin: 1cm; size: A4; }\n  }\n</style>\n</head>\n<body>\n<div class=\"header\">\n  <div class=\"header-left\">\n    <h1>ROUTING SLIP</h1>\n    <div class=\"sub\">V1 - Feb 2026 &nbsp;\u00B7&nbsp; Supreme Court of the Philippines</div>\n  </div>\n  <div class=\"header-right\">\n    <div class=\"label\">Document Control No.</div>\n    <div class=\"ref\">").concat(code, "</div>\n  </div>\n</div>\n<hr/>\n<table class=\"meta\">\n  <tr><td>Document Title:</td><td>").concat(fd.title, "</td></tr>\n  <tr>\n    <td>Document Type:</td>\n    <td>").concat(fd.type, " &nbsp;&nbsp; <strong>Format:</strong> &nbsp;").concat(fd.format === 'HC' ? 'Physical (HC)' : 'Digital (SC)', "</td>\n  </tr>\n  <tr><td>From:</td><td>").concat(fd.from, "</td></tr>\n  <tr><td>Status:</td><td>").concat(fd.status, "</td></tr>\n  <tr><td>Date Received:</td><td>").concat(date, "</td></tr>\n  ").concat(fd.remarks ? "<tr><td>Remarks:</td><td>".concat(fd.remarks, "</td></tr>") : '', "\n</table>\n\n<table class=\"wf\">\n  <tr>\n    <th style=\"width:22px;\"></th>\n    <th style=\"width:28%;\">Status</th>\n    <th style=\"width:28%;\">Handled By</th>\n    <th style=\"width:18%;\">Date</th>\n    <th>Remarks</th>\n  </tr>\n  <tr><td class=\"chk\">\u2610</td><td>Received</td><td>Ira | Beau</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Tracking - Create Slip</td><td>Nante | Dyane</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Sorting</td><td>Roda | Jim</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>For Internal Review</td><td>Atty. Unis | Atty. Kerr | Atty. Jen</td><td></td><td></td></tr>\n  <tr class=\"sec\"><td colspan=\"5\">For Memos</td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Drafting Memo</td><td>Atty. Amy | Atty. Unis | Atty. Kerr | Atty. Jen</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>For DCA Review</td><td>DCA</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>For Revisions</td><td>Atty. Amy | Atty. Unis | Atty. Kerr | Atty. Jen</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>For Printing</td><td>Roda</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Tracking - For DCA Approval and Signature</td><td>DCA</td><td></td><td></td></tr>\n  <tr class=\"sec\"><td colspan=\"5\">Approval and Release</td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>For DCA Signature</td><td>DCA</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Tracking - Released</td><td>Nante | Dyane</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Sorting</td><td>Roda</td><td></td><td></td></tr>\n  <tr class=\"sec\"><td colspan=\"5\">For Physical/Hardcopy</td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Log Outgoing</td><td>Roda</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>For Dispatch</td><td>Irish | Jim</td><td></td><td></td></tr>\n  <tr class=\"sec\"><td colspan=\"5\">For Email</td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Scanning Document</td><td>Beau | Dyane</td><td></td><td></td></tr>\n  <tr><td class=\"chk\">\u2610</td><td>Tracking - Filed</td><td>Jim</td><td></td><td></td></tr>\n</table>\n\n<div style=\"margin-top:8px;\">\n  <strong>Other Remarks:</strong>\n  <div class=\"remarks-box\">").concat(fd.remarks || '', "</div>\n</div>\n\n<div class=\"footer\">\n  <span>Logged by: <strong>").concat(loggedBy, "</strong></span>\n  <span>").concat(date, "</span>\n</div>\n\n<script>window.onload = function(){ window.print(); }</script>\n</body>\n</html>");
        var win = window.open('', '_blank');
        if (win) {
            win.document.write(html);
            win.document.close();
        }
    };
    // ── SharePoint: create list item ─────────────────────────────────────────
    DocumentLogWebPart.prototype._submitToSharePoint = function (fd, attachments) {
        return __awaiter(this, void 0, void 0, function () {
            var body, response, errText, data, itemId, failedAttachments, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this._setStatus('Writing entry to SharePoint list...');
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 8, , 9]);
                        body = JSON.stringify({
                            Title: fd.title,
                            Document_x0020_Type: fd.type,
                            Document_x0020_Format: fd.format,
                            From: fd.from,
                            Status: fd.status,
                            OtherRemarks: fd.remarks,
                        });
                        return [4 /*yield*/, this.context.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: body
                            })];
                    case 2:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 4];
                        return [4 /*yield*/, response.text()];
                    case 3:
                        errText = _a.sent();
                        throw new Error("SharePoint ".concat(response.status, ": ").concat(errText));
                    case 4: return [4 /*yield*/, response.json()];
                    case 5:
                        data = _a.sent();
                        itemId = data.Id;
                        failedAttachments = [];
                        if (!(attachments.length > 0)) return [3 /*break*/, 7];
                        this._setStatus('Uploading attachments...');
                        return [4 /*yield*/, this._uploadAttachments(itemId, attachments)];
                    case 6:
                        failedAttachments = _a.sent();
                        _a.label = 7;
                    case 7:
                        if (failedAttachments.length > 0) {
                            this._showAttachmentWarning(failedAttachments);
                        }
                        this._setStatus('Waiting for reference code to be generated...');
                        this._pollForCode(itemId);
                        return [3 /*break*/, 9];
                    case 8:
                        err_1 = _a.sent();
                        console.error('DocumentLog submit error:', err_1);
                        this._showError('The entry could not be saved to SharePoint. Please try again or contact your administrator.');
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    // ── SharePoint: upload attachments ───────────────────────────────────────
    DocumentLogWebPart.prototype._uploadAttachments = function (itemId, files) {
        return __awaiter(this, void 0, void 0, function () {
            var failed, _i, files_1, file, buffer, encodedName, response, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        failed = [];
                        _i = 0, files_1 = files;
                        _a.label = 1;
                    case 1:
                        if (!(_i < files_1.length)) return [3 /*break*/, 7];
                        file = files_1[_i];
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 5, , 6]);
                        return [4 /*yield*/, file.arrayBuffer()];
                    case 3:
                        buffer = _a.sent();
                        encodedName = encodeURIComponent(file.name.replace(/'/g, "''"));
                        return [4 /*yield*/, this.context.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items(").concat(itemId, ")/AttachmentFiles/add(FileName='").concat(encodedName, "')"), SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-Type': 'application/octet-stream',
                                    'odata-version': ''
                                },
                                body: buffer
                            })];
                    case 4:
                        response = _a.sent();
                        if (!response.ok) {
                            console.error("Attachment upload failed for \"".concat(file.name, "\": ").concat(response.status));
                            failed.push(file.name);
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        err_2 = _a.sent();
                        console.error("Attachment upload error for \"".concat(file.name, "\":"), err_2);
                        failed.push(file.name);
                        return [3 /*break*/, 6];
                    case 6:
                        _i++;
                        return [3 /*break*/, 1];
                    case 7: return [2 /*return*/, failed];
                }
            });
        });
    };
    // ── Poll for code written back by Power Automate ─────────────────────────
    DocumentLogWebPart.prototype._pollForCode = function (itemId) {
        var _this = this;
        var attempts = 0;
        var consecutiveErrors = 0;
        var maxAttempts = 20;
        var maxConsecutiveErrors = 3;
        this.pollInterval = setInterval(function () { return __awaiter(_this, void 0, void 0, function () {
            var response, data, err_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        attempts++;
                        if (attempts > maxAttempts) {
                            clearInterval(this.pollInterval);
                            this._showError('The entry was saved but the reference code was not generated in time. Please check the SharePoint list directly.');
                            return [2 /*return*/];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this.context.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items(").concat(itemId, ")?$select=ReferenceCode"), SPHttpClient.configurations.v1)];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("SharePoint returned ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        consecutiveErrors = 0;
                        if (data.ReferenceCode && data.ReferenceCode.trim() !== '') {
                            clearInterval(this.pollInterval);
                            this._showSuccess(data.ReferenceCode);
                        }
                        else {
                            this._setStatus("Generating code... (".concat(attempts, "/").concat(maxAttempts, ")"));
                        }
                        return [3 /*break*/, 5];
                    case 4:
                        err_3 = _a.sent();
                        console.error('Poll error:', err_3);
                        consecutiveErrors++;
                        if (consecutiveErrors >= maxConsecutiveErrors) {
                            clearInterval(this.pollInterval);
                            this._showError('The entry was saved but the connection was lost while waiting for the reference code. Please check the SharePoint list directly.');
                        }
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); }, 2000);
    };
    // ── Show success state ───────────────────────────────────────────────────
    DocumentLogWebPart.prototype._showSuccess = function (code) {
        this.timers.s3End = Date.now();
        this.stopTick();
        this._el('dl-t3').textContent = this.fmt(this.timers.s3End - this.timers.s3Start);
        this._tc('dl-tc3', 'done');
        this._setProgress('done');
        this._el('dl-code').textContent = code;
        var fd = this._getFormData();
        var rows = [
            ['Reference Code', code],
            ['Logged By', this.context.pageContext.user.displayName],
            ['Document Title', fd.title],
            ['Document Type', fd.type],
            ['Document Format', fd.format === 'HC' ? 'Physical (HC)' : 'Digital (SC)'],
            ['From', fd.from],
            ['Status', fd.status],
            ['Remarks', fd.remarks || '—'],
            ['Date Logged', new Date().toLocaleDateString('en-PH') + ' ' + new Date().toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit' })],
        ];
        this._el('dl-final-table').innerHTML =
            rows.map(function (_a) {
                var k = _a[0], v = _a[1];
                return "<tr><td>".concat(k, "</td><td>").concat(v, "</td></tr>");
            }).join('');
        this._el('dl-ts1').textContent = this.fmt(this.timers.s1End - this.timers.s1Start);
        this._el('dl-ts2').textContent = this.fmt(this.timers.s2End - this.timers.s2Start);
        this._el('dl-ts3').textContent = this.fmt(this.timers.s3End - this.timers.s3Start);
        this._el('dl-loading').style.display = 'none';
        this._el('dl-success').style.display = 'block';
    };
    DocumentLogWebPart.prototype._showError = function (message) {
        this.stopTick();
        this._tc('dl-tc3', '');
        if (message) {
            var notice = this._el('dl-error').querySelector('.dl-notice');
            if (notice)
                notice.innerHTML = "\u2715 &nbsp;".concat(message);
        }
        this._el('dl-loading').style.display = 'none';
        this._el('dl-error').style.display = 'block';
    };
    DocumentLogWebPart.prototype._showAttachmentWarning = function (failedFiles) {
        var names = failedFiles.join(', ');
        var warning = document.createElement('div');
        warning.className = 'dl-notice info';
        warning.innerHTML = "\u26A0 &nbsp;The entry was saved, but these attachments failed to upload: <strong>".concat(names, "</strong>. You can add them manually from the SharePoint list.");
        var loading = this._el('dl-loading');
        loading.parentElement.insertBefore(warning, loading);
    };
    // ── Validation ───────────────────────────────────────────────────────────
    DocumentLogWebPart.prototype._validate = function () {
        var _this = this;
        var checks = [
            { id: 'dl-title', vid: 'dl-v-title' },
            { id: 'dl-type', vid: 'dl-v-type' },
            { id: 'dl-format', vid: 'dl-v-format' },
            { id: 'dl-from', vid: 'dl-v-from' },
            { id: 'dl-status', vid: 'dl-v-status' },
        ];
        var ok = true;
        checks.forEach(function (_a) {
            var id = _a.id, vid = _a.vid;
            var el = _this.domElement.querySelector("#".concat(id));
            var msg = _this.domElement.querySelector("#".concat(vid));
            if (!el.value.trim()) {
                el.classList.add('err');
                msg.classList.add('show');
                ok = false;
            }
            else {
                el.classList.remove('err');
                msg.classList.remove('show');
            }
        });
        return ok;
    };
    // ── Helpers ──────────────────────────────────────────────────────────────
    DocumentLogWebPart.prototype._getFormData = function () {
        return {
            title: this.domElement.querySelector('#dl-title').value.trim(),
            type: this.domElement.querySelector('#dl-type').value,
            format: this.domElement.querySelector('#dl-format').value,
            from: this.domElement.querySelector('#dl-from').value.trim(),
            status: this.domElement.querySelector('#dl-status').value,
            remarks: this.domElement.querySelector('#dl-remarks').value.trim(),
        };
    };
    DocumentLogWebPart.prototype._el = function (id) {
        return this.domElement.querySelector("#".concat(id));
    };
    DocumentLogWebPart.prototype._tc = function (id, state) {
        var el = this._el(id);
        el.className = "dl-tc".concat(state ? ' ' + state : '');
    };
    DocumentLogWebPart.prototype._setStatus = function (msg) {
        var el = this.domElement.querySelector('#dl-status-msg');
        if (el)
            el.textContent = msg;
    };
    DocumentLogWebPart.prototype._showScreen = function (n) {
        this.domElement.querySelectorAll('.dl-screen').forEach(function (s) { return s.classList.remove('active'); });
        this._el("dl-s".concat(n)).classList.add('active');
    };
    DocumentLogWebPart.prototype._setProgress = function (step) {
        var _this = this;
        if (step === 'done') {
            [1, 2, 3].forEach(function (i) { _this._el("dl-ps-".concat(i)).className = 'dl-ps done'; });
            return;
        }
        for (var i = 1; i <= 3; i++) {
            var el = this._el("dl-ps-".concat(i));
            if (i < step)
                el.className = 'dl-ps done';
            else if (i === step)
                el.className = 'dl-ps active';
            else
                el.className = 'dl-ps';
        }
    };
    DocumentLogWebPart.prototype._reset = function () {
        var _this = this;
        ['dl-title', 'dl-from', 'dl-remarks'].forEach(function (id) {
            return _this.domElement.querySelector("#".concat(id)).value = '';
        });
        ['dl-type', 'dl-format', 'dl-status'].forEach(function (id) {
            return _this.domElement.querySelector("#".concat(id)).value = '';
        });
        this._el('dl-attach-list').innerHTML = '';
        this.timers = { s1Start: Date.now() };
        this.startTick('dl-t1', this.timers.s1Start);
        this._tc('dl-tc1', 'running');
        this._tc('dl-tc2', '');
        this._tc('dl-tc3', '');
        this._el('dl-t1').textContent = '0s';
        this._el('dl-t2').textContent = '—';
        this._el('dl-t3').textContent = '—';
        this._el('dl-loading').style.display = 'block';
        this._el('dl-success').style.display = 'none';
        this._el('dl-error').style.display = 'none';
        this._setProgress(1);
        this._showScreen(1);
    };
    Object.defineProperty(DocumentLogWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    DocumentLogWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: 'Document Log Settings' },
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
                                    label: 'List Name',
                                    placeholder: 'Document Log Tracking'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return DocumentLogWebPart;
}(BaseClientSideWebPart));
export default DocumentLogWebPart;
//# sourceMappingURL=DocumentLogWebPart.js.map