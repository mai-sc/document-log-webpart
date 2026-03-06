define(["@microsoft/sp-core-library","@microsoft/sp-http","@microsoft/sp-webpart-base","@microsoft/sp-property-pane"], (__WEBPACK_EXTERNAL_MODULE__878__, __WEBPACK_EXTERNAL_MODULE__272__, __WEBPACK_EXTERNAL_MODULE__134__, __WEBPACK_EXTERNAL_MODULE__723__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 878:
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__878__;

/***/ }),

/***/ 272:
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__272__;

/***/ }),

/***/ 723:
/*!**********************************************!*\
  !*** external "@microsoft/sp-property-pane" ***!
  \**********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__723__;

/***/ }),

/***/ 134:
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__134__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!**************************************************************************!*\
  !*** ./lib/webparts/documentStatusUpdate/DocumentStatusUpdateWebPart.js ***!
  \**************************************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-core-library */ 878);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-http */ 272);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 134);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-property-pane */ 723);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__);
var __extends = (undefined && undefined.__extends) || (function () {
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
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
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




var DocumentStatusUpdateWebPart = /** @class */ (function (_super) {
    __extends(DocumentStatusUpdateWebPart, _super);
    function DocumentStatusUpdateWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.currentItem = null;
        return _this;
    }
    Object.defineProperty(DocumentStatusUpdateWebPart.prototype, "siteUrl", {
        get: function () { return this.properties.siteUrl || this.context.pageContext.web.absoluteUrl; },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(DocumentStatusUpdateWebPart.prototype, "listName", {
        get: function () { return this.properties.listName || 'Document Log Tracking'; },
        enumerable: false,
        configurable: true
    });
    // ── Render ──────────────────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype.render = function () {
        var user = this.context.pageContext.user;
        this.domElement.innerHTML = "\n      <style>\n        .dsu-wrap { font-family: 'Segoe UI', sans-serif; max-width: 700px; margin: 0 auto; padding: 24px 16px; color: #1c1a16; }\n        .dsu-header { text-align: center; margin-bottom: 28px; padding-bottom: 20px; border-bottom: 2px solid #1c1a16; }\n        .dsu-header h2 { font-size: 22px; font-weight: 700; margin: 0 0 6px; }\n        .dsu-header p { font-size: 13px; color: #7a7368; margin: 0; }\n\n        .dsu-card { background: #fff; border: 1px solid #d8d2c8; padding: 28px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }\n        .dsu-card-title { font-size: 15px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid #d8d2c8; display: flex; align-items: center; gap: 8px; }\n        .dsu-card-title .bar { width: 3px; height: 14px; background: #c9a84c; flex-shrink: 0; }\n\n        .dsu-field { margin-bottom: 18px; }\n        .dsu-field:last-of-type { margin-bottom: 0; }\n        .dsu-label { display: block; font-size: 11px; font-weight: 600; color: #7a7368; margin-bottom: 5px; letter-spacing: 0.07em; text-transform: uppercase; }\n        .dsu-label .req { color: #8b4513; }\n        .dsu-input, .dsu-select, .dsu-textarea { width: 100%; padding: 9px 12px; border: 1px solid #d8d2c8; background: #f7f5f0; font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #1c1a16; outline: none; box-sizing: border-box; transition: border-color 0.15s; }\n        .dsu-input:focus, .dsu-select:focus, .dsu-textarea:focus { border-color: #8b4513; background: #fff; }\n        .dsu-textarea { resize: vertical; min-height: 80px; line-height: 1.5; }\n\n        .dsu-notice { padding: 11px 14px; font-size: 13px; display: flex; gap: 8px; align-items: flex-start; margin-bottom: 18px; line-height: 1.5; }\n        .dsu-notice.info { background: #f5ece4; border-left: 3px solid #c8773a; color: #8b4513; }\n        .dsu-notice.success { background: #eaf3ed; border-left: 3px solid #2d6a4f; color: #2d6a4f; }\n        .dsu-notice.error { background: #fdf5f5; border-left: 3px solid #c0392b; color: #c0392b; }\n\n        .dsu-btn { padding: 10px 22px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.15s; letter-spacing: 0.03em; }\n        .dsu-btn-primary { background: #8b4513; color: #fff; }\n        .dsu-btn-primary:hover { background: #6d3410; }\n        .dsu-btn-primary:disabled { background: #d8d2c8; color: #7a7368; cursor: not-allowed; }\n        .dsu-btn-outline { background: transparent; color: #1c1a16; border: 1px solid #b0a898; }\n        .dsu-btn-outline:hover { border-color: #1c1a16; }\n        .dsu-btn-success { background: #2d6a4f; color: #fff; }\n        .dsu-btn-success:hover { background: #1e4d38; }\n        .dsu-btn-row { display: flex; justify-content: space-between; align-items: center; margin-top: 24px; gap: 10px; flex-wrap: wrap; }\n\n        .dsu-rtable { width: 100%; border-collapse: collapse; }\n        .dsu-rtable tr { border-bottom: 1px solid #d8d2c8; }\n        .dsu-rtable tr:last-child { border-bottom: none; }\n        .dsu-rtable td { padding: 9px 0; font-size: 13px; vertical-align: top; }\n        .dsu-rtable td:first-child { color: #7a7368; font-size: 11px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; width: 36%; padding-right: 14px; padding-top: 11px; }\n\n        .dsu-search-row { display: flex; gap: 10px; align-items: flex-start; }\n        .dsu-search-row .dsu-input { flex: 1; }\n\n        .dsu-spinner { width: 16px; height: 16px; border: 2px solid #d8d2c8; border-top-color: #8b4513; border-radius: 50%; animation: dsuspin 0.75s linear infinite; flex-shrink: 0; display: inline-block; vertical-align: middle; }\n        @keyframes dsuspin { to { transform: rotate(360deg); } }\n\n        .dsu-screen { display: none; }\n        .dsu-screen.active { display: block; }\n\n        .dsu-code { display: inline-block; background: #1c1a16; color: #f7f3e8; font-family: 'Courier New', monospace; font-size: 18px; font-weight: 700; letter-spacing: 0.15em; padding: 10px 20px; border: 2px solid #c9a84c; }\n\n        .dsu-status-badge { display: inline-block; padding: 3px 10px; font-size: 11px; font-weight: 600; letter-spacing: 0.05em; text-transform: uppercase; }\n        .dsu-status-badge.st-received { background: #f5ece4; color: #8b4513; }\n        .dsu-status-badge.st-review { background: #e8f0fe; color: #1a56db; }\n        .dsu-status-badge.st-dca { background: #fef3cd; color: #856404; }\n        .dsu-status-badge.st-released { background: #eaf3ed; color: #2d6a4f; }\n        .dsu-status-badge.st-filed { background: #e2e2e2; color: #555; }\n      </style>\n\n      <div class=\"dsu-wrap\">\n        <div class=\"dsu-header\">\n          <h2>Document Status Update</h2>\n          <p>Logged in as <strong>".concat(user.displayName, "</strong> &nbsp;&middot;&nbsp; ").concat(user.email, "</p>\n        </div>\n\n        <!-- Screen 1: Search -->\n        <div class=\"dsu-screen active\" id=\"dsu-s1\">\n          <div class=\"dsu-card\">\n            <div class=\"dsu-card-title\"><span class=\"bar\"></span>Search Document</div>\n            <div class=\"dsu-field\">\n              <label class=\"dsu-label\">Reference Code <span class=\"req\">*</span></label>\n              <div class=\"dsu-search-row\">\n                <input id=\"dsu-search\" type=\"text\" class=\"dsu-input\" placeholder=\"e.g. RCM-SC-0001\" />\n                <button id=\"dsu-btn-search\" class=\"dsu-btn dsu-btn-primary\">Search</button>\n              </div>\n            </div>\n            <div id=\"dsu-search-status\" style=\"margin-top:12px;\"></div>\n          </div>\n        </div>\n\n        <!-- Screen 2: Detail + Update -->\n        <div class=\"dsu-screen\" id=\"dsu-s2\">\n          <div class=\"dsu-card\">\n            <div class=\"dsu-card-title\"><span class=\"bar\"></span>Document Details</div>\n            <table class=\"dsu-rtable\" id=\"dsu-detail-table\"></table>\n          </div>\n          <div class=\"dsu-card\">\n            <div class=\"dsu-card-title\"><span class=\"bar\"></span>Update Status</div>\n            <div class=\"dsu-field\">\n              <label class=\"dsu-label\">Status <span class=\"req\">*</span></label>\n              <select id=\"dsu-status\" class=\"dsu-select\">\n                <option value=\"Received\">Received</option>\n                <option value=\"In Progress\">In Progress</option>\n                <option value=\"For DCA Approval and Signature\">For DCA Approval and Signature</option>\n                <option value=\"Released\">Released</option>\n                <option value=\"Filed\">Filed</option>\n              </select>\n            </div>\n            <div class=\"dsu-field\">\n              <label class=\"dsu-label\">Other Remarks</label>\n              <textarea id=\"dsu-remarks\" class=\"dsu-textarea\" placeholder=\"Optional notes...\"></textarea>\n            </div>\n            <div class=\"dsu-btn-row\">\n              <button id=\"dsu-btn-back\" class=\"dsu-btn dsu-btn-outline\">&larr; Search Again</button>\n              <button id=\"dsu-btn-submit\" class=\"dsu-btn dsu-btn-primary\">Update Status &rarr;</button>\n            </div>\n          </div>\n        </div>\n\n        <!-- Screen 3: Success -->\n        <div class=\"dsu-screen\" id=\"dsu-s3\">\n          <div class=\"dsu-card\">\n            <div class=\"dsu-notice success\">&check; &nbsp;Document status updated successfully.</div>\n            <div style=\"text-align:center;margin:18px 0;\">\n              <div style=\"font-family:'Courier New',monospace;font-size:10px;letter-spacing:0.16em;text-transform:uppercase;color:#7a7368;margin-bottom:8px;\">Reference Code</div>\n              <div class=\"dsu-code\" id=\"dsu-success-code\">\u2014</div>\n            </div>\n            <table class=\"dsu-rtable\" id=\"dsu-success-table\"></table>\n            <div class=\"dsu-btn-row\">\n              <button id=\"dsu-btn-new\" class=\"dsu-btn dsu-btn-outline\">+ Search Another</button>\n              <button id=\"dsu-btn-copy\" class=\"dsu-btn dsu-btn-success\">Copy Code</button>\n            </div>\n          </div>\n        </div>\n      </div>\n    ");
        this._bindEvents();
    };
    // ── Bind events ──────────────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype._bindEvents = function () {
        var _this = this;
        var searchInput = this.domElement.querySelector('#dsu-search');
        var searchBtn = this.domElement.querySelector('#dsu-btn-search');
        searchInput.addEventListener('keydown', function (e) {
            if (e.key === 'Enter')
                _this._doSearch();
        });
        searchBtn.addEventListener('click', function () { return _this._doSearch(); });
        this.domElement.querySelector('#dsu-btn-back').addEventListener('click', function () {
            _this.currentItem = null;
            _this._showScreen(1);
            _this.domElement.querySelector('#dsu-search').value = '';
            _this._el('dsu-search-status').innerHTML = '';
        });
        this.domElement.querySelector('#dsu-btn-submit').addEventListener('click', function () { return _this._doUpdate(); });
        this.domElement.querySelector('#dsu-btn-new').addEventListener('click', function () {
            _this.currentItem = null;
            _this._showScreen(1);
            _this.domElement.querySelector('#dsu-search').value = '';
            _this._el('dsu-search-status').innerHTML = '';
        });
        this.domElement.querySelector('#dsu-btn-copy').addEventListener('click', function () {
            var code = _this._el('dsu-success-code').textContent || '';
            navigator.clipboard.writeText(code)
                .then(function () { return alert('Copied: ' + code); })
                .catch(function () {
                window.prompt('Could not copy automatically. Copy the code below:', code);
            });
        });
    };
    // ── Search ──────────────────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype._doSearch = function () {
        return __awaiter(this, void 0, void 0, function () {
            var input, code, statusEl, searchBtn, filterCode, response, data, items, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        input = this.domElement.querySelector('#dsu-search');
                        code = input.value.trim();
                        statusEl = this._el('dsu-search-status');
                        searchBtn = this.domElement.querySelector('#dsu-btn-search');
                        if (!code) {
                            statusEl.innerHTML = '<div class="dsu-notice error">&cross; &nbsp;Please enter a reference code.</div>';
                            return [2 /*return*/];
                        }
                        searchBtn.disabled = true;
                        statusEl.innerHTML = '<div style="display:flex;align-items:center;gap:10px;font-size:13px;color:#7a7368;"><div class="dsu-spinner"></div>Searching...</div>';
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        filterCode = encodeURIComponent(code);
                        return [4 /*yield*/, this.context.spHttpClient.get("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items?$filter=ReferenceCode eq '").concat(filterCode, "'&$select=Id,Title,ReferenceCode,Status,From,Document_x0020_Type,Document_x0020_Format,OtherRemarks"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__.SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                }
                            })];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("SharePoint returned ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        items = data.value;
                        if (!items || items.length === 0) {
                            statusEl.innerHTML = "<div class=\"dsu-notice error\">&cross; &nbsp;No document found with reference code <strong>".concat(this._escapeHtml(code), "</strong>. Please check and try again.</div>");
                            searchBtn.disabled = false;
                            return [2 /*return*/];
                        }
                        this.currentItem = items[0];
                        this._showDetail();
                        searchBtn.disabled = false;
                        return [3 /*break*/, 5];
                    case 4:
                        err_1 = _a.sent();
                        console.error('DocumentStatusUpdate search error:', err_1);
                        statusEl.innerHTML = '<div class="dsu-notice error">&cross; &nbsp;Could not search SharePoint. Please check your connection and try again.</div>';
                        searchBtn.disabled = false;
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    // ── Show detail screen ─────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype._showDetail = function () {
        if (!this.currentItem)
            return;
        var item = this.currentItem;
        var formatLabel = item.Document_x0020_Format === 'HC' ? 'Physical (HC)' : 'Digital (SC)';
        var statusBadge = this._statusBadge(item.Status);
        var rows = [
            ['Reference Code', "<span class=\"dsu-code\" style=\"font-size:14px;padding:5px 12px;\">".concat(this._escapeHtml(item.ReferenceCode), "</span>")],
            ['Document Title', this._escapeHtml(item.Title)],
            ['Document Type', this._escapeHtml(item.Document_x0020_Type)],
            ['Document Format', formatLabel],
            ['From', this._escapeHtml(item.From)],
            ['Current Status', statusBadge],
            ['Remarks', this._escapeHtml(item.OtherRemarks || '') || '&mdash;'],
        ];
        this._el('dsu-detail-table').innerHTML =
            rows.map(function (_a) {
                var k = _a[0], v = _a[1];
                return "<tr><td>".concat(k, "</td><td>").concat(v, "</td></tr>");
            }).join('');
        // Pre-fill form
        this.domElement.querySelector('#dsu-status').value = item.Status;
        this.domElement.querySelector('#dsu-remarks').value = item.OtherRemarks || '';
        this._showScreen(2);
    };
    // ── Update ──────────────────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype._doUpdate = function () {
        return __awaiter(this, void 0, void 0, function () {
            var submitBtn, newStatus, newRemarks, body, response, errText, err_2, card, existing, notice;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.currentItem)
                            return [2 /*return*/];
                        submitBtn = this.domElement.querySelector('#dsu-btn-submit');
                        newStatus = this.domElement.querySelector('#dsu-status').value;
                        newRemarks = this.domElement.querySelector('#dsu-remarks').value.trim();
                        submitBtn.disabled = true;
                        submitBtn.textContent = 'Updating...';
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 5, , 6]);
                        body = JSON.stringify({
                            Status: newStatus,
                            OtherRemarks: newRemarks,
                        });
                        return [4 /*yield*/, this.context.spHttpClient.post("".concat(this.siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items(").concat(this.currentItem.Id, ")"), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_1__.SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=nometadata',
                                    'odata-version': '',
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
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
                    case 4:
                        this._showSuccess(newStatus, newRemarks);
                        return [3 /*break*/, 6];
                    case 5:
                        err_2 = _a.sent();
                        console.error('DocumentStatusUpdate update error:', err_2);
                        card = this.domElement.querySelectorAll('#dsu-s2 .dsu-card')[1];
                        existing = card.querySelector('.dsu-notice.error');
                        if (existing)
                            existing.remove();
                        notice = document.createElement('div');
                        notice.className = 'dsu-notice error';
                        notice.innerHTML = '&cross; &nbsp;Could not update the document. Please try again or contact your administrator.';
                        card.insertBefore(notice, card.firstChild);
                        submitBtn.disabled = false;
                        submitBtn.textContent = 'Update Status →';
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    // ── Success screen ─────────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype._showSuccess = function (newStatus, newRemarks) {
        if (!this.currentItem)
            return;
        var item = this.currentItem;
        var formatLabel = item.Document_x0020_Format === 'HC' ? 'Physical (HC)' : 'Digital (SC)';
        this._el('dsu-success-code').textContent = item.ReferenceCode;
        var rows = [
            ['Document Title', this._escapeHtml(item.Title)],
            ['Document Type', this._escapeHtml(item.Document_x0020_Type)],
            ['Document Format', formatLabel],
            ['From', this._escapeHtml(item.From)],
            ['Previous Status', this._statusBadge(item.Status)],
            ['New Status', this._statusBadge(newStatus)],
            ['Remarks', this._escapeHtml(newRemarks) || '&mdash;'],
            ['Updated By', this._escapeHtml(this.context.pageContext.user.displayName)],
            ['Updated On', new Date().toLocaleDateString('en-PH') + ' ' + new Date().toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit' })],
        ];
        this._el('dsu-success-table').innerHTML =
            rows.map(function (_a) {
                var k = _a[0], v = _a[1];
                return "<tr><td>".concat(k, "</td><td>").concat(v, "</td></tr>");
            }).join('');
        this._showScreen(3);
    };
    // ── Helpers ─────────────────────────────────────────────────────────────
    DocumentStatusUpdateWebPart.prototype._el = function (id) {
        return this.domElement.querySelector("#".concat(id));
    };
    DocumentStatusUpdateWebPart.prototype._showScreen = function (n) {
        this.domElement.querySelectorAll('.dsu-screen').forEach(function (s) { return s.classList.remove('active'); });
        this._el("dsu-s".concat(n)).classList.add('active');
    };
    DocumentStatusUpdateWebPart.prototype._escapeHtml = function (text) {
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    };
    DocumentStatusUpdateWebPart.prototype._statusBadge = function (status) {
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
        return "<span class=\"dsu-status-badge ".concat(cls, "\">").concat(this._escapeHtml(status), "</span>");
    };
    Object.defineProperty(DocumentStatusUpdateWebPart.prototype, "dataVersion", {
        // ── SPFx lifecycle ──────────────────────────────────────────────────────
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    DocumentStatusUpdateWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: 'Document Status Update Settings' },
                    groups: [
                        {
                            groupName: 'Configuration',
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__.PropertyPaneTextField)('siteUrl', {
                                    label: 'SharePoint Site URL',
                                    description: 'Leave blank to use the current site',
                                    placeholder: 'https://tenant.sharepoint.com/sites/your-site'
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__.PropertyPaneTextField)('listName', {
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
    return DocumentStatusUpdateWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (DocumentStatusUpdateWebPart);

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=document-status-update-web-part.js.map