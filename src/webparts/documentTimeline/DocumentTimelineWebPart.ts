import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export interface IDocumentTimelineWebPartProps {
  siteUrl: string;
  listName: string;
  statusLogListName: string;
}

interface IDocumentItem {
  Id: number;
  Title: string;
  ReferenceCode: string;
  Status: string;
  From: string;
  Document_x0020_Type: string;
  Document_x0020_Format: string;
  OtherRemarks: string;
  Created: string;
}

interface IStatusLogEntry {
  Title: string;
  FromStatus: string;
  ToStatus: string;
  DurationHours: number;
  ChangedBy: string;
  DocumentTitle: string;
  Created: string;
}

export default class DocumentTimelineWebPart extends BaseClientSideWebPart<IDocumentTimelineWebPartProps> {

  private get siteUrl(): string { return this.properties.siteUrl || this.context.pageContext.web.absoluteUrl; }
  private get listName(): string { return this.properties.listName || 'Document Log Tracking'; }
  private get statusLogListName(): string { return this.properties.statusLogListName || 'Document Status Log'; }

  // ── Render ──────────────────────────────────────────────────────────────
  public render(): void {
    const user = this.context.pageContext.user;

    this.domElement.innerHTML = `
      <style>
        .dtl-wrap { font-family: 'Segoe UI', sans-serif; max-width: 700px; margin: 0 auto; padding: 24px 16px; color: #1c1a16; }
        .dtl-header { text-align: center; margin-bottom: 28px; padding-bottom: 20px; border-bottom: 2px solid #1c1a16; }
        .dtl-header h2 { font-size: 22px; font-weight: 700; margin: 0 0 6px; }
        .dtl-header p { font-size: 13px; color: #7a7368; margin: 0; }

        .dtl-card { background: #fff; border: 1px solid #d8d2c8; padding: 28px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
        .dtl-card-title { font-size: 15px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid #d8d2c8; display: flex; align-items: center; gap: 8px; }
        .dtl-card-title .bar { width: 3px; height: 14px; background: #c9a84c; flex-shrink: 0; }

        .dtl-field { margin-bottom: 18px; }
        .dtl-label { display: block; font-size: 11px; font-weight: 600; color: #7a7368; margin-bottom: 5px; letter-spacing: 0.07em; text-transform: uppercase; }
        .dtl-label .req { color: #8b4513; }
        .dtl-input { width: 100%; padding: 9px 12px; border: 1px solid #d8d2c8; background: #f7f5f0; font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #1c1a16; outline: none; box-sizing: border-box; transition: border-color 0.15s; }
        .dtl-input:focus { border-color: #8b4513; background: #fff; }

        .dtl-notice { padding: 11px 14px; font-size: 13px; display: flex; gap: 8px; align-items: flex-start; margin-bottom: 18px; line-height: 1.5; }
        .dtl-notice.info { background: #f5ece4; border-left: 3px solid #c8773a; color: #8b4513; }
        .dtl-notice.success { background: #eaf3ed; border-left: 3px solid #2d6a4f; color: #2d6a4f; }
        .dtl-notice.error { background: #fdf5f5; border-left: 3px solid #c0392b; color: #c0392b; }

        .dtl-btn { padding: 10px 22px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.15s; letter-spacing: 0.03em; }
        .dtl-btn-primary { background: #8b4513; color: #fff; }
        .dtl-btn-primary:hover { background: #6d3410; }
        .dtl-btn-primary:disabled { background: #d8d2c8; color: #7a7368; cursor: not-allowed; }
        .dtl-btn-outline { background: transparent; color: #1c1a16; border: 1px solid #b0a898; }
        .dtl-btn-outline:hover { border-color: #1c1a16; }
        .dtl-btn-row { display: flex; justify-content: space-between; align-items: center; margin-top: 24px; gap: 10px; flex-wrap: wrap; }

        .dtl-rtable { width: 100%; border-collapse: collapse; }
        .dtl-rtable tr { border-bottom: 1px solid #d8d2c8; }
        .dtl-rtable tr:last-child { border-bottom: none; }
        .dtl-rtable td { padding: 9px 0; font-size: 13px; vertical-align: top; }
        .dtl-rtable td:first-child { color: #7a7368; font-size: 11px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; width: 36%; padding-right: 14px; padding-top: 11px; }

        .dtl-search-row { display: flex; gap: 10px; align-items: flex-start; }
        .dtl-search-row .dtl-input { flex: 1; }

        .dtl-spinner { width: 16px; height: 16px; border: 2px solid #d8d2c8; border-top-color: #8b4513; border-radius: 50%; animation: dtlspin 0.75s linear infinite; flex-shrink: 0; display: inline-block; vertical-align: middle; }
        @keyframes dtlspin { to { transform: rotate(360deg); } }

        .dtl-screen { display: none; }
        .dtl-screen.active { display: block; }

        .dtl-code { display: inline-block; background: #1c1a16; color: #f7f3e8; font-family: 'Courier New', monospace; font-size: 22px; font-weight: 700; letter-spacing: 0.18em; padding: 12px 24px; border: 2px solid #c9a84c; }

        /* Timeline */
        .dtl-timeline { position: relative; padding: 0; margin: 0; list-style: none; }
        .dtl-timeline::before { content: ''; position: absolute; left: 15px; top: 0; bottom: 0; width: 2px; background: #d8d2c8; }

        .dtl-tl-node { position: relative; padding: 0 0 28px 44px; }
        .dtl-tl-node:last-child { padding-bottom: 0; }
        .dtl-tl-dot { position: absolute; left: 8px; top: 2px; width: 16px; height: 16px; border-radius: 50%; border: 2px solid #2d6a4f; background: #2d6a4f; z-index: 1; }
        .dtl-tl-node.current .dtl-tl-dot { border-color: #8b4513; background: #8b4513; animation: dtlpulse 2s infinite; }
        .dtl-tl-node.origin .dtl-tl-dot { border-color: #7a7368; background: #7a7368; }
        @keyframes dtlpulse { 0%,100%{box-shadow:0 0 0 0 rgba(139,69,19,0.4)} 50%{box-shadow:0 0 0 8px rgba(139,69,19,0)} }

        .dtl-tl-content { background: #f7f5f0; border: 1px solid #d8d2c8; padding: 14px 16px; }
        .dtl-tl-node.current .dtl-tl-content { border-color: #8b4513; background: #fdf8f3; }
        .dtl-tl-node.origin .dtl-tl-content { border-color: #b0a898; background: #f7f5f0; }

        .dtl-tl-transition { font-size: 14px; font-weight: 700; margin-bottom: 6px; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
        .dtl-tl-arrow { color: #8b4513; font-weight: 400; }
        .dtl-tl-meta { font-size: 11px; color: #7a7368; line-height: 1.6; }
        .dtl-tl-meta span { margin-right: 14px; }

        .dtl-tl-duration { display: inline-block; font-family: 'Courier New', monospace; font-size: 11px; font-weight: 700; background: #eaf3ed; color: #2d6a4f; padding: 2px 8px; letter-spacing: 0.04em; }
        .dtl-tl-node.current .dtl-tl-duration { background: #f5ece4; color: #8b4513; }

        .dtl-total-bar { display: flex; align-items: center; justify-content: center; gap: 12px; padding: 16px; background: #1c1a16; color: #f7f3e8; margin-top: 14px; }
        .dtl-total-label { font-size: 10px; font-weight: 600; letter-spacing: 0.14em; text-transform: uppercase; color: #b0a898; }
        .dtl-total-value { font-family: 'Courier New', monospace; font-size: 22px; font-weight: 700; letter-spacing: 0.08em; color: #c9a84c; }

        .dtl-status-badge { display: inline-block; padding: 3px 10px; font-size: 11px; font-weight: 600; letter-spacing: 0.05em; text-transform: uppercase; }
        .dtl-status-badge.st-received { background: #f5ece4; color: #8b4513; }
        .dtl-status-badge.st-review { background: #e8f0fe; color: #1a56db; }
        .dtl-status-badge.st-dca { background: #fef3cd; color: #856404; }
        .dtl-status-badge.st-released { background: #eaf3ed; color: #2d6a4f; }
        .dtl-status-badge.st-filed { background: #e2e2e2; color: #555; }

        .dtl-empty-tl { text-align: center; padding: 32px 16px; color: #7a7368; font-size: 13px; }
        .dtl-empty-tl .icon { font-size: 28px; margin-bottom: 8px; }
      </style>

      <div class="dtl-wrap">
        <div class="dtl-header">
          <h2>Document Timeline</h2>
          <p>Logged in as <strong>${this._esc(user.displayName)}</strong> &nbsp;&middot;&nbsp; ${this._esc(user.email)}</p>
        </div>

        <!-- Screen 1: Search -->
        <div class="dtl-screen active" id="dtl-s1">
          <div class="dtl-card">
            <div class="dtl-card-title"><span class="bar"></span>Search Document</div>
            <div class="dtl-field">
              <label class="dtl-label">Reference Code <span class="req">*</span></label>
              <div class="dtl-search-row">
                <input id="dtl-search" type="text" class="dtl-input" placeholder="e.g. RCM-SC-0009" />
                <button id="dtl-btn-search" class="dtl-btn dtl-btn-primary">Search</button>
              </div>
            </div>
            <div id="dtl-search-status" style="margin-top:12px;"></div>
          </div>
        </div>

        <!-- Screen 2: Timeline -->
        <div class="dtl-screen" id="dtl-s2">
          <div class="dtl-card" id="dtl-summary-card"></div>
          <div class="dtl-card">
            <div class="dtl-card-title"><span class="bar"></span>Status Timeline</div>
            <div id="dtl-timeline-body"></div>
          </div>
          <div id="dtl-total-bar"></div>
          <div class="dtl-btn-row" style="margin-top:14px;">
            <button id="dtl-btn-back" class="dtl-btn dtl-btn-outline">&larr; Search Again</button>
          </div>
        </div>
      </div>
    `;

    this._bindEvents();
  }

  // ── Events ──────────────────────────────────────────────────────────────
  private _bindEvents(): void {
    const searchInput = this.domElement.querySelector('#dtl-search') as HTMLInputElement;
    searchInput.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter') this._doSearch();
    });
    this.domElement.querySelector('#dtl-btn-search')!.addEventListener('click', () => this._doSearch());
    this.domElement.querySelector('#dtl-btn-back')!.addEventListener('click', () => {
      this._showScreen(1);
      (this.domElement.querySelector('#dtl-search') as HTMLInputElement).value = '';
      this._el('dtl-search-status').innerHTML = '';
    });
  }

  // ── Search ──────────────────────────────────────────────────────────────
  private async _doSearch(): Promise<void> {
    const input = this.domElement.querySelector('#dtl-search') as HTMLInputElement;
    const code = input.value.trim();
    const statusEl = this._el('dtl-search-status');
    const searchBtn = this.domElement.querySelector('#dtl-btn-search') as HTMLButtonElement;

    if (!code) {
      statusEl.innerHTML = '<div class="dtl-notice error">&cross; &nbsp;Please enter a reference code.</div>';
      return;
    }

    searchBtn.disabled = true;
    statusEl.innerHTML = '<div style="display:flex;align-items:center;gap:10px;font-size:13px;color:#7a7368;"><div class="dtl-spinner"></div>Searching...</div>';

    try {
      const filterCode = encodeURIComponent(code);

      const [docResponse, logResponse]: [SPHttpClientResponse, SPHttpClientResponse] = await Promise.all([
        this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$filter=ReferenceCode eq '${filterCode}'`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } }
        ),
        this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${this.statusLogListName}')/items?$filter=Title eq '${filterCode}'&$orderby=Created asc`,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } }
        )
      ]);

      if (!docResponse.ok) {
        throw new Error(`Document list returned ${docResponse.status}`);
      }
      if (!logResponse.ok) {
        throw new Error(`Status log list returned ${logResponse.status}`);
      }

      const docData = await docResponse.json();
      const logData = await logResponse.json();

      const items: IDocumentItem[] = docData.value;
      if (!items || items.length === 0) {
        statusEl.innerHTML = `<div class="dtl-notice error">&cross; &nbsp;No document found with reference code <strong>${this._esc(code)}</strong>.</div>`;
        searchBtn.disabled = false;
        return;
      }

      const doc = items[0];
      const logs: IStatusLogEntry[] = logData.value || [];

      this._renderTimeline(doc, logs);
      searchBtn.disabled = false;

    } catch (err) {
      console.error('DocumentTimeline search error:', err);
      statusEl.innerHTML = '<div class="dtl-notice error">&cross; &nbsp;Could not search SharePoint. Please check your connection and try again.</div>';
      searchBtn.disabled = false;
    }
  }

  // ── Render timeline ─────────────────────────────────────────────────────
  private _renderTimeline(doc: IDocumentItem, logs: IStatusLogEntry[]): void {
    // Summary card
    const createdDate = new Date(doc.Created);
    const formatLabel = doc.Document_x0020_Format === 'HC' ? 'Physical (HC)' : 'Digital (SC)';
    this._el('dtl-summary-card').innerHTML = `
      <div style="text-align:center;margin-bottom:18px;">
        <div style="font-family:'Courier New',monospace;font-size:10px;letter-spacing:0.16em;text-transform:uppercase;color:#7a7368;margin-bottom:8px;">Reference Code</div>
        <div class="dtl-code">${this._esc(doc.ReferenceCode)}</div>
      </div>
      <table class="dtl-rtable">
        <tr><td>Document Title</td><td>${this._esc(doc.Title)}</td></tr>
        <tr><td>Document Type</td><td>${this._esc(doc.Document_x0020_Type)}</td></tr>
        <tr><td>Document Format</td><td>${formatLabel}</td></tr>
        <tr><td>From</td><td>${this._esc(doc.From)}</td></tr>
        <tr><td>Current Status</td><td>${this._statusBadge(doc.Status)}</td></tr>
        <tr><td>Date Logged</td><td>${this._formatDate(createdDate)}</td></tr>
      </table>
    `;

    // Timeline
    const timelineBody = this._el('dtl-timeline-body');

    if (logs.length === 0) {
      timelineBody.innerHTML = `
        <div class="dtl-empty-tl">
          <div class="icon">&#8986;</div>
          <div>No status transitions recorded yet.</div>
          <div style="margin-top:4px;font-size:12px;">The document is currently at <strong>${this._esc(doc.Status)}</strong>.</div>
        </div>
      `;
      this._el('dtl-total-bar').innerHTML = '';
      this._showScreen(2);
      return;
    }

    let totalDuration = 0;
    let html = '<ul class="dtl-timeline">';

    // Origin node: the initial status (FromStatus of the first log entry)
    const firstLog = logs[0];
    const originDate = new Date(doc.Created);
    html += `
      <li class="dtl-tl-node origin">
        <div class="dtl-tl-dot"></div>
        <div class="dtl-tl-content">
          <div class="dtl-tl-transition">${this._statusBadge(firstLog.FromStatus)}</div>
          <div class="dtl-tl-meta">
            <span>Initial status</span>
            <span>${this._formatDateTime(originDate)}</span>
          </div>
        </div>
      </li>
    `;

    // Transition nodes
    for (let i = 0; i < logs.length; i++) {
      const log = logs[i];
      const isLast = i === logs.length - 1;
      const nodeClass = isLast ? 'current' : '';
      const logDate = new Date(log.Created);
      const durationMinutes = log.DurationHours || 0;
      totalDuration += durationMinutes;

      html += `
        <li class="dtl-tl-node ${nodeClass}">
          <div class="dtl-tl-dot"></div>
          <div class="dtl-tl-content">
            <div class="dtl-tl-transition">
              ${this._statusBadge(log.FromStatus)}
              <span class="dtl-tl-arrow">&rarr;</span>
              ${this._statusBadge(log.ToStatus)}
            </div>
            <div class="dtl-tl-meta">
              <span class="dtl-tl-duration">${this._formatDuration(durationMinutes)}</span>
              <span>by ${this._esc(log.ChangedBy || 'System')}</span>
              <span>${this._formatDateTime(logDate)}</span>
            </div>
          </div>
        </li>
      `;
    }

    html += '</ul>';
    timelineBody.innerHTML = html;

    // Total bar
    this._el('dtl-total-bar').innerHTML = `
      <div class="dtl-total-bar">
        <span class="dtl-total-label">Total Time Across Transitions</span>
        <span class="dtl-total-value">${this._formatDuration(totalDuration)}</span>
      </div>
    `;

    this._showScreen(2);
  }

  // ── Helpers ─────────────────────────────────────────────────────────────
  private _el(id: string): HTMLElement {
    return this.domElement.querySelector(`#${id}`) as HTMLElement;
  }

  private _showScreen(n: number): void {
    this.domElement.querySelectorAll('.dtl-screen').forEach(s => s.classList.remove('active'));
    this._el(`dtl-s${n}`).classList.add('active');
  }

  private _esc(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  private _statusBadge(status: string): string {
    let cls = '';
    switch (status) {
      case 'Received': cls = 'st-received'; break;
      case 'In Progress': cls = 'st-review'; break;
      case 'For DCA Approval and Signature': cls = 'st-dca'; break;
      case 'Released': cls = 'st-released'; break;
      case 'Filed': cls = 'st-filed'; break;
    }
    return `<span class="dtl-status-badge ${cls}">${this._esc(status)}</span>`;
  }

  private _formatDate(d: Date): string {
    return d.toLocaleDateString('en-PH', { year: 'numeric', month: 'long', day: 'numeric' });
  }

  private _formatDateTime(d: Date): string {
    return d.toLocaleDateString('en-PH', { year: 'numeric', month: 'short', day: 'numeric' }) +
      ' ' + d.toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit' });
  }

  private _formatDuration(decimalMinutes: number): string {
    const totalSeconds = Math.round(decimalMinutes * 60);
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;

    if (h > 0) {
      return s > 0 ? `${h}h ${m}m ${s}s` : m > 0 ? `${h}h ${m}m` : `${h}h`;
    }
    if (m > 0) {
      return s > 0 ? `${m}m ${s}s` : `${m}m`;
    }
    return `${s}s`;
  }

  // ── SPFx lifecycle ──────────────────────────────────────────────────────
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }
}
