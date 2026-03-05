import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export interface IDocumentStatusUpdateWebPartProps {
  siteUrl: string;
  listName: string;
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
}

export default class DocumentStatusUpdateWebPart extends BaseClientSideWebPart<IDocumentStatusUpdateWebPartProps> {

  private get siteUrl(): string { return this.properties.siteUrl || this.context.pageContext.web.absoluteUrl; }
  private get listName(): string { return this.properties.listName || 'Document Log Tracking'; }

  private currentItem: IDocumentItem | null = null;

  // ── Render ──────────────────────────────────────────────────────────────
  public render(): void {
    const user = this.context.pageContext.user;

    this.domElement.innerHTML = `
      <style>
        .dsu-wrap { font-family: 'Segoe UI', sans-serif; max-width: 700px; margin: 0 auto; padding: 24px 16px; color: #1c1a16; }
        .dsu-header { text-align: center; margin-bottom: 28px; padding-bottom: 20px; border-bottom: 2px solid #1c1a16; }
        .dsu-header h2 { font-size: 22px; font-weight: 700; margin: 0 0 6px; }
        .dsu-header p { font-size: 13px; color: #7a7368; margin: 0; }

        .dsu-card { background: #fff; border: 1px solid #d8d2c8; padding: 28px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
        .dsu-card-title { font-size: 15px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid #d8d2c8; display: flex; align-items: center; gap: 8px; }
        .dsu-card-title .bar { width: 3px; height: 14px; background: #c9a84c; flex-shrink: 0; }

        .dsu-field { margin-bottom: 18px; }
        .dsu-field:last-of-type { margin-bottom: 0; }
        .dsu-label { display: block; font-size: 11px; font-weight: 600; color: #7a7368; margin-bottom: 5px; letter-spacing: 0.07em; text-transform: uppercase; }
        .dsu-label .req { color: #8b4513; }
        .dsu-input, .dsu-select, .dsu-textarea { width: 100%; padding: 9px 12px; border: 1px solid #d8d2c8; background: #f7f5f0; font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #1c1a16; outline: none; box-sizing: border-box; transition: border-color 0.15s; }
        .dsu-input:focus, .dsu-select:focus, .dsu-textarea:focus { border-color: #8b4513; background: #fff; }
        .dsu-textarea { resize: vertical; min-height: 80px; line-height: 1.5; }

        .dsu-notice { padding: 11px 14px; font-size: 13px; display: flex; gap: 8px; align-items: flex-start; margin-bottom: 18px; line-height: 1.5; }
        .dsu-notice.info { background: #f5ece4; border-left: 3px solid #c8773a; color: #8b4513; }
        .dsu-notice.success { background: #eaf3ed; border-left: 3px solid #2d6a4f; color: #2d6a4f; }
        .dsu-notice.error { background: #fdf5f5; border-left: 3px solid #c0392b; color: #c0392b; }

        .dsu-btn { padding: 10px 22px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.15s; letter-spacing: 0.03em; }
        .dsu-btn-primary { background: #8b4513; color: #fff; }
        .dsu-btn-primary:hover { background: #6d3410; }
        .dsu-btn-primary:disabled { background: #d8d2c8; color: #7a7368; cursor: not-allowed; }
        .dsu-btn-outline { background: transparent; color: #1c1a16; border: 1px solid #b0a898; }
        .dsu-btn-outline:hover { border-color: #1c1a16; }
        .dsu-btn-success { background: #2d6a4f; color: #fff; }
        .dsu-btn-success:hover { background: #1e4d38; }
        .dsu-btn-row { display: flex; justify-content: space-between; align-items: center; margin-top: 24px; gap: 10px; flex-wrap: wrap; }

        .dsu-rtable { width: 100%; border-collapse: collapse; }
        .dsu-rtable tr { border-bottom: 1px solid #d8d2c8; }
        .dsu-rtable tr:last-child { border-bottom: none; }
        .dsu-rtable td { padding: 9px 0; font-size: 13px; vertical-align: top; }
        .dsu-rtable td:first-child { color: #7a7368; font-size: 11px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; width: 36%; padding-right: 14px; padding-top: 11px; }

        .dsu-search-row { display: flex; gap: 10px; align-items: flex-start; }
        .dsu-search-row .dsu-input { flex: 1; }

        .dsu-spinner { width: 16px; height: 16px; border: 2px solid #d8d2c8; border-top-color: #8b4513; border-radius: 50%; animation: dsuspin 0.75s linear infinite; flex-shrink: 0; display: inline-block; vertical-align: middle; }
        @keyframes dsuspin { to { transform: rotate(360deg); } }

        .dsu-screen { display: none; }
        .dsu-screen.active { display: block; }

        .dsu-code { display: inline-block; background: #1c1a16; color: #f7f3e8; font-family: 'Courier New', monospace; font-size: 18px; font-weight: 700; letter-spacing: 0.15em; padding: 10px 20px; border: 2px solid #c9a84c; }

        .dsu-status-badge { display: inline-block; padding: 3px 10px; font-size: 11px; font-weight: 600; letter-spacing: 0.05em; text-transform: uppercase; }
        .dsu-status-badge.st-received { background: #f5ece4; color: #8b4513; }
        .dsu-status-badge.st-review { background: #e8f0fe; color: #1a56db; }
        .dsu-status-badge.st-dca { background: #fef3cd; color: #856404; }
        .dsu-status-badge.st-released { background: #eaf3ed; color: #2d6a4f; }
        .dsu-status-badge.st-filed { background: #e2e2e2; color: #555; }
      </style>

      <div class="dsu-wrap">
        <div class="dsu-header">
          <h2>Document Status Update</h2>
          <p>Logged in as <strong>${user.displayName}</strong> &nbsp;&middot;&nbsp; ${user.email}</p>
        </div>

        <!-- Screen 1: Search -->
        <div class="dsu-screen active" id="dsu-s1">
          <div class="dsu-card">
            <div class="dsu-card-title"><span class="bar"></span>Search Document</div>
            <div class="dsu-field">
              <label class="dsu-label">Reference Code <span class="req">*</span></label>
              <div class="dsu-search-row">
                <input id="dsu-search" type="text" class="dsu-input" placeholder="e.g. RCM-SC-0001" />
                <button id="dsu-btn-search" class="dsu-btn dsu-btn-primary">Search</button>
              </div>
            </div>
            <div id="dsu-search-status" style="margin-top:12px;"></div>
          </div>
        </div>

        <!-- Screen 2: Detail + Update -->
        <div class="dsu-screen" id="dsu-s2">
          <div class="dsu-card">
            <div class="dsu-card-title"><span class="bar"></span>Document Details</div>
            <table class="dsu-rtable" id="dsu-detail-table"></table>
          </div>
          <div class="dsu-card">
            <div class="dsu-card-title"><span class="bar"></span>Update Status</div>
            <div class="dsu-field">
              <label class="dsu-label">Status <span class="req">*</span></label>
              <select id="dsu-status" class="dsu-select">
                <option value="Received">Received</option>
                <option value="In Progress">In Progress</option>
                <option value="For DCA Approval and Signature">For DCA Approval and Signature</option>
                <option value="Released">Released</option>
                <option value="Filed">Filed</option>
              </select>
            </div>
            <div class="dsu-field">
              <label class="dsu-label">Other Remarks</label>
              <textarea id="dsu-remarks" class="dsu-textarea" placeholder="Optional notes..."></textarea>
            </div>
            <div class="dsu-btn-row">
              <button id="dsu-btn-back" class="dsu-btn dsu-btn-outline">&larr; Search Again</button>
              <button id="dsu-btn-submit" class="dsu-btn dsu-btn-primary">Update Status &rarr;</button>
            </div>
          </div>
        </div>

        <!-- Screen 3: Success -->
        <div class="dsu-screen" id="dsu-s3">
          <div class="dsu-card">
            <div class="dsu-notice success">&check; &nbsp;Document status updated successfully.</div>
            <div style="text-align:center;margin:18px 0;">
              <div style="font-family:'Courier New',monospace;font-size:10px;letter-spacing:0.16em;text-transform:uppercase;color:#7a7368;margin-bottom:8px;">Reference Code</div>
              <div class="dsu-code" id="dsu-success-code">—</div>
            </div>
            <table class="dsu-rtable" id="dsu-success-table"></table>
            <div class="dsu-btn-row">
              <button id="dsu-btn-new" class="dsu-btn dsu-btn-outline">+ Search Another</button>
              <button id="dsu-btn-copy" class="dsu-btn dsu-btn-success">Copy Code</button>
            </div>
          </div>
        </div>
      </div>
    `;

    this._bindEvents();
  }

  // ── Bind events ──────────────────────────────────────────────────────────
  private _bindEvents(): void {
    const searchInput = this.domElement.querySelector('#dsu-search') as HTMLInputElement;
    const searchBtn = this.domElement.querySelector('#dsu-btn-search') as HTMLButtonElement;

    searchInput.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter') this._doSearch();
    });
    searchBtn.addEventListener('click', () => this._doSearch());

    this.domElement.querySelector('#dsu-btn-back')!.addEventListener('click', () => {
      this.currentItem = null;
      this._showScreen(1);
      (this.domElement.querySelector('#dsu-search') as HTMLInputElement).value = '';
      this._el('dsu-search-status').innerHTML = '';
    });

    this.domElement.querySelector('#dsu-btn-submit')!.addEventListener('click', () => this._doUpdate());

    this.domElement.querySelector('#dsu-btn-new')!.addEventListener('click', () => {
      this.currentItem = null;
      this._showScreen(1);
      (this.domElement.querySelector('#dsu-search') as HTMLInputElement).value = '';
      this._el('dsu-search-status').innerHTML = '';
    });

    this.domElement.querySelector('#dsu-btn-copy')!.addEventListener('click', () => {
      const code = this._el('dsu-success-code').textContent || '';
      navigator.clipboard.writeText(code)
        .then(() => alert('Copied: ' + code))
        .catch(() => {
          window.prompt('Could not copy automatically. Copy the code below:', code);
        });
    });
  }

  // ── Search ──────────────────────────────────────────────────────────────
  private async _doSearch(): Promise<void> {
    const input = this.domElement.querySelector('#dsu-search') as HTMLInputElement;
    const code = input.value.trim();
    const statusEl = this._el('dsu-search-status');
    const searchBtn = this.domElement.querySelector('#dsu-btn-search') as HTMLButtonElement;

    if (!code) {
      statusEl.innerHTML = '<div class="dsu-notice error">&cross; &nbsp;Please enter a reference code.</div>';
      return;
    }

    searchBtn.disabled = true;
    statusEl.innerHTML = '<div style="display:flex;align-items:center;gap:10px;font-size:13px;color:#7a7368;"><div class="dsu-spinner"></div>Searching...</div>';

    try {
      const filterCode = encodeURIComponent(code);
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$filter=ReferenceCode eq '${filterCode}'&$select=Id,Title,ReferenceCode,Status,From,Document_x0020_Type,Document_x0020_Format,OtherRemarks`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        throw new Error(`SharePoint returned ${response.status}`);
      }

      const data = await response.json();
      const items: IDocumentItem[] = data.value;

      if (!items || items.length === 0) {
        statusEl.innerHTML = `<div class="dsu-notice error">&cross; &nbsp;No document found with reference code <strong>${this._escapeHtml(code)}</strong>. Please check and try again.</div>`;
        searchBtn.disabled = false;
        return;
      }

      this.currentItem = items[0];
      this._showDetail();
      searchBtn.disabled = false;

    } catch (err) {
      console.error('DocumentStatusUpdate search error:', err);
      statusEl.innerHTML = '<div class="dsu-notice error">&cross; &nbsp;Could not search SharePoint. Please check your connection and try again.</div>';
      searchBtn.disabled = false;
    }
  }

  // ── Show detail screen ─────────────────────────────────────────────────
  private _showDetail(): void {
    if (!this.currentItem) return;
    const item = this.currentItem;

    const formatLabel = item.Document_x0020_Format === 'HC' ? 'Physical (HC)' : 'Digital (SC)';
    const statusBadge = this._statusBadge(item.Status);

    const rows = [
      ['Reference Code', `<span class="dsu-code" style="font-size:14px;padding:5px 12px;">${this._escapeHtml(item.ReferenceCode)}</span>`],
      ['Document Title', this._escapeHtml(item.Title)],
      ['Document Type', this._escapeHtml(item.Document_x0020_Type)],
      ['Document Format', formatLabel],
      ['From', this._escapeHtml(item.From)],
      ['Current Status', statusBadge],
      ['Remarks', this._escapeHtml(item.OtherRemarks || '') || '&mdash;'],
    ];

    this._el('dsu-detail-table').innerHTML =
      rows.map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('');

    // Pre-fill form
    (this.domElement.querySelector('#dsu-status') as HTMLSelectElement).value = item.Status;
    (this.domElement.querySelector('#dsu-remarks') as HTMLTextAreaElement).value = item.OtherRemarks || '';

    this._showScreen(2);
  }

  // ── Update ──────────────────────────────────────────────────────────────
  private async _doUpdate(): Promise<void> {
    if (!this.currentItem) return;

    const submitBtn = this.domElement.querySelector('#dsu-btn-submit') as HTMLButtonElement;
    const newStatus = (this.domElement.querySelector('#dsu-status') as HTMLSelectElement).value;
    const newRemarks = (this.domElement.querySelector('#dsu-remarks') as HTMLTextAreaElement).value.trim();

    submitBtn.disabled = true;
    submitBtn.textContent = 'Updating...';

    try {
      const body = JSON.stringify({
        Status: newStatus,
        OtherRemarks: newRemarks,
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${this.currentItem.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`SharePoint ${response.status}: ${errText}`);
      }

      this._showSuccess(newStatus, newRemarks);

    } catch (err) {
      console.error('DocumentStatusUpdate update error:', err);
      // Re-show the detail card with error
      const card = this.domElement.querySelectorAll('#dsu-s2 .dsu-card')[1];
      const existing = card.querySelector('.dsu-notice.error');
      if (existing) existing.remove();
      const notice = document.createElement('div');
      notice.className = 'dsu-notice error';
      notice.innerHTML = '&cross; &nbsp;Could not update the document. Please try again or contact your administrator.';
      card.insertBefore(notice, card.firstChild);
      submitBtn.disabled = false;
      submitBtn.textContent = 'Update Status →';
    }
  }

  // ── Success screen ─────────────────────────────────────────────────────
  private _showSuccess(newStatus: string, newRemarks: string): void {
    if (!this.currentItem) return;

    const item = this.currentItem;
    const formatLabel = item.Document_x0020_Format === 'HC' ? 'Physical (HC)' : 'Digital (SC)';

    this._el('dsu-success-code').textContent = item.ReferenceCode;

    const rows = [
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
      rows.map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('');

    this._showScreen(3);
  }

  // ── Helpers ─────────────────────────────────────────────────────────────
  private _el(id: string): HTMLElement {
    return this.domElement.querySelector(`#${id}`) as HTMLElement;
  }

  private _showScreen(n: number): void {
    this.domElement.querySelectorAll('.dsu-screen').forEach(s => s.classList.remove('active'));
    this._el(`dsu-s${n}`).classList.add('active');
  }

  private _escapeHtml(text: string): string {
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
    return `<span class="dsu-status-badge ${cls}">${this._escapeHtml(status)}</span>`;
  }

  // ── SPFx lifecycle ──────────────────────────────────────────────────────
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Document Status Update Settings' },
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
  }
}
