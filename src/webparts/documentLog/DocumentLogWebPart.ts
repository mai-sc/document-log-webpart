import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export interface IDocumentLogWebPartProps {
  siteUrl: string;
  listName: string;
}

export default class DocumentLogWebPart extends BaseClientSideWebPart<IDocumentLogWebPartProps> {

  private get siteUrl(): string { return this.properties.siteUrl || this.context.pageContext.web.absoluteUrl; }
  private get listName(): string { return this.properties.listName || 'Document Log Tracking'; }

  private pollInterval: ReturnType<typeof setInterval> | null = null;

  public dispose(): void {
    if (this.tick) clearInterval(this.tick);
    if (this.pollInterval) clearInterval(this.pollInterval);
    this.tick = null;
    this.pollInterval = null;
    super.dispose();
  }

  // ── Timers ──────────────────────────────────────────────────────────────
  private timers: Record<string, number> = {};
  private tick: ReturnType<typeof setInterval> | null = null;

  private fmt(ms: number): string {
    if (!ms && ms !== 0) return '—';
    const s = Math.round(ms / 1000);
    return s < 60 ? `${s}s` : `${Math.floor(s / 60)}m ${s % 60}s`;
  }

  private startTick(elId: string, t0: number): void {
    if (this.tick) clearInterval(this.tick);
    this.tick = setInterval(() => {
      const el = this.domElement.querySelector(`#${elId}`);
      if (el) el.textContent = this.fmt(Date.now() - t0);
    }, 1000);
  }

  private stopTick(): void {
    if (this.tick) clearInterval(this.tick);
  }

  // ── Render ──────────────────────────────────────────────────────────────
  public render(): void {
    const user = this.context.pageContext.user;

    this.domElement.innerHTML = `
      <style>
        .dl-wrap { font-family: 'Segoe UI', sans-serif; max-width: 700px; margin: 0 auto; padding: 24px 16px; color: #1c1a16; }
        .dl-header { text-align: center; margin-bottom: 28px; padding-bottom: 20px; border-bottom: 2px solid #1c1a16; }
        .dl-header h2 { font-size: 22px; font-weight: 700; margin: 0 0 6px; }
        .dl-header p { font-size: 13px; color: #7a7368; margin: 0; }

        .dl-progress { display: flex; align-items: center; margin-bottom: 20px; }
        .dl-ps { display: flex; align-items: center; gap: 6px; flex: 1; }
        .dl-ps:not(:last-child)::after { content: ''; flex: 1; height: 1px; background: #d8d2c8; margin: 0 6px; }
        .dl-dot { width: 26px; height: 26px; border-radius: 50%; border: 2px solid #d8d2c8; display: flex; align-items: center; justify-content: center; font-size: 11px; color: #7a7368; flex-shrink: 0; background: #fff; transition: all 0.2s; }
        .dl-ps-label { font-size: 11px; color: #7a7368; white-space: nowrap; }
        .dl-ps.active .dl-dot { border-color: #8b4513; background: #8b4513; color: #fff; }
        .dl-ps.active .dl-ps-label { color: #8b4513; font-weight: 600; }
        .dl-ps.done .dl-dot { border-color: #2d6a4f; background: #2d6a4f; color: #fff; }
        .dl-ps.done .dl-ps-label { color: #2d6a4f; }

        .dl-timers { display: flex; gap: 8px; margin-bottom: 20px; flex-wrap: wrap; }
        .dl-tc { font-family: 'Courier New', monospace; font-size: 10px; padding: 4px 10px; background: #fff; border: 1px solid #d8d2c8; color: #7a7368; display: flex; align-items: center; gap: 5px; }
        .dl-tc .dot { width: 5px; height: 5px; border-radius: 50%; background: #d8d2c8; }
        .dl-tc.running .dot { background: #c8773a; animation: dlblink 1s infinite; }
        .dl-tc.done .dot { background: #2d6a4f; }
        @keyframes dlblink { 0%,100%{opacity:1} 50%{opacity:0.2} }

        .dl-card { background: #fff; border: 1px solid #d8d2c8; padding: 28px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
        .dl-card-title { font-size: 15px; font-weight: 700; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 1px solid #d8d2c8; display: flex; align-items: center; gap: 8px; }
        .dl-card-title .bar { width: 3px; height: 14px; background: #c9a84c; flex-shrink: 0; }

        .dl-field { margin-bottom: 18px; }
        .dl-field:last-of-type { margin-bottom: 0; }
        .dl-label { display: block; font-size: 11px; font-weight: 600; color: #7a7368; margin-bottom: 5px; letter-spacing: 0.07em; text-transform: uppercase; }
        .dl-label .req { color: #8b4513; }
        .dl-input, .dl-select, .dl-textarea { width: 100%; padding: 9px 12px; border: 1px solid #d8d2c8; background: #f7f5f0; font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #1c1a16; outline: none; box-sizing: border-box; transition: border-color 0.15s; }
        .dl-input:focus, .dl-select:focus, .dl-textarea:focus { border-color: #8b4513; background: #fff; }
        .dl-input.err, .dl-select.err, .dl-textarea.err { border-color: #c0392b; background: #fdf5f5; }
        .dl-textarea { resize: vertical; min-height: 80px; line-height: 1.5; }
        .dl-row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
        .dl-val { font-size: 11px; color: #c0392b; margin-top: 4px; display: none; }
        .dl-val.show { display: block; }

        .dl-notice { padding: 11px 14px; font-size: 13px; display: flex; gap: 8px; align-items: flex-start; margin-bottom: 18px; line-height: 1.5; }
        .dl-notice.info { background: #f5ece4; border-left: 3px solid #c8773a; color: #8b4513; }
        .dl-notice.success { background: #eaf3ed; border-left: 3px solid #2d6a4f; color: #2d6a4f; }
        .dl-notice.error { background: #fdf5f5; border-left: 3px solid #c0392b; color: #c0392b; }

        .dl-btn { padding: 10px 22px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600; border: none; cursor: pointer; transition: all 0.15s; letter-spacing: 0.03em; }
        .dl-btn-primary { background: #8b4513; color: #fff; }
        .dl-btn-primary:hover { background: #6d3410; }
        .dl-btn-primary:disabled { background: #d8d2c8; color: #7a7368; cursor: not-allowed; }
        .dl-btn-outline { background: transparent; color: #1c1a16; border: 1px solid #b0a898; }
        .dl-btn-outline:hover { border-color: #1c1a16; }
        .dl-btn-success { background: #2d6a4f; color: #fff; }
        .dl-btn-success:hover { background: #1e4d38; }
        .dl-btn-row { display: flex; justify-content: space-between; align-items: center; margin-top: 24px; gap: 10px; flex-wrap: wrap; }

        .dl-rtable { width: 100%; border-collapse: collapse; }
        .dl-rtable tr { border-bottom: 1px solid #d8d2c8; }
        .dl-rtable tr:last-child { border-bottom: none; }
        .dl-rtable td { padding: 9px 0; font-size: 13px; vertical-align: top; }
        .dl-rtable td:first-child { color: #7a7368; font-size: 11px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; width: 36%; padding-right: 14px; padding-top: 11px; }

        .dl-code-wrap { text-align: center; margin: 18px 0; }
        .dl-code-eye { font-family: 'Courier New', monospace; font-size: 10px; letter-spacing: 0.16em; text-transform: uppercase; color: #7a7368; margin-bottom: 8px; }
        .dl-code { display: inline-block; background: #1c1a16; color: #f7f3e8; font-family: 'Courier New', monospace; font-size: 30px; font-weight: 700; letter-spacing: 0.2em; padding: 18px 32px; border: 3px solid #c9a84c; }

        .dl-time-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 18px; }
        .dl-tbox { border: 1px solid #d8d2c8; padding: 12px; text-align: center; background: #f7f5f0; }
        .dl-tbox .tbl { font-family: 'Courier New', monospace; font-size: 9px; letter-spacing: 0.1em; text-transform: uppercase; color: #7a7368; margin-bottom: 5px; }
        .dl-tbox .tbv { font-family: 'Courier New', monospace; font-size: 17px; font-weight: 700; color: #2d6a4f; }

        .dl-spinner { width: 16px; height: 16px; border: 2px solid #d8d2c8; border-top-color: #8b4513; border-radius: 50%; animation: dlspin 0.75s linear infinite; flex-shrink: 0; }
        @keyframes dlspin { to { transform: rotate(360deg); } }
        .dl-status-row { display: flex; align-items: center; gap: 12px; font-family: 'Courier New', monospace; font-size: 12px; color: #7a7368; padding: 10px 0; }

        .dl-screen { display: none; }
        .dl-screen.active { display: block; }

        .dl-attach-area { border: 1px dashed #b0a898; padding: 14px; text-align: center; background: #f7f5f0; cursor: pointer; }
        .dl-attach-area:hover { border-color: #8b4513; }
        .dl-attach-label { font-size: 13px; color: #7a7368; }
        .dl-attach-label span { color: #8b4513; font-weight: 600; text-decoration: underline; }
        .dl-attach-list { margin-top: 8px; }
        .dl-attach-item { font-size: 12px; font-family: 'Courier New', monospace; padding: 4px 0; border-bottom: 1px solid #d8d2c8; display: flex; justify-content: space-between; align-items: center; }
        .dl-attach-item button { background: none; border: none; cursor: pointer; color: #7a7368; font-size: 11px; }
        .dl-attach-item button:hover { color: #c0392b; }
      </style>

      <div class="dl-wrap">
        <div class="dl-header">
          <h2>Document Log</h2>
          <p>Logged in as <strong>${user.displayName}</strong> &nbsp;·&nbsp; ${user.email}</p>
        </div>

        <div class="dl-progress" id="dl-progress">
          <div class="dl-ps active" id="dl-ps-1"><div class="dl-dot">1</div><span class="dl-ps-label">Details</span></div>
          <div class="dl-ps" id="dl-ps-2"><div class="dl-dot">2</div><span class="dl-ps-label">Review</span></div>
          <div class="dl-ps" id="dl-ps-3"><div class="dl-dot">3</div><span class="dl-ps-label">Log Code</span></div>
        </div>

        <div class="dl-timers">
          <div class="dl-tc running" id="dl-tc1"><span class="dot"></span>Filling — <span id="dl-t1">0s</span></div>
          <div class="dl-tc" id="dl-tc2"><span class="dot"></span>Review — <span id="dl-t2">—</span></div>
          <div class="dl-tc" id="dl-tc3"><span class="dot"></span>Processing — <span id="dl-t3">—</span></div>
        </div>

        <!-- Screen 1: Form -->
        <div class="dl-screen active" id="dl-s1">
          <div class="dl-card">
            <div class="dl-card-title"><span class="bar"></span>Document Information</div>
            <div class="dl-field">
              <label class="dl-label">Document Title <span class="req">*</span></label>
              <input id="dl-title" type="text" class="dl-input" placeholder="e.g. Re: Administrative Matter No. 25-01-001" />
              <div class="dl-val" id="dl-v-title">Document title is required.</div>
            </div>
            <div class="dl-row2">
              <div class="dl-field">
                <label class="dl-label">Document Type <span class="req">*</span></label>
                <select id="dl-type" class="dl-select">
                  <option value="">— Select —</option>
                  <option value="RCM">RCM</option>
                  <option value="MISC">MISC</option>
                </select>
                <div class="dl-val" id="dl-v-type">Please select a document type.</div>
              </div>
              <div class="dl-field">
                <label class="dl-label">Document Format <span class="req">*</span></label>
                <select id="dl-format" class="dl-select">
                  <option value="">— Select —</option>
                  <option value="HC">Physical (HC)</option>
                  <option value="SC">Digital (SC)</option>
                </select>
                <div class="dl-val" id="dl-v-format">Please select a document format.</div>
              </div>
            </div>
            <div class="dl-row2">
              <div class="dl-field">
                <label class="dl-label">From <span class="req">*</span></label>
                <input id="dl-from" type="text" class="dl-input" placeholder="Sender name or office" />
                <div class="dl-val" id="dl-v-from">Please indicate the sender.</div>
              </div>
              <div class="dl-field">
                <label class="dl-label">Status <span class="req">*</span></label>
                <select id="dl-status" class="dl-select">
                  <option value="">— Select —</option>
                  <option>Received</option>
                  <option>For Internal Review</option>
                  <option>DCA Review</option>
                  <option>Released</option>
                  <option>Filed</option>
                </select>
                <div class="dl-val" id="dl-v-status">Please select a status.</div>
              </div>
            </div>
            <div class="dl-field">
              <label class="dl-label">Other Remarks</label>
              <textarea id="dl-remarks" class="dl-textarea" placeholder="Optional notes..."></textarea>
            </div>
            <div class="dl-field">
              <label class="dl-label">Attachments <span style="color:#7a7368;font-weight:400;text-transform:none;letter-spacing:0;">(optional)</span></label>
              <div class="dl-attach-area" id="dl-attach-area">
                <div class="dl-attach-label">Drop files here or <span>browse</span></div>
              </div>
              <div class="dl-attach-list" id="dl-attach-list"></div>
            </div>
            <div class="dl-btn-row">
              <span style="font-size:11px;color:#7a7368;font-family:'Courier New',monospace;">Fields marked <span style="color:#8b4513">*</span> are required</span>
              <button id="dl-btn-review" class="dl-btn dl-btn-primary">Review Entry →</button>
            </div>
          </div>
        </div>

        <!-- Screen 2: Review -->
        <div class="dl-screen" id="dl-s2">
          <div class="dl-card">
            <div class="dl-card-title"><span class="bar"></span>Review Before Submitting</div>
            <div class="dl-notice info">ℹ️ &nbsp;Please verify all details before submitting.</div>
            <table class="dl-rtable" id="dl-review-table"></table>
            <div class="dl-btn-row">
              <button id="dl-btn-back" class="dl-btn dl-btn-outline">← Edit Details</button>
              <button id="dl-btn-submit" class="dl-btn dl-btn-primary">Submit & Generate Code →</button>
            </div>
          </div>
        </div>

        <!-- Screen 3: Result -->
        <div class="dl-screen" id="dl-s3">
          <div class="dl-card" id="dl-loading">
            <div class="dl-card-title"><span class="bar"></span>Processing</div>
            <div class="dl-notice info">Submitting to SharePoint and waiting for the log code...</div>
            <div class="dl-status-row">
              <div class="dl-spinner"></div>
              <span id="dl-status-msg">Writing entry to document list...</span>
            </div>
          </div>
          <div class="dl-card" id="dl-success" style="display:none;">
            <div class="dl-notice success">✓ &nbsp;Document logged successfully. Reference code generated.</div>
            <div class="dl-code-wrap">
              <div class="dl-code-eye">Document Reference Code</div>
              <div class="dl-code" id="dl-code">—</div>
            </div>
            <div style="margin-top:24px;">
              <div class="dl-card-title"><span class="bar"></span>Log Summary</div>
              <table class="dl-rtable" id="dl-final-table"></table>
            </div>
            <div class="dl-time-grid">
              <div class="dl-tbox"><div class="tbl">Filling</div><div class="tbv" id="dl-ts1">—</div></div>
              <div class="dl-tbox"><div class="tbl">Review</div><div class="tbv" id="dl-ts2">—</div></div>
              <div class="dl-tbox"><div class="tbl">Processing</div><div class="tbv" id="dl-ts3">—</div></div>
            </div>
            <div class="dl-btn-row" style="margin-top:20px;">
              <button id="dl-btn-new" class="dl-btn dl-btn-outline">+ New Entry</button>
              <div style="display:flex;gap:10px;">
                <button id="dl-btn-copy" class="dl-btn dl-btn-outline">Copy Code</button>
                <button id="dl-btn-print" class="dl-btn dl-btn-success">🖨 Print Routing Slip</button>
              </div>
            </div>
          </div>
          <div class="dl-card" id="dl-error" style="display:none;">
            <div class="dl-notice error">✕ &nbsp;Something went wrong. The entry was saved but the code could not be retrieved. Please check your SharePoint list directly.</div>
            <div class="dl-btn-row">
              <button id="dl-btn-retry-new" class="dl-btn dl-btn-outline">Start New Entry</button>
            </div>
          </div>
        </div>

      </div>
    `;

    this._bindEvents();
    this.timers = { s1Start: Date.now() };
    this.startTick('dl-t1', this.timers.s1Start);
  }

  // ── Bind all events ──────────────────────────────────────────────────────
  private _bindEvents(): void {
    let attachments: File[] = [];

    const attachArea = this.domElement.querySelector('#dl-attach-area') as HTMLElement;
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.multiple = true;
    fileInput.style.display = 'none';
    attachArea.appendChild(fileInput);
    attachArea.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      Array.from(fileInput.files || []).forEach((f: File) => {
        if (!attachments.find((a: File) => a.name === f.name)) attachments.push(f);
      });
      renderAttachments();
    });

    const renderAttachments = (): void => {
      const list = this.domElement.querySelector('#dl-attach-list') as HTMLElement;
      list.innerHTML = attachments.map((f: File) =>
        `<div class="dl-attach-item">
          <span>📎 ${f.name}</span>
          <button data-name="${f.name}">✕ remove</button>
        </div>`
      ).join('');
      list.querySelectorAll('button').forEach(btn => {
        btn.addEventListener('click', () => {
          attachments = attachments.filter((a: File) => a.name !== btn.getAttribute('data-name'));
          renderAttachments();
        });
      });
    };

    this.domElement.querySelector('#dl-btn-review')!.addEventListener('click', () => {
      if (!this._validate()) return;
      const fd = this._getFormData();

      this.timers.s1End = Date.now();
      this.stopTick();
      this._el('dl-t1').textContent = this.fmt(this.timers.s1End - this.timers.s1Start);
      this._tc('dl-tc1', 'done');

      this.timers.s2Start = Date.now();
      this.startTick('dl-t2', this.timers.s2Start);
      this._tc('dl-tc2', 'running');

      const rows = [
        ['Logged By', this.context.pageContext.user.displayName],
        ['Document Title', fd.title],
        ['Document Type', fd.type],
        ['Document Format', fd.format === 'HC' ? 'Physical (HC)' : 'Digital (SC)'],
        ['From', fd.from],
        ['Status', fd.status],
        ['Remarks', fd.remarks || '—'],
        ['Attachments', attachments.length ? attachments.map((a: File) => a.name).join(', ') : 'None'],
      ];
      this._el('dl-review-table').innerHTML =
        rows.map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('');

      this._setProgress(2);
      this._showScreen(2);
    });

    this.domElement.querySelector('#dl-btn-back')!.addEventListener('click', () => {
      this.timers.s1End = undefined as unknown as number;
      this.startTick('dl-t1', this.timers.s1Start);
      this._tc('dl-tc1', 'running');
      this._tc('dl-tc2', '');
      this._el('dl-t2').textContent = '—';
      this.timers.s2Start = undefined as unknown as number;
      this._setProgress(1);
      this._showScreen(1);
    });

    this.domElement.querySelector('#dl-btn-submit')!.addEventListener('click', () => {
      const fd = this._getFormData();

      this.timers.s2End = Date.now();
      this.stopTick();
      this._el('dl-t2').textContent = this.fmt(this.timers.s2End - this.timers.s2Start);
      this._tc('dl-tc2', 'done');

      this.timers.s3Start = Date.now();
      this.startTick('dl-t3', this.timers.s3Start);
      this._tc('dl-tc3', 'running');

      this._setProgress(3);
      this._showScreen(3);

      this._submitToSharePoint(fd, attachments);
    });

    this.domElement.querySelector('#dl-btn-new')!.addEventListener('click', () => this._reset());
    this.domElement.querySelector('#dl-btn-retry-new')?.addEventListener('click', () => this._reset());

    this.domElement.querySelector('#dl-btn-copy')!.addEventListener('click', () => {
      const code = this._el('dl-code').textContent || '';
      navigator.clipboard.writeText(code)
        .then(() => alert('Copied: ' + code))
        .catch(() => {
          window.prompt('Could not copy automatically. Copy the code below:', code);
        });
    });

    this.domElement.querySelector('#dl-btn-print')!.addEventListener('click', () => {
      const code = this._el('dl-code').textContent || '';
      const fd = this._getFormData();
      this._printSlip(code, fd);
    });
  }

  // ── Print routing slip ───────────────────────────────────────────────────
  private _printSlip(code: string, fd: Record<string, string>): void {
    const date = new Date().toLocaleDateString('en-PH', { year: 'numeric', month: 'long', day: 'numeric' });
    const loggedBy = this.context.pageContext.user.displayName;

    const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Routing Slip - ${code}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 11px; color: #000; padding: 20px 28px; max-width: 800px; margin: 0 auto; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px; }
  .header-left h1 { font-size: 22px; font-weight: 700; }
  .header-left .sub { font-size: 10px; color: #666; margin-top: 2px; }
  .header-right { text-align: right; }
  .header-right .label { font-size: 10px; color: #555; }
  .ref { font-family: 'Courier New', monospace; font-size: 20px; font-weight: 700; letter-spacing: 0.1em; border: 2px solid #c9a84c; background: #1c1a16; color: #f7f3e8; padding: 6px 14px; display: inline-block; margin-top: 3px; }
  hr { border: none; border-top: 2px solid #000; margin: 8px 0; }
  .meta { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
  .meta td { padding: 3px 6px 3px 0; vertical-align: top; }
  .meta td:first-child { font-weight: 700; white-space: nowrap; width: 130px; }
  table.wf { width: 100%; border-collapse: collapse; font-size: 10px; }
  table.wf th { background: #d9d9d9; border: 1px solid #000; padding: 5px 6px; text-align: left; font-size: 10px; }
  table.wf td { border: 1px solid #000; padding: 5px 6px; vertical-align: middle; }
  table.wf td.chk { text-align: center; width: 22px; font-size: 13px; }
  table.wf tr.sec td { background: #d9d9d9; font-weight: 700; }
  .remarks-box { border: 1px solid #000; padding: 8px; min-height: 48px; margin-top: 4px; }
  .footer { margin-top: 10px; font-size: 10px; color: #555; display: flex; justify-content: space-between; }
  @media print {
    body { padding: 10px 16px; }
    @page { margin: 1cm; size: A4; }
  }
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <h1>ROUTING SLIP</h1>
    <div class="sub">V1 - Feb 2026 &nbsp;·&nbsp; Supreme Court of the Philippines</div>
  </div>
  <div class="header-right">
    <div class="label">Document Control No.</div>
    <div class="ref">${code}</div>
  </div>
</div>
<hr/>
<table class="meta">
  <tr><td>Document Title:</td><td>${fd.title}</td></tr>
  <tr>
    <td>Document Type:</td>
    <td>${fd.type} &nbsp;&nbsp; <strong>Format:</strong> &nbsp;${fd.format === 'HC' ? 'Physical (HC)' : 'Digital (SC)'}</td>
  </tr>
  <tr><td>From:</td><td>${fd.from}</td></tr>
  <tr><td>Status:</td><td>${fd.status}</td></tr>
  <tr><td>Date Received:</td><td>${date}</td></tr>
  ${fd.remarks ? `<tr><td>Remarks:</td><td>${fd.remarks}</td></tr>` : ''}
</table>

<table class="wf">
  <tr>
    <th style="width:22px;"></th>
    <th style="width:28%;">Status</th>
    <th style="width:28%;">Handled By</th>
    <th style="width:18%;">Date</th>
    <th>Remarks</th>
  </tr>
  <tr><td class="chk">☐</td><td>Received</td><td>Ira | Beau</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>Tracking - Create Slip</td><td>Nante | Dyane</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>Sorting</td><td>Roda | Jim</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>For Internal Review</td><td>Atty. Unis | Atty. Kerr | Atty. Jen</td><td></td><td></td></tr>
  <tr class="sec"><td colspan="5">For Memos</td></tr>
  <tr><td class="chk">☐</td><td>Drafting Memo</td><td>Atty. Amy | Atty. Unis | Atty. Kerr | Atty. Jen</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>For DCA Review</td><td>DCA</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>For Revisions</td><td>Atty. Amy | Atty. Unis | Atty. Kerr | Atty. Jen</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>Tracking - DCA Review</td><td>Nante</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>For Printing</td><td>Roda</td><td></td><td></td></tr>
  <tr class="sec"><td colspan="5">Approval and Release</td></tr>
  <tr><td class="chk">☐</td><td>For DCA Signature</td><td>DCA</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>Tracking - Released</td><td>Nante</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>Sorting</td><td>Roda</td><td></td><td></td></tr>
  <tr class="sec"><td colspan="5">For Physical/Hardcopy</td></tr>
  <tr><td class="chk">☐</td><td>Log Outgoing</td><td>Roda</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>For Dispatch</td><td>Irish | Jim</td><td></td><td></td></tr>
  <tr class="sec"><td colspan="5">For Email</td></tr>
  <tr><td class="chk">☐</td><td>Scanning Document</td><td>Beau | Dyane</td><td></td><td></td></tr>
  <tr><td class="chk">☐</td><td>Tracking - Filed</td><td>Jim</td><td></td><td></td></tr>
</table>

<div style="margin-top:8px;">
  <strong>Other Remarks:</strong>
  <div class="remarks-box">${fd.remarks || ''}</div>
</div>

<div class="footer">
  <span>Logged by: <strong>${loggedBy}</strong></span>
  <span>${date}</span>
</div>

<script>window.onload = function(){ window.print(); }<\/script>
</body>
</html>`;

    const win = window.open('', '_blank');
    if (win) {
      win.document.write(html);
      win.document.close();
    }
  }

  // ── SharePoint: create list item ─────────────────────────────────────────
  private async _submitToSharePoint(fd: Record<string, string>, attachments: File[]): Promise<void> {
    this._setStatus('Writing entry to SharePoint list...');

    try {
      const body = JSON.stringify({
        Title: fd.title,
        Document_x0020_Type: fd.type,
        Document_x0020_Format: fd.format,
        From: fd.from,
        Status: fd.status,
        OtherRemarks: fd.remarks,
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`SharePoint ${response.status}: ${errText}`);
      }

      const data = await response.json();
      const itemId: number = data.Id;

      let failedAttachments: string[] = [];
      if (attachments.length > 0) {
        this._setStatus('Uploading attachments...');
        failedAttachments = await this._uploadAttachments(itemId, attachments);
      }

      if (failedAttachments.length > 0) {
        this._showAttachmentWarning(failedAttachments);
      }

      this._setStatus('Waiting for reference code to be generated...');
      this._pollForCode(itemId);

    } catch (err) {
      console.error('DocumentLog submit error:', err);
      this._showError('The entry could not be saved to SharePoint. Please try again or contact your administrator.');
    }
  }

  // ── SharePoint: upload attachments ───────────────────────────────────────
  private async _uploadAttachments(itemId: number, files: File[]): Promise<string[]> {
    const failed: string[] = [];
    for (const file of files) {
      try {
        const buffer = await file.arrayBuffer();
        const response = await this.context.spHttpClient.post(
          `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            },
            body: buffer
          }
        );
        if (!response.ok) {
          console.error(`Attachment upload failed for "${file.name}": ${response.status}`);
          failed.push(file.name);
        }
      } catch (err) {
        console.error(`Attachment upload error for "${file.name}":`, err);
        failed.push(file.name);
      }
    }
    return failed;
  }

  // ── Poll for code written back by Power Automate ─────────────────────────
  private _pollForCode(itemId: number): void {
    let attempts = 0;
    let consecutiveErrors = 0;
    const maxAttempts = 20;
    const maxConsecutiveErrors = 3;

    this.pollInterval = setInterval(async () => {
      attempts++;
      if (attempts > maxAttempts) {
        clearInterval(this.pollInterval!);
        this._showError('The entry was saved but the reference code was not generated in time. Please check the SharePoint list directly.');
        return;
      }

      try {
        const response = await this.context.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})?$select=ReferenceCode`,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          throw new Error(`SharePoint returned ${response.status}`);
        }

        const data = await response.json();
        consecutiveErrors = 0;

        if (data.ReferenceCode && data.ReferenceCode.trim() !== '') {
          clearInterval(this.pollInterval!);
          this._showSuccess(data.ReferenceCode);
        } else {
          this._setStatus(`Generating code... (${attempts}/${maxAttempts})`);
        }
      } catch (err) {
        console.error('Poll error:', err);
        consecutiveErrors++;
        if (consecutiveErrors >= maxConsecutiveErrors) {
          clearInterval(this.pollInterval!);
          this._showError('The entry was saved but the connection was lost while waiting for the reference code. Please check the SharePoint list directly.');
        }
      }
    }, 2000);
  }

  // ── Show success state ───────────────────────────────────────────────────
  private _showSuccess(code: string): void {
    this.timers.s3End = Date.now();
    this.stopTick();
    this._el('dl-t3').textContent = this.fmt(this.timers.s3End - this.timers.s3Start);
    this._tc('dl-tc3', 'done');
    this._setProgress('done');

    this._el('dl-code').textContent = code;

    const fd = this._getFormData();
    const rows = [
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
      rows.map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('');

    this._el('dl-ts1').textContent = this.fmt(this.timers.s1End - this.timers.s1Start);
    this._el('dl-ts2').textContent = this.fmt(this.timers.s2End - this.timers.s2Start);
    this._el('dl-ts3').textContent = this.fmt(this.timers.s3End - this.timers.s3Start);

    this._el('dl-loading').style.display = 'none';
    this._el('dl-success').style.display = 'block';
  }

  private _showError(message?: string): void {
    this.stopTick();
    this._tc('dl-tc3', '');
    if (message) {
      const notice = this._el('dl-error').querySelector('.dl-notice');
      if (notice) notice.innerHTML = `✕ &nbsp;${message}`;
    }
    this._el('dl-loading').style.display = 'none';
    this._el('dl-error').style.display = 'block';
  }

  private _showAttachmentWarning(failedFiles: string[]): void {
    const names = failedFiles.join(', ');
    const warning = document.createElement('div');
    warning.className = 'dl-notice info';
    warning.innerHTML = `⚠ &nbsp;The entry was saved, but these attachments failed to upload: <strong>${names}</strong>. You can add them manually from the SharePoint list.`;
    const loading = this._el('dl-loading');
    loading.parentElement!.insertBefore(warning, loading);
  }

  // ── Validation ───────────────────────────────────────────────────────────
  private _validate(): boolean {
    const checks = [
      { id: 'dl-title',  vid: 'dl-v-title'  },
      { id: 'dl-type',   vid: 'dl-v-type'   },
      { id: 'dl-format', vid: 'dl-v-format' },
      { id: 'dl-from',   vid: 'dl-v-from'   },
      { id: 'dl-status', vid: 'dl-v-status' },
    ];
    let ok = true;
    checks.forEach(({ id, vid }) => {
      const el = this.domElement.querySelector(`#${id}`) as HTMLInputElement;
      const msg = this.domElement.querySelector(`#${vid}`) as HTMLElement;
      if (!el.value.trim()) {
        el.classList.add('err');
        msg.classList.add('show');
        ok = false;
      } else {
        el.classList.remove('err');
        msg.classList.remove('show');
      }
    });
    return ok;
  }

  // ── Helpers ──────────────────────────────────────────────────────────────
  private _getFormData(): Record<string, string> {
    return {
      title:   (this.domElement.querySelector('#dl-title')   as HTMLInputElement).value.trim(),
      type:    (this.domElement.querySelector('#dl-type')    as HTMLSelectElement).value,
      format:  (this.domElement.querySelector('#dl-format')  as HTMLSelectElement).value,
      from:    (this.domElement.querySelector('#dl-from')    as HTMLInputElement).value.trim(),
      status:  (this.domElement.querySelector('#dl-status')  as HTMLSelectElement).value,
      remarks: (this.domElement.querySelector('#dl-remarks') as HTMLTextAreaElement).value.trim(),
    };
  }

  private _el(id: string): HTMLElement {
    return this.domElement.querySelector(`#${id}`) as HTMLElement;
  }

  private _tc(id: string, state: 'running' | 'done' | ''): void {
    const el = this._el(id);
    el.className = `dl-tc${state ? ' ' + state : ''}`;
  }

  private _setStatus(msg: string): void {
    const el = this.domElement.querySelector('#dl-status-msg');
    if (el) el.textContent = msg;
  }

  private _showScreen(n: number): void {
    this.domElement.querySelectorAll('.dl-screen').forEach(s => s.classList.remove('active'));
    this._el(`dl-s${n}`).classList.add('active');
  }

  private _setProgress(step: number | 'done'): void {
    if (step === 'done') {
      [1, 2, 3].forEach(i => { this._el(`dl-ps-${i}`).className = 'dl-ps done'; });
      return;
    }
    for (let i = 1; i <= 3; i++) {
      const el = this._el(`dl-ps-${i}`);
      if (i < step) el.className = 'dl-ps done';
      else if (i === step) el.className = 'dl-ps active';
      else el.className = 'dl-ps';
    }
  }

  private _reset(): void {
    ['dl-title', 'dl-from', 'dl-remarks'].forEach(id =>
      (this.domElement.querySelector(`#${id}`) as HTMLInputElement).value = '');
    ['dl-type', 'dl-format', 'dl-status'].forEach(id =>
      (this.domElement.querySelector(`#${id}`) as HTMLSelectElement).value = '');
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
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }
}