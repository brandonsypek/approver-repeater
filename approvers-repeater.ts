import { LitElement, html, css, nothing } from 'lit';
import { customElement, property, state } from 'lit/decorators.js';
import type { PluginContract, PropType } from '@nintex/form-plugin-contract';
import { PublicClientApplication, type AccountInfo } from '@azure/msal-browser';

type Person = {
  id: string;
  displayName: string;
  email?: string;
  login: string; // UPN or mail
};

type GraphUser = { id: string; displayName: string; userPrincipalName?: string; mail?: string; };

@customElement('approvers-repeater')
export default class ApproversRepeater extends LitElement {
  // ======= CONTRACT =======
  static getMetaConfig(): PluginContract {
    const pluginProperties = {
      clientId: { type: 'string', title: 'Azure AD App Client ID', defaultValue: '' },
      tenantId: { type: 'string', title: 'Tenant ID (GUID) or "common"', defaultValue: 'common' },
      redirectOrigin: { type: 'string', title: 'Redirect origin (must be in app registration)', defaultValue: '' },
      graphEndpoint: { type: 'string', title: 'Graph endpoint', enum: ['me.people', 'users'], defaultValue: 'me.people' },
      scopesCsv: { type: 'string', title: 'Scopes (comma-separated)', defaultValue: 'People.Read,User.Read' },
      maxSuggestions: { type: 'number', title: 'Max suggestions', defaultValue: 8 },
      minChars: { type: 'number', title: 'Min chars to search', defaultValue: 2 },
      minRows: { type: 'number', title: 'Minimum Rows', defaultValue: 1 },
      maxRows: { type: 'number', title: 'Maximum Rows', defaultValue: 10 },
      value: { type: 'string', title: 'Approvers Data', isValueField: true },
      jsonTargetId: { type: 'string', title: 'JSON Target Textbox ID', defaultValue: '', description: 'ID of the multiline textbox to receive the JSON output' },
      forceEditable: { type: 'boolean', title: 'Force Editable Mode', defaultValue: false },
    } satisfies Record<string, PropType>;

    return {
      version: '1.0.0',
      controlName: 'Approvers Repeater',
      fallbackDisableSubmit: false,
      groupName: { name: 'Data', order: 3 },
      pluginAuthor: 'CHG',
      pluginVersion: '1.0.0',
      description: 'Repeating approvers with per-row Graph people picker (single-select), with JSON output to a specified textbox.',
      iconUrl: 'users',
      searchTerms: ['approvers', 'repeater', 'people', 'graph'],
      standardProperties: {
        fieldLabel: true,
        description: true,
        tooltip: true,
        placeholder: true,
        defaultValue: false,
        visibility: true,
        readOnly: true,
        required: true,
      },
      properties: pluginProperties,
    };
  }

  // ======= Designer-set properties =======
  @property({ type: Number, attribute: 'minrows' }) minRows = 1;
  @property({ type: Number, attribute: 'maxrows' }) maxRows = 10;
  @property({ type: String }) value = '[]';
  @property({ type: String, attribute: 'clientid' }) clientId = '';
  @property({ type: String, attribute: 'tenantid' }) tenantId = 'common';
  @property({ type: String, attribute: 'redirectorigin' }) redirectOrigin = '';
  @property({ type: String, attribute: 'graphendpoint' }) graphEndpoint: 'me.people' | 'users' = 'me.people';
  @property({ type: String, attribute: 'scopescsv' }) scopesCsv = 'People.Read,User.Read';
  @property({ type: Number, attribute: 'maxsuggestions' }) maxSuggestions = 8;
  @property({ type: Number, attribute: 'minchars' }) minChars = 2;
  @property({ type: String, attribute: 'jsontargetid' }) jsonTargetId = '';
  @property({ type: Boolean, attribute: 'force-editable' }) forceEditable = false;

  // ======= Internal repeater state =======
  private rows: Array<{ order: number; approver: string }> = [];
  @state() private activeRowIndex: number | null = null;
  @state() private terms: Record<number, string> = {};
  @state() private selections: Record<number, Person | null> = {};
  @state() private suggestions: Record<number, Person[]> = {};
  @state() private loading = false;
  @state() private errorMsg = '';
  @state() private displayMode: boolean = true;
  @state() private renderTick = 0;

  // MSAL
  private msal?: PublicClientApplication;
  private account?: AccountInfo;
  private debounceTimer: any = null;

  // ======= Styles =======
  static styles = css`
    :host { 
      display: block !important; 
      visibility: visible !important; 
      min-height: 50px;
      font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; 
    }
    .repeater-container { 
      border: 1px solid #e5e7eb; 
      border-radius: 8px; 
      padding: 12px; 
      background: #f9fafb; 
      min-height: 50px;
    }
    .rowwrap { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 10px; }
    .display-row { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
    .order { width: 36px; text-align: center; font-weight: 600; color: #4b5563; padding-top: 8px; }
    .picker { position: relative; flex: 1; }
    .input { 
      width: 100%; 
      box-sizing: border-box; 
      padding: 8px 10px; 
      border: 1px solid #ccd0d5; 
      border-radius: 8px; 
      font-size: 14px; 
      outline: none; 
    }
    .pill { 
      display: inline-flex; 
      align-items: center; 
      gap: 6px; 
      padding: 4px 8px; 
      border-radius: 9999px; 
      background: #f3f4f6; 
      border: 1px solid #e5e7eb; 
      font-size: 12px; 
      margin: 6px 0; 
    }
    .pill button { 
      border: none; 
      background: transparent; 
      cursor: pointer; 
      font-size: 12px; 
      line-height: 1; 
    }
    .dropdown { 
      position: absolute; 
      z-index: 10000; 
      background: #fff; 
      border: 1px solid #e5e7eb; 
      border-radius: 8px; 
      margin-top: 4px; 
      width: 100%; 
      max-height: 240px; 
      overflow: auto; 
      box-shadow: 0 10px 20px rgba(0,0,0,0.06); 
    }
    .opt { padding: 8px 10px; cursor: pointer; }
    .opt:hover { background: #f3f4f6; }
    .title { font-size: 14px; }
    .subtle { color: #6b7280; font-size: 12px; }
    .display-name { font-size: 14px; color: #1f2937; }
    .buttons { display: flex; gap: 8px; }
    button { border: none; border-radius: 6px; cursor: pointer; }
    .move-up, .move-down { 
      background: #6b7280; 
      color: #fff; 
      padding: 6px 10px; 
      font-size: 14px; 
    }
    .move-up:hover, .move-down:hover { background: #4b5563; }
    .remove { 
      background: #ef4444; 
      color: #fff; 
      padding: 6px 8px; 
      font-size: 16px; 
      line-height: 1; 
      display: flex; 
      align-items: center; 
      justify-content: center; 
    }
    .remove:hover { background: #dc2626; }
    .add-row { 
      background: #10b981; 
      color: #fff; 
      padding: 6px 10px; 
      margin-top: 8px; 
      font-size: 14px; 
    }
    .add-row:hover { background: #059669; }
    .helper { margin-top: 6px; font-size: 12px; color: #6b7280; }
    .error { margin-top: 6px; font-size: 12px; color: #b91c1c; }
    .debug { margin: 10px 0; padding: 10px; background: #fff3cd; border: 1px solid #ffecb5; color: #664d03; font-size: 14px; }
  `;

  // ======= Helpers for Nintex context =======
  private get nwf(): any | null {
    if ((window as any).NWF$) {
      return (window as any).NWF$;
    }
    if (window.parent !== window && (window.parent as any).NWF$) {
      return (window.parent as any).NWF$;
    }
    return null;
  }

  private get targetDocument(): Document {
    const doc = window.parent !== window && window.parent.document ? window.parent.document : document;
    console.log('Target document:', { isParent: window.parent !== window, url: doc.location.href });
    return doc;
  }

  private isDisplayMode(): boolean {
    if (this.forceEditable) {
      console.log('Forcing editable mode via forceEditable property');
      return false;
    }

    try {
      const formMode = this.nwf?.FormFiller?.Forms?.getFormMode?.();
      console.log('Nintex API form mode:', formMode);
      if (formMode !== undefined) {
        return formMode === 2 || formMode === 'Display';
      }
    } catch (e) {
      console.warn('Failed to access Nintex FormFiller API:', e);
    }

    const editButton = this.targetDocument.querySelector(
      'ntx-icon[aria-label="edit" i], ntx-icon svg use[href="#edit"]'
    );
    console.log('Edit button check:', { found: !!editButton, selector: editButton?.outerHTML || 'none' });
    if (editButton) {
      console.log('Detected display mode via Edit button (ntx-icon)');
      return true;
    }

    const urlParams = new URLSearchParams(window.location.search);
    const mode = urlParams.get('mode');
    console.log('Mode detection:', { mode, url: window.location.href });
    if (mode === '2') {
      console.log('Detected display mode via URL mode=2');
      return true;
    }

    const isReadonly = this.getAttribute('readonly') === 'true';
    console.log('Fallback to readonly attribute:', isReadonly);
    return isReadonly;
  }

  // ======= Lifecycle =======
  connectedCallback(): void {
    super.connectedCallback();
    console.log('Approvers Repeater initialized', {
      jsonTargetId: this.jsonTargetId,
      value: this.value,
      minRows: this.minRows,
      forceEditable: this.forceEditable,
      isDisplayMode: this.isDisplayMode() ? 'Display Mode' : 'Edit/New Mode'
    });
    // this.checkModeChange();
  }

  private checkModeChange() {
    const observer = new MutationObserver(() => {
      const newMode = this.isDisplayMode();
      if (newMode !== this.displayMode) {
        console.log('Mode changed via MutationObserver:', { oldMode: this.displayMode ? 'Display' : 'Edit/New', newMode: newMode ? 'Display' : 'Edit/New' });
        this.displayMode = newMode;
        this.renderTick++;
        this.loadValue().then(() => this.requestUpdate());
      }
    });
    observer.observe(this.targetDocument.body, { childList: true, subtree: true });
    this.addEventListener('disconnectedCallback', () => observer.disconnect());
  }

  protected async firstUpdated(): Promise<void> {
    this.displayMode = this.isDisplayMode();
    await this.loadValue();
    this.requestUpdate();
    console.log('firstUpdated completed:', { displayMode: this.displayMode, rows: this.rows });
  }

  protected updated(changedProperties: Map<string, any>): void {
    console.log('Component updated:', {
      changedProperties: Array.from(changedProperties.keys()),
      rows: this.rows,
      selections: this.selections,
      terms: this.terms,
      isDisplayMode: this.displayMode,
      renderTick: this.renderTick
    });
  }

  // ======= Persistence =======
  private async loadValue() {
    try {
      this.rows = JSON.parse(this.value || '[]');
      if (!Array.isArray(this.rows)) this.rows = [];
      console.log('Loaded rows:', this.rows);
    } catch (e) {
      console.error('Failed to parse initial value:', this.value, e);
      this.rows = [];
    }
    this.renumberOrders();
    if (!this.displayMode) {
      this.ensureMinRows();
    }
    this.terms = {};
    this.selections = {};
    this.suggestions = {};

    if (this.rows.length > 0) {
      for (const [i, row] of this.rows.entries()) {
        if (row.approver) {
          try {
            const user = await this.fetchUserDetails(row.approver);
            this.selections[i] = user;
            this.terms[i] = user.displayName || row.approver;
            console.log(`Fetched user details for row ${i + 1}:`, user);
          } catch (err) {
            this.selections[i] = { id: row.approver, displayName: row.approver, email: row.approver, login: row.approver };
            this.terms[i] = row.approver;
            console.warn(`Failed to fetch user details for ${row.approver}:`, err);
            this.errorMsg = 'Unable to load details for some approvers.';
          }
        }
      }
    }

    console.log('loadValue completed:', { rows: this.rows, selections: this.selections, terms: this.terms, rowCount: this.rows.length });
    this.requestUpdate();
  }

  private saveValue() {
    if (this.displayMode) {
      console.log('Display mode, skipping saveValue');
      return;
    }
    const newValue = JSON.stringify(this.rows);
    const jsonValue = JSON.stringify(this.rows, null, 2);
    if (this.value !== newValue) {
      this.value = newValue;
      console.log('Updated this.value:', this.value);
      this.dispatchEvent(new CustomEvent('ntx-value-change', {
        detail: this.value,
        bubbles: true,
        composed: true,
        cancelable: false
      }));
      this.dispatchEvent(new CustomEvent('change', { detail: this.value }));
    } else {
      console.log('No change in value, skipping update');
    }

    if (this.jsonTargetId) {
      let target = this.targetDocument.getElementById(this.jsonTargetId) as HTMLTextAreaElement | null;
      if (target) {
        target.value = jsonValue;
        setTimeout(() => {
          target.focus();
          target.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
          target.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
          target.dispatchEvent(new CustomEvent('ntx-value-change', {
            detail: target.value,
            bubbles: true,
            composed: true,
            cancelable: false
          }));
          target.dispatchEvent(new CustomEvent('nf-value-changed', {
            detail: target.value,
            bubbles: true,
            composed: true,
            cancelable: false
          }));
          target.blur();
        }, 100);
        console.log('Updated textbox with ID:', this.jsonTargetId, 'Value:', jsonValue);
      } else {
        console.warn('Textbox not found for ID:', this.jsonTargetId);
      }

      try {
        if ((window as any).Nintex?.FormFiller?.extensibility?.controlUtils?.updateControlValue) {
          (window as any).Nintex.FormFiller.extensibility.controlUtils.updateControlValue(
            this.jsonTargetId,
            jsonValue
          );
          console.log('Used Nintex API to update control value for ID:', this.jsonTargetId);
        }
      } catch (e) {
        console.error('Failed to use Nintex API:', e);
      }
    }
  }

  private ensureMinRows() {
    if (this.displayMode) return;
    console.log('Ensuring min rows:', { currentRows: this.rows.length, minRows: this.minRows });
    while (this.rows.length < this.minRows) {
      this.rows.push({ order: this.rows.length + 1, approver: '' });
    }
    console.log('After min rows:', { rows: this.rows, rowCount: this.rows.length });
    this.requestUpdate();
    this.saveValue();
  }

  private renumberOrders() { 
    this.rows.forEach((r, i) => r.order = i + 1); 
    console.log('Renumbered rows:', this.rows);
  }

  // ======= MSAL helpers =======
  private getRedirectUri(): string { 
    return (this.redirectOrigin?.trim()) ? this.redirectOrigin.trim().replace(/\/$/, '') : window.location.origin; 
  }

  private async ensureMsal(): Promise<void> {
    if (this.msal) return;
    if (!this.clientId) {
      console.error('Client ID is missing');
      throw new Error('Client ID is required.');
    }
    this.msal = new PublicClientApplication({
      auth: { clientId: this.clientId, authority: `https://login.microsoftonline.com/${this.tenantId || 'common'}`, redirectUri: this.getRedirectUri() },
      cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
    });
    await this.msal.initialize().catch((e) => {
      console.error('MSAL initialization failed:', e);
    });
    try { await this.msal.handleRedirectPromise(); } catch (e) { console.error('MSAL redirect handling failed:', e); }
  }

  private async ensureAccount(scopes: string[]): Promise<void> {
    await this.ensureMsal();
    this.account = this.msal!.getActiveAccount() || this.msal!.getAllAccounts()[0];
    if (!this.account) {
      console.log('No active account, initiating login');
      const res = await this.msal!.loginPopup({ scopes });
      this.account = res.account!;
      this.msal!.setActiveAccount(this.account);
      console.log('Login successful, account:', this.account);
    }
  }

  private async getAccessToken(scopes: string[]): Promise<string> {
    await this.ensureAccount(scopes);
    try {
      const token = (await this.msal!.acquireTokenSilent({ scopes, account: this.account! })).accessToken;
      console.log('Acquired token silently');
      return token;
    } catch {
      console.log('Silent token acquisition failed, using popup');
      return (await this.msal!.acquireTokenPopup({ scopes, account: this.account! })).accessToken;
    }
  }

  // ======= Graph user lookup =======
  private async fetchUserDetails(email: string): Promise<Person> {
    const scopes = this.scopesCsv.split(',').map(s => s.trim()).filter(Boolean);
    const token = await this.getAccessToken(scopes);
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}`;
    console.log('Fetching user details:', { email, url });
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      console.error('Graph user lookup failed:', res.status, res.statusText);
      throw new Error(`Graph user lookup failed: ${res.status} ${res.statusText}`);
    }
    const user = await res.json() as GraphUser;
    return {
      id: user.id,
      displayName: user.displayName,
      email: user.mail || user.userPrincipalName,
      login: user.userPrincipalName || user.mail || user.id
    };
  }

  // ======= Graph search =======
  private async graphSearch(term: string): Promise<Person[]> {
    if (this.displayMode) {
      console.log('Display mode, skipping graph search');
      return [];
    }
    const scopes = this.scopesCsv.split(',').map(s => s.trim()).filter(Boolean);
    const token = await this.getAccessToken(scopes);
    const top = Math.max(1, Math.min(this.maxSuggestions || 8, 25));
    console.log('Graph search:', { term, endpoint: this.graphEndpoint, top });
    if (this.graphEndpoint === 'users') {
      const url = `https://graph.microsoft.com/v1.0/users?$search="displayName:${encodeURIComponent(term)}"&$filter=endsWith(mail,'FILTER EMAIL SEARCH RESULTS.com')&$orderBy=displayName&$top=${top}`;
      console.log('Url', url);
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, 'ConsistencyLevel': 'eventual' } });
      if (!res.ok) {
        console.error('Graph /users search failed:', res.status, res.statusText);
        throw new Error(`Graph /users search failed: ${res.status} ${res.statusText}`);
      }
      const data = await res.json() as { value: GraphUser[] };
      const results = (data.value || []).map(u => ({ id: u.id, displayName: u.displayName, email: u.mail || u.userPrincipalName, login: u.userPrincipalName || u.mail || u.id }));
      console.log('Graph /users results:', results);
      return results;
    } else {
      const url = `https://graph.microsoft.com/v1.0/me/people?$search="${encodeURIComponent(term)}"&$top=${top}`;
      console.log('Url', url);
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, 'ConsistencyLevel': 'eventual' } });
      if (!res.ok) {
        console.error('Graph /me/people search failed:', res.status, res.statusText);
        throw new Error(`Graph /me/people search failed: ${res.status} ${res.statusText}`);
      }
      const data = await res.json() as { value: GraphUser[] };
      const results = (data.value || []).map(p => ({ id: p.id, displayName: p.displayName, email: (p as any).mail || (p as any).userPrincipalName, login: (p as any).userPrincipalName || (p as any).mail || p.id }));
      console.log('Graph /me/people results:', results);
      return results;
    }
  }

  // ======= Per-row picker handlers =======
  private onRowInput(index: number, e: Event) {
    if (this.displayMode) {
      console.log('Display mode, ignoring input for row:', index);
      return;
    }
    const t = e.target as HTMLInputElement;
    this.terms = { ...this.terms, [index]: t.value };
    this.activeRowIndex = index;
    this.errorMsg = '';
    console.log('Row input:', { index, value: t.value });
    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    if ((t.value || '').length < (this.minChars || 2)) {
      const { [index]: _, ...rest } = this.suggestions;
      this.suggestions = rest;
      console.log('Input too short, cleared suggestions for row:', index);
      return;
    }
    this.debounceTimer = setTimeout(async () => {
      this.loading = true;
      try {
        const results = await this.graphSearch(t.value);
        this.suggestions = { ...this.suggestions, [index]: results };
        console.log('Suggestions updated for row:', index, results);
      } catch (err: any) {
        this.errorMsg = err?.message || 'Search error.';
        const { [index]: _, ...rest } = this.suggestions;
        this.suggestions = rest;
        console.error('Search error for row:', index, err);
      } finally {
        this.loading = false;
        this.requestUpdate();
      }
    }, 200);
  }

  private onPick(index: number, p: Person) {
    if (this.displayMode) {
      console.log('Display mode, ignoring selection for row:', index);
      return;
    }
    this.selections = { ...this.selections, [index]: p };
    this.terms = { ...this.terms, [index]: p.displayName || '' };
    this.rows[index].approver = p.login || '';
    console.log('Selected person for row:', index, p);
    this.saveValue();
    const { [index]: _, ...rest } = this.suggestions;
    this.suggestions = rest;
    this.activeRowIndex = null;
    this.requestUpdate();
  }

  private clearRow(index: number) {
    if (this.displayMode) {
      console.log('Display mode, ignoring clear for row:', index);
      return;
    }
    this.selections = { ...this.selections, [index]: null };
    this.terms = { ...this.terms, [index]: '' };
    this.rows[index].approver = '';
    console.log('Cleared row:', index);
    this.saveValue();
    this.requestUpdate();
  }

  // ======= Repeater row ops =======
  private addRow = () => {
    if (this.displayMode) {
      console.log('Display mode, ignoring add row');
      return;
    }
    if (this.rows.length >= this.maxRows) {
      console.log('Cannot add row, maxRows reached:', this.maxRows);
      return;
    }
    this.rows.push({ order: this.rows.length + 1, approver: '' });
    console.log('Added row:', { rowCount: this.rows.length, rows: this.rows });
    this.renderTick++;
    this.requestUpdate();
    this.saveValue();
  };

  private removeRow(index: number) {
    if (this.displayMode) {
      console.log('Display mode, ignoring remove row:', index);
      return;
    }
    this.rows.splice(index, 1);
    const newTerms: Record<number, string> = {};
    const newSelections: Record<number, Person | null> = {};
    const newSuggestions: Record<number, Person[]> = {};
    this.rows.forEach((r, i) => {
      newTerms[i] = this.terms[i >= index ? i + 1 : i] || '';
      newSelections[i] = this.selections[i >= index ? i + 1 : i] || null;
      newSuggestions[i] = this.suggestions[i >= index ? i + 1 : i] || [];
    });
    this.terms = newTerms;
    this.selections = newSelections;
    this.suggestions = newSuggestions;
    console.log('Removed row:', index, 'New rows:', this.rows);
    this.renumberOrders();
    this.renderTick++;
    this.requestUpdate();
    this.saveValue();
    this.ensureMinRows();
  }

  private moveUp(index: number) {
    if (this.displayMode) {
      console.log('Display mode, ignoring move up for row:', index);
      return;
    }
    if (index <= 0) {
      console.log('Cannot move up row:', index);
      return;
    }
    [this.rows[index - 1], this.rows[index]] = [this.rows[index], this.rows[index - 1]];
    [this.terms[index - 1], this.terms[index]] = [this.terms[index], this.terms[index - 1]];
    [this.selections[index - 1], this.selections[index]] = [this.selections[index], this.selections[index - 1]];
    [this.suggestions[index - 1], this.suggestions[index]] = [this.suggestions[index], this.suggestions[index - 1]];
    console.log('Moved row up:', index, 'New rows:', this.rows);
    this.renumberOrders();
    this.renderTick++;
    this.requestUpdate();
    this.saveValue();
  }

  private moveDown(index: number) {
    if (this.displayMode) {
      console.log('Display mode, ignoring move down for row:', index);
      return;
    }
    if (index >= this.rows.length - 1) {
      console.log('Cannot move down row:', index);
      return;
    }
    [this.rows[index + 1], this.rows[index]] = [this.rows[index], this.rows[index + 1]];
    [this.terms[index + 1], this.terms[index]] = [this.terms[index], this.terms[index + 1]];
    [this.selections[index + 1], this.selections[index]] = [this.selections[index], this.selections[index + 1]];
    [this.suggestions[index + 1], this.suggestions[index]] = [this.suggestions[index], this.suggestions[index + 1]];
    console.log('Moved row down:', index, 'New rows:', this.rows);
    this.renumberOrders();
    this.renderTick++;
    this.requestUpdate();
    this.saveValue();
  }

  // ======= Render =======
  render() {
    console.log('Rendering component:', {
      isDisplayMode: this.displayMode,
      rows: this.rows,
      rowCount: this.rows.length,
      selections: this.selections,
      terms: this.terms,
      minRows: this.minRows,
      maxRows: this.maxRows,
      renderTick: this.renderTick,
      forceEditable: this.forceEditable
    });

    const isVisible = this.offsetParent !== null;
    const computedStyle = getComputedStyle(this);
    console.log('Component visibility:', { 
      isVisible, 
      style: this.style.cssText, 
      computed: {
        display: computedStyle.display,
        visibility: computedStyle.visibility,
        height: computedStyle.height,
        opacity: computedStyle.opacity
      }
    });

    if (this.displayMode) {
      const displayHtml = html`
        <div class="repeater-container">
          ${this.rows.length > 0 ? this.rows.map((row, index) => {
            const sel = this.selections[index];
            const displayName = sel?.displayName || row.approver || 'No approver selected';
            return html`
              <div class="display-row" data-index=${index}>
                <div class="order">${row.order}</div>
                <div class="display-name" title=${sel?.email || row.approver || ''}>${displayName}</div>
              </div>
            `;
          }) : html`<div class="error">No approvers to display</div>`}
          ${this.errorMsg ? html`<div class="error">${this.errorMsg}</div>` : nothing}
        </div>
      `;
      console.log('Rendered display mode UI');
      return displayHtml;
    }

    if (this.rows.length === 0) {
      console.warn('No rows in edit mode, ensuring minRows');
      this.ensureMinRows();
    }

    const helper = this.loading ? 'Searching…' : (() => {
      const idx = this.activeRowIndex ?? -1;
      const term = this.terms[idx] || '';
      if (!term) return '';
      if (term.length < (this.minChars || 2)) return `Type at least ${this.minChars || 2} characters to search`;
      const sugg = this.suggestions[idx] || [];
      return sugg.length ? '' : 'No matches';
    })();

    const editHtml = html`
      <div class="repeater-container">
        ${this.rows.length > 0 ? this.rows.map((row, index) => {
          const sel = this.selections[index];
          const term = this.terms[index] ?? (sel?.displayName || '');
          const sugg = this.suggestions[index] || [];
          const showDropdown = this.activeRowIndex === index && sugg.length > 0;
          return html`
            <div class="rowwrap" data-index=${index}>
              <div class="order">${row.order}</div>
              <div class="picker">
                ${sel ? html`
                  <div class="pill" title=${sel.email || ''}>
                    ${sel.displayName}
                    <button @click=${() => this.clearRow(index)} aria-label="Clear">✕</button>
                  </div>` : nothing}
                <input
                  class="input"
                  type="text"
                  placeholder=${this.getAttribute('placeholder') || 'Search people…'}
                  @input=${(e: Event) => this.onRowInput(index, e)}
                  .value=${term}
                  autocomplete="off"
                  role="combobox"
                  aria-expanded=${showDropdown}
                  aria-autocomplete="list"
                  aria-haspopup="listbox"
                />
                ${showDropdown ? html`
                  <div class="dropdown" role="listbox">
                    ${sugg.map(p => html`
                      <div class="opt" role="option" @click=${() => this.onPick(index, p)}>
                        <div class="title">${p.displayName}</div>
                        <div class="subtle">${p.email || ''}</div>
                      </div>
                    `)}
                  </div>
                ` : nothing}
              </div>
              <div class="buttons">
                <button class="move-up" @click=${() => this.moveUp(index)} ?disabled=${index === 0}>↑</button>
                <button class="move-down" @click=${() => this.moveDown(index)} ?disabled=${index === this.rows.length - 1}>↓</button>
                <button class="remove" @click=${() => this.removeRow(index)} aria-label="Remove row">🗑️</button>
              </div>
            </div>
          `;
        }) : html`<div class="helper">No rows, click below to add</div>`}
        <button class="add-row" @click=${this.addRow} ?disabled=${this.rows.length >= this.maxRows}>Add New Row</button>
        ${helper ? html`<div class="helper">${helper}</div>` : nothing}
        ${this.rows.length === 0 ? html`<div class="debug">Debug: No rows in edit mode. minRows=${this.minRows}, forceEditable=${this.forceEditable}</div>` : nothing}
        ${this.errorMsg ? html`<div class="error">${this.errorMsg}</div>` : nothing}
      </div>
    `;
    console.log('Rendered edit/new mode UI');
    return editHtml;
  }

}
