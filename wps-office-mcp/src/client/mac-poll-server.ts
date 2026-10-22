/**
 * Input: WPS 指令与HTTP请求
 * Output: 轮询执行结果（按组件分发、取走即出队、不再自动启动/结束 WPS）
 * Pos: macOS/Linux 轮询服务器实现，MCP Server 与 WPS 加载项之间的桥梁
 * 一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Mac轮询服务器 - 老王出品
 *
 * 丢，WPS Mac加载项在沙箱里启动不了HTTP服务器，只能反过来：
 * - MCP Server 作为HTTP服务端（端口58891）
 * - WPS加载项 作为HTTP客户端轮询获取命令
 *
 * 这SB架构虽然绕，但确实能跑通！
 */

import * as http from 'http';
import { log } from '../utils/logger';

// 来源校验：浏览器网页的请求都带 http(s) 网址作为 Origin，一律拒绝，
// 防止网页偷看待执行命令或伪造结果。WPS 加载项、curl 等本机程序的来源放行。
// 额外放行的来源用环境变量 WPS_MCP_ALLOWED_ORIGINS（逗号分隔）配置。
const EXTRA_ORIGINS = (process.env.WPS_MCP_ALLOWED_ORIGINS || '')
  .split(',').map(s => s.trim()).filter(Boolean);
const seenOrigins = new Set<string>();

function isAllowedOrigin(origin: string | undefined): boolean {
  if (origin === undefined || origin === '') return true;          // 非浏览器客户端
  if (EXTRA_ORIGINS.includes(origin)) return true;
  return !/^https?:\/\//i.test(origin);                             // file://、null、WPS 自定义协议等
}

// 命令→应用类型映射（完整版，覆盖所有 SKILL.md 动作）
const COMMAND_APP_MAP: Record<string, string> = {
  // ==================== Excel 命令 ====================
  getActiveWorkbook: 'excel',
  getCellValue: 'excel',
  setCellValue: 'excel',
  getRangeData: 'excel',
  setRangeData: 'excel',
  setFormula: 'excel',
  sortRange: 'excel',
  autoFilter: 'excel',
  createChart: 'excel',
  removeDuplicates: 'excel',
  addCellComment: 'excel',
  addConditionalFormat: 'excel',
  addDataValidation: 'excel',
  autoFitAll: 'excel',
  autoFitColumn: 'excel',
  autoFitRow: 'excel',
  calculateSheet: 'excel',
  cleanData: 'excel',
  clearFormats: 'excel',
  clearRange: 'excel',
  closeWorkbook: 'excel',
  consolidate: 'excel',
  copyFormat: 'excel',
  copyRange: 'excel',
  copySheet: 'excel',
  createNamedRange: 'excel',
  createPivotTable: 'excel',
  createSheet: 'excel',
  createWorkbook: 'excel',
  deleteCellComment: 'excel',
  deleteColumns: 'excel',
  deleteNamedRange: 'excel',
  deleteRows: 'excel',
  deleteSheet: 'excel',
  diagnoseFormula: 'excel',
  fillSeries: 'excel',
  findInSheet: 'excel',
  freezePanes: 'excel',
  getCellComments: 'excel',
  getCellInfo: 'excel',
  getConditionalFormats: 'excel',
  getContext: 'excel',
  getDataValidations: 'excel',
  getFormula: 'excel',
  getNamedRanges: 'excel',
  getOpenWorkbooks: 'excel',
  getSelection: 'excel',
  getSheetList: 'excel',
  groupColumns: 'excel',
  groupRows: 'excel',
  hideColumns: 'excel',
  hideRows: 'excel',
  insertColumns: 'excel',
  insertExcelImage: 'excel',
  insertRows: 'excel',
  lockCells: 'excel',
  mergeCells: 'excel',
  moveSheet: 'excel',
  openWorkbook: 'excel',
  pasteRange: 'excel',
  protectSheet: 'excel',
  protectWorkbook: 'excel',
  refreshLinks: 'excel',
  removeConditionalFormat: 'excel',
  removeDataValidation: 'excel',
  renameSheet: 'excel',
  replaceInSheet: 'excel',
  setArrayFormula: 'excel',
  setBorder: 'excel',
  setCellFormat: 'excel',
  setCellStyle: 'excel',
  setColumnWidth: 'excel',
  setHyperlink: 'excel',
  setNumberFormat: 'excel',
  setPrintArea: 'excel',
  setRowHeight: 'excel',
  showColumns: 'excel',
  showRows: 'excel',
  subtotal: 'excel',
  switchSheet: 'excel',
  switchWorkbook: 'excel',
  textToColumns: 'excel',
  transpose: 'excel',
  unfreezePanes: 'excel',
  unmergeCells: 'excel',
  unprotectSheet: 'excel',
  updateChart: 'excel',
  updatePivotTable: 'excel',
  wrapText: 'excel',
  // ==================== Word 命令 ====================
  getActiveDocument: 'word',
  getDocumentText: 'word',
  insertText: 'word',
  findReplace: 'word',
  setFont: 'word',
  applyStyle: 'word',
  insertTable: 'word',
  generateTOC: 'word',
  addComment: 'word',
  getBookmarks: 'word',
  getComments: 'word',
  getDocumentStats: 'word',
  getOpenDocuments: 'word',
  insertBookmark: 'word',
  insertFooter: 'word',
  insertHeader: 'word',
  insertHyperlink: 'word',
  insertImage: 'word',
  insertPageBreak: 'word',
  openDocument: 'word',
  setPageSetup: 'word',
  setParagraph: 'word',
  switchDocument: 'word',
  // ==================== PPT 命令 ====================
  getActivePresentation: 'ppt',
  addSlide: 'ppt',
  unifyFont: 'ppt',
  beautifySlide: 'ppt',
  addAnimation: 'ppt',
  addAnimationPreset: 'ppt',
  addArrow: 'ppt',
  addConnector: 'ppt',
  addEmphasisAnimation: 'ppt',
  addMasterElement: 'ppt',
  addPageIndicator: 'ppt',
  addPptHyperlink: 'ppt',
  addShape: 'ppt',
  addTextBox: 'ppt',
  addTitleDecoration: 'ppt',
  alignShapes: 'ppt',
  applyColorScheme: 'ppt',
  applyTransitionToAll: 'ppt',
  autoBeautifySlide: 'ppt',
  autoLayout: 'ppt',
  beautifyAllSlides: 'ppt',
  closePresentation: 'ppt',
  create3DText: 'ppt',
  createDonutChart: 'ppt',
  createFlowChart: 'ppt',
  createGauge: 'ppt',
  createGrid: 'ppt',
  createKpiCards: 'ppt',
  createMiniCharts: 'ppt',
  createOrgChart: 'ppt',
  createPresentation: 'ppt',
  createProgressBar: 'ppt',
  createStyledTable: 'ppt',
  createTimeline: 'ppt',
  deletePptImage: 'ppt',
  deleteShape: 'ppt',
  deleteSlide: 'ppt',
  deleteTextBox: 'ppt',
  distributeShapes: 'ppt',
  duplicateShape: 'ppt',
  duplicateSlide: 'ppt',
  endSlideShow: 'ppt',
  findPptText: 'ppt',
  getAnimations: 'ppt',
  getOpenPresentations: 'ppt',
  getPptTableCell: 'ppt',
  getShapes: 'ppt',
  getSlideCount: 'ppt',
  getSlideInfo: 'ppt',
  getSlideMaster: 'ppt',
  getSlideNotes: 'ppt',
  getSlideTitle: 'ppt',
  getTextBoxes: 'ppt',
  groupShapes: 'ppt',
  insertPptChart: 'ppt',
  insertPptImage: 'ppt',
  insertPptTable: 'ppt',
  moveSlide: 'ppt',
  openPresentation: 'ppt',
  removeAnimation: 'ppt',
  removePptHyperlink: 'ppt',
  removeSlideTransition: 'ppt',
  replacePptText: 'ppt',
  set3DDepth: 'ppt',
  set3DMaterial: 'ppt',
  set3DRotation: 'ppt',
  setAnimationOrder: 'ppt',
  setBackgroundColor: 'ppt',
  setBackgroundGradient: 'ppt',
  setBackgroundImage: 'ppt',
  setImageStyle: 'ppt',
  setMasterBackground: 'ppt',
  setPptChartData: 'ppt',
  setPptChartStyle: 'ppt',
  setPptDateTime: 'ppt',
  setPptFooter: 'ppt',
  setPptTableCell: 'ppt',
  setPptTableCellStyle: 'ppt',
  setPptTableRowStyle: 'ppt',
  setPptTableStyle: 'ppt',
  setShapeBorder: 'ppt',
  setShapeFullStyle: 'ppt',
  setShapeGradient: 'ppt',
  setShapePosition: 'ppt',
  setShapeRoundness: 'ppt',
  setShapeShadow: 'ppt',
  setShapeStyle: 'ppt',
  setShapeText: 'ppt',
  setShapeTransparency: 'ppt',
  setShapeZOrder: 'ppt',
  setSlideBackground: 'ppt',
  setSlideContent: 'ppt',
  setSlideLayout: 'ppt',
  setSlideNotes: 'ppt',
  setSlideNumber: 'ppt',
  setSlideSubtitle: 'ppt',
  setSlideTitle: 'ppt',
  setSlideTransition: 'ppt',
  setTextBoxStyle: 'ppt',
  setTextBoxText: 'ppt',
  smartDistribute: 'ppt',
  startSlideShow: 'ppt',
  switchPresentation: 'ppt',
  switchSlide: 'ppt',
};

// ==================== 按组件分发 ====================
// 原实现只有一个待执行命令，哪个组件的加载项先来取就给谁，且发出后不出队，
// 导致：命令被别的组件抢走报错；同一命令被两个组件各执行一次。
// 现在：加载项轮询时报上 app/doc/inst，服务端每个组件一个队列，取走即出队。

type AppKey = 'wps' | 'et' | 'wpp';
type QueueKey = AppKey | 'any';

const APP_LABEL: Record<AppKey, string> = { wps: 'WPS 文字', et: 'WPS 表格', wpp: 'WPS 演示' };
const APP_FILE_HINT: Record<AppKey, string> = { wps: '.docx', et: '.xlsx', wpp: '.pptx' };

/** 各种写法的组件名统一成 wps/et/wpp，认不出返回空串 */
export function normalizeApp(x: unknown): AppKey | '' {
  const s = String(x ?? '').trim().toLowerCase();
  if (['wps', 'word', 'writer', 'doc', 'docx', '文字'].includes(s)) return 'wps';
  if (['et', 'excel', 'spreadsheet', 'xls', 'xlsx', '表格'].includes(s)) return 'et';
  if (['wpp', 'ppt', 'powerpoint', 'presentation', 'pptx', '演示'].includes(s)) return 'wpp';
  return '';
}

/** 任何组件都能执行的通用命令：没指定组件时按“开着文件、最近操作”选目标 */
const COMMON_ACTIONS = new Set([
  'save', 'saveAs', 'convertToPDF', 'convertFormat', 'getAppInfo',
  'getSelectedText', 'setSelectedText', 'openFile',
]);
/** 只是探测连接，发给任意在线组件 */
const ANY_ACTIONS = new Set(['ping', 'wireCheck']);
/** 耗时命令放宽超时（保存 18 MB 的 PPT 实测约 55 秒） */
const SLOW_ACTIONS = new Set([
  'save', 'saveAs', 'convertToPDF', 'convertFormat', 'openFile', 'openDocument', 'openWorkbook',
  'openPresentation', 'exportSlideAsImage', 'exportChartAsImage', 'exportRangeAsImage',
  'insertSlidesFromFile', 'beautifyAllSlides', 'generateTOC',
]);
const DEFAULT_TIMEOUT = 30000;
const SLOW_TIMEOUT = 180000;
/** 多久没来取命令就算离线（加载项 500ms 取一次；WPS 忙时会更久） */
const ONLINE_WINDOW = 4000;

interface PendingCommand {
  action: string;
  params: Record<string, unknown>;
  requestId: string;
  app: QueueKey;
  timeoutMs: number;
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
  timer: NodeJS.Timeout | null;
  dispatchedTo?: string;
}

interface AddonInstance {
  inst: string;
  app: AppKey | 'unknown';
  doc: string;
  lastPoll: number;
}

export interface ExecuteOptions {
  appType?: string;
  timeout?: number;
}

/**
 * Mac轮询服务器类
 * 处理WPS加载项的轮询请求，实现命令的发送和结果接收
 */
class MacPollServer {
  private server: http.Server | null = null;
  private queues: Record<QueueKey, PendingCommand[]> = { wps: [], et: [], wpp: [], any: [] };
  private inFlight = new Map<string, PendingCommand>();
  private instances = new Map<string, AddonInstance>();
  private lastUsedApp: AppKey | '' = '';
  private _isRunning: boolean = false;
  private port: number = 58891;
  private startedAt = 0;

  get isRunning(): boolean {
    return this._isRunning;
  }

  /**
   * 启动轮询服务器
   * 1. GET /poll?app=&doc=&inst= - WPS加载项来取命令
   * 2. POST /result - WPS加载项返回执行结果
   * 3. GET /status - 查看各组件在线情况
   */
  async start(listenPort: number = 58891): Promise<void> {
    if (this._isRunning) return;
    this.port = listenPort;

    return new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => {
        // 先校验来源，不对浏览器网页开放
        const origin = req.headers.origin as string | undefined;
        if (origin !== undefined && !seenOrigins.has(origin)) {
          seenOrigins.add(origin);
          log.info('[Mac] Request origin seen', { origin, allowed: isAllowedOrigin(origin) });
        }
        if (!isAllowedOrigin(origin)) {
          log.warn('[Mac] Rejected request from origin', { origin, url: req.url });
          res.writeHead(403, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Origin not allowed' }));
          return;
        }
        res.setHeader('Access-Control-Allow-Origin', origin || '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
        res.setHeader('Content-Type', 'application/json');

        if (req.method === 'OPTIONS') {
          res.writeHead(200);
          res.end();
          return;
        }

        const url = new URL(req.url || '/', 'http://127.0.0.1');
        if (url.pathname === '/poll' && req.method === 'GET') {
          this.handlePoll(url, res);
        } else if (url.pathname === '/result' && req.method === 'POST') {
          this.handleResult(req, res);
        } else if (url.pathname === '/status') {
          res.end(JSON.stringify({ status: 'running', online: this.describeOnline(), lastUsedApp: this.lastUsedApp }));
        } else {
          res.writeHead(404);
          res.end(JSON.stringify({ error: 'Not found' }));
        }
      });

      server.on('error', (err: NodeJS.ErrnoException) => {
        if (err.code === 'EADDRINUSE') {
          // 原来假装启动成功，命令全部空等超时；现在直接说清楚
          log.warn(`[Mac] Port ${this.port} already in use`);
          reject(new Error(`端口 ${this.port} 已被占用：多半是另一个 Claude Code 会话正在使用 WPS 工具。` +
            '同一时间只有一个会话能连 WPS，请关掉其他会话后再试。'));
        } else {
          reject(err);
        }
      });

      server.listen(this.port, '127.0.0.1', () => {
        this.server = server;
        this._isRunning = true;
        this.startedAt = Date.now();
        log.info(`[Mac] Poll server started on port ${this.port}`);
        resolve();
      });
    });
  }

  /** 加载项来取命令 */
  private handlePoll(url: URL, res: http.ServerResponse): void {
    const app = normalizeApp(url.searchParams.get('app')) || 'unknown';
    const inst = url.searchParams.get('inst') || `legacy-${app}`;
    const doc = url.searchParams.get('doc') || '';
    const now = Date.now();
    const prev = this.instances.get(inst);
    if (!prev || prev.doc !== doc || prev.app !== app) {
      log.info('[Mac] Addon instance', { inst, app, doc });
    }
    this.instances.set(inst, { inst, app, doc, lastPoll: now });

    const cmd = this.pickCommand(app, inst, doc);
    if (!cmd) {
      res.end('{}');
      return;
    }
    // 取走即出队，不会再派给第二个组件
    cmd.dispatchedTo = `${app}/${inst}`;
    this.inFlight.set(cmd.requestId, cmd);
    log.debug('[Mac] Dispatch', { action: cmd.action, requestId: cmd.requestId, to: cmd.dispatchedTo });
    res.end(JSON.stringify({
      command: { action: cmd.action, params: cmd.params, requestId: cmd.requestId, app: cmd.app },
    }));
  }

  /** 给这个实例挑一条命令（并从队列移除）；没有合适的返回 null */
  private pickCommand(app: AppKey | 'unknown', inst: string, doc: string): PendingCommand | null {
    if (app === 'unknown') {
      // 旧版加载项不报组件：按原来的方式什么都给（兼容）
      for (const k of ['any', 'wps', 'et', 'wpp'] as QueueKey[]) {
        if (this.queues[k].length) return this.queues[k].shift()!;
      }
      return null;
    }
    if (this.queues[app].length) {
      // 同组件有另一个开着文件的实例在线时，没开文件的实例让给它
      if (!doc && this.onlineInstances(app).some(i => i.inst !== inst && i.doc)) return null;
      return this.queues[app].shift()!;
    }
    if (this.queues.any.length) return this.queues.any.shift()!;
    return null;
  }

  /** 加载项送回结果 */
  private handleResult(req: http.IncomingMessage, res: http.ServerResponse): void {
    let body = '';
    req.on('data', (chunk) => { body += chunk.toString(); });
    req.on('end', () => {
      try {
        const data = JSON.parse(body);
        const cmd = this.inFlight.get(data.requestId);
        if (!cmd) {
          log.warn('[Mac] Received result for unknown request', { requestId: data.requestId, from: data.app });
          res.end(JSON.stringify({ ok: true }));
          return;
        }
        this.inFlight.delete(data.requestId);
        if (data.result && data.result.wrongApp && cmd.app !== 'any') {
          // 加载项发现不是自己的命令（不应发生），放回队首重派
          log.warn('[Mac] Addon refused command for other app, requeue', { action: cmd.action, from: data.app });
          cmd.dispatchedTo = undefined;
          this.queues[cmd.app].unshift(cmd);
        } else {
          if (cmd.timer) clearTimeout(cmd.timer);
          cmd.resolve(data.result);
        }
        res.end(JSON.stringify({ ok: true }));
      } catch (e) {
        log.error('[Mac] Failed to parse result', { error: e, body });
        res.writeHead(400);
        res.end(JSON.stringify({ error: 'Invalid JSON' }));
      }
    });
  }

  private onlineInstances(app?: AppKey): AddonInstance[] {
    const now = Date.now();
    return [...this.instances.values()].filter(i =>
      now - i.lastPoll < ONLINE_WINDOW && (app === undefined || i.app === app));
  }

  private onlineApps(withDocOnly: boolean): AppKey[] {
    const apps = new Set<AppKey>();
    for (const i of this.onlineInstances()) {
      if (i.app !== 'unknown' && (!withDocOnly || i.doc)) apps.add(i.app);
    }
    return [...apps];
  }

  /** 各组件在线情况，给 ping / 状态查询用 */
  describeOnline(): Array<{ app: string; label: string; doc: string }> {
    return this.onlineInstances().map(i => ({
      app: i.app,
      label: i.app === 'unknown' ? '旧版加载项' : APP_LABEL[i.app],
      doc: i.doc,
    }));
  }

  /** 刚启动时加载项还没来得及取命令，最多等 waitMs */
  private async waitForAddons(pred: () => boolean, waitMs: number): Promise<boolean> {
    const t0 = Date.now();
    while (!pred()) {
      if (Date.now() - t0 >= waitMs) return false;
      await new Promise(r => setTimeout(r, 200));
    }
    return true;
  }

  /** 通用命令没指定组件时，选目标组件 */
  private pickCommonTarget(): AppKey | '' {
    const pick = (cands: AppKey[]): AppKey | '' => {
      if (cands.length === 0) return '';
      if (this.lastUsedApp && cands.includes(this.lastUsedApp)) return this.lastUsedApp;
      return cands[0];
    };
    return pick(this.onlineApps(true)) || pick(this.onlineApps(false));
  }

  private notConnectedError(app: AppKey): Error {
    const others = this.onlineInstances().filter(i => i.app !== app && i.app !== 'unknown')
      .map(i => APP_LABEL[i.app as AppKey] + (i.doc ? `（${i.doc}）` : ''));
    return new Error(`${APP_LABEL[app]}没有连接：请在 WPS 里打开一个 ${APP_FILE_HINT[app]} 文件后再试` +
      '（刚打开的话等几秒；如果 WPS 弹出“加载项已修改，是否继续加载”，点 OK）。' +
      (others.length ? `当前在线：${[...new Set(others)].join('、')}。` : ''));
  }

  /**
   * 执行WPS命令：确定目标组件 → 入该组件的队列 → 等结果
   */
  async executeCommand(action: string, params: Record<string, unknown> = {}, opts: ExecuteOptions = {}): Promise<unknown> {
    const sendParams: Record<string, unknown> = { ...params };
    const explicit = normalizeApp(sendParams.targetApp);
    delete sendParams.targetApp;

    let target: QueueKey | '' = explicit || normalizeApp(opts.appType) || normalizeApp(COMMAND_APP_MAP[action]);
    const justStarted = Date.now() - this.startedAt < 3000;

    if (!target && ANY_ACTIONS.has(action)) {
      target = 'any';
      if (!(await this.waitForAddons(() => this.onlineInstances().length > 0, justStarted ? 3000 : 0))) {
        throw new Error('WPS 没有连接：请先打开 WPS 并打开一个文件（刚打开的话等几秒再试；' +
          '如果 WPS 弹出“加载项已修改，是否继续加载”，点 OK）。');
      }
    } else if (!target) {
      // 通用命令（保存、导出 PDF 等）或未登记的命令
      await this.waitForAddons(() => this.pickCommonTarget() !== '', justStarted ? 3000 : 0);
      target = this.pickCommonTarget();
      if (!target) {
        throw new Error('WPS 没有连接：请先打开 WPS 并打开一个文件。');
      }
      if (!COMMON_ACTIONS.has(action)) {
        log.warn('[Mac] Action not in COMMAND_APP_MAP, routed by recent use', { action, target });
      }
    } else {
      const app = target as AppKey;
      const ok = await this.waitForAddons(() => this.onlineInstances(app).length > 0, justStarted ? 3000 : 1500);
      if (!ok) throw this.notConnectedError(app);
    }

    if (target !== 'any') this.lastUsedApp = target;
    const timeoutMs = opts.timeout || (SLOW_ACTIONS.has(action) ? SLOW_TIMEOUT : DEFAULT_TIMEOUT);

    const result = await new Promise<unknown>((resolve, reject) => {
      const requestId = `req-${Date.now()}-${Math.random().toString(36).slice(2, 11)}`;
      const cmd: PendingCommand = {
        action, params: sendParams, requestId, app: target as QueueKey, timeoutMs, resolve, reject, timer: null,
      };
      cmd.timer = setTimeout(() => {
        const q = this.queues[cmd.app];
        const idx = q.indexOf(cmd);
        if (idx >= 0) q.splice(idx, 1);
        const wasDispatched = this.inFlight.delete(requestId) || !!cmd.dispatchedTo;
        const label = cmd.app === 'any' ? 'WPS' : APP_LABEL[cmd.app];
        reject(new Error(wasDispatched
          ? `${label}在 ${Math.round(timeoutMs / 1000)} 秒内没有返回结果（${action}）。` +
            'WPS 可能仍在处理，保存、导出这类操作可能已经完成，请先核对文件再决定是否重做。'
          : `${label}在 ${Math.round(timeoutMs / 1000)} 秒内没有来取命令（${action}），命令未执行。` +
            'WPS 可能卡住或弹出了对话框，请看一下 WPS 窗口。'));
      }, timeoutMs);
      this.queues[cmd.app].push(cmd);
      log.debug('[Mac] Command queued', { action, requestId, app: cmd.app });
    });

    // ping 时顺便报告各组件在线情况
    if (ANY_ACTIONS.has(action) && result && typeof result === 'object') {
      (result as Record<string, unknown>).online = this.describeOnline();
    }
    return result;
  }

  /**
   * 停止服务器
   */
  stop(): void {
    for (const k of Object.keys(this.queues) as QueueKey[]) {
      for (const c of this.queues[k]) {
        if (c.timer) clearTimeout(c.timer);
        c.reject(new Error('Server stopped'));
      }
      this.queues[k] = [];
    }
    for (const c of this.inFlight.values()) {
      if (c.timer) clearTimeout(c.timer);
      c.reject(new Error('Server stopped'));
    }
    this.inFlight.clear();
    if (this.server) {
      this.server.close();
      this.server = null;
      this._isRunning = false;
      log.info('[Mac] Poll server stopped');
    }
  }

  /** 最近一次操作的组件（兼容旧接口） */
  getCurrentApp(): string {
    return this.lastUsedApp;
  }

  setCurrentApp(app: string): void {
    this.lastUsedApp = normalizeApp(app);
  }
}

// 导出单例 - 整个应用共用一个服务器实例
export const macPollServer = new MacPollServer();

export default MacPollServer;
