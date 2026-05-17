/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: 跨平台 WPS 客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS通信客户端 - 老王的跨平台版
 * Windows: 通过PowerShell调用WPS COM接口
 * Mac: 通过反向轮询服务器（MCP Server当服务端，WPS加载项来轮询）
 *
 * 丢，为了兼容Mac老王可是费了老大劲了
 * WPS Mac加载项在沙箱里启动不了HTTP服务器，只能反过来搞！
 */

import { spawn } from 'child_process';
import * as path from 'path';
import * as os from 'os';
import {
  WpsEndpointConfig,
  WpsApiRequest,
  WpsApiResponse,
  WpsAppType,
  WpsClientStatus,
  DocumentInfo,
  WorkbookInfo,
  PresentationInfo,
} from '../types/wps';
import { log, logRequest, logResponse } from '../utils/logger';
import { errorUtils } from '../utils/error';
import { macPollServer } from './mac-poll-server';

// 平台判断
function isMacPlatform() {
  return os.platform() === 'darwin';
}
// const IS_WINDOWS = os.platform() === 'win32';  // 暂时不用，保留备用

// PowerShell脚本路径 (Windows)
const PS_SCRIPT_PATH = path.join(__dirname, '../../scripts/wps-com.ps1');

// Mac轮询服务器端口
const MAC_POLL_PORT = 58891;

/**
 * 执行Mac轮询调用
 * 通过轮询服务器发送命令，等待WPS加载项取走并返回结果
 */
async function execMacPoll(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  log.debug('Executing Mac Poll', { action, params });

  try {
    // 确保轮询服务器已启动
    if (!macPollServer.isRunning) {
      log.info('[Mac] Starting poll server...');
      await macPollServer.start(MAC_POLL_PORT);
    }

    // 通过轮询服务器执行命令
    const result = await macPollServer.executeCommand(action, params);
    return result;
  } catch (error) {
    log.error('Mac Poll call failed', { action, error });
    throw error;
  }
}

/**
 * 执行PowerShell命令 (Windows)
 */
async function execPowerShell(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const paramsJson = JSON.stringify(params);
    const args = [
      '-ExecutionPolicy', 'Bypass',
      '-File', PS_SCRIPT_PATH,
      '-Action', action,
      '-Params', paramsJson
    ];

    log.debug('Executing PowerShell', { action, params });

    const ps = spawn('powershell', args, {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('close', (code) => {
      if (code !== 0 && stderr) {
        log.error('PowerShell error', { stderr, code });
        reject(new Error(stderr));
        return;
      }

      try {
        const result = JSON.parse(stdout.trim());
        resolve(result);
      } catch (e) {
        log.error('Failed to parse PowerShell output', { stdout });
        reject(new Error(`Invalid JSON output: ${stdout}`));
      }
    });

    ps.on('error', (err) => {
      reject(err);
    });
  });
}

/**
 * 统一执行接口 - 根据平台选择调用方式
 * Mac: 反向轮询模式（MCP Server是服务端，WPS加载项来取命令）
 * Windows: PowerShell调用COM接口
 */
async function execWpsAction(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  if (isMacPlatform()) {
    return execMacPoll(action, params);
  } else {
    return execPowerShell(action, params);
  }
}

// 超时时间（毫秒）
const COM_TIMEOUT = 5000;

/**
 * 带超时和重试的WPS调用
 * 注：COM 操作无法真正取消，此实现确保超时后不再等待并快速失败
 */
async function execWpsActionWithRetry(action: string, params: Record<string, unknown> = {}, maxRetries: number = 3): Promise<unknown> {
  let lastError: Error | null = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // 使用 Promise.race 实现超时
      // 注意：无法真正取消 COM 调用，但可确保不会无限等待
      const timeoutPromise = new Promise((_, reject) => {
        setTimeout(() => reject(new Error('COM 调用超时（' + COM_TIMEOUT + 'ms）')), COM_TIMEOUT);
      });

      const result = await Promise.race([
        execWpsAction(action, params),
        timeoutPromise
      ]);

      return result;
    } catch (error) {
      lastError = error as Error;
      const errMsg = error instanceof Error ? error.message : String(error);

      // 超时是快速失败的友好错误，不记录为严重警告
      if (errMsg.includes('超时')) {
        log.info(`WPS call timeout, attempt ${attempt}/${maxRetries}`, { action });
      } else {
        log.warn(`WPS call failed, attempt ${attempt}/${maxRetries}`, { action, error: errMsg });
      }

      if (attempt < maxRetries) {
        // 等待后重试（指数退避）
        await new Promise(resolve => setTimeout(resolve, 500 * attempt));
      }
    }
  }

  throw lastError || new Error('WPS call failed after retries');
}

/**
 * WPS客户端类 - 跨平台通信
 * Windows: PowerShell COM桥接
 * Mac: HTTP调用WPS加载项
 */
export class WpsClient {
  private status: WpsClientStatus;

  constructor(_config?: Partial<WpsEndpointConfig>) {
    this.status = { connected: false };
    const method = isMacPlatform() ? 'HTTP (Mac Addon)' : 'PowerShell COM';
    log.info('WPS Client initialized', { method, platform: os.platform() });
  }

  /**
   * 调用WPS接口（跨平台）
   */
  async invokeAction<T = unknown>(action: string, params: Record<string, unknown> = {}): Promise<WpsApiResponse<T>> {
    const startTime = Date.now();
    logRequest(action, params);

    try {
      const result = await execWpsActionWithRetry(action, params, 3) as WpsApiResponse<T>;
      const duration = Date.now() - startTime;
      logResponse(action, result.success, duration);

      if (result.success) {
        this.status.connected = true;
        this.status.lastHeartbeat = new Date();
      }

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      logResponse(action, false, duration);
      this.status.connected = false;
      throw errorUtils.wrap(error, `WPS COM call failed: ${action}`);
    }
  }

  /**
   * 兼容旧API
   */
  async callApi<T = unknown>(request: WpsApiRequest): Promise<WpsApiResponse<T>> {
    const actionMap: Record<string, string> = {
      'workbook.getActive': 'getActiveWorkbook',
      'cell.getValue': 'getCellValue',
      'cell.setValue': 'setCellValue',
      'range.getData': 'getRangeData',
      'range.setData': 'setRangeData',
      'file.save': 'save',
      'ping': 'ping',
    };
    const action = actionMap[request.method] || request.method;
    return this.invokeAction<T>(action, request.params || {});
  }

  /**
   * 检查WPS连接状态
   */
  async checkConnection(): Promise<boolean> {
    try {
      const result = await this.invokeAction('ping');
      this.status.connected = result.success;
      return result.success;
    } catch {
      this.status.connected = false;
      this.status.error = 'Connection check failed';
      return false;
    }
  }

  /**
   * 获取客户端状态
   */
  getStatus(): WpsClientStatus {
    return { ...this.status };
  }

  // ==================== 表格操作 (WPS表格) ====================

  async getActiveWorkbook(): Promise<WorkbookInfo | null> {
    const response = await this.invokeAction<WorkbookInfo>('getActiveWorkbook');
    return response.success ? response.data || null : null;
  }

  async getCellValue(sheet: string | number, row: number, col: number): Promise<unknown> {
    const response = await this.invokeAction<{ value: unknown }>('getCellValue', { sheet, row, col });
    return response.data?.value;
  }

  async setCellValue(sheet: string | number, row: number, col: number, value: unknown): Promise<boolean> {
    const response = await this.invokeAction('setCellValue', { sheet, row, col, value });
    return response.success;
  }

  async getRangeData(sheet: string | number, range: string): Promise<unknown[][]> {
    const response = await this.invokeAction<{ data: unknown[][] }>('getRangeData', { sheet, range });
    return response.data?.data || [];
  }

  async setRangeData(sheet: string | number, range: string, data: unknown[][]): Promise<boolean> {
    const response = await this.invokeAction('setRangeData', { sheet, range, data });
    return response.success;
  }

  async setFormula(sheet: string | number, row: number, col: number, formula: string): Promise<boolean> {
    const response = await this.invokeAction('setFormula', { sheet, row, col, formula });
    return response.success;
  }

  // ==================== 文档操作 (WPS文字) ====================

  async getActiveDocument(): Promise<DocumentInfo | null> {
    const response = await this.invokeAction<DocumentInfo>('getActiveDocument');
    return response.success ? response.data || null : null;
  }

  async createDocument(): Promise<boolean> {
    const response = await this.invokeAction('createDocument');
    return response.success;
  }

  async insertText(text: string, position?: number): Promise<boolean> {
    const response = await this.invokeAction('insertText', { text, position });
    return response.success;
  }

  async getDocumentText(): Promise<string> {
    const response = await this.invokeAction<{ text: string }>('getDocumentText');
    return response.data?.text || '';
  }

  // ==================== 演示操作 (WPS演示) ====================

  async getActivePresentation(): Promise<PresentationInfo | null> {
    const response = await this.invokeAction<PresentationInfo>('getActivePresentation');
    return response.success ? response.data || null : null;
  }

  async createPresentation(): Promise<boolean> {
    const response = await this.invokeAction('createPresentation');
    return response.success;
  }

  async addSlide(layout?: string): Promise<boolean> {
    const response = await this.invokeAction('addSlide', { layout });
    return response.success;
  }

  // ==================== 通用操作 ====================

  async executeMethod<T = unknown>(
    method: string,
    params?: Record<string, unknown>,
    _appType?: WpsAppType
  ): Promise<WpsApiResponse<T>> {
    return this.invokeAction<T>(method, params);
  }

  async openFile(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('openFile', { path: filePath });
    return response.success;
  }

  async saveFile(_appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('save');
    return response.success;
  }

  async saveFileAs(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('saveAs', { path: filePath });
    return response.success;
  }
}

// 导出单例
export const wpsClient = new WpsClient();

export default WpsClient;
