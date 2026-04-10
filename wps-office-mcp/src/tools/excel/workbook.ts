/**
 * Input: 工作簿管理工具参数
 * Output: 工作簿操作结果
 * Pos: Excel 工作簿管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

/**
 * 打开指定路径的Excel工作簿
 */
export const openWorkbookDefinition: ToolDefinition = {
  name: 'wps_excel_open_workbook',
  description: '打开指定路径的Excel工作簿文件。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '工作簿文件路径',
      },
    },
    required: ['filePath'],
  },
};

export const openWorkbookHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath } = args as { filePath: string };
  if (!filePath) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '请提供工作簿文件路径' }],
      error: '缺少文件路径',
    };
  }
  try {
    const params = { filePath, path: filePath };
    const response = await wpsClient.executeMethod<{ message: string }>(
      'openWorkbook',
      params,
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `打开工作簿失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `工作簿已打开: ${filePath}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `打开工作簿出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取所有已打开的工作簿列表
 */
export const getOpenWorkbooksDefinition: ToolDefinition = {
  name: 'wps_excel_get_open_workbooks',
  description: '获取当前所有已打开的Excel工作簿列表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getOpenWorkbooksHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ workbooks: string[] }>(
      'getOpenWorkbooks',
      {},
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取工作簿列表失败: ${response.error}` }], error: response.error };
    }
    const list = response.data?.workbooks || [];
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已打开的工作簿 (${list.length}个):\n${list.join('\n') || '无'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取工作簿列表出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 切换到指定名称的工作簿
 */
export const switchWorkbookDefinition: ToolDefinition = {
  name: 'wps_excel_switch_workbook',
  description: '切换到指定名称的Excel工作簿。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '工作簿名称',
      },
    },
    required: ['name'],
  },
};

export const switchWorkbookHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };
  if (!name) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '请提供工作簿名称' }], error: '缺少工作簿名称' };
  }
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'switchWorkbook',
      { name },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `切换工作簿失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已切换到工作簿: ${name}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `切换工作簿出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 关闭指定工作簿
 */
export const closeWorkbookDefinition: ToolDefinition = {
  name: 'wps_excel_close_workbook',
  description: '关闭指定的Excel工作簿，可选是否保存。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '工作簿名称，不填则关闭当前工作簿',
      },
      save: {
        type: 'boolean',
        description: '是否保存，默认true',
      },
    },
    required: [],
  },
};

export const closeWorkbookHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, save } = args as { name?: string; save?: boolean };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'closeWorkbook',
      { name, save: save !== false },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `关闭工作簿失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `工作簿已关闭${name ? ': ' + name : ''}${save !== false ? '（已保存）' : '（未保存）'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `关闭工作簿出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 新建空白工作簿
 */
export const createWorkbookDefinition: ToolDefinition = {
  name: 'wps_excel_create_workbook',
  description: '新建一个空白Excel工作簿。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const createWorkbookHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'createWorkbook',
      {},
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `新建工作簿失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '新工作簿已创建' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `新建工作簿出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取指定单元格的值
 */
export const getCellValueDefinition: ToolDefinition = {
  name: 'wps_excel_get_cell_value',
  description: '获取Excel指定单元格的值。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名称' },
      row: { type: 'number', description: '行号（从1开始）' },
      col: { type: 'number', description: '列号（从1开始）' },
    },
    required: ['sheet', 'row', 'col'],
  },
};

export const getCellValueHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { sheet, row, col } = args as { sheet: string; row: number; col: number };
  try {
    const response = await wpsClient.executeMethod<{ value: unknown }>(
      'getCellValue',
      { sheet, row, col },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取单元格值失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `单元格值: ${JSON.stringify(response.data?.value)}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取单元格值出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置指定单元格的值
 */
export const setCellValueDefinition: ToolDefinition = {
  name: 'wps_excel_set_cell_value',
  description: '设置Excel指定单元格的值。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名称' },
      row: { type: 'number', description: '行号（从1开始）' },
      col: { type: 'number', description: '列号（从1开始）' },
      value: { type: 'string', description: '要设置的值' },
    },
    required: ['sheet', 'row', 'col', 'value'],
  },
};

export const setCellValueHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { sheet, row, col, value } = args as { sheet: string; row: number; col: number; value: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setCellValue',
      { sheet, row, col, value },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置单元格值失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `单元格值已设置: ${sheet}!R${row}C${col} = ${value}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置单元格值出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取指定单元格的公式
 */
export const getFormulaDefinition: ToolDefinition = {
  name: 'wps_excel_get_formula',
  description: '获取Excel指定单元格的公式。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名称' },
      cell: { type: 'string', description: '单元格地址，如 A1、B2' },
    },
    required: ['sheet', 'cell'],
  },
};

export const getFormulaHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { sheet, cell } = args as { sheet: string; cell: string };
  try {
    const response = await wpsClient.executeMethod<{ formula: string }>(
      'getFormula',
      { sheet, cell },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取公式失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `单元格 ${cell} 的公式: ${response.data?.formula || '无公式'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取公式出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取单元格详细信息
 */
export const getCellInfoDefinition: ToolDefinition = {
  name: 'wps_excel_get_cell_info',
  description: '获取单元格的详细信息（值、公式、格式等）。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名称' },
      cell: { type: 'string', description: '单元格地址，如 A1' },
    },
    required: ['sheet', 'cell'],
  },
};

export const getCellInfoHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { sheet, cell } = args as { sheet: string; cell: string };
  try {
    const response = await wpsClient.executeMethod<Record<string, unknown>>(
      'getCellInfo',
      { sheet, cell },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取单元格信息失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `单元格 ${cell} 信息:\n${JSON.stringify(response.data, null, 2)}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取单元格信息出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 清除指定范围
 */
export const clearRangeDefinition: ToolDefinition = {
  name: 'wps_excel_clear_range',
  description: '清除指定范围的内容、格式或全部。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '范围地址，如 A1:C10' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前工作表' },
      type: {
        type: 'string',
        description: '清除类型：all（全部）、contents（仅内容）、formats（仅格式）',
        enum: ['all', 'contents', 'formats'],
      },
    },
    required: ['range'],
  },
};

export const clearRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, sheet, type } = args as { range: string; sheet?: string; type?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'clearRange',
      { range, sheet, type: type || 'all' },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `清除范围失败: ${response.error}` }], error: response.error };
    }
    const typeLabel = type === 'contents' ? '内容' : type === 'formats' ? '格式' : '全部';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `范围 ${range} 的${typeLabel}已清除` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `清除范围出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 导出所有工作簿管理相关的Tools
 */
export const workbookTools: RegisteredTool[] = [
  { definition: openWorkbookDefinition, handler: openWorkbookHandler },
  { definition: getOpenWorkbooksDefinition, handler: getOpenWorkbooksHandler },
  { definition: switchWorkbookDefinition, handler: switchWorkbookHandler },
  { definition: closeWorkbookDefinition, handler: closeWorkbookHandler },
  { definition: createWorkbookDefinition, handler: createWorkbookHandler },
  { definition: getCellValueDefinition, handler: getCellValueHandler },
  { definition: setCellValueDefinition, handler: setCellValueHandler },
  { definition: getFormulaDefinition, handler: getFormulaHandler },
  { definition: getCellInfoDefinition, handler: getCellInfoHandler },
  { definition: clearRangeDefinition, handler: clearRangeHandler },
];

export default workbookTools;
