/**
 * Input: 通用操作工具参数
 * Output: 操作结果
 * Pos: 通用操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 通用操作Tools - 保存/连接检测/文本选取等基础模块
 *
 * 包含：
 * - wps_common_save: 保存当前文档
 * - wps_common_save_as: 另存为
 * - wps_common_ping: 检测WPS连接
 * - wps_common_wire_check: 检查通信线路
 * - wps_common_get_app_info: 获取WPS应用信息
 * - wps_common_get_selected_text: 获取选中文本
 * - wps_common_set_selected_text: 替换选中文本
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

// ============================================================
// 1. wps_common_save - 保存当前文档
// ============================================================

export const saveDefinition: ToolDefinition = {
  name: 'wps_common_save',
  description: `保存当前文档。

使用场景：
- "保存一下"
- "Ctrl+S"
- "把修改存起来"

特点：
- 自动保存当前活动文档
- 如果文档从未保存过，可能会提示选择保存路径`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const saveHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>('save', {});

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          { type: 'text', text: `文档保存成功！${response.data?.message || ''}` },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `保存失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `保存出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 2. wps_common_save_as - 另存为
// ============================================================

export const saveAsDefinition: ToolDefinition = {
  name: 'wps_common_save_as',
  description: `将当前文档另存为指定路径和格式。

使用场景：
- "另存为到桌面"
- "换个名字保存"
- "保存一份副本到指定位置"

特点：
- 支持指定完整文件路径
- 可选指定保存格式`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '目标文件完整路径，包含文件名和扩展名',
      },
      format: {
        type: 'string',
        description: '保存格式（可选），如 docx, xlsx, pptx 等',
      },
    },
    required: ['filePath'],
  },
};

export const saveAsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath, format } = args as {
    filePath: string;
    format?: string;
  };

  if (!filePath || filePath.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文件路径不能为空，请指定保存路径' }],
      error: '文件路径为空',
    };
  }

  try {
    const params: Record<string, unknown> = {
      filePath,
      path: filePath,
      outputPath: filePath,
    };
    if (format) {
      params.format = format.toLowerCase().replace(/^\./, '');
    }

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      outputPath: string;
    }>('saveAs', params);

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `另存为成功！\n输出路径: ${response.data.outputPath || filePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `另存为失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `另存为出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 3. wps_common_ping - 检测WPS连接
// ============================================================

export const pingDefinition: ToolDefinition = {
  name: 'wps_common_ping',
  description: `检测WPS应用连接状态。

使用场景：
- "WPS连上了吗"
- "检查一下WPS是否在线"
- "测试WPS连接"

特点：
- 快速检测WPS加载项是否可达
- 返回连接状态信息`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const pingHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>('ping', {});

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          { type: 'text', text: `WPS连接正常！${response.data?.message || 'pong'}` },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `WPS连接失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `WPS连接检测出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 4. wps_common_wire_check - 检查通信线路
// ============================================================

export const wireCheckDefinition: ToolDefinition = {
  name: 'wps_common_wire_check',
  description: `检查与WPS加载项之间的通信线路状态。

使用场景：
- "检查通信状态"
- "诊断连接问题"
- "通信线路是否正常"

特点：
- 比ping更详细的通信诊断
- 返回线路延迟、协议状态等信息`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const wireCheckHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      latency?: number;
      protocol?: string;
    }>('wireCheck', {});

    if (response.success && response.data) {
      const data = response.data;
      let info = '通信线路正常！';
      if (data.latency !== undefined) {
        info += `\n延迟: ${data.latency}ms`;
      }
      if (data.protocol) {
        info += `\n协议: ${data.protocol}`;
      }
      if (data.message) {
        info += `\n${data.message}`;
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: info }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `通信线路检查失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `通信线路检查出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 5. wps_common_get_app_info - 获取WPS应用信息
// ============================================================

export const getAppInfoDefinition: ToolDefinition = {
  name: 'wps_common_get_app_info',
  description: `获取WPS应用的基本信息。

使用场景：
- "WPS是什么版本"
- "查看WPS信息"
- "获取应用状态"

特点：
- 返回WPS版本号、构建信息
- 返回当前打开的文档信息
- 返回运行平台信息`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getAppInfoHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      version?: string;
      build?: string;
      platform?: string;
      activeDocument?: string;
    }>('getAppInfo', {});

    if (response.success && response.data) {
      const data = response.data;
      const lines: string[] = ['WPS应用信息：'];
      if (data.version) lines.push(`版本: ${data.version}`);
      if (data.build) lines.push(`构建号: ${data.build}`);
      if (data.platform) lines.push(`平台: ${data.platform}`);
      if (data.activeDocument) lines.push(`当前文档: ${data.activeDocument}`);
      if (data.message) lines.push(data.message);

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取应用信息失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取应用信息出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 6. wps_common_get_selected_text - 获取选中文本
// ============================================================

export const getSelectedTextDefinition: ToolDefinition = {
  name: 'wps_common_get_selected_text',
  description: `获取当前文档中选中的文本内容。

使用场景：
- "读取我选中的内容"
- "获取当前选区的文字"
- "看看我选了什么"

特点：
- 返回当前选区的纯文本内容
- 适用于Word、Excel、PPT`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getSelectedTextHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      text: string;
    }>('getSelectedText', {});

    if (response.success && response.data) {
      const text = response.data.text;
      if (!text || text.length === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: '当前没有选中任何文本。' }],
        };
      }
      return {
        id: uuidv4(),
        success: true,
        content: [
          { type: 'text', text: `选中的文本内容：\n${text}` },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取选中文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取选中文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 7. wps_common_set_selected_text - 替换选中文本
// ============================================================

export const setSelectedTextDefinition: ToolDefinition = {
  name: 'wps_common_set_selected_text',
  description: `替换当前文档中选中的文本内容。

使用场景：
- "把选中的内容替换成xxx"
- "修改选区的文字"
- "用新内容替换选中部分"

特点：
- 将当前选区的文本替换为指定内容
- 适用于Word、Excel、PPT`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '要替换为的新文本内容',
      },
    },
    required: ['text'],
  },
};

export const setSelectedTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text } = args as { text: string };

  if (text === undefined || text === null) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '替换文本不能为空，请指定要替换的内容' }],
      error: '替换文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>('setSelectedText', { text });

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          { type: 'text', text: `选中文本已替换成功！${response.data?.message || ''}` },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `替换选中文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `替换选中文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 导出所有通用操作Tools
// ============================================================

export const generalTools: RegisteredTool[] = [
  { definition: saveDefinition, handler: saveHandler },
  { definition: saveAsDefinition, handler: saveAsHandler },
  { definition: pingDefinition, handler: pingHandler },
  { definition: wireCheckDefinition, handler: wireCheckHandler },
  { definition: getAppInfoDefinition, handler: getAppInfoHandler },
  { definition: getSelectedTextDefinition, handler: getSelectedTextHandler },
  { definition: setSelectedTextDefinition, handler: setSelectedTextHandler },
];

export default generalTools;
