/**
 * Input: 文档管理工具参数
 * Output: 文档操作结果
 * Pos: Word 文档管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word文档管理Tools
 * 处理文档的打开、切换、获取列表等管理操作
 *
 * 包含：
 * - wps_word_get_open_documents: 获取所有已打开的文档列表
 * - wps_word_switch_document: 切换到指定文档
 * - wps_word_open_document: 打开指定路径的文档
 * - wps_word_get_document_text: 获取文档文本内容
 * - wps_word_insert_header: 设置页眉内容
 * - wps_word_insert_footer: 设置页脚内容
 * - wps_word_generate_doc_toc: 自动生成文档目录
 * - wps_word_insert_section_break: 插入分节符
 * - wps_word_set_line_spacing: 设置行距
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
 * 获取所有已打开的文档列表
 */
export const getOpenDocumentsDefinition: ToolDefinition = {
  name: 'wps_word_get_open_documents',
  description: `获取当前WPS Writer中所有已打开的文档列表。

使用场景：
- "看看现在打开了哪些文档"
- "列出所有打开的Word文件"
- "查看当前文档列表"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getOpenDocumentsHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documents: Array<{ name: string; path: string; active: boolean }>;
    }>(
      'getOpenDocuments',
      {},
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const docs = response.data.documents;
      if (!docs || docs.length === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: '当前没有打开任何文档' }],
        };
      }

      const docList = docs
        .map((doc, i) => `${i + 1}. ${doc.name}${doc.active ? ' (当前活动)' : ''}\n   路径: ${doc.path}`)
        .join('\n');

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已打开的文档列表 (共${docs.length}个):\n${docList}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 切换到指定文档
 */
export const switchDocumentDefinition: ToolDefinition = {
  name: 'wps_word_switch_document',
  description: `切换到指定名称的文档。

使用场景：
- "切换到报告.docx"
- "打开另一个文档窗口"
- "切换到那个合同文档"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要切换到的文档名称',
      },
    },
    required: ['name'],
  },
};

export const switchDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  if (!name || name.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文档名称不能为空！' }],
      error: '文档名称为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchDocument',
      { name },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到文档: ${name}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 打开指定路径的文档
 */
export const openDocumentDefinition: ToolDefinition = {
  name: 'wps_word_open_document',
  description: `打开指定路径的Word文档。

使用场景：
- "打开桌面上的报告.docx"
- "帮我打开这个文件路径的文档"
- "加载指定位置的Word文件"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '要打开的文档文件路径',
      },
    },
    required: ['filePath'],
  },
};

export const openDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath } = args as { filePath: string };

  if (!filePath || filePath.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文件路径不能为空！' }],
      error: '文件路径为空',
    };
  }

  try {
    const params = {
      filePath,
      path: filePath,
    };

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documentName: string;
    }>(
      'openDocument',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档打开成功！\n文件: ${response.data.documentName || filePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `打开文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `打开文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取文档文本内容
 */
export const getDocumentTextDefinition: ToolDefinition = {
  name: 'wps_word_get_document_text',
  description: `获取当前Word文档的文本内容。

使用场景：
- "读取文档内容"
- "获取文档的前100个字符"
- "查看文档从第50到第200个字符的内容"

可指定起始和结束位置来获取部分文本。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      start: {
        type: 'number',
        description: '起始位置（字符索引），默认从头开始',
      },
      end: {
        type: 'number',
        description: '结束位置（字符索引），默认到文档末尾',
      },
    },
  },
};

export const getDocumentTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { start, end } = args as { start?: number; end?: number };

  try {
    const params: Record<string, unknown> = {};
    if (start !== undefined) params.start = start;
    if (end !== undefined) params.end = end;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      text: string;
      length: number;
    }>(
      'getDocumentText',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const rangeInfo = start !== undefined || end !== undefined
        ? `\n范围: ${start ?? 0} - ${end ?? '末尾'}`
        : '';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档文本内容 (${response.data.length}字符)${rangeInfo}:\n\n${response.data.text}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页眉内容
 */
export const insertHeaderDefinition: ToolDefinition = {
  name: 'wps_word_insert_header',
  description: `设置页眉内容。

使用场景：
- "给文档加个页眉"
- "设置页眉为公司名称"
- "修改第2节的页眉"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '页眉文本内容',
      },
      section: {
        type: 'number',
        description: '节编号（从1开始），默认1',
        default: 1,
      },
    },
    required: ['text'],
  },
};

export const insertHeaderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, section = 1 } = args as { text: string; section?: number };

  if (!text) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '页眉文本不能为空！' }],
      error: '页眉文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertHeader',
      { text, section },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `页眉已设置: "${text}" (第${section}节)` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页眉失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页眉出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页脚内容
 */
export const insertFooterDefinition: ToolDefinition = {
  name: 'wps_word_insert_footer',
  description: `设置页脚内容。

使用场景：
- "给文档加个页脚"
- "设置页脚为页码"
- "修改第1节的页脚"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '页脚文本内容',
      },
      section: {
        type: 'number',
        description: '节编号（从1开始），默认1',
        default: 1,
      },
    },
    required: ['text'],
  },
};

export const insertFooterHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, section = 1 } = args as { text: string; section?: number };

  if (!text) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '页脚文本不能为空！' }],
      error: '页脚文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertFooter',
      { text, section },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `页脚已设置: "${text}" (第${section}节)` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页脚失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页脚出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 自动生成文档目录
 */
export const generateDocTocDefinition: ToolDefinition = {
  name: 'wps_word_generate_doc_toc',
  description: `自动生成文档目录。根据文档中的标题样式自动生成目录。

前提条件：文档中的标题必须使用"标题1"、"标题2"等样式。

使用场景：
- "帮我生成目录"
- "在文档开头插入目录"
- "自动生成文档目录"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      levels: {
        type: 'number',
        description: '目录包含的标题级别数，如3表示包含标题1-3，默认3',
        default: 3,
      },
    },
  },
};

export const generateDocTocHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { levels = 3 } = args as { levels?: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'generateTOC',
      { levels },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `目录已生成（包含标题1-${levels}级）` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `生成目录失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `生成目录出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 插入分节符
 */
export const insertSectionBreakDefinition: ToolDefinition = {
  name: 'wps_word_insert_section_break',
  description: `插入分节符，用于将文档分为不同的节，以便对各节应用不同的页面设置。

使用场景：
- "插入一个分节符"
- "从下一页开始新的一节"
- "在这里分节"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      breakType: {
        type: 'string',
        description: '分节符类型：nextPage(下一页)、continuous(连续)、evenPage(偶数页)、oddPage(奇数页)，默认nextPage',
        default: 'nextPage',
      },
    },
  },
};

export const insertSectionBreakHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { breakType = 'nextPage' } = args as { breakType?: string };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertSectionBreak', // NOTE: macOS未实现，仅Windows支持
      { breakType },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `分节符已插入（类型: ${breakType}）` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入分节符失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入分节符出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置行距
 */
export const setLineSpacingDefinition: ToolDefinition = {
  name: 'wps_word_set_line_spacing',
  description: `设置段落行距。

使用场景：
- "把行距设为1.5倍"
- "设置第3段的行距为2倍"
- "调整行距"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      lineSpacing: {
        type: 'number',
        description: '行距值（如1.0、1.5、2.0等）',
      },
      paragraphIndex: {
        type: 'number',
        description: '段落索引（从0开始），不指定则应用于当前段落或全文',
      },
    },
    required: ['lineSpacing'],
  },
};

export const setLineSpacingHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { lineSpacing, paragraphIndex } = args as { lineSpacing: number; paragraphIndex?: number };

  if (lineSpacing === undefined || lineSpacing <= 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行距值必须为正数！' }],
      error: '行距值无效',
    };
  }

  try {
    const params: Record<string, unknown> = { lineSpacing };
    if (paragraphIndex !== undefined) params.paragraphIndex = paragraphIndex;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setLineSpacing', // NOTE: macOS未实现，仅Windows支持
      params,
      WpsAppType.WRITER
    );

    if (response.success) {
      const target = paragraphIndex !== undefined ? `第${paragraphIndex}段` : '当前段落';
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `行距已设置为 ${lineSpacing} 倍（${target}）` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置行距失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置行距出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有文档管理相关的Tools
 */
export const documentTools: RegisteredTool[] = [
  { definition: getOpenDocumentsDefinition, handler: getOpenDocumentsHandler },
  { definition: switchDocumentDefinition, handler: switchDocumentHandler },
  { definition: openDocumentDefinition, handler: openDocumentHandler },
  { definition: getDocumentTextDefinition, handler: getDocumentTextHandler },
  { definition: insertHeaderDefinition, handler: insertHeaderHandler },
  { definition: insertFooterDefinition, handler: insertFooterHandler },
  { definition: generateDocTocDefinition, handler: generateDocTocHandler },
  { definition: insertSectionBreakDefinition, handler: insertSectionBreakHandler },
  { definition: setLineSpacingDefinition, handler: setLineSpacingHandler },
];

export default documentTools;
