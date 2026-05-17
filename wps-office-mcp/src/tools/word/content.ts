/**
 * Input: Word 内容操作参数
 * Output: 文档内容变更结果
 * Pos: Word 内容工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word内容操作Tools - 文档内容编辑模块
 * 处理文档内容的插入、查找替换、表格、段落格式等操作
 *
 * 包含：
 * - wps_word_insert_text: 插入文本到文档
 * - wps_word_find_replace: 查找替换功能
 * - wps_word_insert_table: 插入表格
 * - wps_word_set_paragraph: 设置段落格式
 * - wps_word_get_active_document: 获取当前文档信息
 * - wps_word_insert_image: 插入图片
 * - wps_word_insert_page_break: 插入分页符
 * - wps_word_set_font_style: 设置字体样式
 * - wps_word_insert_comment: 插入批注
 * - wps_word_set_text_color: 设置文字颜色
 * - wps_word_get_paragraphs: 获取文档段落结构
 * - wps_word_find_in_document: 查找文本返回位置（不替换）
 * - wps_word_smart_fill_field: 智能填写模板字段
 * - wps_word_replace_bookmark_content: 替换书签内容
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
 * 插入文本到文档
 * 可以在光标位置、文档开头或结尾插入文本
 */
export const insertTextDefinition: ToolDefinition = {
  name: 'wps_word_insert_text',
  description: `在Word文档中插入文本。

使用场景：
- "在文档开头加个标题"
- "在光标位置插入这段话"
- "在文档末尾添加备注"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '要插入的文本内容',
      },
      position: {
        type: 'string',
        description: '插入位置，可选值: "cursor"(光标位置), "start"(文档开头), "end"(文档结尾)。默认cursor',
        enum: ['cursor', 'start', 'end'],
      },
      style: {
        type: 'string',
        description: '插入后应用的样式，如 "标题 1"、"正文"',
      },
      new_paragraph: {
        type: 'boolean',
        description: '插入后是否新起一段，默认false',
      },
    },
    required: ['text'],
  },
};

export const insertTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, position, style, new_paragraph } = args as {
    text: string;
    position?: string;
    style?: string;
    new_paragraph?: boolean;
  };

  if (!text || text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '要插入的文本不能为空！' }],
      error: '文本内容为空',
    };
  }

  try {
    // 处理换行符，如果需要新段落
    const finalText = new_paragraph ? text + '\n' : text;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      position: string;
      textLength: number;
    }>(
      'insertText',
      {
        text: finalText,
        position: position || 'cursor',
        style,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const positionText =
        position === 'start' ? '文档开头' :
        position === 'end' ? '文档结尾' : '光标位置';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文本插入成功！\n位置: ${positionText}\n字符数: ${response.data.textLength}${style ? `\n应用样式: ${style}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 查找替换功能
 * 老王最爱用的功能之一，批量替换文本太爽了
 */
export const findReplaceDefinition: ToolDefinition = {
  name: 'wps_word_find_replace',
  description: `在Word文档中查找并替换文本。

使用场景：
- "把所有的'公司'替换成'集团'"
- "把文档里的错别字改过来"
- "批量替换某个词"

支持选项：
- 区分大小写
- 全字匹配
- 全部替换或仅替换一处`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      find_text: {
        type: 'string',
        description: '要查找的文本',
      },
      replace_text: {
        type: 'string',
        description: '替换为的文本，如果只是查找不替换，可以不填',
      },
      replace_all: {
        type: 'boolean',
        description: '是否全部替换，默认true',
      },
      match_case: {
        type: 'boolean',
        description: '是否区分大小写，默认false',
      },
      match_whole_word: {
        type: 'boolean',
        description: '是否全字匹配，默认false',
      },
    },
    required: ['find_text'],
  },
};

export const findReplaceHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { find_text, replace_text, replace_all, match_case, match_whole_word } = args as {
    find_text: string;
    replace_text?: string;
    replace_all?: boolean;
    match_case?: boolean;
    match_whole_word?: boolean;
  };

  if (!find_text || find_text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '查找文本不能为空！' }],
      error: '查找文本为空',
    };
  }

  try {
    // 如果没有替换文本，就只是查找
    const isReplaceMode = replace_text !== undefined && replace_text !== null;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      findText: string;
      replaceText: string;
      count: number;
    }>(
      'findReplace',
      {
        findText: find_text,
        replaceText: isReplaceMode ? replace_text : '',
        replaceAll: replace_all !== false, // 默认true
        matchCase: match_case || false,
        matchWholeWord: match_whole_word || false,
        replaceMode: isReplaceMode,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const result = response.data;

      if (isReplaceMode) {
        if (result.count === 0) {
          return {
            id: uuidv4(),
            success: true,
            content: [
              {
                type: 'text',
                text: `未找到 "${find_text}"，没有进行替换`,
              },
            ],
          };
        }
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `替换完成！\n查找: "${find_text}"\n替换为: "${replace_text}"\n替换了 ${result.count} 处`,
            },
          ],
        };
      } else {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `查找完成！\n"${find_text}" 在文档中出现了 ${result.count} 次`,
            },
          ],
        };
      }
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `查找替换失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `查找替换出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 插入表格到文档光标位置
 */
export const insertTableDefinition: ToolDefinition = {
  name: 'wps_word_insert_table',
  description: '在Word文档光标位置插入表格',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      rows: { type: 'number', description: '表格行数' },
      cols: { type: 'number', description: '表格列数' },
    },
    required: ['rows', 'cols'],
  },
};

export const insertTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { rows, cols } = args as { rows: number; cols: number };
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'insertTable',
      { rows, cols },
      WpsAppType.WRITER
    );
    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `已插入 ${rows}×${cols} 表格` }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入表格失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入表格出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置当前段落格式
 */
export const setParagraphDefinition: ToolDefinition = {
  name: 'wps_word_set_paragraph',
  description: '设置当前段落格式（对齐方式、行间距等）',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      alignment: { type: 'string', description: '对齐方式: left/center/right/justify' },
      lineSpacing: { type: 'number', description: '行间距倍数，如1.5、2' },
    },
  },
};

export const setParagraphHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const params = args as { alignment?: string; lineSpacing?: number };
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'setParagraph',
      params,
      WpsAppType.WRITER
    );
    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: '段落格式已设置' }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置段落失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置段落出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取当前活动文档信息
 */
export const getActiveDocumentDefinition: ToolDefinition = {
  name: 'wps_word_get_active_document',
  description: '获取当前WPS Writer活动文档的基本信息',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getActiveDocumentHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      name: string;
      path: string;
      pageCount: number;
      wordCount: number;
    }>(
      'getActiveDocument',
      {},
      WpsAppType.WRITER
    );
    if (response.success && response.data) {
      const d = response.data;
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `当前文档: ${d.name}\n路径: ${d.path}\n页数: ${d.pageCount}\n字数: ${d.wordCount}` }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档信息失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取文档信息出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 在文档中插入图片
 */
export const insertImageDefinition: ToolDefinition = {
  name: 'wps_word_insert_image',
  description: `在Word文档中插入图片。

使用场景：
- "在文档中插入一张图片"
- "把这个截图放到文档里"
- "在光标位置插入logo"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      imagePath: {
        type: 'string',
        description: '图片文件路径',
      },
      width: {
        type: 'number',
        description: '图片宽度（磅），可选',
      },
      height: {
        type: 'number',
        description: '图片高度（磅），可选',
      },
    },
    required: ['imagePath'],
  },
};

export const insertImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { imagePath, width, height } = args as {
    imagePath: string;
    width?: number;
    height?: number;
  };

  if (!imagePath || imagePath.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '图片路径不能为空！' }],
      error: '图片路径为空',
    };
  }

  try {
    // 跨平台参数对齐：macOS/Windows 底层均优先读取 params.path，同时保留 imagePath/filePath 别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertImage',
      { imagePath, path: imagePath, filePath: imagePath, width, height },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片插入成功！\n路径: ${imagePath}${width ? `\n宽度: ${width}磅` : ''}${height ? `\n高度: ${height}磅` : ''}`,
          },
        ],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有内容操作相关的Tools
 */
export const insertPageBreakDefinition: ToolDefinition = {
  name: 'wps_word_insert_page_break',
  description: '在文档光标位置插入分页符',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const insertPageBreakHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'insertPageBreak',
      {},
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? '分页符已插入' : (response.data as any)?.message || '插入失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入分页符出错: ${e.message}` }], error: e.message };
  }
};

export const setFontStyleDefinition: ToolDefinition = {
  name: 'wps_word_set_font_style',
  description: '设置选中文字的字体样式属性',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      fontName: { type: 'string', description: '字体名称' },
      fontSize: { type: 'number', description: '字号' },
      bold: { type: 'boolean', description: '是否加粗' },
      italic: { type: 'boolean', description: '是否斜体' },
    },
    required: [],
  },
};

export const setFontStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'setFont',
      args,
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? '字体已设置' : (response.data as any)?.message || '设置失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置字体出错: ${e.message}` }], error: e.message };
  }
};

/**
 * 插入批注到文档选中内容
 */
export const insertCommentDefinition: ToolDefinition = {
  name: 'wps_word_insert_comment',
  description: '在Word文档选中内容处插入批注',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: { type: 'string', description: '批注内容' },
    },
    required: ['text'],
  },
};

export const insertCommentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text } = args as { text: string };
  if (!text || text.trim() === '') {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '批注内容不能为空！' }], error: '批注内容为空' };
  }
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'addComment',
      { text },
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? '批注已插入' : (response.data as any)?.message || '插入失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入批注出错: ${e.message}` }], error: e.message };
  }
};

/**
 * 设置选中文字的颜色
 */
export const setTextColorDefinition: ToolDefinition = {
  name: 'wps_word_set_text_color',
  description: '设置Word文档中选中文字的颜色',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      color: { type: 'string', description: '颜色值，如 "#FF0000"、"red"' },
    },
    required: ['color'],
  },
};

export const setTextColorHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { color } = args as { color: string };
  if (!color || color.trim() === '') {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '颜色值不能为空！' }], error: '颜色值为空' };
  }
  try {
    const response = await wpsClient.executeMethod<{ success: boolean; message: string }>(
      'setTextColor', // NOTE: macOS未实现，仅Windows支持
      { color },
      WpsAppType.WRITER
    );
    return {
      id: uuidv4(),
      success: response.success,
      content: [{ type: 'text', text: response.success ? `文字颜色已设置为 ${color}` : (response.data as any)?.message || '设置失败' }],
    };
  } catch (e: any) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置文字颜色出错: ${e.message}` }], error: e.message };
  }
};

/**
 * 获取文档段落结构
 * 返回段落的文本、样式、位置信息，用于了解文档结构和识别填写位置
 */
export const getParagraphsDefinition: ToolDefinition = {
  name: 'wps_word_get_paragraphs',
  description: `获取Word文档的段落结构信息，返回每段的文本、样式和字符位置。

使用场景：
- "了解文档的结构"
- "查看文档有哪些段落"
- "帮我看看模板里有哪些需要填写的位置"
- 在填写模板前，先读取文档结构以识别填写位置

返回信息包括：段落索引、文本内容、样式名称、字符起止位置。
支持分页获取（startParagraph/endParagraph），默认返回前50段。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      start_paragraph: {
        type: 'number',
        description: '起始段落索引（从1开始），默认1',
      },
      end_paragraph: {
        type: 'number',
        description: '结束段落索引，默认为起始+49',
      },
    },
    required: [],
  },
};

export const getParagraphsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { start_paragraph, end_paragraph } = args as {
    start_paragraph?: number;
    end_paragraph?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      paragraphs: Array<{ index: number; text: string; style: string; start: number; end: number }>;
      totalCount: number;
      returnedCount: number;
    }>(
      'getDocumentParagraphs',
      {
        startParagraph: start_paragraph,
        endParagraph: end_paragraph,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const data = response.data;
      const paragraphs = data.paragraphs;
      const totalCount = data.totalCount;
      const returnedCount = data.returnedCount;
      const lines = paragraphs.map(
        (p: { index: number; text: string; style: string; start: number; end: number }) => `[${p.index}] (${p.style}) ${p.text}`
      );
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档段落结构（共${totalCount}段，返回${returnedCount}段）：\n${lines.join('\n')}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取段落结构失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取段落结构出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 查找文本返回位置（不替换）
 * 返回匹配文本的位置信息和上下文，供后续精确操作使用
 */
export const findInDocumentDefinition: ToolDefinition = {
  name: 'wps_word_find_in_document',
  description: `在Word文档中查找文本并返回位置信息，不执行替换操作。

使用场景：
- "找一下'项目名称'在文档的哪个位置"
- "看看文档里有哪些地方需要填写"
- 在使用smartFillField之前，先用此工具定位关键字

返回信息包括：匹配文本、字符起止位置、所在段落索引、上下文（前后50字符）。
与findReplace不同，此工具仅查找不替换，返回位置信息供后续操作使用。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      find_text: {
        type: 'string',
        description: '要查找的文本',
      },
      match_case: {
        type: 'boolean',
        description: '是否区分大小写，默认false',
      },
      match_whole_word: {
        type: 'boolean',
        description: '是否全字匹配，默认false',
      },
      max_results: {
        type: 'number',
        description: '最大返回结果数，默认20',
      },
    },
    required: ['find_text'],
  },
};

export const findInDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { find_text, match_case, match_whole_word, max_results } = args as {
    find_text: string;
    match_case?: boolean;
    match_whole_word?: boolean;
    max_results?: number;
  };

  if (!find_text || find_text.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '查找文本不能为空！' }],
      error: '查找文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      results: Array<{ text: string; start: number; end: number; paragraphIndex: number; context: string }>;
      count: number;
      findText: string;
    }>(
      'findInDocument',
      {
        findText: find_text,
        matchCase: match_case || false,
        matchWholeWord: match_whole_word || false,
        maxResults: max_results || 20,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const data = response.data;
      const results = data.results;
      const count = data.count;
      const findText = data.findText;
      if (count === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: `未找到 "${findText}"` }],
        };
      }
      const lines = results.map(
        (r: { text: string; start: number; end: number; paragraphIndex: number; context: string }) => `  段落${r.paragraphIndex} [${r.start}-${r.end}]: ...${r.context}...`
      );
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `找到 "${findText}" 共${count}处：\n${lines.join('\n')}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `查找失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `查找出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 智能填写模板字段
 * 自动判断关键字附近的填写模式（下划线/冒号后/标签后/占位符），在正确位置插入内容并保持格式
 */
export const smartFillFieldDefinition: ToolDefinition = {
  name: 'wps_word_smart_fill_field',
  description: `智能填写Word模板中的字段。自动识别关键字附近的填写模式，在正确位置插入内容并保持原有格式。

使用场景：
- "把项目名称填写为'XX信息化项目'"
- "填写建设单位为'XX公司'"
- "模板里的项目编号填上'2026-001'"
- 任何需要在模板文档中"填写"而非"替换"的场景

支持的填写模式（fillMode，默认auto自动判断）：
- auto: 自动判断（推荐）
- underline: 关键字后有下划线___，替换下划线为填写内容
- afterColon: 关键字后有冒号（：或:），在冒号后插入
- afterLabel: 关键字是标签，直接在关键字后插入
- placeholder: 关键字被{}或【】包裹，替换整个占位符

重要：模板填写场景应优先使用此工具，而非findReplace。findReplace会删除关键字本身并可能破坏格式。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      keyword: {
        type: 'string',
        description: '要填写的关键字（如"项目名称"、"建设单位"）',
      },
      value: {
        type: 'string',
        description: '要填写的值（如"XX信息化项目"）',
      },
      fill_mode: {
        type: 'string',
        description: '填写模式: auto(自动判断), underline(下划线), afterColon(冒号后), afterLabel(标签后), placeholder(占位符)。默认auto',
        enum: ['auto', 'underline', 'afterColon', 'afterLabel', 'placeholder'],
      },
    },
    required: ['keyword', 'value'],
  },
};

export const smartFillFieldHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { keyword, value, fill_mode } = args as {
    keyword: string;
    value: string;
    fill_mode?: string;
  };

  if (!keyword || keyword.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '关键字不能为空！' }],
      error: '关键字为空',
    };
  }

  if (!value) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '填写值不能为空！' }],
      error: '填写值为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      keyword: string;
      value: string;
      fillMode: string;
      result: string;
    }>(
      'smartFillField',
      {
        keyword,
        value,
        fillMode: fill_mode || 'auto',
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const data = response.data;
      const fillMode = data.fillMode;
      const result = data.result;
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `填写完成！\n关键字: "${keyword}"\n填写值: "${value}"\n填写模式: ${fillMode}\n结果: ${result}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `填写失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `填写出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 替换书签内容
 * 通过书签名定位并替换内容，保持书签位置格式
 */
export const replaceBookmarkContentDefinition: ToolDefinition = {
  name: 'wps_word_replace_bookmark_content',
  description: `替换Word文档中书签的内容。通过书签名定位，替换书签范围内的文本，保持原有格式。

使用场景：
- "把书签'project_name'的内容改为'XX项目'"
- 模板文档使用书签标记填写位置时使用

注意：替换后书签会自动重建，不会丢失。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '书签名称',
      },
      text: {
        type: 'string',
        description: '要替换为的文本内容',
      },
    },
    required: ['name', 'text'],
  },
};

export const replaceBookmarkContentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, text } = args as {
    name: string;
    text: string;
  };

  if (!name || name.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '书签名称不能为空！' }],
      error: '书签名称为空',
    };
  }

  if (!text && text !== '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '替换文本不能为空！' }],
      error: '替换文本为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      name: string;
      text: string;
      start: number;
      end: number;
    }>(
      'replaceBookmarkContent',
      {
        name,
        text,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `书签内容替换完成！\n书签: "${name}"\n新内容: "${text}"`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `替换书签内容失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `替换书签内容出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有内容操作相关的Tools
 */
export const contentTools: RegisteredTool[] = [
  { definition: insertTextDefinition, handler: insertTextHandler },
  { definition: findReplaceDefinition, handler: findReplaceHandler },
  { definition: insertTableDefinition, handler: insertTableHandler },
  { definition: setParagraphDefinition, handler: setParagraphHandler },
  { definition: getActiveDocumentDefinition, handler: getActiveDocumentHandler },
  { definition: insertImageDefinition, handler: insertImageHandler },
  { definition: insertPageBreakDefinition, handler: insertPageBreakHandler },
  { definition: setFontStyleDefinition, handler: setFontStyleHandler },
  { definition: insertCommentDefinition, handler: insertCommentHandler },
  { definition: setTextColorDefinition, handler: setTextColorHandler },
  { definition: getParagraphsDefinition, handler: getParagraphsHandler },
  { definition: findInDocumentDefinition, handler: findInDocumentHandler },
  { definition: smartFillFieldDefinition, handler: smartFillFieldHandler },
  { definition: replaceBookmarkContentDefinition, handler: replaceBookmarkContentHandler },
];

export default contentTools;
