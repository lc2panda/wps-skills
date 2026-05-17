/**
 * Input: Word 工具定义
 * Output: Word 工具注册数组
 * Pos: Word Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word Tools入口 - Word工具汇总模块
 * 整合格式化、内容操作和文档管理的所有Tools
 *
 * 这里导出的是所有Word相关的Tools，注册的时候一把梭哈
 */

import { RegisteredTool } from '../../types/tools';
import { formatTools } from './format';
import { contentTools } from './content';
import { documentTools } from './document';

/**
 * 所有Word相关的Tools
 * 包含：
 * - 格式化Tools: apply_style, set_font, generate_toc, insert_bookmark, set_page_setup
 * - 内容Tools: insert_text, find_replace, insert_table, set_paragraph, get_active_document, insert_image, insert_page_break, set_font_style, insert_comment, set_text_color, get_paragraphs, find_in_document, smart_fill_field, replace_bookmark_content
 * - 文档管理Tools: get_open_documents, switch_document, open_document, get_document_text, insert_header, insert_footer, generate_doc_toc, insert_section_break, set_line_spacing
 */
export const wordTools: RegisteredTool[] = [
  ...formatTools,
  ...contentTools,
  ...documentTools,
];

// 分别导出，方便按需使用
export { formatTools } from './format';
export { contentTools } from './content';
export { documentTools } from './document';

// 导出单独的定义和处理器，方便测试
export {
  applyStyleDefinition,
  applyStyleHandler,
  setFontDefinition,
  setFontHandler,
  generateTocDefinition,
  generateTocHandler,
} from './format';

export {
  insertTextDefinition,
  insertTextHandler,
  findReplaceDefinition,
  findReplaceHandler,
  insertTableDefinition,
  insertTableHandler,
  setParagraphDefinition,
  setParagraphHandler,
  getActiveDocumentDefinition,
  getActiveDocumentHandler,
  insertImageDefinition,
  insertImageHandler,
  insertPageBreakDefinition,
  insertPageBreakHandler,
  setFontStyleDefinition,
  setFontStyleHandler,
  insertCommentDefinition,
  insertCommentHandler,
  setTextColorDefinition,
  setTextColorHandler,
  getParagraphsDefinition,
  getParagraphsHandler,
  findInDocumentDefinition,
  findInDocumentHandler,
  smartFillFieldDefinition,
  smartFillFieldHandler,
  replaceBookmarkContentDefinition,
  replaceBookmarkContentHandler,
} from './content';

export {
  getOpenDocumentsDefinition,
  getOpenDocumentsHandler,
  switchDocumentDefinition,
  switchDocumentHandler,
  openDocumentDefinition,
  openDocumentHandler,
  getDocumentTextDefinition,
  getDocumentTextHandler,
  insertHeaderDefinition,
  insertHeaderHandler,
  insertFooterDefinition,
  insertFooterHandler,
  generateDocTocDefinition,
  generateDocTocHandler,
  insertSectionBreakDefinition,
  insertSectionBreakHandler,
  setLineSpacingDefinition,
  setLineSpacingHandler,
} from './document';

export default wordTools;
