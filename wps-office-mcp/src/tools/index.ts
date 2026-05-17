/**
 * Input: Tool 定义集合
 * Output: Tool 注册数组
 * Pos: MCP Tools 总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tools总入口 - MCP工具汇总注册模块
 * 整合Excel、Word、PPT、Common的所有Tools
 *
 * 使用方法：
 * import { allTools } from './tools';
 * toolRegistry.registerAll(allTools);
 *
 * 或者按需导入：
 * import { excelTools, wordTools, pptTools, commonTools } from './tools';
 */

import { RegisteredTool } from '../types/tools';
import { excelTools } from './excel';
import { wordTools } from './word';
import { pptTools } from './ppt';
import { commonTools } from './common';

/**
 * 所有MCP Tools集合（共230个，另有12个内置工具在mcp-server.ts中注册，全局共242个）
 *
 * Excel (82个):
 *   公式(6): set_formula, generate_formula, diagnose_formula, set_array_formula, recalculate, auto_sum
 *   数据(12): read_range, write_range, clean_data, remove_duplicates, sort_range, find_replace, insert_row, add_comment, protect_sheet, set_conditional_format, hide_column, protect_workbook
 *   图表(4): create_chart, update_chart, export_chart_as_image, export_range_as_image
 *   透视表(2): create_pivot_table, update_pivot_table
 *   工作表(16): create_sheet, delete_sheet, rename_sheet, copy_sheet, get_sheet_list, switch_sheet, move_sheet, get_selection, delete_row, insert_column, delete_column, freeze_panes, auto_fill, set_named_range, set_zoom, save_workbook
 *   格式化(10): set_cell_format, set_cell_style, set_border, set_number_format, merge_cells, unmerge_cells, set_column_width, set_row_height, auto_fit_column, auto_fit_row
 *   工作簿(10): open_workbook, get_open_workbooks, switch_workbook, close_workbook, create_workbook, get_cell_value, set_cell_value, get_formula, get_cell_info, clear_range
 *   高级数据(7): auto_filter, copy_range, paste_range, fill_series, transpose, text_to_columns, subtotal
 *
 * Word (27个):
 *   格式化(5): apply_style, set_font, generate_toc, insert_bookmark, set_page_setup
 *   内容(10): insert_text, find_replace, insert_table, set_paragraph, get_active_document, insert_image, add_comment, insert_hyperlink, get_word_count, insert_page_break
 *   文档管理(9): get_open_documents, switch_document, open_document, get_document_text, insert_header, insert_footer, generate_doc_toc, create_document, save_document
 *
 * PPT (112个):
 *   幻灯片(5): add_slide, beautify, unify_font, set_font_color, align_objects
 *   幻灯片操作(22): delete_slide, duplicate_slide, move_slide, get_slide_count, get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes, add_shape, set_shape_style, add_textbox, set_slide_title, insert_image, set_shape_text, set_animation, set_background, set_slide_size, set_transition, add_chart, set_shape_fill, add_speaker_notes
 *   演示文稿管理(8): create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation, set_slide_theme, copy_slide, insert_slide_image
 *   文本框(7): delete_textbox, get_textboxes, set_textbox_text, set_textbox_style, get_slide_title, set_slide_subtitle, set_slide_content
 *
 * Common (9个):
 *   转换(2): convert_to_pdf, convert_format
 *   （其余7个工具详见 common/ 子模块）
 *
 * 内置工具（12个，在 mcp-server.ts 中注册）:
 *   wps_check_connection, wps_get_active_document, wps_insert_text, wps_get_active_workbook,
 *   wps_get_cell_value, wps_set_cell_value, wps_get_active_presentation, wps_execute_method,
 *   wps_cache_data, wps_get_cached_data, wps_list_cache, wps_clear_cache
 */
export const allTools: RegisteredTool[] = [
  ...excelTools,
  ...wordTools,
  ...pptTools,
  ...commonTools,
];

// 按应用类型分别导出
export { excelTools } from './excel';
export { wordTools } from './word';
export { pptTools } from './ppt';
export { commonTools } from './common';

// Excel相关导出
export {
  formulaTools,
  dataTools,
  setFormulaDefinition,
  setFormulaHandler,
  generateFormulaDefinition,
  generateFormulaHandler,
  diagnoseFormulaDefinition,
  diagnoseFormulaHandler,
  readRangeDefinition,
  readRangeHandler,
  writeRangeDefinition,
  writeRangeHandler,
  cleanDataDefinition,
  cleanDataHandler,
  removeDuplicatesDefinition,
  removeDuplicatesHandler,
} from './excel';

// Word相关导出
export {
  formatTools,
  contentTools,
  applyStyleDefinition,
  applyStyleHandler,
  setFontDefinition,
  setFontHandler,
  generateTocDefinition,
  generateTocHandler,
  insertTextDefinition,
  insertTextHandler,
  findReplaceDefinition,
  findReplaceHandler,
} from './word';

// PPT相关导出
export {
  slideTools,
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
} from './ppt';

// Common相关导出
export {
  convertTools,
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './common';

/**
 * 获取所有Tool的数量
 */
export const getToolCount = (): number => allTools.length;

/**
 * 获取按应用分类的Tool数量
 */
export const getToolCountByApp = (): { excel: number; word: number; ppt: number; common: number } => ({
  excel: excelTools.length,
  word: wordTools.length,
  ppt: pptTools.length,
  common: commonTools.length,
});

export default allTools;
