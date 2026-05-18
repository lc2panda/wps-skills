/**
 * WPS Office COM Actions 索引
 * 共 238 个 COM Actions（afterColon/afterLabel 已移除，改用 smartFillField fillMode）
 *
 * 工具名称映射说明：
 * - Gateway 使用短名称（如 setFont、addSlide）进行索引和搜索
 * - 执行时通过 wpsClient.executeMethod() 动态调用 WPS COM 方法
 * - WPS Client 支持动态方法名执行，无需显式映射
 *
 * 验证状态：
 * - verified: 已通过实际测试验证可用
 * - indexed: 仅索引，无直接实现（通过 executeMethod 动态调用）
 * - stub: 有占位实现，尚未完整测试
 */

// 验证状态枚举
export type VerificationStatus = 'verified' | 'indexed' | 'stub';

export interface ToolIndexItem {
  name: string;
  description: string;
  keywords: string[];
  category: string;
  appType: WpsAppType;
  paramsSchema: Record<string, ToolParamSchema>;
  status?: VerificationStatus;
}

import { wpsClient } from '../../client/wps-client';
import { ToolCallResult, ToolHandler, ToolInputSchema } from '../../types/tools';
import { WpsAppType } from '../../types/wps';
import { allTools } from '../index';

export interface ToolParamSchema {
  type: 'string' | 'number' | 'boolean' | 'object' | 'array';
  description?: string;
  required: boolean;
  default?: unknown;
  enum?: string[];
}

// ==================== Schema 自动生成 ====================
// 从注册工具定义自动生成 COM_ACTIONS 的 paramsSchema，消除双重定义
// 纯 PS1 工具（无对应注册工具）仍使用硬编码 paramsSchema 作为兜底

function reverseParamMap(forward: Record<string, string>): Record<string, string> {
  const rev: Record<string, string> = {};
  for (const [k, v] of Object.entries(forward)) rev[v] = k;
  return rev;
}

function inputSchemaToParamsSchema(
  inputSchema: ToolInputSchema,
  revMap?: Record<string, string>
): Record<string, ToolParamSchema> {
  const requiredSet = new Set(inputSchema.required || []);
  const params: Record<string, ToolParamSchema> = {};
  for (const [key, prop] of Object.entries(inputSchema.properties || {})) {
    const mappedKey = revMap?.[key] ?? key;
    params[mappedKey] = {
      type: prop.type,
      description: prop.description,
      required: requiredSet.has(key),
    };
    if (prop.enum) params[mappedKey].enum = prop.enum;
  }
  return params;
}

// ==================== 工具验证状态 ====================
// verified: PowerShell 脚本 + wps-client.ts 中已完整实现并测试
// indexed: 仅在 Gateway 索引中存在，通过 executeMethod 动态调用
// stub: 有占位实现，尚未完整测试

// 将 wps_xxx_xxx 风格的短名称转为驼峰 (set_font → setFont)
function toCamelCase(snake: string): string {
  return snake.replace(/_(.)/g, (_, c) => c.toUpperCase());
}

// 工具注册名 → appType 映射（从 name 前缀推断，值与 WpsAppType 一致）
function appTypeFromName(name: string): string {
  const prefix = name.match(/^wps_(word|excel|ppt|common)_/i)?.[1]?.toLowerCase();
  if (prefix === 'excel') return 'et';
  if (prefix === 'ppt') return 'wpp';
  return 'wps';
}

// 逐工具参数名映射：COM_ACTIONS 的 camelCase 参数名 → TS handler 的期望参数名
// 仅在 COM schema 与 handler inputSchema 参数名不一致时需要
const HANDLER_PARAM_MAP: Record<string, Record<string, string>> = {
  setFont: { fontName: 'font_name', fontSize: 'font_size' },
  applyStyle: { styleName: 'style_name' },
  beautifySlide: { slideIndex: 'slide_index', colorScheme: 'color_scheme', beautifyAll: 'beautify_all' },
  findInDocument: { text: 'find_text', matchCase: 'match_case', matchWholeWord: 'match_whole_word', maxResults: 'max_results' },
  replaceBookmarkContent: { content: 'text' },
  smartFillField: { fillMode: 'fill_mode' },
  getDocumentParagraphs: { startParagraph: 'start_paragraph', endParagraph: 'end_paragraph' },
};

// 无工具需要跳过 handler 路由（所有 COM schema 参数名已与 handler inputSchema 对齐）
const HANDLER_SKIP = new Set<string>();

// 从注册工具中预先构建 handler 映射表：COM 短名 + appType → TS handler
// 使用 "name|appType" 复合键解决跨应用同名冲突（如 insertImage 同时存在于 Word 和 PPT）
// 使 executeTool 能根据 TOOLS_INDEX 中的 appType 精确路由到正确的 handler
const HANDLER_MAP = new Map<string, ToolHandler>();
for (const tool of allTools) {
  const shortName = tool.definition.name.replace(/^wps_(word|excel|ppt|common)_/i, '');
  const camelName = toCamelCase(shortName);
  const appType = appTypeFromName(tool.definition.name);
  const key = `${camelName}|${appType}`;
  if (HANDLER_MAP.has(key)) {
    console.warn(`[HANDLER_MAP] 重复键：${key}（来自 ${tool.definition.name}）`);
  }
  HANDLER_MAP.set(key, tool.handler);
}
// 处理命名不遵循 wps_{word|excel|ppt|common}_ 约定的工具
// wps_convert_to_pdf 无 common 段，toCamelCase 产生 wpsConvertToPdf 而非 convertToPDF
const pdfTool = allTools.find(t => t.definition.name === 'wps_convert_to_pdf');
if (pdfTool) {
  HANDLER_MAP.set('convertToPDF|wps', pdfTool.handler);
}
const paragraphsTool = allTools.find(t => t.definition.name === 'wps_word_get_paragraphs');
if (paragraphsTool) {
  HANDLER_MAP.set('getDocumentParagraphs|wps', paragraphsTool.handler);
}

// 构建 schema 映射表：COM 短名 → paramsSchema
// 优先从注册工具定义自动生成，消除 COM_ACTIONS 的 paramsSchema 双重定义
const SCHEMA_MAP = new Map<string, Record<string, ToolParamSchema>>();
for (const tool of allTools) {
  const shortName = tool.definition.name.replace(/^wps_(word|excel|ppt|common)_/i, '');
  const camelName = toCamelCase(shortName);
  const revMap = reverseParamMap(HANDLER_PARAM_MAP[camelName] || {});
  const schema = inputSchemaToParamsSchema(tool.definition.inputSchema, revMap);
  if (Object.keys(schema).length > 0) SCHEMA_MAP.set(camelName, schema);
}
if (paragraphsTool) {
  const schema = inputSchemaToParamsSchema(paragraphsTool.definition.inputSchema);
  if (Object.keys(schema).length > 0) SCHEMA_MAP.set('getDocumentParagraphs', schema);
}

const VERIFIED_TOOLS = new Set([
  // === Common ===
  'ping', 'wireCheck', 'getAppInfo', 'getContext', 'getOpenDocuments', 'getOpenPresentations',
  'switchDocument', 'switchPresentation', 'switchWorkbook', 'convertToPDF', 'convertFormat',
  'trim', 'underline', 'placeholder',
  // === Word ===
  'getActiveDocument', 'getDocumentText', 'getSelectedText', 'setSelectedText', 'save',
  'saveAs', 'openFile', 'openDocument', 'createDocument', 'closeDocument',
  'setFont', 'setParagraph', 'applyStyle', 'generateTOC',
  'insertBookmark', 'getBookmarks', 'replaceBookmarkContent',
  'findInDocument', 'findReplace',
  'getDocumentParagraphs', 'getDocumentStats',
  'insertTable', 'insertImage', 'setPageSetup', 'insertHeader', 'insertFooter',
  'insertHyperlink', 'insertPageBreak', 'setHyperlink',
  'smartFillField', 'addComment', 'getComments', 'insertText',
  'switchDocument', 'getOpenDocuments',
  // === Excel ===
  'getActiveWorkbook', 'getCellValue', 'setCellValue', 'getRangeData', 'setRangeData',
  'setFormula', 'getFormula', 'setArrayFormula', 'diagnoseFormula',
  'createSheet', 'deleteSheet', 'renameSheet', 'copySheet', 'getSheetList', 'switchSheet', 'moveSheet',
  'createPivotTable', 'updatePivotTable',
  'createChart', 'updateChart', 'createDonutChart', 'createFlowChart', 'createGauge', 'createGrid', 'createKpiCards', 'createMiniCharts',
  'setCellFormat', 'setCellStyle', 'setBorder', 'copyFormat', 'clearFormats',
  'addConditionalFormat', 'removeConditionalFormat', 'getConditionalFormats',
  'addDataValidation', 'removeDataValidation', 'getDataValidations',
  'mergeCells', 'unmergeCells', 'setColumnWidth', 'setRowHeight',
  'autoFitColumn', 'autoFitRow', 'autoFitAll', 'setNumberFormat', 'wrapText', 'setPrintArea',
  'getSelection', 'clearRange', 'insertRows', 'insertColumns', 'deleteRows', 'deleteColumns',
  'hideRows', 'hideColumns', 'showRows', 'showColumns', 'groupRows', 'groupColumns',
  'freezePanes', 'unfreezePanes', 'findInSheet', 'replaceInSheet',
  'copyRange', 'pasteRange', 'fillSeries', 'transpose', 'textToColumns', 'subtotal',
  'createNamedRange', 'deleteNamedRange', 'getNamedRanges',
  'addCellComment', 'deleteCellComment', 'getCellComments',
  'protectSheet', 'unprotectSheet', 'protectWorkbook',
  'insertExcelImage', 'lockCells', 'openWorkbook', 'getOpenWorkbooks', 'createWorkbook',
  'cleanData', 'removeDuplicates', 'sortRange', 'autoFilter', 'getCellInfo',
  'refreshLinks', 'consolidate', 'calculateSheet', 'getExcelContext', 'generateFormula',
  'remove_duplicates', 'unify_date', 'closeWorkbook',
  // === PPT ===
  'getActivePresentation', 'createPresentation', 'openPresentation', 'closePresentation',
  'getSlideCount', 'addSlide', 'deleteSlide', 'duplicateSlide', 'moveSlide', 'switchSlide',
  'getSlideInfo', 'getSlideTitle', 'getSlideNotes', 'setSlideTitle', 'setSlideSubtitle',
  'setSlideContent', 'setSlideNotes', 'setSlideBackground', 'setBackgroundColor',
  'setBackgroundGradient', 'setBackgroundImage',
  'setSlideTransition', 'removeSlideTransition', 'applyTransitionToAll',
  'addAnimation', 'removeAnimation', 'setAnimationOrder', 'getAnimations',
  'addAnimationPreset', 'addEmphasisAnimation',
  'beautifySlide', 'beautifyAllSlides', 'autoBeautifySlide', 'unifyFont',
  'addShape', 'deleteShape', 'duplicateShape', 'getShapes', 'alignShapes', 'groupShapes', 'distributeShapes',
  'addTextBox', 'setTextBoxText', 'getTextBoxes', 'setTextBoxStyle', 'deleteTextBox',
  'insertPptImage', 'deletePptImage', 'setImageStyle',
  'insertPptChart', 'setPptChartData', 'setPptChartStyle',
  'insertPptTable', 'getPptTableCell', 'setPptTableCell', 'setPptTableStyle',
  'setPptTableCellStyle', 'setPptTableRowStyle',
  'setShapeStyle', 'setShapeBorder', 'setShapeShadow', 'setShapeGradient', 'setShapePosition',
  'setShapeZOrder', 'setShapeFullStyle', 'setShapeRoundness', 'setShapeTransparency', 'setShapeText',
  'addConnector', 'addArrow', 'setSlideLayout', 'setSlideNumber', 'setPptDateTime', 'setPptFooter',
  'setMasterBackground', 'getSlideMaster', 'applyColorScheme',
  'addPptHyperlink', 'removePptHyperlink', 'findPptText', 'replacePptText',
  'addTitleDecoration', 'addPageIndicator', 'addMasterElement',
  'startSlideShow', 'endSlideShow', 'autoLayout', 'smartDistribute',
  'create3DText', 'set3DDepth', 'set3DMaterial', 'set3DRotation',
  // Export 工具 (Issue #15)
  'exportChartAsImage', 'exportRangeAsImage', 'exportSlideAsImage',
]);

// 计算索引中每个工具的验证状态
function getToolStatus(toolName: string): VerificationStatus {
  return VERIFIED_TOOLS.has(toolName) ? 'verified' : 'indexed';
}

const COM_ACTIONS: ToolIndexItem[] = [
  // Word 操作 (~35)
  { name: 'getActiveDocument', description: '获取当前活动的 WPS 文档信息', keywords: ['文档', '当前'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'getDocumentText', description: '获取文档全部文本内容', keywords: ['文本', '内容'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'getSelectedText', description: '获取选中的文本', keywords: ['选中', '选区'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'setSelectedText', description: '设置选中的文本内容', keywords: ['设置', '文本'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '文本内容', required: true } } },
  { name: 'save', description: '保存当前文档', keywords: ['保存'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'saveAs', description: '另存文档', keywords: ['另存'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { filePath: { type: 'string', description: '文件路径', required: true }, format: { type: 'string', description: '格式', required: false } } },
  { name: 'openFile', description: '打开文档', keywords: ['打开'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { filePath: { type: 'string', description: '文件路径', required: true } } },
  { name: 'createDocument', description: '新建空白文档', keywords: ['新建'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'closeDocument', description: '关闭当前文档', keywords: ['关闭'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'openDocument', description: '打开文档', keywords: ['打开'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { filePath: { type: 'string', description: '文件路径', required: true } } },
  { name: 'setFont', description: '设置字体格式', keywords: ['字体', '格式'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { fontName: { type: 'string', description: '字体名称', required: false }, fontSize: { type: 'number', description: '字号', required: false }, bold: { type: 'boolean', description: '加粗', required: false }, italic: { type: 'boolean', description: '斜体', required: false } } },
  { name: 'setParagraph', description: '设置段落格式', keywords: ['段落', '对齐'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { alignment: { type: 'string', description: '对齐方式', required: false } } },
  { name: 'applyStyle', description: '应用文档样式', keywords: ['样式'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { styleName: { type: 'string', description: '样式名称', required: true } } },
  { name: 'generateTOC', description: '生成目录', keywords: ['目录'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'insertBookmark', description: '插入书签', keywords: ['书签'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { name: { type: 'string', description: '书签名称', required: true } } },
  { name: 'getBookmarks', description: '获取文档中所有书签', keywords: ['书签'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'replaceBookmarkContent', description: '替换书签内容', keywords: ['书签', '替换'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { name: { type: 'string', description: '书签名称', required: true }, content: { type: 'string', description: '替换内容', required: true } } },
  { name: 'findInDocument', description: '在文档中查找文本', keywords: ['查找'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '查找文本', required: true }, matchCase: { type: 'boolean', description: '区分大小写，默认false', required: false }, matchWholeWord: { type: 'boolean', description: '全词匹配，默认false', required: false }, maxResults: { type: 'number', description: '最大返回结果数，默认10', required: false } } },
  { name: 'findReplace', description: '查找替换', keywords: ['替换'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { find: { type: 'string', description: '查找内容', required: true }, replace: { type: 'string', description: '替换为', required: true } } },
  { name: 'getDocumentParagraphs', description: '获取文档段落列表', keywords: ['段落'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { startParagraph: { type: 'number', description: '起始段落索引（从1开始），默认1', required: false }, endParagraph: { type: 'number', description: '结束段落索引，默认起始+49', required: false } } },
  { name: 'getDocumentStats', description: '获取文档统计信息', keywords: ['统计'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'insertTable', description: '插入表格', keywords: ['表格'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { rows: { type: 'number', description: '行数', required: true }, cols: { type: 'number', description: '列数', required: true } } },
  { name: 'insertImage', description: '插入图片', keywords: ['图片'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { imagePath: { type: 'string', description: '图片文件路径', required: true }, width: { type: 'number', description: '图片宽度（磅），可选', required: false }, height: { type: 'number', description: '图片高度（磅），可选', required: false } } },
  { name: 'setPageSetup', description: '设置页面布局', keywords: ['页面'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { orientation: { type: 'string', description: '页面方向: "portrait"(纵向) 或 "landscape"(横向)', enum: ['portrait', 'landscape'], required: false }, marginTop: { type: 'number', description: '上边距（磅值）', required: false }, marginBottom: { type: 'number', description: '下边距（磅值）', required: false }, marginLeft: { type: 'number', description: '左边距（磅值）', required: false }, marginRight: { type: 'number', description: '右边距（磅值）', required: false } } },
  { name: 'insertHeader', description: '插入页眉', keywords: ['页眉'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '页眉内容', required: true } } },
  { name: 'insertFooter', description: '插入页脚', keywords: ['页脚'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '页脚内容', required: true } } },
  { name: 'insertHyperlink', description: '插入超链接', keywords: ['超链接'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '链接文本', required: true }, address: { type: 'string', description: '链接地址', required: true } } },
  { name: 'insertPageBreak', description: '插入分页符', keywords: ['分页'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'setHyperlink', description: '设置超链接', keywords: ['超链接'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '链接文本', required: true }, address: { type: 'string', description: '链接地址', required: true } } },
  { name: 'smartFillField', description: '智能填写模板字段', keywords: ['模板', '填写'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { keyword: { type: 'string', description: '关键字', required: true }, value: { type: 'string', description: '填写值', required: true }, fillMode: { type: 'string', description: '填写模式: auto(自动), underline(下划线), afterColon(冒号后), afterLabel(标签后), placeholder(占位符)', enum: ['auto', 'underline', 'afterColon', 'afterLabel', 'placeholder'], required: false } } },
  { name: 'addComment', description: '添加批注', keywords: ['批注'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '批注内容', required: true } } },
  { name: 'getComments', description: '获取文档批注列表', keywords: ['批注'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'insertText', description: '插入文本', keywords: ['文本', '插入'], category: 'word', appType: WpsAppType.WRITER, paramsSchema: { text: { type: 'string', description: '文本内容', required: true } } },

  // Excel 操作 (~120)
  { name: 'getActiveWorkbook', description: '获取当前工作簿信息', keywords: ['工作簿', '当前'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'getCellValue', description: '获取单元格值', keywords: ['单元格', '值'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true } } },
  { name: 'setCellValue', description: '设置单元格值', keywords: ['单元格', '写入'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true }, value: { type: 'string', description: '值', required: true } } },
  { name: 'getRangeData', description: '读取区域数据', keywords: ['区域', '数据'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'setRangeData', description: '写入区域数据', keywords: ['写入'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '起始单元格', required: true }, data: { type: 'array', description: '数据', required: true } } },
  { name: 'setFormula', description: '设置公式', keywords: ['公式'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true }, formula: { type: 'string', description: '公式', required: true } } },
  { name: 'getFormula', description: '获取公式', keywords: ['公式'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '单元格', required: true } } },
  { name: 'setArrayFormula', description: '设置数组公式', keywords: ['数组'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true }, formula: { type: 'string', description: '公式', required: true } } },
  { name: 'diagnoseFormula', description: '诊断公式错误', keywords: ['诊断', '公式'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { formula: { type: 'string', description: '公式', required: true } } },
  { name: 'createSheet', description: '创建工作表', keywords: ['工作表', '新建'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '工作表名称', required: true } } },
  { name: 'deleteSheet', description: '删除工作表', keywords: ['工作表', '删除'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '工作表名称', required: true } } },
  { name: 'renameSheet', description: '重命名工作表', keywords: ['重命名'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { oldName: { type: 'string', description: '原名称', required: true }, newName: { type: 'string', description: '新名称', required: true } } },
  { name: 'copySheet', description: '复制工作表', keywords: ['复制'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '源工作表', required: true }, newName: { type: 'string', description: '新名称', required: true } } },
  { name: 'getSheetList', description: '���取���作表列表', keywords: ['列表'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'switchSheet', description: '切换工作表', keywords: ['切换'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '工作表名称', required: true } } },
  { name: 'moveSheet', description: '移动工作表', keywords: ['移动'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '工作表名称', required: true }, position: { type: 'number', description: '位置', required: true } } },
  { name: 'createPivotTable', description: '创建透视表', keywords: ['透视表'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { sourceRange: { type: 'string', description: '数据源', required: true }, destination: { type: 'string', description: '目标位置', required: true } } },
  { name: 'updatePivotTable', description: '更新透视表', keywords: ['透视表'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { tableName: { type: 'string', description: '透视表名称', required: true } } },
  { name: 'createChart', description: '创建图表', keywords: ['图表'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '数据区域', required: true }, chartType: { type: 'string', description: '图表类型', required: false } } },
  { name: 'updateChart', description: '更新图表', keywords: ['图表'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { chartName: { type: 'string', description: '图表名称', required: true } } },
  { name: 'createDonutChart', description: '创建环形图', keywords: ['图表', '环形'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'createFlowChart', description: '创建流程图', keywords: ['图表', '流程'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'createGauge', description: '创建仪表图', keywords: ['图表', '仪表'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'createGrid', description: '创建网格图', keywords: ['图表', '网格'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'createKpiCards', description: '创建 KPI 卡片', keywords: ['KPI'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'createMiniCharts', description: '创建迷你图', keywords: ['迷你图'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'setCellFormat', description: '设置单元格格式', keywords: ['格式'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true }, numberFormat: { type: 'string', description: '数字格式', required: false } } },
  { name: 'setCellStyle', description: '设置单元格样式', keywords: ['样式'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'setBorder', description: '设置边框', keywords: ['边框'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'copyFormat', description: '复制格式', keywords: ['复制'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { source: { type: 'string', description: '源区域', required: true }, target: { type: 'string', description: '目标区域', required: true } } },
  { name: 'clearFormats', description: '清除格式', keywords: ['清除'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'addConditionalFormat', description: '添加条件格式', keywords: ['条件'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'removeConditionalFormat', description: '删除条件格式', keywords: ['条件'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'getConditionalFormats', description: '获取条件格式', keywords: ['条件'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'addDataValidation', description: '添加数据验证', keywords: ['验证'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'removeDataValidation', description: '删除数据验证', keywords: ['验证'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'getDataValidations', description: '获取数据验证列表', keywords: ['验证'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'mergeCells', description: '合并单元格', keywords: ['合并'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'unmergeCells', description: '取消合并单元��', keywords: ['取消合并'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'setColumnWidth', description: '设置列宽', keywords: ['列宽'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { column: { type: 'string', description: '列', required: true }, width: { type: 'number', description: '宽度', required: true } } },
  { name: 'setRowHeight', description: '设置行高', keywords: ['行高'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true }, height: { type: 'number', description: '高度', required: true } } },
  { name: 'autoFitColumn', description: '自动调整列宽', keywords: ['自动'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { column: { type: 'string', description: '列', required: true } } },
  { name: 'autoFitRow', description: '自动调整行高', keywords: ['自动'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true } } },
  { name: 'autoFitAll', description: '自动调整所有行列', keywords: ['自动'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'setNumberFormat', description: '设置数字格式', keywords: ['数字'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true }, format: { type: 'string', description: '格式', required: true } } },
  { name: 'wrapText', description: '自动换行', keywords: ['换行'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'setPrintArea', description: '设置打印区域', keywords: ['打印'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'getSelection', description: '获取选区', keywords: ['选区'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'clearRange', description: '清除区域内容', keywords: ['清除'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'insertRows', description: '插入行', keywords: ['插入', '行'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true } } },
  { name: 'insertColumns', description: '插入列', keywords: ['插入', '列'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { column: { type: 'string', description: '列', required: true } } },
  { name: 'deleteRows', description: '删除行', keywords: ['删除', '行'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true } } },
  { name: 'deleteColumns', description: '删除列', keywords: ['删除', '列'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { column: { type: 'string', description: '列', required: true } } },
  { name: 'hideRows', description: '隐藏行', keywords: ['隐藏'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true } } },
  { name: 'hideColumns', description: '隐藏列', keywords: ['隐藏'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { column: { type: 'string', description: '列', required: true } } },
  { name: 'showRows', description: '取消隐藏行', keywords: ['显示'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true } } },
  { name: 'showColumns', description: '取消隐藏列', keywords: ['显示'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { column: { type: 'string', description: '列', required: true } } },
  { name: 'groupRows', description: '组合行', keywords: ['组合'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { startRow: { type: 'number', description: '起始行', required: true }, endRow: { type: 'number', description: '结束行', required: true } } },
  { name: 'groupColumns', description: '组合列', keywords: ['组合'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { startColumn: { type: 'string', description: '起始列', required: true }, endColumn: { type: 'string', description: '结束列', required: true } } },
  { name: 'freezePanes', description: '冻结窗格', keywords: ['冻结'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: false } } },
  { name: 'unfreezePanes', description: '取消冻结窗格', keywords: ['冻结'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'findInSheet', description: '在工作表中查找', keywords: ['查找'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { text: { type: 'string', description: '查找内容', required: true } } },
  { name: 'replaceInSheet', description: '在工作表中替换', keywords: ['替换'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { find: { type: 'string', description: '查找', required: true }, replace: { type: 'string', description: '替换', required: true } } },
  { name: 'copyRange', description: '复制区域', keywords: ['复制'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { source: { type: 'string', description: '源区域', required: true }, target: { type: 'string', description: '目标', required: true } } },
  { name: 'pasteRange', description: '粘贴区域', keywords: ['粘贴'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { target: { type: 'string', description: '目标', required: true } } },
  { name: 'fillSeries', description: '填充序列', keywords: ['填充'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'transpose', description: '转置数据', keywords: ['转置'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'textToColumns', description: '文本分列', keywords: ['分列'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'subtotal', description: '分类汇总', keywords: ['汇总'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'createNamedRange', description: '定义名称', keywords: ['名称'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '名称', required: true }, range: { type: 'string', description: '区域', required: true } } },
  { name: 'deleteNamedRange', description: '删除名称', keywords: ['名称'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { name: { type: 'string', description: '名称', required: true } } },
  { name: 'getNamedRanges', description: '获取名称列表', keywords: ['名称'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'addCellComment', description: '添加批注', keywords: ['批注'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true } } },
  { name: 'deleteCellComment', description: '删除批注', keywords: ['批注'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true } } },
  { name: 'getCellComments', description: '获取批注列表', keywords: ['批注'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'protectSheet', description: '保护工作表', keywords: ['保护'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { password: { type: 'string', description: '密码', required: false } } },
  { name: 'unprotectSheet', description: '取消保护工作表', keywords: ['保护'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { password: { type: 'string', description: '密码', required: false } } },
  { name: 'protectWorkbook', description: '保护工作簿', keywords: ['保护'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { password: { type: 'string', description: '密码', required: false } } },
  { name: 'insertExcelImage', description: '插入图片', keywords: ['图片'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { filePath: { type: 'string', description: '图片路径', required: true } } },
  { name: 'lockCells', description: '锁定单元格', keywords: ['锁定'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'openWorkbook', description: '打开工作簿', keywords: ['打开'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { filePath: { type: 'string', description: '文件路径', required: true } } },
  { name: 'getOpenWorkbooks', description: '获取打开的工作簿列表', keywords: ['工作簿'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'createWorkbook', description: '新建工作簿', keywords: ['新建'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'cleanData', description: '清洗数据', keywords: ['清洗'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'removeDuplicates', description: '删除重复项', keywords: ['去重'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'sortRange', description: '排序区域', keywords: ['排序'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'autoFilter', description: '自动筛选', keywords: ['筛选'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'getCellInfo', description: '获取单元格信息', keywords: ['单元格'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true } } },
  { name: 'refreshLinks', description: '刷新链接', keywords: ['链接'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'consolidate', description: '合并计算', keywords: ['合并'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'calculateSheet', description: '重新计算工作表', keywords: ['计算'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'getExcelContext', description: '获取 Excel 上下文', keywords: ['上下文'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'generateFormula', description: '根据自然语言生成公式', keywords: ['公式', '生成'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { description: { type: 'string', description: '用自然语言描述要计算的逻辑', required: true } } },

  // PPT 操作 (~80)
  { name: 'getActivePresentation', description: '获取当前演示文稿信息', keywords: ['演示', '当前'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'createPresentation', description: '新建演示文稿', keywords: ['新建'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'openPresentation', description: '打开演示文稿', keywords: ['打开'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { filePath: { type: 'string', description: '文件路径', required: true } } },
  { name: 'closePresentation', description: '关闭演示文稿', keywords: ['关闭'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'getSlideCount', description: '获取幻灯片数量', keywords: ['数量'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addSlide', description: '添加幻灯片', keywords: ['添加'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { layout: { type: 'string', description: '布局', required: false } } },
  { name: 'deleteSlide', description: '删除幻灯片', keywords: ['删除'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true } } },
  { name: 'duplicateSlide', description: '复制幻灯片', keywords: ['复制'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true } } },
  { name: 'moveSlide', description: '移动幻灯片', keywords: ['移动'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { fromIndex: { type: 'number', description: '原位置', required: true }, toIndex: { type: 'number', description: '新位置', required: true } } },
  { name: 'switchSlide', description: '切换幻灯片', keywords: ['切换'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '目标幻灯片索引（从1开始）', required: true } } },
  { name: 'getSlideInfo', description: '获取幻灯片信息', keywords: ['信息'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true } } },
  { name: 'getSlideTitle', description: '获取幻灯片标题', keywords: ['标题'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true } } },
  { name: 'getSlideNotes', description: '获取幻灯片备注', keywords: ['备注'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true } } },
  { name: 'setSlideTitle', description: '设置幻灯片标题', keywords: ['标题'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始），默认1', required: false }, title: { type: 'string', description: '标题文本', required: true } } },
  { name: 'setSlideSubtitle', description: '设置幻灯片副标题', keywords: ['副标题'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true }, subtitle: { type: 'string', description: '副标题内容', required: true } } },
  { name: 'setSlideContent', description: '设置幻灯片内容', keywords: ['内容'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true }, content: { type: 'string', description: '正文内容', required: true } } },
  { name: 'setSlideNotes', description: '设置幻灯片备注', keywords: ['备注'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片索引（从1开始）', required: true }, notes: { type: 'string', description: '备注内容', required: true } } },
  { name: 'setSlideBackground', description: '设置幻灯片背景', keywords: ['背景'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true }, background: { type: 'object', description: '背景配置对象，包含 type/color/colors/imagePath/pattern 等字段（也可使用旧平铺参数 color/imagePath）', required: false } } },
  { name: 'setBackgroundColor', description: '设置背景颜色', keywords: ['背景'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { color: { type: 'string', description: '颜色', required: true } } },
  { name: 'setBackgroundGradient', description: '设置背景渐变', keywords: ['背景', '渐变'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setBackgroundImage', description: '设置背景图片', keywords: ['背景'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setSlideTransition', description: '设置切换效果', keywords: ['切换'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true }, effect: { type: 'string', description: '切换效果名称', required: true }, advanceAfter: { type: 'number', description: '自动翻页间隔时间（秒），不填则手动翻页', required: false }, sound: { type: 'string', description: '切换时播放的声音文件路径（可选）', required: false } } },
  { name: 'removeSlideTransition', description: '删除切换效果', keywords: ['切换'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true } } },
  { name: 'applyTransitionToAll', description: '应用切换到全部', keywords: ['切换'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { effect: { type: 'string', description: '效果', required: true } } },
  { name: 'addAnimation', description: '添加动画', keywords: ['动画'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true }, shapeIndex: { type: 'number', description: '形状索引（从1开始）', required: true }, effect: { type: 'string', description: '动画效果名称，如 "fadeIn"、"flyIn"、"wipe" 等', required: true }, trigger: { type: 'string', description: '触发方式', enum: ['onClick', 'withPrevious', 'afterPrevious'], required: false } } },
  { name: 'removeAnimation', description: '删除动画', keywords: ['动画'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true }, animationIndex: { type: 'number', description: '动画索引（从1开始）', required: true } } },
  { name: 'setAnimationOrder', description: '设置动画顺序', keywords: ['动画'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true }, animationIndex: { type: 'number', description: '当前动画索引（从1开始）', required: true }, newOrder: { type: 'number', description: '新的播放顺序位置（从1开始）', required: true } } },
  { name: 'getAnimations', description: '获取动画列表', keywords: ['动画'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true } } },
  { name: 'addAnimationPreset', description: '添加预设动画', keywords: ['动画'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addEmphasisAnimation', description: '添加强调动画', keywords: ['动画'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'beautifySlide', description: '美化幻灯片', keywords: ['美化'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '要美化的幻灯片页码，不填则美化当前页', required: false }, colorScheme: { type: 'string', description: '配色方案', enum: ['business', 'tech', 'creative', 'minimal'], required: false }, font: { type: 'string', description: '统一使用的字体，如 "微软雅黑"、"思源黑体"', required: false }, beautifyAll: { type: 'boolean', description: '是否美化所有幻灯片，默认false只美化指定页', required: false } } },
  { name: 'beautifyAllSlides', description: '美化所有幻灯片', keywords: ['美化'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'autoBeautifySlide', description: '自动美化幻灯片', keywords: ['美化'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '幻灯片页码（从1开始）', required: true } } },
  { name: 'unifyFont', description: '统一字体', keywords: ['字体'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { fontName: { type: 'string', description: '字体名称', required: true } } },
  { name: 'addShape', description: '添加形状', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { shapeType: { type: 'string', description: '形状类型', required: true } } },
  { name: 'deleteShape', description: '删除形状', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'duplicateShape', description: '复制形状', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'getShapes', description: '获取形状列表', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'alignShapes', description: '对齐形状', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { alignment: { type: 'string', description: '对齐方式', required: true } } },
  { name: 'groupShapes', description: '组合形状', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'distributeShapes', description: '分布形状', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addTextBox', description: '添加文本框', keywords: ['文本框'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { text: { type: 'string', description: '文本内容', required: true } } },
  { name: 'setTextBoxText', description: '设置文本框文本', keywords: ['文本框'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'getTextBoxes', description: '获取文本框列表', keywords: ['文本框'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setTextBoxStyle', description: '设置文本框样式', keywords: ['文本框'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'deleteTextBox', description: '删除文本框', keywords: ['文本框'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'insertPptImage', description: '插入图片', keywords: ['图片'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { filePath: { type: 'string', description: '图片路径', required: true } } },
  { name: 'deletePptImage', description: '删除图片', keywords: ['图片'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'setImageStyle', description: '设置图片样式', keywords: ['图片'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'insertPptChart', description: '插入图表', keywords: ['图表'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { chartType: { type: 'string', description: '图表类型', required: true } } },
  { name: 'setPptChartData', description: '设置图表数据', keywords: ['图表'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'setPptChartStyle', description: '设置图表样式', keywords: ['图表'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'insertPptTable', description: '插入表格', keywords: ['表格'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { rows: { type: 'number', description: '行数', required: true }, cols: { type: 'number', description: '列数', required: true } } },
  { name: 'getPptTableCell', description: '获取表格单元格', keywords: ['表格'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { tableIndex: { type: 'number', description: '表格索引', required: true }, row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true } } },
  { name: 'setPptTableCell', description: '设置表格单元格', keywords: ['表格'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { tableIndex: { type: 'number', description: '表格索引', required: true }, row: { type: 'number', description: '行号', required: true }, col: { type: 'number', description: '列号', required: true }, text: { type: 'string', description: '文本', required: true } } },
  { name: 'setPptTableStyle', description: '设置表格样式', keywords: ['表格'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { tableIndex: { type: 'number', description: '索引', required: true } } },
  { name: 'setPptTableCellStyle', description: '设置表格单元格样式', keywords: ['表格'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setPptTableRowStyle', description: '设置表格行样式', keywords: ['表格'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeStyle', description: '设置形状样式', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'setShapeBorder', description: '设置形状边框', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeShadow', description: '设置形状阴影', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeGradient', description: '设置形状渐变', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapePosition', description: '设置形状位置', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '索引', required: true } } },
  { name: 'setShapeZOrder', description: '设置形状层次', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeFullStyle', description: '设置形状完整样式', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeRoundness', description: '设置形状圆角', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeTransparency', description: '设置形状透明度', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setShapeText', description: '设置形状文本', keywords: ['形状'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addConnector', description: '添加连接符', keywords: ['连接'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addArrow', description: '添加箭头', keywords: ['箭头'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setSlideLayout', description: '设置幻灯片布局', keywords: ['布局'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { index: { type: 'number', description: '页码', required: true } } },
  { name: 'setSlideNumber', description: '设置幻灯片页码', keywords: ['页码'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setPptDateTime', description: '设置日期时间', keywords: ['日期'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setPptFooter', description: '设置页脚', keywords: ['页脚'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'setMasterBackground', description: '设置母版背景', keywords: ['母版'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'getSlideMaster', description: '获取母版', keywords: ['母版'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'applyColorScheme', description: '应用配色方案', keywords: ['配色'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addPptHyperlink', description: '添加超链接', keywords: ['超链接'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'removePptHyperlink', description: '删除超链接', keywords: ['超链接'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'findPptText', description: '查找文本', keywords: ['查找'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'replacePptText', description: '替换文本', keywords: ['替换'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addTitleDecoration', description: '添加标题装饰', keywords: ['标题'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addPageIndicator', description: '添加页码指示器', keywords: ['页码'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'addMasterElement', description: '添加母版元素', keywords: ['母版'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'startSlideShow', description: '开始幻灯片放映', keywords: ['放映'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'endSlideShow', description: '结束幻灯片放映', keywords: ['放映'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'autoLayout', description: '自动布局', keywords: ['布局'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'smartDistribute', description: '智能分布', keywords: ['分布'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'create3DText', description: '创建 3D 文字', keywords: ['3D'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'set3DDepth', description: '设置 3D 深度', keywords: ['3D'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'set3DMaterial', description: '设置 3D 材质', keywords: ['3D'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'set3DRotation', description: '设置 3D 旋转', keywords: ['3D'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: {} },

  // 通用操作 (~10)
  { name: 'convertToPDF', description: '转换为 PDF', keywords: ['PDF'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: { outputPath: { type: 'string', description: '输出路径', required: false } } },
  { name: 'convertFormat', description: '格式转换', keywords: ['转换'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: { targetFormat: { type: 'string', description: '目标格式', required: true } } },
  { name: 'getAppInfo', description: '获取应用信息', keywords: ['应用'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'getContext', description: '获取当前上下文', keywords: ['上下文'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'getOpenDocuments', description: '获取打开的文档列表', keywords: ['文档'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'getOpenPresentations', description: '获取打开的演示文稿列表', keywords: ['演示'], category: 'common', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'switchDocument', description: '切换文档', keywords: ['切换'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'switchPresentation', description: '切换演示文稿', keywords: ['切换'], category: 'common', appType: WpsAppType.PRESENTATION, paramsSchema: {} },
  { name: 'switchWorkbook', description: '切换工作簿', keywords: ['切换'], category: 'common', appType: WpsAppType.SPREADSHEET, paramsSchema: {} },
  { name: 'ping', description: '检测 MCP Server 连接状态', keywords: ['检测', 'ping', '连接'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'wireCheck', description: '检测 WPS COM 连接状态', keywords: ['检测', 'COM', '连接'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'placeholder', description: '占位符', keywords: ['占位符'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'trim', description: '去除前后空格', keywords: ['空格'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'underline', description: '添加下划线', keywords: ['下划线'], category: 'common', appType: WpsAppType.WRITER, paramsSchema: {} },
  { name: 'remove_duplicates', description: '删除重复行（cleanData 子命令）', keywords: ['去重', '重复行'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'unify_date', description: '统一日期格式（cleanData 子命令）', keywords: ['日期', '统一'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域', required: true } } },
  { name: 'closeWorkbook', description: '关闭工作簿（可选择是否保存）', keywords: ['工作簿', '关闭', '保存'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { save: { type: 'boolean', description: '保存后再关闭', required: false } } },
  // Export 工具
  { name: 'exportChartAsImage', description: '将图表导出为位图图片（PNG/JPG/GIF/BMP）', keywords: ['图表', '导出', '图片'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { chartName: { type: 'string', description: '图表名称', required: true }, outputPath: { type: 'string', description: '输出路径', required: true }, format: { type: 'string', description: '图片格式', required: false }, sheet: { type: 'string', description: '工作表', required: false } } },
  { name: 'exportRangeAsImage', description: '将区域导出为位图图片（PNG/JPG/GIF/BMP）', keywords: ['区域', '导出', '图片'], category: 'excel', appType: WpsAppType.SPREADSHEET, paramsSchema: { range: { type: 'string', description: '区域地址', required: true }, outputPath: { type: 'string', description: '输出路径', required: true }, format: { type: 'string', description: '图片格式', required: false }, sheet: { type: 'string', description: '工作表', required: false } } },
  { name: 'exportSlideAsImage', description: '将幻灯片导出为位图图片（PNG/JPG/GIF/BMP）', keywords: ['幻灯片', '导出', '图片'], category: 'ppt', appType: WpsAppType.PRESENTATION, paramsSchema: { slideIndex: { type: 'number', description: '页码', required: true }, outputPath: { type: 'string', description: '输出路径', required: true }, format: { type: 'string', description: '图片格式', required: false }, width: { type: 'number', description: '宽度', required: false }, height: { type: 'number', description: '高度', required: false } } },
];

export const TOOLS_INDEX: ToolIndexItem[] = COM_ACTIONS.map(tool => ({
  ...tool,
  status: getToolStatus(tool.name),
}));

export interface SearchOptions {
  query: string;
  category?: string;
  limit?: number;
}

export interface SearchResult {
  total: number;
  results: Array<{
    name: string;
    description: string;
    category: string;
    appType: string;
    params: Record<string, ToolParamSchema>;
    example: string;
  }>;
  next_steps: string;
}

export function searchTools(options: SearchOptions): SearchResult {
  const { query, category, limit = 10 } = options;
  let filtered = TOOLS_INDEX;
  if (category) filtered = filtered.filter(t => t.category === category);
  if (query) {
    const q = query.toLowerCase();
    filtered = filtered.filter(t =>
      t.name.toLowerCase().includes(q) ||
      t.description.toLowerCase().includes(q) ||
      t.keywords.some(k => k.toLowerCase().includes(q))
    );
  }
  if (query) {
    const q = query.toLowerCase();
    filtered = filtered.sort((a, b) => {
      const aExact = a.name.toLowerCase() === q ? 1 : 0;
      const bExact = b.name.toLowerCase() === q ? 1 : 0;
      if (aExact !== bExact) return bExact - aExact;
      const aStart = a.name.toLowerCase().startsWith(q) ? 1 : 0;
      const bStart = b.name.toLowerCase().startsWith(q) ? 1 : 0;
      if (aStart !== bStart) return bStart - aStart;
      return a.name.localeCompare(b.name);
    });
  }
  const results = filtered.slice(0, limit).map(tool => ({
    name: tool.name,
    description: tool.description.split('\n')[0],
    category: tool.category,
    appType: tool.appType,
    params: SCHEMA_MAP.get(tool.name) ?? tool.paramsSchema,
    example: `wps_office_execute('${tool.name}', {...})`,
  }));
  return { total: filtered.length, results, next_steps: "使用 wps_office_execute 执行" };
}

export interface ExecuteOptions {
  tool_name: string;
  arguments: Record<string, unknown>;
}

export async function executeTool(options: ExecuteOptions): Promise<ToolCallResult> {
  const { tool_name, arguments: args } = options;
  const indexItem = TOOLS_INDEX.find(t => t.name === tool_name);
  if (!indexItem) {
    return {
      id: '',
      success: false,
      content: [{ type: 'text', text: JSON.stringify({ error: `工具 "${tool_name}" 不存在`, suggestion: '使用 wps_office_search 查找可用工具' }) }]
    };
  }
  // 优先使用 TS handler（带参数校验、类型安全、详细错误信息）
  // 用 "name|appType" 复合键精确匹配，处理跨应用同名工具
  const camelName = toCamelCase(tool_name);
  const handlerKey = `${tool_name}|${indexItem.appType}`;
  const fallbackKey = `${camelName}|${indexItem.appType}`;
  const handler = HANDLER_MAP.get(handlerKey) ?? HANDLER_MAP.get(fallbackKey);
  if (handler && !HANDLER_SKIP.has(tool_name)) {
    // 应用逐工具参数名映射（将 COM 参数名转为 handler 期望的参数名）
    const paramMap = HANDLER_PARAM_MAP[tool_name];
    const mappedArgs = paramMap && Object.keys(paramMap).length > 0
      ? Object.fromEntries(
          Object.entries(args).map(([k, v]) => [paramMap[k] ?? k, v])
        )
      : args;
    try {
      return await handler(mappedArgs);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      return {
        id: '',
        success: false,
        content: [{ type: 'text', text: JSON.stringify({ error: `TS handler 执行失败`, details: errorMessage, tool: tool_name, params: args }) }]
      };
    }
  }
  // 无 TS handler，直接透传 PS1（兜底）
  try {
    const result = await wpsClient.executeMethod(tool_name, args as Record<string, unknown>, indexItem.appType);
    return {
      id: '',
      success: result.success,
      content: [{ type: 'text', text: JSON.stringify({ result, tool_name, appType: indexItem.appType }) }]
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const errorStack = error instanceof Error ? error.stack : undefined;
    return {
      id: '',
      success: false,
      content: [{
        type: 'text',
        text: JSON.stringify({
          error: `执行工具 "${tool_name}" 失败`,
          details: errorMessage,
          tool: tool_name,
          appType: indexItem.appType,
          params: args,
          // 仅在非生产环境包含堆栈
          ...(process.env.NODE_ENV !== 'production' && { stack: errorStack })
        }, null, 2)
      }]
    };
  }
}

export default { TOOLS_INDEX, searchTools, executeTool };