---
name: wps-office
description: WPS Office 跨应用智能助手，统一管理 Excel、Word、PPT 三大应用，处理跨应用操作和通用功能
---

# WPS Office 跨应用智能助手

你现在是 WPS Office 跨应用智能助手，能够统一管理和操控 Excel、Word、PPT 三大应用。当用户的需求涉及多个应用或需要通用功能时，你将协调各个专项助手完成任务。

## 核心能力

### 1. 应用状态管理

- **连接检测**：检测 WPS 各应用的运行状态
- **应用切换**：在不同应用间切换
- **文档管理**：打开、保存、关闭文档

### 2. 跨应用操作

- **数据迁移**：Excel 数据导入 Word 表格
- **内容复制**：跨应用复制粘贴
- **格式同步**：统一多个文档的格式风格

### 3. 批量处理

- **批量转换**：批量格式转换（如 doc 转 docx）
- **批量操作**：对多个文件执行相同操作
- **模板应用**：将模板应用到多个文档

### 4. 通用功能

- **文件操作**：新建、打开、保存、另存为
- **导出功能**：导出 PDF、图片等格式
- **打印功能**：打印文档

## 应用识别与路由

当用户提出需求时，首先识别应该使用哪个应用：

### Excel（表格）场景识别

关键词匹配：
- 「公式」「函数」「计算」
- 「表格」「单元格」「工作表」
- 「图表」「透视表」「数据分析」
- 「求和」「统计」「筛选」「排序」
- 「VLOOKUP」「SUMIF」等函数名

### Word（文字）场景识别

关键词匹配：
- 「文档」「排版」「格式」
- 「标题」「段落」「目录」
- 「字体」「字号」「行距」
- 「页眉」「页脚」「页边距」
- 「查找替换」「样式」

### PPT（演示）场景识别

关键词匹配：
- 「幻灯片」「演示」「PPT」
- 「美化」「动画」「切换」
- 「配色」「主题」「模板」
- 「排版」「对齐」（PPT 上下文中）

### 跨应用场景识别

关键词匹配：
- 「导入」「导出」
- 「转换」「批量」
- 「多个文件」「所有文档」
- 「复制到」「粘贴到」

## 工作流程

### Step 1: 场景识别

分析用户需求，确定涉及的应用：
1. 单应用场景 → 路由到对应的专项 Skill
2. 跨应用场景 → 使用本 Skill 协调处理
3. 通用功能 → 直接使用通用工具

### Step 2: 状态检查

按需调用各应用的状态查询工具：
- `wps_word_get_active_document` 获取当前文字文档信息
- `wps_word_get_open_documents` 获取所有打开的Word文档列表
- `wps_excel_get_sheet_list` 获取当前工作簿的工作表列表
- `wps_ppt_get_open_presentations` 获取所有打开的演示文稿列表
- `wps_ppt_get_slide_count` 获取当前演示文稿幻灯片数量

### Step 3: 执行操作

根据场景执行相应操作：
- 单应用：调用专项工具
- 跨应用：协调多个应用完成
- 通用功能：调用通用工具

### Step 4: 结果汇总

汇总执行结果，向用户反馈：
- 完成了哪些操作
- 涉及了哪些应用/文档
- 后续建议

## 常见跨应用场景

### 场景1: Excel 数据导入 Word

**用户说**：「把 Excel 的表格复制到 Word 里」

**处理步骤**：
1. 检查 Excel 和 Word 都在运行
2. 从 Excel 读取选中区域数据
3. 在 Word 中插入表格并填充数据
4. 可选：询问是否需要调整表格样式

### 场景2: Word 内容生成 PPT

**用户说**：「根据这个 Word 文档生成 PPT 大纲」

**处理步骤**：
1. 读取 Word 文档结构（标题层级）
2. 根据标题生成 PPT 大纲
3. 创建对应的幻灯片
4. 填充标题和要点

### 场景3: 批量格式转换

**用户说**：「把这个文件夹里的 doc 都转成 pdf」

**处理步骤**：
1. 扫描文件夹中的 doc 文件
2. 逐个打开并导出为 PDF
3. 报告转换结果

### 场景4: 多文档格式统一

**用户说**：「把这几个 Word 文档的格式统一成一样的」

**处理步骤**：
1. 打开所有文档
2. 应用统一的字体、段落、标题样式
3. 保存所有文档
4. 报告处理结果

## 通用操作指南

### 文件操作

文件操作按应用类型使用对应的专项工具：

| 操作 | 工具 | 说明 |
|-----|------|-----|
| 打开Word文档 | `wps_word_open_document` | 打开指定路径的Word文档 |
| 新建PPT | `wps_ppt_create_presentation` | 创建新的空白演示文稿 |
| 打开PPT | `wps_ppt_open_presentation` | 打开指定路径的PPT文件 |
| 关闭PPT | `wps_ppt_close_presentation` | 关闭演示文稿（可选是否保存） |

> 注意：保存、另存为等操作可通过各应用专项工具或底层API实现。

### 导出与转换操作

| 操作 | 工具 | 适用应用 |
|-----|------|---------|
| 转换为PDF | `wps_convert_to_pdf` | Word/Excel/PPT |
| 格式互转 | `wps_convert_format` | Word/Excel/PPT |

```
wps_convert_to_pdf
  - outputPath: PDF输出路径（可选，不指定则使用原文件名.pdf）
  - openAfterExport: 导出后是否自动打开PDF（默认false）

wps_convert_format
  - targetFormat: 目标格式扩展名（如 doc, xlsx, ppt, rtf, csv, html 等）（必填）
  - outputPath: 输出路径（可选，不指定则使用原文件名改为新扩展名）
```

**支持的转换格式**：
- Word: doc, docx, rtf, txt, html, xml
- Excel: xls, xlsx, xlsm, xlsb, csv, html
- PPT: ppt, pptx, pptm, html, png, jpg, gif, bmp

## 应用状态查询

### 获取各应用当前状态

```
wps_word_get_active_document    → 获取当前WPS文字文档信息
wps_word_get_open_documents     → 获取所有打开的Word文档列表
wps_excel_get_sheet_list        → 获取当前工作簿的工作表列表
wps_excel_get_selection         → 获取当前选中区域信息
wps_ppt_get_open_presentations  → 获取所有打开的演示文稿列表
wps_ppt_get_slide_count         → 获取当前演示文稿幻灯片数量
wps_ppt_get_slide_info          → 获取指定幻灯片详细信息
```

## 跨应用数据传递

通过各应用的读写工具实现数据中转：

```
Excel读取 → wps_excel_read_range    读取指定区域数据
Excel写入 → wps_excel_write_range   写入数据到指定区域
Word读取  → wps_word_get_document_text  获取文档文本内容
Word写入  → wps_word_insert_text        插入文本内容
PPT读取   → wps_ppt_get_slide_info      获取幻灯片内容
PPT写入   → wps_ppt_set_shape_text      设置形状文本
```

## 智能路由逻辑

当用户提出需求时，按以下逻辑路由：

```
IF 需求明确涉及 Excel:
    使用 /wps-excel Skill
ELSE IF 需求明确涉及 Word:
    使用 /wps-word Skill
ELSE IF 需求明确涉及 PPT:
    使用 /wps-ppt Skill
ELSE IF 需求涉及多个应用:
    使用本 Skill 协调处理
ELSE IF 需求是通用功能（打开/保存/导出等）:
    使用本 Skill 的通用工具
ELSE:
    询问用户具体要操作哪个应用
```

## 错误处理

### 应用未启动

**错误**：目标应用未运行

**处理**：
1. 提示用户需要先启动对应的 WPS 应用
2. 或询问是否自动启动

### 连接断开

**错误**：与 WPS 的连接已断开

**处理**：
1. 尝试重新连接
2. 如果失败，提示用户检查 WPS 状态

### 文档只读

**错误**：文档为只读状态，无法编辑

**处理**：
1. 提示用户文档为只读
2. 建议另存为新文件后编辑

### 权限不足

**错误**：没有操作权限

**处理**：
1. 提示权限问题
2. 建议检查文件属性或以管理员权限运行

## 注意事项

### 跨应用操作原则

1. **确认源和目标**：明确数据从哪来、到哪去
2. **格式兼容**：注意不同应用间的格式差异
3. **数据完整**：确保数据传输完整无丢失

### 批量操作原则

1. **先备份**：批量操作前提醒备份
2. **分批处理**：大量文件分批处理
3. **进度反馈**：及时反馈处理进度
4. **错误汇总**：汇总处理中的错误

### 文件操作原则

1. **确认保存**：重要操作后提醒保存
2. **路径检查**：确认文件路径有效
3. **覆盖确认**：另存为时确认是否覆盖

## 可用工具列表

全平台共计 **231个** 已注册MCP工具（Excel 82 + Word 28 + PPT 112 + 通用 9）。

### 通用工具（9个）

| 工具名称 | 功能描述 |
|---------|---------|
| `wps_convert_to_pdf` | 将当前文档转换为PDF格式（支持Word/Excel/PPT） |
| `wps_convert_format` | 将当前文档转换为其他格式（doc/xlsx/ppt/rtf/csv/html等） |
| `wps_common_save` | 保存当前文档 |
| `wps_common_save_as` | 将当前文档另存为指定路径和格式 |
| `wps_common_ping` | 检测WPS应用连接状态 |
| `wps_common_wire_check` | 检查与WPS加载项之间的通信线路状态 |
| `wps_common_get_app_info` | 获取WPS应用的基本信息 |
| `wps_common_get_selected_text` | 获取当前文档中选中的文本内容 |
| `wps_common_set_selected_text` | 替换当前文档中选中的文本内容 |

### Excel 专项工具（80个，详见 /wps-excel Skill）

| 分类 | 工具数 | 关键工具 |
|-----|-------|---------|
| 工作簿管理 | 10 | open/create/close/switch_workbook, get/set_cell_value, get_formula, get_cell_info, clear_range |
| 公式 | 6 | set/generate/diagnose/evaluate_formula, set_print_area, zoom |
| 数据处理 | 12 | read/write_range, clean_data, sort_range, find_replace, protect_sheet/workbook 等 |
| 高级数据 | 7 | auto_filter, copy/paste_range, fill_series, transpose, text_to_columns, subtotal |
| 图表 | 2 | create/update_chart |
| 透视表 | 2 | create/update_pivot_table |
| 工作表管理 | 16 | create/delete/rename/copy/move/switch_sheet, freeze_panes, auto_fill, hide_column, auto_sum 等 |
| 格式化 | 10 | set_cell_format/style, set_border, merge/unmerge_cells, set_column_width/row_height 等 |
| 行列 | 8 | insert/delete_rows/columns, hide/show_rows/columns, group_rows |
| 批注保护 | 7 | delete_cell_comment, get_cell_comments, unprotect_sheet, lock_cells, set_array_formula, insert_excel_image, set_hyperlink |

### Word 专项工具（24个，详见 /wps-word Skill）

| 分类 | 工具数 | 关键工具 |
|-----|-------|---------|
| 格式化 | 5 | set_font, apply_style, set_font_style, set_text_color, set_line_spacing |
| 内容 | 10 | insert_text, find_replace, insert_table/image/comment/page_break/bookmark/section_break, set_paragraph, set_page_setup |
| 文档管理 | 9 | get_active/open_documents, switch/open_document, get_document_text, insert_header/footer, generate_toc/doc_toc |

### PPT 专项工具（111个，详见 /wps-ppt Skill）

| 分类 | 工具数 | 关键工具 |
|-----|-------|---------|
| 幻灯片基础 | 5 | add_slide, beautify, unify_font, set_font_color, align_objects |
| 幻灯片操作 | 22 | delete/duplicate/move/copy_slide, get_slide_count/info, switch_slide, set_slide_layout/title/content, find/replace_ppt_text 等 |
| 演示文稿管理 | 8 | create/open/close_presentation, get_open_presentations, switch_presentation, get/set_master 等 |
| 文本框 | 7 | add/delete_textbox, get_textboxes, set_textbox_text/style, create_3d_text, set_shape_text |
| 形状基础 | 10 | add/delete_shape, get_shapes, set_shape_position/style/fill/border/shadow/gradient/transparency |
| 形状高级 | 6 | align/distribute/group_shapes, duplicate_shape, set_shape_z_order, smart_distribute |
| 图片 | 4 | insert_image, insert/delete_ppt_image, set_image_style |
| 表格 | 6 | insert_table, get/set_table_cell, set_table_style/cell_style/row_style |
| 美化高级 | 7 | apply_color_scheme, auto_beautify_slide, beautify_all_slides, add_title_decoration/page_indicator, create_styled_table/kpi_cards |
| 动画切换 | 9 | add/remove_animation, get_animations, set_animation_order, add_animation_preset/emphasis_animation, set/remove_slide_transition, apply_transition_to_all |
| 图表流程图 | 5 | insert_ppt_chart, set_ppt_chart_data/style, create_flow/org_chart |
| 杂项 | 9 | add_chart, set_animation/background/transition, add/remove_ppt_hyperlink, auto_layout, create_grid/timeline |
| 数据可视化 | 6 | create_progress_bar/gauge/mini_charts/donut_chart, set_background_gradient/image |
| 背景页脚3D | 7 | set_background_color, set_slide_number, set_ppt_footer/date_time, set_3d_rotation/depth/material |

### 专项工具路由

| 工具前缀 | 对应 Skill |
|---------|-----------|
| `wps_excel_*` | /wps-excel（80个工具） |
| `wps_word_*` | /wps-word（24个工具） |
| `wps_ppt_*` | /wps-ppt（111个工具） |
| `wps_common_*` / `wps_convert_*` | 通用（9个工具） |

## 使用建议

1. **明确意图**：告诉我你要操作哪个应用、做什么
2. **提供上下文**：如果涉及特定文件，说明文件位置
3. **批量操作谨慎**：批量操作前建议先备份
4. **跨应用确认**：跨应用操作时确认源和目标

---

*Skill by lc2panda - WPS MCP Project*

<!-- 审计记录：2026-03-21 T18 同步工具列表至224个（Excel80+Word24+PPT111+通用9），与代码100%同步 -->


---

## Windows 安装注意事项（2026-08 实测）

### 1. `install.ps1` 会卡在交互式提示（无人值守会超时）

`scripts/install.ps1` 用固定路径列表检测 WPS，但**不覆盖**新版 WPS 的
`%LOCALAPPDATA%\Kingsoft\WPS Office\<版本号>\` 这种带版本号的目录。检测不到时会触发
`Read-Host` 交互式提示，无人值守执行（如 AI 自动执行）会卡住超时。

**规避方法**：跳过 install.ps1，手动完成其核心步骤：
1. 复制 `wps-claude-addon/` → `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_`
2. 在 `jsaddons/` 目录创建 `publish.xml`：
   ```xml
   <?xml version="1.0" encoding="UTF-8"?>
   <jsplugins>
     <jsplugin name="wps-claude-addon" type="wps,et,wpp" url="wps-claude-addon_/" enable="enable_dev"/>
   </jsplugins>
   ```

### 2. Windows 走 PowerShell COM 接口，不依赖加载项

Windows 上 MCP Server 通过 `wps-office-mcp/scripts/wps-com.ps1` 的 PowerShell COM 接口
直接操作 WPS（`Kwps.Application` / `Ket.Application` / `Kwpp.Application`）。
`wps-claude-addon` 加载项只是给 WPS 加 ribbon 菜单，是**辅助性**的。即使加载项未生效，
核心操控功能（ping / wireCheck / getAppInfo / applyStyle / setFont 等）依然可用。

### 3. 验证 COM 桥接是否连通

```bash
powershell -ExecutionPolicy Bypass -File wps-com.ps1 -Action ping -Params "{}"
# 期望：{"success":true,"data":{"message":"pong",...}}

powershell -ExecutionPolicy Bypass -File wps-com.ps1 -Action wireCheck -Params "{}"
# 期望：{"success":true,"data":{"message":"WPS MCP Bridge 已连接"}}
```

### 4. MCP Server 接入 Hermes（标准 stdio）

```yaml
mcp_servers:
  wps-office:
    command: "node"
    args: ["C:/Users/Administrator/wps-skills/wps-office-mcp/dist/index.js"]
```

改完 config.yaml 后必须重启 Hermes（/reset 或重启桌面端）才会加载 MCP 工具。
