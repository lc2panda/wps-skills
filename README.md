# WPS Skills

> 此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。

AI编程助手通过自然语言操控WPS Office的MCP工具集。

## 项目定位

本项目是MCP Server + Skills框架，让AI助手（Claude Code/Cursor/Augment等）能操控WPS Office。
- 三层工具体系：231直接注册 + 14内置 + 240 Gateway COM Actions
- 支持Excel(82工具) / Word(28工具) / PPT(112工具) / 通用(9工具)
- 支持macOS、Windows、Linux

## 前提条件

- 已安装 WPS Office（https://www.wps.cn/ 或 https://www.wps.com/）
- Node.js >= 18.0.0
- Git

## 自动安装（AI执行）

以下步骤由AI助手自动执行。需要人类交互的步骤标注 ⚠️。

### 步骤1: 克隆项目

```bash
git clone https://github.com/lc2panda/wps-skills.git
cd wps-skills
```

如果项目已存在，跳过此步。

### 步骤2: 安装依赖并编译

```bash
cd wps-office-mcp
npm install
npm run build
cd ..
```

### 步骤3: 配置MCP Server

根据使用的AI工具，将以下配置写入对应文件。注意将路径替换为实际的项目绝对路径。

**Claude Code** — 写入 `~/.claude/settings.json`：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/你的路径/wps-skills/wps-office-mcp/dist/index.js"]
    }
  }
}
```

**Cursor** — 写入项目根目录 `.cursor/mcp.json`：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/你的路径/wps-skills/wps-office-mcp/dist/index.js"]
    }
  }
}
```

**Augment / 其他MCP兼容IDE** — 参考各IDE的MCP Server配置文档，使用相同的command和args。

### 步骤4: 安装WPS加载项

⚠️ 需要人工操作（AI无法直接操作WPS应用）：

```bash
# macOS
bash scripts/auto-install-mac.sh

# Windows (PowerShell)
powershell scripts/install.ps1

# Linux
bash scripts/install.sh
```

⚠️ 安装后必须重启WPS Office才能生效。

### 步骤5: 安装Skills（仅Claude Code需要）

```bash
# 创建skills目录（如不存在）
mkdir -p ~/.claude/skills

# 创建符号链接
ln -sf "$(pwd)/skills/wps-excel" ~/.claude/skills/wps-excel
ln -sf "$(pwd)/skills/wps-word" ~/.claude/skills/wps-word
ln -sf "$(pwd)/skills/wps-ppt" ~/.claude/skills/wps-ppt
ln -sf "$(pwd)/skills/wps-office" ~/.claude/skills/wps-office
```

### 步骤6: 验证安装

```bash
# 验证MCP Server可启动
node wps-office-mcp/dist/index.js &
# 应看到 "MCP Server started successfully" 日志
kill %1 2>/dev/null
```

## 架构

```
Skills层(SKILL.md自然语言指导)
  ↓ OpenCode/Claude选择Agent+Skill
Agent层(4个专业subagent)
  ↓ 按需调用
MCP Server层(三层次工具体系)
  ├── 内置工具(14个): 基础连接/数据缓存/万能方法
  ├── Gateway工具(2个): search + execute (发现与执行)
  └── COM Actions(240个): 通过Gateway按需加载
  ↓ wpsClient.executeMethod()
执行层
  ├── macOS: wps-claude-assistant (240 action, HTTP轮询)
  └── Windows: wps-com.ps1 (240 action, COM接口)
```

### 工具渐进式加载

**解决的问题**：传统MCP Server将所有工具一次性注册，数量多(200+)时会导致AI助手工具发现困难、上下文膨胀、调用决策变慢。

**Gateway 模式**（本PR核心改造）：

```
用户请求 → wps_office_search(关键词搜索) → 返回匹配工具列表 → wps_office_execute(执行)
```

| 层次 | 数量 | 说明 |
|------|------|------|
| 内置工具 | 14个 | 连接检测、单元格读写、缓存、万能方法调用等高频基础操作 |
| Gateway 搜索 | `wps_office_search` | 通过关键词/分类搜索COM Action，返回工具名+参数schema |
| Gateway 执行 | `wps_office_execute` | 执行搜索到的COM Action，参数按schema传入 |
| COM Actions | 240个 | 底层完整API集合，覆盖Excel/Word/PPT全部操作 |

**优势**：
1. **轻量注册**：MCP协议只注册少量内置工具+2个Gateway，避免200+工具注册
2. **按需发现**：AI助手通过自然语言关键词发现可用工具，而非从庞大列表中手动挑选
3. **动态扩展**：新增COM Action只需更新索引，无需修改MCP注册逻辑
4. **降低延迟**：减少MCP初始化时的工具列表传输量，加快连接速度

## 工具清单

| 应用 | 工具数 | 主要能力 |
|------|--------|---------|
| Excel | 82 | 公式生成/诊断/数据处理/图表导出/透视表/工作表/格式/保护 |
| Word | 28 | 样式/字体/段落/目录/页眉页脚/查找替换/模板填写/书签/批注 |
| PPT | 112 | 幻灯片/形状/图片/表格/导出/美化/动画/图表/3D/数据可视化 |
| 通用 | 9 | 保存/连接检测/文本选取/格式转换/PDF导出 |
| 内置 | 14 | 连接检查/万能方法调用/数据缓存/搜索/执行 |

## Skills + Agents

| Skills（自然语言指导） | 核心能力 |
|------|---------|
| wps-word | 文档排版/样式管理/目录生成/**模板填写**/查找替换 |
| wps-excel | 公式生成/数据清洗/图表透视表/数据分析 |
| wps-ppt | 幻灯片美化/动画设置/内容生成/排版优化 |
| wps-office | 跨应用协调/格式转换/批量处理 |

| Agents（专业subagent） | 模式 | 专精领域 |
|------|------|---------|
| wps-word | subagent | Word文档排版/格式/模板填写 |
| wps-excel | subagent | Excel公式/数据分析/图表 |
| wps-ppt | subagent | PPT幻灯片美化/动画 |
| wps-expert | primary | 跨应用协调/复杂任务规划 |

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| MCP连接失败 | 确认 `npm install && npm run build` 已执行，检查dist/index.js存在 |
| WPS未响应 | 重启WPS Office，确认加载项已安装 |
| "arguments error" | 重新运行安装脚本，重启WPS |
| Linux找不到插件 | 查看INSTALL.md中的Linux专用指南 |
| 工具调用返回null | 确认WPS中已打开对应类型的文档 |

## 许可证

MIT
