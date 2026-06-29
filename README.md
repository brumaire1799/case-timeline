# 案件时间轴生成器 (Case Timeline Generator)

根据案件材料或整理好的时间线文本，提取关键时间节点，生成横向时间轴（HTML网页和PPT两种格式）。用于案情梳理、客户汇报、庭上演示。

---

## 输入模式

### 模式 A：结构化时间线文本

律师直接给出已整理好的结构化时间线，每行包含：
- **date**：时间
- **party**：归属方
- **event**：事件描述

直接跳到「第二步：确认当事人组别和上下方」。

### 模式 B：原始案件材料

律师提供原始文档（PDF、Word、Excel、JPG/PNG 图片、聊天记录截图等），或口述案情。需要 AI 从中提取时间节点。

**支持的格式**（全部通过内置 Read 工具处理，无需额外安装组件）：
- PDF → 用 Read 工具读取（含扫描件 OCR 视觉理解）
- Word (.docx) → 用 Read 工具读取
- Excel (.xlsx) → 用 Read 工具读取
- 图片 (.jpg/.png) → 用 Read 工具读取（OCR 提取图中文字）
- 口述 → 逐条记录关键事件和时间，追问具体日期

**提取后必须经律师确认**，律师可手动增删改事件、调整时间和归属方。

---

## 工作流程

### 第一步：接收输入

- 模式 A：直接进入第二步
- 模式 B：读取材料，提取时间节点

### 第二步：确认当事人组别和上下方

> **支持两方或多方当事人。** 案件可能涉及原告一/原告二、被告一/被告二等多方，每方独立配置颜色和位置。

1. **列出所有当事人**：从事件中统计唯一的 party 值，展示清单，询问是否有遗漏或需合并的当事人
2. **配置颜色和位置**：为每方自动分配默认颜色（上方暖色系，下方冷色/暗色系）和默认位置，以表格展示供用户确认调整
3. **两方场景简化**：如只有两方，按"原告 vs 被告 / 上诉人 vs 被上诉人 / 申请人 vs 被申请人"三选一快速确认

### 第三步：确认时间轴（模式 B 必须；模式 A 可选）

将提取的时间轴以表格展示，询问用户是否有遗漏、时间/描述/归属方是否需要调整。**等待用户确认后再进入下一步。**

### 第四步：选择输出格式

- **HTML 网页**（推荐）：横向时间轴网页，每页最多 18 个事件。深色背景白色幻灯片，适合投屏。支持键盘翻页（← →）、全屏（F）、一键打印 PDF（Ctrl+P）
- **PPT**：圆角矩形卡片式布局，黑体 8pt 左对齐，每页最多 16 个事件。每个文本框可独立编辑

默认同时生成两种格式。

### 第五步：生成时间轴

HTML 可直接展示，PPT 用于微调时间轴布局和内容。

#### 生成命令

```bash
# HTML（纯 Node.js，无需依赖）
node scripts/generate_timeline_html.js events.json 案件时间轴.html

# PPT（需要 python-pptx）
~/miniconda3/bin/python3 scripts/generate_timeline_pptx.py events.json 案件时间轴.pptx
```

#### 输入格式（events.json）

```json
{
  "title": "（2025）沪XXXX民初XXX号 案件大事记时间轴",
  "parties": [
    {"name": "原告一", "color": "#CC0000", "position": "top"},
    {"name": "原告二", "color": "#E85D3F", "position": "top"},
    {"name": "被告一", "color": "#333333", "position": "bottom"},
    {"name": "被告二", "color": "#1A5276", "position": "bottom"},
    {"name": "法院", "color": "#666666", "position": "bottom"}
  ],
  "events": [
    {"date": "2025年1月10日", "event": "原告一与被告一签订《采购合同》", "party": "原告一"},
    {"date": "2025年2月15日", "event": "被告二将设备装车发货", "party": "被告二"}
  ]
}
```

> **向后兼容**：也支持旧的两方格式（`topParty`/`bottomParty`），脚本会自动转换。

---

## 文件结构

```
timeline/
├── SKILL.md                              # Skill 详细定义
├── README.md                             # 项目说明（本文件）
├── events.json                           # 示例事件数据
├── scripts/
│   ├── generate_timeline_html.js         # HTML 生成器（Node.js）
│   └── generate_timeline_pptx.py         # PPT 生成器（python-pptx）
└── references/
    ├── timeline-extraction.md            # 时间节点提取指南
    └── party-color-palette.md            # 多方配色参考
```
