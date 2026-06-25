# 案件时间轴生成器 (Case Timeline Generator)

将案件材料转化为清晰的法律时间轴，支持生成 **HTML 网页**（投屏演示/打印PDF）和 **可编辑 PPT**（提交法院/进一步修改）。

## 适用场景

- 案情梳理、庭审准备
- 客户汇报、可视化沟通
- 法院提交（打印 PDF 或提交 PPT）

## 快速开始

### 1. 准备事件数据

创建 `events.json`：

```json
{
  "title": "（2025）沪XXXX民初XXX号 案件大事记时间轴",
  "topParty": "原告",
  "bottomParty": "被告",
  "events": [
    {"date": "2025年1月10日", "event": "原告与被告签订《采购合同》", "party": "原告"},
    {"date": "2025年2月15日", "event": "被告将设备装车发货", "party": "被告"}
  ]
}
```

### 2. 生成 HTML 时间轴

```bash
node scripts/generate_timeline_html.js events.json 案件时间轴.html
```

纯 Node.js，无需安装依赖。浏览器打开即可：
- ← → 键盘翻页
- F 键全屏投屏
- Ctrl+P 打印 PDF 提交法院

### 3. 生成 PPT 时间轴

```bash
pip install python-pptx
python3 scripts/generate_timeline_pptx.py events.json 案件时间轴.pptx
```

生成可编辑的 PPTX 文件，每个文本框均可独立修改。

## 项目结构

```
├── SKILL.md                          # Claude Code Skill 定义
├── references/
│   └── timeline-extraction.md        # 案件材料时间节点提取指南
├── scripts/
│   ├── generate_timeline_html.js     # HTML 生成器
│   └── generate_timeline_pptx.py     # PPT 生成器（python-pptx）
└── events.json                       # 示例数据
```

## 作为 Claude Code Skill 使用

在 Claude Code 中，将此项目作为 Skill 加载后：
- **模式 A**：直接提供结构化时间线文本，快速生成时间轴
- **模式 B**：上传 PDF/Word/图片/聊天记录等原始材料，AI 自动提取时间节点，律师确认后生成

详见 [SKILL.md](SKILL.md)。

## License

MIT
