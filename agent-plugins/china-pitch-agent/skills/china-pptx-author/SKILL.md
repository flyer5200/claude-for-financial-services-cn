---
name: china-pptx-author
description: Create PowerPoint presentations for A-share investment analysis, pitch decks, and client reports. Adapts the original pptx-author skill for Chinese business presentation conventions, A-share content, and domestic data visualizations. Triggers on "A股PPT制作", "投资PPT", "create presentation China", "pitch deck slides", "制作PPT", or "PowerPoint [company/topic]".
---

# china-pptx-author

## Purpose

Create professional **A股投资分析PPT** — structured PowerPoint presentations for Chinese equity research, client meetings, and IC presentations.

## Data Sources

### Primary: AkShare MCP

```python
get_index_data("000001")              → Market overview charts
get_quote(ticker)                     → Price charts
get_industry_stocks(industry="...")    → Peer comparison charts
```

### Secondary Sources
- 巨潮 — company data for charts
- Wind / Choice — professional charts
- Research team — market data

## Workflow

### Step 1: Define Presentation Purpose

**Common presentation types:**

| Type | Audience | Length | Key Content |
|------|----------|--------|-------------|
| 深度报告 (Deep dive) | Internal IC | 30-50 slides | Full analysis |
| 首次覆盖 (Initiation) | Internal | 20-30 slides | Company overview + thesis |
| 行业研究 (Sector) | Internal | 15-25 slides | Industry overview |
| 客户路演 (Roadshow) | Clients | 15-20 slides | Investment thesis |
| 晨会汇报 (Morning meeting) | Internal | 5-10 slides | Daily strategy |
| IC汇报 (IC presentation) | Investment committee | 20-30 slides | Investment case |

### Step 2: Structure the Deck

**Standard A-share research deck structure:**

```
1. 投资摘要 (Investment Summary) — 1 slide
2. 投资逻辑 (Investment Thesis) — 1-2 slides
3. 公司概览 (Company Overview) — 2-3 slides
4. 行业分析 (Industry Analysis) — 3-5 slides
5. 财务分析 (Financial Analysis) — 3-5 slides
6. 盈利预测 (Earnings Forecast) — 2-3 slides
7. 估值分析 (Valuation) — 2-3 slides
8. 风险提示 (Risk Factors) — 1 slide
9. 投资建议 (Recommendation) — 1 slide
```

### Step 3: Design Standards

**Chinese business presentation design:**

| Element | Standard |
|---------|----------|
| 模板 | Firm template (corporate colors) |
| 字体 | 微软雅黑 / 思源黑体 (preferred) |
| 字号 | Title 32-44pt, Body 18-24pt |
| 配色 | Blue (primary), gray (secondary) |
| Logo | Firm logo top-right, company logo |
| 页码 | Bottom center |
| 日期 | Footer |
| 免责声明 | Last slide |

### Step 4: Content Guidelines

**Slide content rules:**

| Rule | Guideline |
|------|-----------|
| 每页一个主题 | One idea per slide |
| 标题先行 | Clear headline at top |
| 数据可视化 | Charts > tables > text |
| 中文为主 | Primary language Chinese |
| 术语统一 | Consistent terminology |
| 数据来源 | Cite data source at bottom |

### Step 5: Common Slide Types

**Slide templates:**

**Cover slide:**
```
[Company Logo]
[公司名称]
[报告标题]
[分析师] | [日期]
[机构名称]
```

**Executive summary:**
```
投资摘要

核心观点:
• [Point 1]
• [Point 2]
• [Point 3]

目标价: ¥XX (X% upside)
评级: [买入/增持/中性/减持]
```

**Financial chart:**
```
营业收入及增长

[Bar + line chart]
• 柱状图: 营业收入 (亿元)
• 折线图: 同比增长率 (%)

数据来源: 公司年报, AkShare
```

**Peer comparison:**
```
可比公司估值比较

[Table]
公司 | 市值(亿) | P/E | P/B | P/S | EV/EBITDA

数据来源: AkShare, Wind
```

**Valuation:**
```
估值与目标价

[Football field chart]
目标价: ¥XX - ¥XX
当前价: ¥XX
上行空间: X%

[Methodology table]
```

### Step 6: Chart Standards

**Chart types for A-share presentations:**

| Chart | Use | Tool |
|-------|-----|------|
| 柱状图 (Bar) | Revenue, profit comparison | Excel embedded |
| 折线图 (Line) | Trends, growth rates | Excel embedded |
| 饼图 (Pie) | Market share, allocation | Excel embedded |
| 瀑布图 (Waterfall) | Earnings bridge | Excel embedded |
| 足球场 (Football field) | Valuation range | Excel embedded |
| 股价走势 (Stock chart) | Price history | Excel embedded |
| 散点图 (Scatter) | Comps analysis | Excel embedded |

**Chart formatting:**
- Title clear and descriptive
- Axis labels in Chinese
- Data labels where appropriate
- Source cited at bottom
- Consistent colors across deck

### Step 7: Regulatory Slides

**Required slides:**

**Risk factors (风险提示):**
```
风险提示

一、市场风险
   股票价格受宏观经济、市场情绪等因素影响

二、行业风险
   行业政策变化、竞争加剧

三、公司风险
   经营风险、管理层变动、重大诉讼

四、其他风险
   [Specific risks]

本报告仅供内部参考, 不构成投资建议。
```

**Disclaimer (免责声明):**
```
免责声明

本报告由[机构]发布, 仅供内部参考, 不构成任何投资建议。
报告中的信息来源于我们认为可靠的渠道, 但不保证其准确性。
投资者应独立做出投资决策, 并承担相应风险。
```

### Step 8: Quality Checks

Before finalizing:
- [ ] All slides have titles
- [ ] Data consistent across slides
- [ ] Charts labeled and sourced
- [ ] No typos or formatting errors
- [ ] Consistent terminology
- [ ] Regulatory slides included
- [ ] File naming convention followed
- [ ] Template applied correctly

## China-Specific Considerations

### Terminology

| English | Chinese (Standard) |
|---------|-------------------|
| Revenue | 营业收入 |
| Net income | 归母净利润 |
| Gross margin | 毛利率 |
| EBITDA | EBITDA |
| EPS | 每股收益 |
| ROE | ROE |
| Target price | 目标价 |
| Rating | 评级 (买入/增持/中性/减持) |
| Upside | 上行空间 |

### Common Presentation Conventions

| Convention | Practice |
|-----------|----------|
| 评级体系 | 买入/增持/中性/减持/卖出 |
| 目标价 | 12-month target price |
| 一致预期 | Consensus estimates |
| 同比 | Year-over-year |
| 环比 | Quarter-over-quarter |
| 千万元 | Unit for large numbers |

## Quality Checks

Before delivering:
- [ ] Purpose clearly defined
- [ ] Structure appropriate
- [ ] Content accurate
- [ ] Design consistent
- [ ] Data sourced
- [ ] Charts clear
- [ ] Regulatory slides included
- [ ] No errors
