# 在途商旅 Word Master Template｜五套风格 × 八类版式

## 目标

同一份产品内容数据，不改变 `title / summary / screenshot / features / values / steps / metrics`，通过 `--style` 切换五套真正不同的版式系统。

## 五套风格

| 风格 | 核心阅读逻辑 | 适用内容 |
|---|---|---|
| A 专业蓝白·证据型 | 截图证据 → 功能说明 → 价值 | 通用产品功能、技术标 |
| B 现代青绿·双栏型 | 界面 → 核心能力 → 价值 | SaaS、系统集成、移动协同 |
| C 稳重靛蓝·报告型 | 能力摘要 → 证据 → 详细功能 → 管理价值 | BI、数据、风控、集团管理 |
| D 商务金棕·杂志型 | 大图 → 编号说明 → 价值 | 服务、运营、实施、客户成功 |
| E 治理绿白·卡片型 | 能力卡 → 界面证据 → 管理价值 | 合规、政策、预算、风控 |

## 八类页面

### 01 标准功能页 `standard`

同一内容在五套风格下使用不同的主视觉结构：
- A：大截图全宽，功能表随后
- B：截图左 + 核心能力右
- C：三张能力卡在前，截图在后
- D：截图左 + 编号功能说明右
- E：3×N 能力卡在前，截图在后

### 02 截图标注页 `annotated`

用于真实 PC 截图、局部功能标号、评审证据对应。内容区采用“截图 + ①②③说明”，不把说明文字烧进图片。

### 03 PC+移动端协同 `pc_mobile`

不是简单两张图并排，而是根据风格改变比例与阅读关系：A 60/40、B 50/50、C 上图下标签、D 75/25、E 能力说明后双端证据。

### 04 移动端多截图 `mobile_gallery`

A/C：2列证据阅读；B/E：3列移动端能力墙；D：逐条“手机截图 + 业务说明”。

### 05 流程页 `process`

A：横向 5 步；B/E：3列阶段型；C：纵向编号报告型；D：4段流程型。关键界面作为流程证据附在后面。

### 06 数据看板 `dashboard`

A：看板 → 指标；B：指标卡 → 看板；C：看板 → 指标条；D：看板 → 业务说明；E：治理指标 → 数据证据。

### 07 能力总览 `overview`

A/B/E：能力卡矩阵；C：报告式能力地图；D：编号长条能力列表。

### 08 能力矩阵 `matrix`

适合招标响应、产品能力覆盖、PC/APP/管理端/数据能力对照。所有单元格为 Word 原生可编辑文本。

## 内容与版式分离

数据只负责“是什么”，版式负责“怎么呈现”。例如：

```text
page.type = standard
page.title = 系统登录
page.features = [...]
page.values = [...]
```

渲染：

```bash
python src/generate_master_manual.py --data data/master_demo.json --style clean_blue
python src/generate_master_manual.py --data data/master_demo.json --style teal_grid
python src/generate_master_manual.py --data data/master_demo.json --style indigo_report
python src/generate_master_manual.py --data data/master_demo.json --style amber_standard
python src/generate_master_manual.py --data data/master_demo.json --style green_governance
python src/generate_master_manual.py --data data/master_demo.json --all-styles
```

## 下一阶段

1. 用真实在途商旅截图替换 demo 占位。
2. 为每类页面增加 PC/手机截图自动比例检测。
3. 增加截图标注编号组件，并自动生成对应说明。
4. 增加页眉、页脚、章节编号、目录、页码和分页控制。
5. 在稳定版式后再接入完整 136+ 页内容库，避免“先批量生成、后发现版式不成立”。
