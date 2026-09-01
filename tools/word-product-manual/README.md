# 在途商旅 Word 企业产品手册母版

面向 TMC / 商旅 SaaS 的 A4 Word 产品功能手册生成器。

## 为什么采用这套方案

本工程不是把 PPT 版式搬进 Word，而是以 **Word 原生可编辑** 为第一原则，解决商旅 SaaS 手册中最典型的排版问题：

- PC 后台截图尺寸不统一
- 手机 App 截图需要等宽排列
- 同一张截图需要配编号、功能解释
- 一页需要同时容纳截图、功能说明、应用价值
- 136 页以上内容需要保持统一的样式和分页逻辑
- 最终需要同时适用于客户产品手册、投标文件和 PDF

## 技术路线

核心工具：

- [`docxtpl`](https://github.com/elapouya/python-docx-template)：以 `.docx` 作为真正的视觉模板，通过 Jinja2 变量批量生成文档。官方文档明确支持图片、表格、页眉页脚等 Word 原生内容。
- [`python-docx`](https://github.com/python-openxml/python-docx)：创建和修改 `.docx`，负责图片尺寸、表格、样式、页眉页脚等精细控制。

## 设计原则

1. A4 竖版，不按 PPT 设计。
2. 截图是页面主体，但不是整页图片。
3. PC 截图、手机截图、组合截图分别处理。
4. 功能解释采用结构化能力卡/表格，而不是大段说明。
5. 截图与文字均保持 Word 可编辑。
6. 统一页边距、标题层级、图片边框、表格、页眉页脚。
7. 页面组件可重复使用，后续只替换数据和截图。

## 已包含的页面组件

- `standard`：标准功能页，截图 + 功能说明 + 应用价值
- `pc_annotated`：PC 截图 + 编号标注 + 右侧/下方解释
- `mobile_gallery`：1~3 个手机截图等宽排列
- `pc_mobile`：PC 后台 + 手机端组合
- `steps`：多截图步骤/流程页
- `dashboard`：BI / 风控 / 数据驾驶舱
- `process`：审批、预订、对账、结算等业务流程
- `overview`：章节/能力总览页

## 快速开始

```bash
python -m venv .venv
# Windows
.venv\\Scripts\\activate
# macOS / Linux
# source .venv/bin/activate

pip install -r requirements.txt
python src/generate_manual.py --data data/sample_manual.json --output output/zaitu_product_manual_demo.docx
```

首次运行会自动创建 `templates/zaitu_master_template.docx`，因此不需要手工制作一个空白 Word 模板。

## 推荐工作方式

后续正式制作时，建议把视觉规范固定在 `templates/zaitu_master_template.docx`，把每页的产品内容放到 JSON/YAML 数据中。这样修改截图或文字时，不需要重新排版整本手册。

## 目录

```text
word-product-manual/
├─ templates/             # Word 母版；首次运行自动生成
├─ data/                  # 产品内容数据
├─ assets/                # PC / 手机截图、图标、品牌资源
├─ src/                   # 生成器
├─ output/                # 生成的 docx
├─ requirements.txt
└─ README.md
```
