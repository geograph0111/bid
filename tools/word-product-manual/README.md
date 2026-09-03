# 在途商旅 Word 企业产品手册母版

面向 TMC / 商旅 SaaS 的 A4 Word 产品功能手册与投标文件产品能力章节生成器。

## 本次版本重点：5 套可切换母版

这不是把 PPT 版式搬进 Word，而是以 **Word 原生可编辑 + 投标评审友好 + SaaS 截图证据化** 为第一原则。生成器现在提供 5 套不同视觉与排版逻辑的 Master Style：

1. `clean_blue`：专业蓝白，通用技术标主母版
2. `teal_grid`：现代青绿，系统能力 / 集成 / SaaS 能力
3. `indigo_report`：稳重靛蓝，数据 / BI / 风控 / 管理后台
4. `amber_standard`：商务金棕，服务 / 实施 / 运营保障
5. `green_governance`：治理绿白，合规 / 政策 / 风控 / 企业治理

五套样式共用同一套内容数据，因此后续可以只切换 `--style`，不重新整理产品内容。

## 页面组件

- `standard`：标准功能页
- `mobile_gallery`：1~3 个手机截图等宽排列
- `pc_mobile`：PC 管理端 + 手机员工端协同
- `steps` / `process`：业务流程与步骤
- `dashboard`：BI / 风控 / 数据驾驶舱
- `overview`：章节 / 能力总览

## 投标文件设计原则

- A4 竖版，适合 Word / PDF / 打印
- 截图是功能证据，不是装饰
- 核心文字保持 Word 可编辑、可搜索、可复制
- PC / 手机截图统一比例和尺寸逻辑
- 一页一主题，避免为了省页数把内容压成小字
- 表格、能力卡、流程、指标等均作为可复用组件
- 不使用整页海报式图片，不做 PPT 化排版

## 运行

```bash
python src/generate_manual.py --data data/sample_manual.json --style clean_blue
```

生成其他样式：

```bash
python src/generate_manual.py --data data/sample_manual.json --style teal_grid
python src/generate_manual.py --data data/sample_manual.json --style indigo_report
python src/generate_manual.py --data data/sample_manual.json --style amber_standard
python src/generate_manual.py --data data/sample_manual.json --style green_governance
```

一次生成 5 套 Demo：

```bash
python src/generate_manual.py --data data/sample_manual.json --all-styles
```

## 商旅 SaaS 截图适配

正式使用时，把 PC 后台、移动端、BI 看板、流程页面截图放入 `assets/screenshots/`，再通过数据文件绑定到对应页面。后续计划增加自动裁切、PC/手机识别、编号标注和正式 `.docx` 视觉母版。

## 目录

```text
word-product-manual/
├─ templates/             # Word 母版
├─ data/                  # 产品内容数据
├─ assets/                # PC / 手机截图、图标、品牌资源
├─ src/                   # 生成器
├─ docs/                  # 母版设计规范
├─ output/                # 生成的 docx
└─ requirements.txt
```
