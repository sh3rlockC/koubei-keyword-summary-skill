# koubei-keyword-summary-skill

[![Release](https://img.shields.io/github/v/release/sh3rlockC/koubei-keyword-summary-skill)](https://github.com/sh3rlockC/koubei-keyword-summary-skill/releases)
[![Downloads](https://img.shields.io/github/downloads/sh3rlockC/koubei-keyword-summary-skill/total)](https://github.com/sh3rlockC/koubei-keyword-summary-skill/releases)
[![License](https://img.shields.io/github/license/sh3rlockC/koubei-keyword-summary-skill)](https://github.com/sh3rlockC/koubei-keyword-summary-skill/blob/main/LICENSE)

一个可复用的 OpenClaw Skill，用于读取汽车之家与懂车帝口碑 Excel，提炼高频优缺点、方向摘要、代表性原句，并默认输出**双平台整合摘要 Excel**。

## 适用场景

适用于这类任务：

- 已有汽车之家口碑 Excel，需要自动整理“最满意 / 最不满意”摘要
- 已有懂车帝口碑 Excel，需要从长评中归纳正向 / 负向方向
- 想把**汽车之家 + 懂车帝**两份口碑整合成一份汇报版 Excel
- 想直接生成适合业务、产品、老板看的**一页纸总结**
- 想同时保留统计口径、代表性原句、平台差异和产品机会点

## 核心能力

### 单平台模式（汽车之家）
- 读取包含 `最满意` / `最不满意` 的原始口碑 Excel
- 输出三口径方向统计、关键词、主题、业务摘要

### 双平台模式（推荐）
- 同时读取：
  - 汽车之家原始口碑 Excel
  - 懂车帝原始口碑 Excel
- 自动输出：
  - 一页纸总结
  - 总览摘要
  - 跨平台对比
  - 综合业务摘要
  - 产品机会点
  - 汽车之家满意/不满意摘要
  - 懂车帝正向/负向摘要

## 仓库结构

```text
skill/
├── SKILL.md
├── references/
└── scripts/
    └── summarize_koubei_excel.py
```

## 安装

### 方式 1：下载 release 中的 `.skill` 包

前往 Releases 页面下载：

- [Releases](https://github.com/sh3rlockC/koubei-keyword-summary-skill/releases)

然后按你的 OpenClaw / Skill 安装方式导入即可。

### 方式 2：直接使用脚本

## 用法

### 单平台模式（汽车之家）

```bash
python3 skill/scripts/summarize_koubei_excel.py \
  --input /path/to/ZJ口碑_车型.xlsx \
  --output /path/to/车型_汽车之家口碑摘要.xlsx \
  --model-name 车型名
```

### 双平台整合模式（推荐）

```bash
python3 skill/scripts/summarize_koubei_excel.py \
  --autohome-input /path/to/ZJ口碑_车型.xlsx \
  --dcd-input /path/to/DCD口碑_车型.xlsx \
  --output /path/to/车型_双平台口碑摘要.xlsx \
  --model-name 车型名
```

## 双平台输出内容

双平台模式默认生成这些 sheet：

1. `一页纸总结`
2. `总览摘要`
3. `跨平台对比`
4. `综合业务摘要`
5. `产品机会点`
6. `汽车之家_满意摘要`
7. `汽车之家_不满意摘要`
8. `懂车帝_正向摘要`
9. `懂车帝_负向摘要`

## 一页纸总结包含什么

自动放在最终 Excel 最前面，适合直接汇报。包含：

- 样本基础
- 一句话结论
- 核心优点 Top5
- 核心槽点 Top5
- 平台差异观察
- 汇报口径建议

## 单平台输出内容

汽车之家单平台模式默认生成这些 sheet：

1. `总览摘要`
2. `业务摘要`
3. `产品机会点`
4. `最满意方向统计`
5. `最不满意方向统计`
6. `最满意方向摘要`
7. `最不满意方向摘要`
8. `最满意主题`
9. `最不满意主题`
10. `最满意高频词`
11. `最不满意高频词`

## 统计口径说明

### 主方向口径
每条口碑只归到 1~3 个主方向。适合：
- 汇报
- 排序
- 接近人工阅读习惯的口径

### 条目口径
一条口碑命中某方向记 1 次。适合：
- 稳定统计
- 跨车型横向比较

### 语句归并口径
把口碑拆成更细的表达片段，再对相似语句归并累计。适合：
- 做深度洞察
- 看问题密度
- 发现高频细问题

## 示例

```bash
python3 skill/scripts/summarize_koubei_excel.py \
  --autohome-input ./ZJ口碑_启源A06.xlsx \
  --dcd-input ./DCD口碑_启源A06.xlsx \
  --output ./启源A06_双平台口碑摘要.xlsx \
  --model-name 启源A06
```

## Release 内容

每个 Release 默认包含：

- `.skill` 包
- GitHub 自动生成的源码压缩包

## License

MIT

## Changelog & Releases

- User-visible changes are tracked in [`CHANGELOG.md`](./CHANGELOG.md).
- For a new release, update the `Unreleased` section first, then cut the versioned release.
- GitHub Release notes should match the same user-visible changes, not just raw commit history.

