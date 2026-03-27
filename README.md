# koubei-keyword-summary-skill

[![Release](https://img.shields.io/github/v/release/sh3rlockC/koubei-keyword-summary-skill.)](https://github.com/sh3rlockC/koubei-keyword-summary-skill./releases)
[![Downloads](https://img.shields.io/github/downloads/sh3rlockC/koubei-keyword-summary-skill./total)](https://github.com/sh3rlockC/koubei-keyword-summary-skill./releases)
[![License](https://img.shields.io/github/license/sh3rlockC/koubei-keyword-summary-skill.)](https://github.com/sh3rlockC/koubei-keyword-summary-skill./blob/main/LICENSE)


一个可复用的 OpenClaw Skill，用于读取汽车口碑 Excel，提炼“最满意 / 最不满意”的方向、主题、关键词、业务摘要，并导出正式交付版 Excel。

## 适用场景

适用于这类任务：

- 已有汽车口碑采集结果 Excel，需要自动整理摘要
- 想从“最满意 / 最不满意”中提炼高频方向
- 想输出适合业务、产品、市场团队阅读的总结表
- 想同时保留统计口径与代表性原句，便于复核

## 核心能力

- 读取口碑 Excel 中的 `最满意` / `最不满意`
- 输出三口径统计：
  - **主方向口径**：每条口碑只归到 1~3 个主方向，最接近人工阅读习惯
  - **条目口径**：一条口碑命中某方向记 1 次，适合稳定汇报
  - **语句归并口径**：拆句后对相似表达归并计数，适合深入分析
- 自动生成：
  - 总览摘要
  - 业务摘要
  - 产品机会点
  - 最满意/最不满意方向统计
  - 最满意/最不满意方向摘要
  - 最满意/最不满意主题
  - 最满意/最不满意高频词

## 仓库结构

```text
skill/
├── SKILL.md
└── scripts/
    └── summarize_koubei_excel.py
```

## 安装

### 方式 1：下载 release 中的 `.skill` 包

前往 Releases 页面下载：

- [`koubei-keyword-summary.skill`](https://github.com/sh3rlockC/koubei-keyword-summary-skill/releases/latest)

然后按你的 OpenClaw / Skill 安装方式导入即可。

### 方式 2：直接使用脚本

```bash
python3 skill/scripts/summarize_koubei_excel.py \
  --input /path/to/koubei.xlsx \
  --output /path/to/output.xlsx \
  --model-name 启源A06
```

## 输出内容

生成的 Excel 默认包含这些 sheet：

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

## 推荐阅读顺序

如果你是第一次看结果，建议按这个顺序：

1. `总览摘要`
2. `业务摘要`
3. `产品机会点`
4. `最满意/最不满意方向统计`
5. `最满意/最不满意方向摘要`

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

## 示例命令

```bash
python3 skill/scripts/summarize_koubei_excel.py \
  --input ./ZJ启源A06最满意or最不满意_全量.xlsx \
  --output ./启源A06_口碑关键词摘要_正式交付版.xlsx \
  --model-name 启源A06
```

## Release 内容

每个 Release 默认包含：

- `.skill` 包
- GitHub 自动生成的源码压缩包

## License

MIT