# koubei-keyword-summary-skill

一个可复用的 OpenClaw Skill，用于读取汽车口碑 Excel，提炼“最满意 / 最不满意”的方向、主题、关键词、业务摘要，并导出正式交付版 Excel。

## 功能

- 从口碑 Excel 读取 `最满意` / `最不满意`
- 输出三口径统计：
  - 主方向口径
  - 条目口径
  - 语句归并口径
- 生成以下工作表：
  - 总览摘要
  - 业务摘要
  - 产品机会点
  - 最满意/最不满意方向统计
  - 最满意/最不满意方向摘要
  - 最满意/最不满意主题
  - 最满意/最不满意高频词

## 目录结构

```text
skill/
├── SKILL.md
└── scripts/
    └── summarize_koubei_excel.py
```

## 用法

```bash
python3 skill/scripts/summarize_koubei_excel.py \
  --input /path/to/koubei.xlsx \
  --output /path/to/output.xlsx \
  --model-name 启源A06
```

## 输出说明

推荐优先阅读：

1. `总览摘要`
2. `业务摘要`
3. `产品机会点`

其中：
- **主方向口径** 更接近人工阅读习惯
- **条目口径** 更适合稳定汇报
- **语句归并口径** 更适合深入分析

## 发布内容

Release 中会附带：
- `.skill` 包
- 源码 zip / tar.gz（GitHub 自动生成）

## License

MIT
