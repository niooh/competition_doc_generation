> 自动从 docx 生成竞赛答卷相关文件的工具。支持从 `.docx` 赛题文档中提取小题结构，生成带公式的批改工作簿。

## Quick Start

### 1. 安装依赖

你需要有：
- Python 3.11+
- 包管理器 [uv](https://github.com/astral-sh/uv)
- Windows 10+ 环境（如果需要支持通过 Microsoft Office 将 docx 转为 pdf）

然后在 pwsh 运行以下命令安装需要的包：

```pwsh
uv sync
```

如果你在 wsl2 中，需要使用以下方式调用命令：

```bash
pwsh.exe -Command <cmd>
```

### 2. 准备试题文档

- 在项目根目录下创建 `doc/`，将所有 `.docx` 试题文件放入此文件夹。
- 试题文档需遵循模板要求：
  - 使用 Word 多级列表（自动编号），最终会被解析为 `1-1`、`2-1-1` 等形式。
  - 每道小题末尾附带分值，格式为中文括号，例如“（5分）”，且分值括号必须位于行尾。

### 3. 生成表格

```pwsh
uv run xlsx
# 或 python -m src.generate_xlsx
```

<details>
<summary>生成 <code>dist/final.xlsx</code>，即为批改表格。</summary>

生成表格包含以下部分：

| Sheet 名称 | 用途 |
|-----------|------|
| 答卷基本信息 | 填写用户名/ID、选题情况、查看批改完成状态，附有使用说明 |
| 各题分表 | 每题一个 sheet，自动填入小题号与满分，得分处可手动填写 |
| 总分统计 | 自动计算全卷总分、最高三题总分及排名 |
</details>

### 4. 合并生成 PDF

```pwsh
uv run merge
# 或 python -m src.merge_docx_to_pdf
```

<details>
<summary>生成 <code>dist/merged.pdf</code>，即为批改表格。</summary>
包含封面、目录、正文及灰色倾斜水印。
</details>

## Develop & Contribute

```pwsh
uv sync --group dev
uvx ty check .  # 检查类型
```

## Plan

- [x] 导出表格
- [x] 合并文档目录
- [ ] 检查文档格式
