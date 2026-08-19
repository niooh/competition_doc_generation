> 从符合模板格式的一系列竞赛试题 .docx 自动生成相关文件的工具。支持从赛题文档中提取小题结构，生成批改统计表格、合并问卷等。

## Quick Start

### 1. 安装依赖

你需要有：
- 包管理器 [uv](https://github.com/astral-sh/uv)
- Python 3.11+（如果没装，uv 会自动安装）
- Windows 10+ 环境，支持通过 Microsoft Office 将 docx 转为 pdf

在 `pwsh / powershell` 运行以下命令，安装需要的包：

```pwsh
uv sync
```

### 2. 准备试题文档

- 在项目根目录下创建 `doc/`，将所有 `.docx` 试题文件放入此文件夹，或者一个特定的子文件夹 (如 `doc/初中组/`、`doc/高中组/`、`doc/大学组/`)。
- 可以用与 .docx 文件同层级的 `order.txt` 指定合并的顺序，每行一个名字。
- 试题文档需遵循模板要求：
  - 使用 Word 多级列表自动编号。
  - 每道小题末尾附带分值，格式为中文括号，例如“（5分）”，且分值括号必须位于行尾。
- 在 `src/config.py` 定义年份和组别。

### 3. 生成文档

#### 合并生成 PDF

```pwsh
uv run merge
# 或 python -m src.merge_docx_to_pdf
```

<details>
<summary>生成 <code>dist/merged.pdf</code>，即为试题文档。</summary>
包含封面、目录、正文，并添加水印。
封面的周期表图片来源于 IUPAC《Periodic Table of the Elements》2022 年 5 月 4 日版。原始文件见<a href="https://iupac.org/wp-content/uploads/2022/07/IUPAC_Periodic_Table-04May22_CRA.pdf">链接</a>。
</details>

#### 生成表格

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

### 4. 初步检查 DOCX 格式（实验性，可选）

```pwsh
uv run check
# 或 python -m src.utils.format_checker
```

## Develop & Contribute

如果你想为本代码库做出贡献，请阅读 [Contributing Guide](./.github/CONTRIBUTING.md)。

## Plan

- [x] 导出表格
- [x] 合并文档目录
- [ ] 检查文档格式
