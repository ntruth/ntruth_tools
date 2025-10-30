# TXT 文案转 Excel 工具

本仓库提供一个基于 Python 的小工具，可将 TXT 文案按照需求文档中
的规则整理后，批量填充到指定的 Excel 模板第一列中。工具同时支持
命令行与图形界面两种使用方式，便于在 Windows 环境下快速整理营销
文案、素材文档等文本数据。

## 功能特性

- **换行自动加逗号**：同一段落中每遇到一次换行都会插入中文逗号“，”。
- **空行分组**：以空行为分隔，将文案划分为独立的文案单元。
- **模板写入**：使用固定的 Excel 模板，只填充第一列，其余格式保持不变。
- **GUI/CLI 双模式**：既可双击运行图形界面，也支持脚本化处理。

## 环境准备

1. 安装 Python 3.8 及以上版本（Windows 默认自带 Tkinter）。
2. 安装依赖：

   ```bash
   pip install openpyxl
   ```

工具内置了一个简单的 Excel 模板，会在第一次运行时自动生成到
`templates/txt_to_excel_template.xlsx`。你也可以通过运行
`python -m templates` 将模板导出到任意位置，或者替换成自定义的模板。

## 图形界面使用方法

1. 双击或通过命令 `python txt_to_excel_tool.py --gui` 启动图形界面。
2. 在界面中依次选择 TXT 文案文件、Excel 模板（默认为仓库模板）和输出
   Excel 文件保存位置。
3. 点击“开始转换”，程序会按照规则生成结果并提示写入的文案条数。

## 命令行使用方法

在命令行中执行：

```bash
python txt_to_excel_tool.py --txt 输入.txt --output 输出.xlsx
```

可选参数：

- `--template`：指定自定义的 Excel 模板路径，默认使用仓库内置模板。
- `--start-row`：设置写入 Excel 的起始行（默认为 1）。

示例：

```bash
python txt_to_excel_tool.py --txt data/sample.txt \
    --template templates/txt_to_excel_template.xlsx \
    --output dist/result.xlsx
```

执行完成后，处理后的每个文案单元都会放在 Excel 的第一列中，行内
换行被中文逗号替换，空行则作为文案分隔。

## 文案处理规则回顾

1. TXT 文案以 UTF-8 编码读取。
2. 同一文案段落的多行内容用中文逗号串联。
3. 遇到空行就结束当前文案段并写入 Excel。
4. 只会写入 Excel 的第一列，其余列保持模板原样。

如需扩展更多功能（如批量选择、列映射等），可以在 `txt_to_excel_tool.py`
基础上进一步开发。
