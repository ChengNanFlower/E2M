# Excel 工作表转 Markdown

将 Excel 文件（.xlsx）中的所有工作表批量转换为一个格式整齐的 Markdown 文件。每个工作表对应一个二级标题章节，表格列宽自动对齐。

## 环境要求

- Python 3.7+
- openpyxl

```bash
pip install openpyxl
```

## 用法

```bash
# 基本用法：生成与输入文件同名的 .md 文件
python3 excel_to_md.py example.xlsx

# 指定输出文件
python3 excel_to_md.py data.xlsx output.md

# 指定文档一级标题
python3 excel_to_md.py data.xlsx -t "文档"

# 查看帮助
python3 excel_to_md.py -h
```

## 特性

- **自动识别标题行**：如果工作表首行只有少量内容（如合并单元格的表格名称），会自动识别为标题并以斜体展示
- **忽略空行**：自动跳过工作表中的空白行
- **列宽对齐**：根据每列内容自动计算最佳列宽，生成整齐的表格
- **批量处理**：一次性导出所有工作表

## 项目结构

```
.
├── excel_to_md.py                           # 转换脚本
└── README.md                                # 本文件
```

## 许可

MIT
