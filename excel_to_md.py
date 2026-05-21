#!/usr/bin/env python3
"""将 Excel 文件中的所有工作表批量转换为 Markdown 表格文件。

用法:
    python3 excel_to_md.py input.xlsx
    python3 excel_to_md.py input.xlsx output.md
    python3 excel_to_md.py input.xlsx -t "文档标题"
"""

import argparse
import os
import sys

import openpyxl


def is_empty_row(row_values):
    """判断整行是否为空（所有单元格都是 None 或空字符串）。"""
    return all(v is None or str(v).strip() == "" for v in row_values)


def row_values(ws, row_num, max_col):
    """读取工作表中某一行的所有单元格值，None 转为空字符串。"""
    return [ws.cell(row=row_num, column=c).value for c in range(1, max_col + 1)]


def format_md_table(headers, data_rows):
    """将表头和数据行格式化为对齐的 Markdown 表格。"""
    all_rows = [headers] + data_rows
    str_rows = [[str(v) if v is not None else "" for v in row] for row in all_rows]

    # 计算每列的最大宽度
    col_widths = [max(len(row[c]) for row in str_rows) for c in range(len(headers))]
    # 分隔线至少需要 3 个字符
    col_widths = [max(w, 3) for w in col_widths]

    lines = []
    # 表头行
    header_cells = [str_rows[0][c].ljust(col_widths[c]) for c in range(len(headers))]
    lines.append("| " + " | ".join(header_cells) + " |")
    # 分隔线
    sep_cells = ["-" * col_widths[c] for c in range(len(headers))]
    lines.append("| " + " | ".join(sep_cells) + " |")
    # 数据行
    for row in str_rows[1:]:
        cells = [row[c].ljust(col_widths[c]) for c in range(len(headers))]
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


def process_sheet(ws):
    """从工作表中提取标题行、表头和数据行。

    自动识别标题行：如果第一行只有 1-2 个非空单元格且第二行非空单元格较多，
    则认为第一行是标题，第二行是表头。
    """
    max_col = ws.max_column
    all_data = []
    for r in range(1, ws.max_row + 1):
        vals = row_values(ws, r, max_col)
        if not is_empty_row(vals):
            all_data.append(vals)

    if not all_data:
        return None, None, []

    # 检测首行是否为标题行
    title = None
    start_idx = 0
    if len(all_data) >= 2:
        first_row = all_data[0]
        non_empty_count = sum(1 for v in first_row if v is not None and str(v).strip() != "")
        second_row = all_data[1]
        second_non_empty = sum(1 for v in second_row if v is not None and str(v).strip() != "")
        if non_empty_count <= 2 and second_non_empty >= 3:
            title = str(first_row[0]) if first_row[0] else None
            start_idx = 1

    headers = [str(v) if v is not None else "" for v in all_data[start_idx]]
    data_rows = all_data[start_idx + 1:]

    return title, headers, data_rows


def main():
    parser = argparse.ArgumentParser(
        description="将 Excel 文件中的所有工作表转换为 Markdown 表格文件"
    )
    parser.add_argument("input", help="输入的 Excel 文件路径（.xlsx）")
    parser.add_argument("output", nargs="?", help="输出的 Markdown 文件路径（可选，默认与输入同名）")
    parser.add_argument("-t", "--title", help="Markdown 文档的一级标题（可选，默认使用文件名）")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"错误：文件不存在 — {args.input}", file=sys.stderr)
        sys.exit(1)

    # 确定输出路径
    if args.output:
        output_path = args.output
    else:
        base = os.path.splitext(os.path.basename(args.input))[0]
        output_path = f"{base}.md"

    # 确定文档标题
    doc_title = args.title if args.title else os.path.splitext(os.path.basename(args.input))[0]

    wb = openpyxl.load_workbook(args.input)

    sections = []

    for name in wb.sheetnames:
        ws = wb[name]
        title, headers, data_rows = process_sheet(ws)

        if not headers or not data_rows:
            continue

        lines = []
        lines.append(f"## {name}")
        lines.append("")

        display_title = title if title else name
        if display_title:
            lines.append(f"*{display_title}*")
            lines.append("")

        lines.append(format_md_table(headers, data_rows))
        lines.append("")
        sections.append("\n".join(lines))

    wb.close()

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"# {doc_title}\n\n")
        f.write("\n".join(sections))

    print(f"完成！共 {len(sections)} 个工作表已写入 {output_path}")


if __name__ == "__main__":
    main()
