#!/usr/bin/env python3
"""
测试报告生成器
用法: python generate_report.py <markdown_file> [--pdf]
"""

import re
import sys
import os
import subprocess
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_PATH = SCRIPT_DIR / "template.html"
EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def find_browser():
    for p in EDGE_PATHS:
        if os.path.exists(p):
            return p
    return None


def escape_html(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def parse_markdown(md_text):
    lines = md_text.strip().split("\n")
    result = {
        "title": "",
        "status": "",
        "command": "",
        "meta": "",
        "test_cases": [],
    }

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if line.startswith("# "):
            title_match = re.search(r"# (.+?) — (.+?) — (.+)", line)
            if title_match:
                result["title"] = title_match.group(1)
                result["status"] = title_match.group(3)
            else:
                result["title"] = line[2:]
                result["status"] = "测试报告"

        elif line.startswith("```bash"):
            i += 1
            cmd_lines = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                cmd_lines.append(lines[i].strip())
                i += 1
            result["command"] = "\n".join(cmd_lines)

        elif line.startswith("**模型**") or line.startswith("**Model**"):
            result["meta"] = f"<p>{escape_html(line)}</p>"

        elif line.startswith("### TC-"):
            tc = parse_test_case(lines, i)
            if tc:
                result["test_cases"].append(tc)
                i = tc["end_index"]
                continue

        i += 1

    return result


def parse_test_case(lines, start_idx):
    line = lines[start_idx].strip()
    match = re.match(r"### (TC-\d+)(.*)", line)
    if not match:
        return None

    tc_id = match.group(1)
    tc_title_raw = match.group(2).strip()
    tc_title = re.sub(r"[✅❌⚠️]", "", tc_title_raw).strip()
    passed = "✅" in tc_title_raw

    tc = {
        "id": tc_id,
        "title": f"{tc_id} {tc_title}",
        "passed": passed,
        "summary_lines": [],
        "table_header": [],
        "table_rows": [],
        "end_index": start_idx + 1,
    }

    i = start_idx + 1
    while i < len(lines):
        line = lines[i].strip()

        if line.startswith("### ") or (line.startswith("# ") and not line.startswith("### ")):
            break

        if line.startswith("|") and not tc["table_header"]:
            tc["table_header"] = [c.strip() for c in line.split("|")[1:-1]]
            i += 1
            if i < len(lines) and re.match(r"\|[-|]+\|", lines[i].strip()):
                i += 1
            continue

        if line.startswith("|") and tc["table_header"]:
            row = [c.strip() for c in line.split("|")[1:-1]]
            tc["table_rows"].append(row)
            i += 1
            continue

        if line:
            tc["summary_lines"].append(line)

        i += 1

    tc["end_index"] = i
    return tc


def render_table(header, rows):
    if not header:
        return ""
    html = '<table><thead><tr>'
    for h in header:
        html += f"<th>{escape_html(h)}</th>"
    html += "</tr></thead><tbody>"
    for row in rows:
        html += "<tr>"
        for cell in row:
            cell_html = escape_html(cell)
            if "✅" in cell:
                cell_html = cell_html.replace("✅", '<span class="pass">✅</span>')
            elif "❌" in cell:
                cell_html = cell_html.replace("❌", '<span class="fail">❌</span>')
            elif re.match(r"^\d+(\.\d+)?$", cell.strip()):
                cell_html = f'<span class="highlight">{cell_html}</span>'
            html += f"<td>{cell_html}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html


def render_test_case(tc):
    check = "✅" if tc["passed"] else "❌"
    html = f'<div class="test-case">\n'
    html += f'  <h2><span class="check">{check}</span> {escape_html(tc["title"])}</h2>\n'

    if tc["summary_lines"]:
        html += '  <div class="summary">\n'
        for line in tc["summary_lines"]:
            line_html = escape_html(line)
            line_html = re.sub(
                r"(\d+(?:\.\d+)?)",
                r'<span class="highlight">\1</span>',
                line_html,
            )
            html += f"    <p>{line_html}</p>\n"
        html += "  </div>\n"

    if tc["table_header"]:
        html += render_table(tc["table_header"], tc["table_rows"])
        html += "\n"

    html += "</div>\n"
    return html


def generate_html(parsed):
    template = TEMPLATE_PATH.read_text(encoding="utf-8")

    meta_html = parsed["meta"]
    if not meta_html:
        meta_html = "<p>测试报告</p>"

    test_cases_html = "\n".join(render_test_case(tc) for tc in parsed["test_cases"])

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    footer = f"<p>报告生成时间：{now}</p>"

    html = template.replace("{{TITLE}}", parsed["title"])
    html = html.replace("{{STATUS}}", parsed["status"])
    html = html.replace("{{META}}", meta_html)
    html = html.replace("{{COMMAND}}", escape_html(parsed["command"]))
    html = html.replace("{{TEST_CASES}}", test_cases_html)
    html = html.replace("{{FOOTER}}", footer)

    return html


def generate_pdf(html_path, pdf_path):
    browser = find_browser()
    if not browser:
        print("未找到 Chrome/Edge 浏览器，跳过 PDF 生成")
        return False

    try:
        result = subprocess.run(
            [
                browser,
                "--headless",
                "--disable-gpu",
                f"--print-to-pdf={pdf_path}",
                "--no-pdf-header-footer",
                f"file:///{html_path}",
            ],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if os.path.exists(pdf_path):
            print(f"PDF 已生成: {pdf_path}")
            return True
        else:
            print(f"PDF 生成失败: {result.stderr}")
            return False
    except Exception as e:
        print(f"PDF 生成异常: {e}")
        return False


def main():
    if len(sys.argv) < 2:
        print("用法: python generate_report.py <markdown_file> [--pdf]")
        print("示例: python generate_report.py test_report.md --pdf")
        sys.exit(1)

    md_path = Path(sys.argv[1])
    gen_pdf = "--pdf" in sys.argv

    if not md_path.exists():
        print(f"文件不存在: {md_path}")
        sys.exit(1)

    md_text = md_path.read_text(encoding="utf-8")
    parsed = parse_markdown(md_text)
    html = generate_html(parsed)

    html_path = md_path.with_suffix(".html")
    html_path.write_text(html, encoding="utf-8")
    print(f"HTML 已生成: {html_path}")

    if gen_pdf:
        pdf_path = md_path.with_suffix(".pdf")
        generate_pdf(str(html_path.resolve()), str(pdf_path))


if __name__ == "__main__":
    main()
