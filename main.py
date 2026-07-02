#!/usr/bin/env python3
"""
论文格式转换器 - 核心引擎
将 Markdown 论文稿按指定学术格式规范转换为排版规范的 DOCX 文档。

用法：
    python main.py --input paper.md --output paper.docx --format gbt7714
    python main.py --input paper.md --output paper.docx --format apa7 --custom-options custom.json
    python main.py --list-formats
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

# 中文格式后处理（首行缩进、标题间距等）
try:
    from docx import Document
    from docx.shared import Pt, Twips
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_PYTHON_DOCX = True
except ImportError:
    HAS_PYTHON_DOCX = False

# ============================================================
# 论文格式规范 → md-to-docx 参数映射
# 数据来源：GB/T 7713.1-2006、GB/T 7714-2025、河北大学学位论文规范、
#           APA 7th、MLA 9th、Chicago Manual of Style、IMRaD
# ============================================================

FORMAT_PRESETS = {
    # ---- 中文学位论文 (GB/T 7713.1 + GB/T 7714-2025 + 高校通用参数) ----
    "gbt7714": {
        "display_name": "GB/T 7714-2025 学位论文格式（中文）",
        "description": "依据GB/T 7713.1-2006结构框架 + GB/T 7714-2025引用规范 + 中文学位论文排版标准（宋体正文、黑体标题、小四号、1.5倍行距、首行缩进2字符、两端对齐）",
        "options": {
            "documentType": "report",
            "style": {
                "fontFamily": "宋体",
                "paragraphSize": 24,           # 小四号 ≈ 12pt (half-pts)
                "heading1Size": 32,             # 一级标题三号 ≈ 16pt
                "heading2Size": 28,             # 二级标题四号 ≈ 14pt
                "heading3Size": 24,             # 三级标题小四 ≈ 12pt
                "lineSpacing": 1.5,
                "headingSpacing": 12,           # 标题段前段后间距
                "paragraphSpacing": 0,
                "paragraphAlignment": "JUSTIFIED",
                "headingAlignment": "LEFT",
            },
            "template": {
                "page": {
                    "margin": {
                        "top": 1440,            # 2.54cm → twips
                        "bottom": 1440,
                        "left": 1800,           # 3.17cm → twips
                        "right": 1800,
                    }
                }
            },
        },
        "sections": [
            {"type": "cover", "title": "封面"},
            {"type": "abstract", "title": "摘要"},
            {"type": "abstract_en", "title": "Abstract"},
            {"type": "toc", "title": "目录"},
            {"type": "body", "title": "正文"},
            {"type": "references", "title": "参考文献"},
            {"type": "appendix", "title": "附录"},
            {"type": "acknowledgement", "title": "致谢"},
        ],
        "csl_style": "gb-t-7714-2025-numeric",
        "chinese_formatting": True,  # 启用中文格式后处理（首行缩进等）
    },

    # ---- APA 7th Edition ----
    "apa7": {
        "display_name": "APA 7th Edition",
        "description": "美国心理学会第7版格式，社会科学、教育学、心理学常用。标题页、作者-日期引用、参考文献悬挂缩进",
        "options": {
            "fontFamily": "Times New Roman",
            "paragraphSize": 12,
            "lineSpacing": 2.0,
            "pageMarginTop": 2.54,
            "pageMarginBottom": 2.54,
            "pageMarginLeft": 2.54,
            "pageMarginRight": 2.54,
        },
        "sections": [
            {"type": "title_page", "title": ""},
            {"type": "abstract", "title": "Abstract"},
            {"type": "body", "title": ""},
            {"type": "references", "title": "References"},
            {"type": "appendix", "title": "Appendix"},
        ],
    },

    # ---- MLA 9th Edition ----
    "mla9": {
        "display_name": "MLA 9th Edition",
        "description": "现代语言协会第9版格式，文学、语言、文化研究等人文学科常用。作者-页码引用、无独立标题页",
        "options": {
            "fontFamily": "Times New Roman",
            "paragraphSize": 12,
            "lineSpacing": 2.0,
            "pageMarginTop": 2.54,
            "pageMarginBottom": 2.54,
            "pageMarginLeft": 2.54,
            "pageMarginRight": 2.54,
        },
        "sections": [
            {"type": "body", "title": ""},
            {"type": "works_cited", "title": "Works Cited"},
        ],
    },

    # ---- Chicago Manual of Style (Notes-Bibliography) ----
    "chicago": {
        "display_name": "Chicago Manual of Style (Notes-Bibliography)",
        "description": "芝加哥格式注释-参考文献体系，历史学、艺术史、人类学常用。脚注/尾注 + 参考文献列表",
        "options": {
            "fontFamily": "Times New Roman",
            "paragraphSize": 12,
            "lineSpacing": 2.0,
            "pageMarginTop": 2.54,
            "pageMarginBottom": 2.54,
            "pageMarginLeft": 2.54,
            "pageMarginRight": 2.54,
        },
        "sections": [
            {"type": "title_page", "title": ""},
            {"type": "body", "title": ""},
            {"type": "bibliography", "title": "Bibliography"},
        ],
    },

    # ---- IMRaD 国际学术论文结构 ----
    "imrad": {
        "display_name": "IMRaD 国际学术论文结构",
        "description": "Introduction-Methods-Results-Discussion 标准科学论文结构。自然科学、医学、工程学投稿常用",
        "options": {
            "fontFamily": "Times New Roman",
            "paragraphSize": 11,
            "lineSpacing": 1.5,
            "pageMarginTop": 2.54,
            "pageMarginBottom": 2.54,
            "pageMarginLeft": 2.54,
            "pageMarginRight": 2.54,
        },
        "sections": [
            {"type": "title_page", "title": ""},
            {"type": "abstract", "title": "Abstract"},
            {"type": "introduction", "title": "Introduction"},
            {"type": "methods", "title": "Methods"},
            {"type": "results", "title": "Results"},
            {"type": "discussion", "title": "Discussion"},
            {"type": "conclusion", "title": "Conclusion"},
            {"type": "references", "title": "References"},
        ],
    },
}


def list_formats():
    """打印所有支持的格式规范"""
    print("支持的论文格式规范:\n")
    for key, preset in FORMAT_PRESETS.items():
        print(f"  {key:<15s}  {preset['display_name']}")
        print(f"  {'':15s}  {preset['description']}")
        print()


def build_options(format_name: str, custom_options: dict = None) -> dict:
    """根据格式名称构建 md-to-docx options JSON。

    从预设中加载排版参数，再与自定义 options 合并（自定义项覆写预设）。
    """
    if format_name not in FORMAT_PRESETS:
        available = ", ".join(FORMAT_PRESETS.keys())
        raise ValueError(f"未知格式 '{format_name}'。可用格式: {available}")

    preset = FORMAT_PRESETS[format_name]
    options = dict(preset["options"])

    if custom_options:
        options.update(custom_options)

    return options


def apply_chinese_formatting(docx_path: str, verbose: bool = False) -> None:
    """对生成的 DOCX 执行中文排版后处理。

    覆盖 md-to-docx 不支持的中文排版规范：
    1. 正文段落首行缩进 2 字符
    2. 标题保持无缩进、两端对齐确保
    """
    if not HAS_PYTHON_DOCX:
        if verbose:
            print("[警告] python-docx 不可用，跳过中文格式后处理")
        return

    import re
    doc = Document(docx_path)

    # 中文标题前缀模式
    heading_patterns = re.compile(
        r'^('
        r'[一二三四五六七八九十]+[\、\.\，]|'          # 一、 二、 等
        r'[（\(][一二三四五六七八九十]+[）\)]|'         # （一）(二) 等
        r'\d+[\、\.\．]|'                              # 1. 2. 等
        r'[\(（]\d+[）\)]|'                             # (1) （2）等
        r'第[一二三四五六七八九十百千]+[章节篇条]|'    # 第一章 等
        r'前言|引言|绪论|摘要|Abstract|目录|参考文献|附录|致谢|结语|结论'
        r')'
    )

    # 中文字号: 小四=12pt, 2字符缩进=24pt=480twips
    first_line_indent_twips = 480

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        pf = para.paragraph_format

        # 判断是否为标题段落
        is_heading = False

        # 检查 run 格式特征：标题通常加粗且字号 ≥ 14pt
        for run in para.runs:
            if run.font.size and run.font.size >= Pt(14):
                is_heading = True
                break
            if run.bold and run.font.size and run.font.size >= Pt(13):
                is_heading = True
                break

        # 检查文本前缀是否匹配标题模式
        if not is_heading and heading_patterns.match(text):
            is_heading = True

        # 全文标题（无序号的首行）通过加粗+居中等特征判断
        # 若全文第一段且加粗，视为标题
        if not is_heading and para is doc.paragraphs[0]:
            for run in para.runs:
                if run.bold:
                    is_heading = True
                    break

        if is_heading:
            # 标题：无首行缩进，左对齐
            pf.first_line_indent = None
            if pf.alignment is None:
                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            # 正文：首行缩进2字符，两端对齐
            if pf.first_line_indent is None:
                pf.first_line_indent = Twips(first_line_indent_twips)
            if pf.alignment is None:
                pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.save(docx_path)
    if verbose:
        print(f"[后处理] 中文格式已应用（首行缩进2字符、两端对齐、标题无缩进）")


def run_md_to_docx(input_path: str, output_path: str, options: dict, verbose: bool = False) -> str:
    """调用 npx @mohtasham/md-to-docx 执行转换。

    Args:
        input_path: 输入 Markdown 文件路径
        output_path: 输出 DOCX 文件路径
        options: 排版参数 dict
        verbose: 是否输出详细日志

    Returns:
        生成的 DOCX 文件路径

    Raises:
        RuntimeError: 转换失败
        subprocess.TimeoutExpired: 超时
    """
    # 将 options 写入临时 JSON 文件
    tmp_dir = tempfile.mkdtemp()
    options_path = os.path.join(tmp_dir, "options.json")

    with open(options_path, "w", encoding="utf-8") as f:
        json.dump(options, f, ensure_ascii=False, indent=2)

    if verbose:
        print(f"[配置] 排版参数:\n{json.dumps(options, ensure_ascii=False, indent=2)}")

    cmd = [
        "npx", "@mohtasham/md-to-docx",
        input_path,
        output_path,
        "--options", options_path,
    ]

    if verbose:
        print(f"[执行] {' '.join(cmd)}")

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        error_msg = (result.stderr or result.stdout or "未知错误").strip()
        # npx 安装日志不算错误
        if " installed" in error_msg and result.returncode == 0:
            pass
        else:
            raise RuntimeError(
                f"md-to-docx 转换失败 (code={result.returncode}): {error_msg}"
            )

    if verbose and result.stdout:
        print(f"[输出] {result.stdout.strip()}")

    # 验证输出文件
    output_file = Path(output_path)
    if not output_file.exists():
        raise RuntimeError(f"转换过程完成但输出文件未找到: {output_path}")

    return str(output_file.resolve())


def main():
    parser = argparse.ArgumentParser(
        description="论文格式转换器 — Markdown → 规范 DOCX",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  python main.py -i paper.md -o paper.docx -f gbt7714\n"
            "  python main.py -i paper.md -o paper.docx -f apa7 -c custom.json\n"
            "  python main.py --list-formats\n"
        ),
    )
    parser.add_argument("--input", "-i", required=True, help="输入 Markdown 文件路径")
    parser.add_argument("--output", "-o", required=True, help="输出 DOCX 文件路径")
    parser.add_argument(
        "--format", "-f",
        default="gbt7714",
        help=f"格式规范（默认: gbt7714）。可用: {', '.join(FORMAT_PRESETS.keys())}",
    )
    parser.add_argument(
        "--custom-options", "-c",
        help="自定义 JSON options 文件路径，可覆写预设中的任何参数",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="详细输出日志")
    parser.add_argument("--list-formats", action="store_true", help="列出所有支持的格式规范")

    args = parser.parse_args()

    # 列出格式
    if args.list_formats:
        list_formats()
        return

    # 验证输入文件
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误: 输入文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    # 加载自定义 options
    custom_options = None
    if args.custom_options:
        custom_path = Path(args.custom_options)
        if not custom_path.exists():
            print(f"错误: 自定义 options 文件不存在: {args.custom_options}", file=sys.stderr)
            sys.exit(1)
        try:
            with open(custom_path, "r", encoding="utf-8") as f:
                custom_options = json.load(f)
        except json.JSONDecodeError as e:
            print(f"错误: 自定义 options JSON 解析失败: {e}", file=sys.stderr)
            sys.exit(1)

    try:
        options = build_options(args.format, custom_options)
        result_path = run_md_to_docx(
            str(input_path.resolve()),
            str(Path(args.output).resolve()),
            options,
            args.verbose,
        )

        preset = FORMAT_PRESETS.get(args.format, {})
        format_name = preset.get("display_name", args.format)

        # 中文格式后处理（首行缩进等）
        if preset.get("chinese_formatting"):
            try:
                apply_chinese_formatting(result_path, args.verbose)
                print(f"   中文格式: 首行缩进2字符、两端对齐 ✅")
            except Exception as e:
                print(f"   ⚠️ 中文格式后处理异常: {e}")

        print(f"✅ 转换完成")
        print(f"   输出: {result_path}")
        print(f"   格式: {format_name}")

        # 摘要关键参数
        style = options.get("style", {})
        key_params = {k: style[k] for k in
                      ["fontFamily", "paragraphSize", "lineSpacing"]
                      if k in style}
        if key_params:
            print(f"   参数: {json.dumps(key_params, ensure_ascii=False)}")

    except ValueError as e:
        print(f"❌ 参数错误: {e}", file=sys.stderr)
        sys.exit(1)
    except subprocess.TimeoutExpired:
        print("❌ 转换超时（120秒），请检查网络或输入文件大小", file=sys.stderr)
        sys.exit(1)
    except RuntimeError as e:
        print(f"❌ 转换失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
