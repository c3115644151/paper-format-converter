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

# ============================================================
# 论文格式规范 → md-to-docx 参数映射
# 数据来源：GB/T 7713.1-2006、GB/T 7714-2025、河北大学学位论文规范、
#           APA 7th、MLA 9th、Chicago Manual of Style、IMRaD
# ============================================================

FORMAT_PRESETS = {
    # ---- 中文学位论文 (GB/T 7713.1 + GB/T 7714-2025 + 高校通用参数) ----
    "gbt7714": {
        "display_name": "GB/T 7714-2025 学位论文格式",
        "description": "依据GB/T 7713.1-2006结构框架 + GB/T 7714-2025引用规范 + 高校通用排版参数（正文字体宋体、标题黑体、小四号正文、1.5倍行距）",
        "options": {
            "fontFamily": "宋体",
            "paragraphSize": 12,          # 小四号 ≈ 12pt
            "headingFontFamily": "黑体",
            "headingSize": 15,            # 一级标题小三 ≈ 15pt
            "lineSpacing": 1.5,
            "pageMarginTop": 2.54,        # 上 2.54cm
            "pageMarginBottom": 2.54,     # 下 2.54cm
            "pageMarginLeft": 3.17,       # 左 3.17cm（装订侧）
            "pageMarginRight": 3.17,      # 右 3.17cm
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
        "csl_style": "gb-t-7714-2025-numeric",  # GB/T 7714 顺序编码制
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

        print(f"✅ 转换完成")
        print(f"   输出: {result_path}")
        print(f"   格式: {format_name}")

        # 摘要关键参数
        key_params = {k: options[k] for k in
                      ["fontFamily", "paragraphSize", "lineSpacing"]
                      if k in options}
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
