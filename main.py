"""
AI 學術修訂合併器（AI Revision Merger）
==========================================
Windows 桌面工具：將 AI 產生的 Markdown 修訂稿自動合併到原始 Word 文件，
保留 Microsoft Word 原生追蹤修訂。

兩種入口：
  GUI：  python main.py
  CLI：  python main.py --cli --markdown revised.md --output out.docx
          python main.py --cli --original old.docx --markdown revised.md
"""

import sys
import os
import argparse

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


def run_cli(args):
    """CLI 模式：提供給 Codex / 其他工具直接呼叫"""
    from core.markdown_converter import MarkdownConverter
    from core.insdel_to_docx import InsDelToDocx
    from core.word_merger import WordMerger
    from utils.file_utils import suggest_output_path, validate_docx, validate_markdown, get_temp_path

    exit_code = 0

    # ── 讀取 Markdown ──
    md_path = args.markdown
    if os.path.exists(md_path):
        with open(md_path, "r", encoding="utf-8") as f:
            md_content = f.read()
        print(f"[CLI] 已讀取 Markdown: {md_path} ({len(md_content)} 字元)")
    else:
        md_content = md_path  # 直接當作文字內容
        print(f"[CLI] 使用內嵌 Markdown 文字 ({len(md_content)} 字元)")

    valid, msg = validate_markdown(md_content)
    if not valid:
        print(f"[CLI] ✗ Markdown 驗證失敗: {msg}")
        return 1

    # ── 判斷輸出路徑 ──
    if args.output:
        output_path = args.output
    elif args.original:
        output_path = suggest_output_path(args.original)
    else:
        output_path = None

    has_insdel = "<ins>" in md_content or "<del>" in md_content

    # ── 管線 1：含 <ins>/<del> → 直接追蹤修訂 ──
    if has_insdel:
        print("[CLI] 偵測到 <ins>/<del> 標記 → 使用 InsDelToDocx 管線")
        converter = InsDelToDocx(author=args.author or "AI Revision")
        wm = WordMerger()

        # 若未指定輸出，使用 --original 同目錄或桌面
        if not output_path:
            import tempfile
            output_path = os.path.join(tempfile.gettempdir(), "tracked_revision.docx")

        success, msg, final_output_path = converter.convert(md_content, output_path)
        if success:
            output_path = final_output_path or output_path
            print(f"[CLI] ✅ {msg}")
            display_ok, display_msg = wm.ensure_revision_display_colors()
            if display_ok:
                print(f"[CLI] ✅ {display_msg}")
            else:
                print(f"[CLI] ⚠ {display_msg}")
        else:
            print(f"[CLI] ❌ {msg}")
            exit_code = 1

    # ── 管線 2：有 old.docx → Pandoc + Word Compare ──
    elif args.original:
        print("[CLI] 管線: Markdown → Pandoc → Word Compare")

        # 步驟 1: Pandoc
        print("[CLI] [1/2] Pandoc 轉換 Markdown → DOCX...")
        temp_docx = get_temp_path(prefix="cli_merge_")

        ref_docx = args.reference_docx or None
        mc = MarkdownConverter()
        success, msg, final_temp_docx = mc.convert_to_docx(md_content, temp_docx, reference_docx=ref_docx)
        if not success:
            print(f"[CLI] ❌ Pandoc 失敗: {msg}")
            return 1
        temp_docx = final_temp_docx or temp_docx

        print("[CLI] [2/2] Word 原生文件比較...")
        if not output_path:
            output_path = suggest_output_path(args.original)

        wm = WordMerger()
        success, msg, final_output_path = wm.compare_documents(
            original_path=args.original,
            revised_path=temp_docx,
            output_path=output_path,
            author_name=args.author or "AI Revision Merger",
        )

        try:
            os.unlink(temp_docx)
        except OSError:
            pass

        if success:
            output_path = final_output_path or output_path
            print(f"[CLI] ✅ {msg}")
            display_ok, display_msg = wm.ensure_revision_display_colors()
            if display_ok:
                print(f"[CLI] ✅ {display_msg}")
            else:
                print(f"[CLI] ⚠ {display_msg}")
        else:
            print(f"[CLI] ❌ {msg}")
            exit_code = 1

    # ── 管線 3：純 Markdown → DOCX ──
    else:
        print("[CLI] 管線: Markdown → Pandoc → DOCX (純轉換)")

        if not output_path:
            import tempfile
            output_path = os.path.join(tempfile.gettempdir(), "output.docx")

        mc = MarkdownConverter()
        success, msg, final_output_path = mc.convert_to_docx(
            md_content, output_path,
            reference_docx=args.reference_docx or None,
        )
        if success:
            output_path = final_output_path or output_path
            print(f"[CLI] ✅ {msg}")
        else:
            print(f"[CLI] ❌ {msg}")
            exit_code = 1

    # ── 結果摘要 ──
    if exit_code == 0:
        print(f"[CLI] 輸出: {output_path}")
    return exit_code


def main():
    parser = argparse.ArgumentParser(
        description="AI 學術修訂合併器 — CLI 模式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
  # 含 <ins>/<del> 的 Markdown → 追蹤修訂 DOCX
  python main.py --cli --markdown revised.md --output result.docx

  # Markdown vs 原始 old.docx → 追蹤修訂 (Word Compare)
  python main.py --cli --original old.docx --markdown revised.md

  # 純 Markdown → DOCX（無追蹤修訂）
  python main.py --cli --markdown revised.md --output out.docx
""",
    )
    parser.add_argument("--cli", action="store_true", help="使用 CLI 模式（預設為 GUI）")
    parser.add_argument("--original", "-o", help="原始 Word 檔案路徑 (old.docx)")
    parser.add_argument("--markdown", "-m", help="Markdown 檔案路徑，或直接傳入 Markdown 文字")
    parser.add_argument("--output", help="輸出 DOCX 路徑（若不指定則自動建議）")
    parser.add_argument("--reference-docx", help="選用：Word 模板 (reference.docx)")
    parser.add_argument("--author", help="追蹤修訂作者名稱")

    args = parser.parse_args()

    if args.cli:
        # CLI 模式
        if not args.markdown:
            print("❌ CLI 模式需指定 --markdown 參數")
            sys.exit(1)
        sys.exit(run_cli(args))
    else:
        # GUI 模式（預設）
        from gui.app import main as gui_main
        gui_main()


if __name__ == "__main__":
    main()
