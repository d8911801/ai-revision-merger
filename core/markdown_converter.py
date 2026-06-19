"""
Markdown to DOCX Converter
使用 Pandoc 將 Markdown 轉換為 Word 文件（DOCX）

支援:
- reference.docx 模板套用
- 標題層級保留
- 段落格式保留
- 中文排版支援
"""

import subprocess
import tempfile
import os
from pathlib import Path
from utils.file_utils import resolve_output_conflict


class MarkdownConverter:
    """將 Markdown 文字轉換為 Word 文件 (.docx)"""

    def __init__(self, pandoc_path: str = "pandoc"):
        self.pandoc_path = pandoc_path

    def convert_to_docx(
        self,
        markdown_text: str,
        output_path: str,
        reference_docx: str | None = None,
    ) -> tuple[bool, str, str | None]:
        """
        將 Markdown 文字轉為 DOCX 檔案。

        Args:
            markdown_text: Markdown 文字內容
            output_path: 輸出 DOCX 的路徑
            reference_docx: 選用，Word 模板路徑（用來自訂樣式）

        Returns:
            (success: bool, message: str)
        """
        # 寫入暫存 Markdown 檔
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                suffix=".md",
                encoding="utf-8",
                delete=False,
            ) as f:
                f.write(markdown_text)
                temp_md_path = f.name
        except OSError as e:
            return False, f"無法建立暫存 Markdown 檔案：{e}", None

        try:
            result = self._run_pandoc(temp_md_path, output_path, reference_docx)
            return result
        finally:
            # 清理暫存 Markdown
            try:
                os.unlink(temp_md_path)
            except OSError:
                pass

    def _run_pandoc(
        self,
        input_path: str,
        output_path: str,
        reference_docx: str | None = None,
    ) -> tuple[bool, str, str | None]:
        """執行 Pandoc 轉換"""
        final_output_path, renamed = resolve_output_conflict(output_path)

        # 確保輸出目錄存在
        output_dir = os.path.dirname(final_output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        cmd = [
            self.pandoc_path,
            input_path,
            "-f", "markdown",
            "-t", "docx",
            "-o", final_output_path,
            # 中文排版友善選項
            "--wrap=none",
            "--markdown-headings=setext",
        ]

        # 如果有參考模板，加入參數
        if reference_docx and os.path.exists(reference_docx):
            cmd.extend(["--reference-doc", reference_docx])

        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )

            if proc.returncode != 0:
                error_msg = proc.stderr.strip() if proc.stderr else "Pandoc 轉換失敗（未知錯誤）"
                return False, f"Pandoc 錯誤 (code {proc.returncode}): {error_msg}", None

            # 驗證輸出檔案存在
            if not os.path.exists(final_output_path):
                return False, "Pandoc 完成但未產出檔案", None

            if renamed:
                return True, f"轉換成功；原輸出檔被占用，已改存為：{final_output_path}", final_output_path
            return True, "轉換成功", final_output_path

        except FileNotFoundError:
            return False, (
                "找不到 Pandoc。請確認已安裝 Pandoc 並加入系統 PATH。\n"
                "下載：https://pandoc.org/installing.html"
            ), None
        except subprocess.TimeoutExpired:
            return False, "Pandoc 執行逾時（超過 120 秒）", None
        except Exception as e:
            return False, f"Pandoc 執行錯誤：{e}", None

    @staticmethod
    def _Sanitize_Markdown(markdown_text: str) -> str:
        """清理 AI 產出的 Markdown（移除 HTML、程式碼框等）"""
        import re

        # 移除 HTML 標籤
        markdown_text = re.sub(r'<[^>]+>', '', markdown_text)

        # 注意：保留 fenced code blocks 可能包含有用內容
        # 此處根據需求可選擇保留或移除
        # 目前保留，讓 Pandoc 自行處理

        return markdown_text
