"""
Word Document Merger (via COM Automation)
利用 Microsoft Word 原生「文件比較」功能，
將 AI 修訂稿合併到原始文件，產出帶有原生追蹤修訂的 DOCX。

核心 API: Document.Compare()
參考: https://learn.microsoft.com/en-us/office/vba/api/word.document.compare
"""

import os
import time
from pathlib import Path
from utils.file_utils import resolve_output_conflict

# pywin32 COM client（執行期載入）
try:
    import pythoncom
    import win32com.client
    from win32com.client import constants as wdConst
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    pythoncom = None
    win32com = None
    wdConst = None


class WordMerger:
    """
    使用 Word COM Automation 比較兩個 DOCX，
    產出帶原生追蹤修訂的 compared.docx。
    """

    def __init__(self):
        self._app = None
        self._com_initialized = False

    # ── Compare 參數常數（WD_COMPARE_GRANULARITY enum）────────────────
    GRANULARITY_CHAR = 1
    GRANULARITY_WORD = 2

    # ── 狀態常數 ──
    STATE_READY = 0
    STATE_COMPARING = 1
    STATE_DONE = 2
    STATE_ERROR = -1

    @property
    def is_available(self) -> bool:
        """檢查 COM 是否可用（端是否安裝 Word 與 pywin32）"""
        return COM_AVAILABLE

    def compare_documents(
        self,
        original_path: str,
        revised_path: str,
        output_path: str,
        author_name: str = "AI Revision Merger",
        granularity: int | None = None,
    ) -> tuple[bool, str, str | None]:
        """
        比較兩個 DOCX 並產出帶追蹤修訂的結果。

        Args:
            original_path: 原始 Word 檔案路徑
            revised_path:  AI 修訂後的新 Word 檔案路徑
            output_path:   輸出 compared.docx 路徑
            author_name:   追蹤修訂顯示的作者名稱
            granularity:   比較精細度 (1=字元級, 2=單詞級)

        Returns:
            (success: bool, message: str)
        """
        if not COM_AVAILABLE:
            return False, (
                "Word COM Automation 不可用。\n"
                "請確認：\n"
                "1. 已安裝 Microsoft Word\n"
                "2. 已安裝 pywin32 (pip install pywin32)"
            ), None

        # 驗證輸入檔案
        for label, path in [("原始檔案", original_path), ("修訂稿", revised_path)]:
            if not os.path.exists(path):
                return False, f"{label}不存在：{path}", None

        if granularity is None:
            granularity = self.GRANULARITY_WORD

        final_output_path, renamed = resolve_output_conflict(output_path)

        # 確保輸出目錄存在
        output_dir = os.path.dirname(final_output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        try:
            # 初始化 COM
            self._init_com()

            # 啟動 Word（Visible=False 後台執行）
            self._app = win32com.client.Dispatch("Word.Application")
            self._app.Visible = False
            self._app.DisplayAlerts = 0  # wdAlertsNone
            self._app.ScreenUpdating = False

            # 開啟原始文件
            original_doc = self._app.Documents.Open(
                os.path.abspath(original_path),
            )

            # ── 核心：呼叫原生 Compare ──
            # Document.Compare(
            #     Name,                    # 要比較的檔案
            #     AuthorName,              # 修訂作者名稱
            #     CompareTarget,           # 比較目標
            #     DetectFormatChanges,     # 偵測格式變更
            #     Granularity              # 比較精細度
            # )
            try:
                original_doc.Compare(
                    os.path.abspath(revised_path),
                    author_name,
                    wdConst.wdCompareTargetNew if wdConst else 1,
                    True,     # DetectFormatChanges
                    granularity,
                )
            except Exception as e:
                # 關閉文件並清理
                self._safe_close_doc(original_doc)
                self._quit_word()
                return False, f"Word Compare 執行失敗：{e}", None

            # 接受所有修訂後另存（保留 Track Changes）
            self._finalize_document(original_doc)

            # 另存為輸出檔案
            output_abs = os.path.abspath(final_output_path)
            try:
                original_doc.SaveAs2(output_abs, FileFormat=16)  # wdFormatDocumentDefault
            except Exception:
                # Fallback: try SaveAs
                original_doc.SaveAs(output_abs, FileFormat=16)

            # 關閉文件
            original_doc.Close(SaveChanges=False)

            # 關閉 Word
            self._quit_word()

            # 驗證輸出
            if os.path.exists(final_output_path):
                if renamed:
                    return True, f"追蹤修訂文件已產出；原輸出檔被占用，已改存為：{final_output_path}", final_output_path
                return True, f"追蹤修訂文件已產出：{final_output_path}", final_output_path
            else:
                return False, "文件比較完成但未產出檔案", None

        except Exception as e:
            self._quit_word()
            return False, f"Word 操作異常：{e}", None

    def ensure_revision_display_colors(
        self,
        inserted_color: int | None = None,
        deleted_color: int | None = None,
        inserted_mark: int | None = None,
        deleted_mark: int | None = None,
    ) -> tuple[bool, str]:
        """
        設定 Word 全域的修訂顯示樣式。
        這是應用程式層級偏好，不是單一文件層級設定。
        """
        if not COM_AVAILABLE:
            return False, "Word COM Automation 不可用，無法調整修訂顯示色彩"

        inserted_color = inserted_color if inserted_color is not None else (getattr(wdConst, "wdBlue", 2) if wdConst else 2)
        deleted_color = deleted_color if deleted_color is not None else (getattr(wdConst, "wdRed", 6) if wdConst else 6)
        inserted_mark = inserted_mark if inserted_mark is not None else 1
        deleted_mark = deleted_mark if deleted_mark is not None else 1

        try:
            self._init_com()
            self._app = win32com.client.Dispatch("Word.Application")
            self._app.Visible = False
            self._app.DisplayAlerts = 0

            opts = self._app.Options
            opts.InsertedTextColor = inserted_color
            opts.DeletedTextColor = deleted_color
            opts.InsertedTextMark = inserted_mark
            opts.DeletedTextMark = deleted_mark

            try:
                self._app.NormalTemplate.Save()
            except Exception:
                pass

            return True, "已將 Word 修訂顯示設定為插入藍色、刪除紅色"
        except Exception as e:
            return False, f"設定 Word 修訂顯示失敗：{e}"
        finally:
            self._quit_word()

    def _init_com(self):
        """初始化 COM 執行緒支援"""
        if not self._com_initialized:
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
            except Exception:
                pass  # 可能已經初始化

    def _uninit_com(self):
        """釋放 COM 資源"""
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialized = False

    def _safe_close_doc(self, doc, save_changes: bool = False):
        """安全關閉 Word 文件"""
        try:
            doc.Close(SaveChanges=save_changes)
        except Exception:
            pass

    def _quit_word(self):
        """安全關閉 Word"""
        try:
            if self._app is not None:
                self._app.Quit(SaveChanges=False)
        except Exception:
            pass
        finally:
            self._app = None
            self._uninit_com()

    def _finalize_document(self, doc):
        """
        文件後處理：確保比較結果正確顯示。
        預設行為：接受所有修訂（不保留）或保留全部。
        目前保留所有 Track Changes，不做任何接受/拒絕。
        """
        # 預設不自動接受任何修訂，保留完整 Track Changes
        # 如需自動接受格式變更，可在這裡擴充
        pass
