"""
AI Revision Merger - Desktop GUI
簡潔的 tkinter 桌面介面：選擇檔案 → 貼入 Markdown → 一鍵生成
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, font as tkfont

# 專案根目錄加入路徑
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.markdown_converter import MarkdownConverter
from core.word_merger import WordMerger
from core.insdel_to_docx import InsDelToDocx
from utils.file_utils import (
    suggest_output_path,
    validate_docx,
    validate_markdown,
    get_temp_path,
)


class RevisionMergerApp:
    """AI 學術修訂合併器主視窗"""

    TITLE = "AI 學術修訂合併器"
    WINDOW_WIDTH = 1400
    WINDOW_HEIGHT = 1100

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(self.TITLE)
        self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
        self.root.minsize(1000, 700)

        # 置中視窗
        self._center_window()

        # 狀態變數
        self.original_path = tk.StringVar()
        self.revised_path = tk.StringVar()      # 修訂後 Word 檔案（直接 Compare 用）
        self.reference_path = tk.StringVar()
        self.markdown_text = ""
        self.is_processing = False

        # 核心模組
        self.converter = MarkdownConverter()
        self.merger = WordMerger()
        self.insdel = InsDelToDocx()

        # 建立介面
        self._build_ui()

        # 關閉視窗時清理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI 建構 ──────────────────────────────────────────────

    def _build_ui(self):
        """建立所有 GUI 元件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="12")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── 區塊 1：選擇原始 Word ──
        self._build_file_section(main_frame, 0)

        # ── 區塊 2：選擇修訂後 Word（直接 Compare 用）──
        self._build_revised_section(main_frame, 1)

        # ── 區塊 3：選擇參考模板（可選）──
        self._build_reference_section(main_frame, 2)

        # ── 區塊 4：Markdown 輸入區 ──
        self._build_markdown_section(main_frame, 3)

        # ── 區塊 5：執行按鈕 ──
        self._build_action_section(main_frame, 4)

        # ── 區塊 6：狀態／日誌輸出 ──
        self._build_status_section(main_frame, 5)

    def _build_file_section(self, parent, row):
        """選擇原始 Word 檔案區塊"""
        frame = ttk.LabelFrame(parent, text="原始 Word 檔案（選用）", padding="8")
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        # 路徑顯示
        path_entry = ttk.Entry(
            frame,
            textvariable=self.original_path,
            font=("Microsoft JhengHei", 28),
        )
        path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        # 選擇按鈕
        btn_select = ttk.Button(
            frame,
            text="選擇檔案...",
            command=self._select_original,
            width=16,
        )
        btn_select.grid(row=0, column=1, sticky="e")

        # 提示文字
        hint = ttk.Label(
            frame,
            text="若有 old.docx 則產出追蹤修訂；不選則只做 Markdown → DOCX 轉換",
            font=("Microsoft JhengHei", 26),
            foreground="gray",
        )
        hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

    def _build_revised_section(self, parent, row):
        """選擇修訂後 Word 檔案區塊（直接 Compare 用）"""
        frame = ttk.LabelFrame(parent, text="修訂後 Word 檔案（選用）", padding="8")
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        path_entry = ttk.Entry(
            frame,
            textvariable=self.revised_path,
            font=("Microsoft JhengHei", 28),
        )
        path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        btn_select = ttk.Button(
            frame,
            text="選擇檔案...",
            command=self._select_revised,
            width=16,
        )
        btn_select.grid(row=0, column=1, sticky="e")

        hint = ttk.Label(
            frame,
            text="若有修訂後的 .docx，可與原始檔案直接 Compare（不需 Markdown）",
            font=("Microsoft JhengHei", 26),
            foreground="gray",
        )
        hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

    def _build_reference_section(self, parent, row):
        """選擇參考模板區塊（可選）"""
        frame = ttk.LabelFrame(parent, text="Word 模板（選用）", padding="8")
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        path_entry = ttk.Entry(
            frame,
            textvariable=self.reference_path,
            font=("Microsoft JhengHei", 28),
        )
        path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        btn_select = ttk.Button(
            frame,
            text="選擇模板...",
            command=self._select_reference,
            width=16,
        )
        btn_select.grid(row=0, column=1, sticky="e")

        # 提示文字
        hint = ttk.Label(
            frame,
            text="選用：reference.docx 樣式模板，用於自訂輸出格式",
            font=("Microsoft JhengHei", 26),
            foreground="gray",
        )
        hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

    def _build_markdown_section(self, parent, row):
        """Markdown 輸入區塊"""
        frame = ttk.LabelFrame(parent, text="AI 修訂稿（Markdown）", padding="8")
        frame.grid(row=row, column=0, sticky="nsew", pady=(0, 8))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # 捲動文字區域
        self.md_text = scrolledtext.ScrolledText(
            frame,
            wrap=tk.WORD,
            font=("Consolas", 26),
            undo=True,
            relief=tk.SUNKEN,
            borderwidth=1,
        )
        self.md_text.grid(row=0, column=0, sticky="nsew")
        self.md_text.bind("<<Modified>>", self._on_markdown_modified)

        # 右下角字數顯示
        self.char_count_var = tk.StringVar(value="字數：0")
        count_label = ttk.Label(
            frame,
            textvariable=self.char_count_var,
            font=("Microsoft JhengHei", 26),
            foreground="gray",
        )
        count_label.grid(row=1, column=0, sticky="e", pady=(2, 0))

        # 讓這個區塊可隨視窗伸縮
        parent.rowconfigure(row, weight=1)

    def _build_action_section(self, parent, row):
        """執行按鈕區塊"""
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=(4, 8))

        # 生成按鈕（主要操作）
        self.btn_generate = ttk.Button(
            frame,
            text="開始處理",
            command=self._start_merge,
            width=24,
        )
        self.btn_generate.pack(side=tk.LEFT, padx=(0, 12))

        # 進度條
        self.progress = ttk.Progressbar(
            frame,
            mode="indeterminate",
            length=200,
        )
        self.progress.pack(side=tk.LEFT, padx=(0, 12))

        # 輸出目錄快捷按鈕
        self.btn_open_dir = ttk.Button(
            frame,
            text="開啟輸出目錄",
            command=self._open_output_dir,
            width=18,
            state=tk.DISABLED,
        )
        self.btn_open_dir.pack(side=tk.RIGHT)

        # 記錄上次輸出目錄
        self._last_output_dir = None

    def _build_status_section(self, parent, row):
        """狀態／日誌輸出區塊"""
        frame = ttk.LabelFrame(parent, text="處理狀態", padding="6")
        frame.grid(row=row, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)

        self.status_text = scrolledtext.ScrolledText(
            frame,
            height=4,
            wrap=tk.WORD,
            font=("Consolas", 26),
            state=tk.DISABLED,
            relief=tk.SUNKEN,
            borderwidth=1,
        )
        self.status_text.grid(row=0, column=0, sticky="ew")

    # ── 事件處理 ──────────────────────────────────────────────

    def _center_window(self):
        """將視窗置中"""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _select_original(self):
        """選擇原始 Word 檔案"""
        path = filedialog.askopenfilename(
            title="選擇原始 Word 檔案",
            filetypes=[
                ("Word 文件", "*.docx"),
                ("所有檔案", "*.*"),
            ],
        )
        if path:
            self.original_path.set(path)
            self._log(f"已選擇原始檔案：{os.path.basename(path)}")

    def _select_revised(self):
        """選擇修訂後 Word 檔案"""
        path = filedialog.askopenfilename(
            title="選擇修訂後 Word 檔案",
            filetypes=[
                ("Word 文件", "*.docx"),
                ("所有檔案", "*.*"),
            ],
        )
        if path:
            self.revised_path.set(path)
            self._log(f"已選擇修訂後檔案：{os.path.basename(path)}")

    def _select_reference(self):
        """選擇 Word 模板"""
        path = filedialog.askopenfilename(
            title="選擇 Word 模板（reference.docx）",
            filetypes=[
                ("Word 文件", "*.docx"),
                ("所有檔案", "*.*"),
            ],
        )
        if path:
            self.reference_path.set(path)
            self._log(f"已選擇模板：{os.path.basename(path)}")

    def _on_markdown_modified(self, event=None):
        """Markdown 內容變更時更新字數"""
        content = self.md_text.get("1.0", tk.END)
        self.markdown_text = content
        char_count = len(content.replace("\n", "").replace(" ", ""))
        self.char_count_var.set(f"字數：{char_count}")
        self.md_text.edit_modified(False)

    def _open_output_dir(self):
        """開啟上次輸出目錄"""
        if self._last_output_dir and os.path.exists(self._last_output_dir):
            os.startfile(self._last_output_dir)

    # ── 核心處理流程 ──────────────────────────────────────────

    def _start_merge(self):
        """開始合併流程（在背景執行）"""
        if self.is_processing:
            messagebox.showwarning("處理中", "正在處理中，請稍候...")
            return

        # 驗證輸入
        errors = self._validate_inputs()
        if errors:
            messagebox.showerror("輸入錯誤", "\n".join(errors))
            return

        # 禁用按鈕，啟用進度條
        self._set_ui_state(processing=True)
        self._clear_log()

        # 在背景執行緒處理
        thread = threading.Thread(target=self._run_merge, daemon=True)
        thread.start()

    def _validate_inputs(self) -> list[str]:
        """驗證所有輸入，回傳錯誤訊息列表"""
        errors = []

        # 檢查原始檔案（可選）
        orig_path = self.original_path.get().strip()
        if orig_path:
            valid, msg = validate_docx(orig_path)
            if not valid:
                errors.append(f"原始檔案：{msg}")

        # 檢查修訂後 Word（可選）
        revised_docx = self.revised_path.get().strip()
        if revised_docx:
            valid, msg = validate_docx(revised_docx)
            if not valid:
                errors.append(f"修訂後檔案：{msg}")

        # 若已有兩個 Word 檔案 → 直接 Compare，不需 Markdown
        if orig_path and revised_docx:
            return errors

        # 檢查 Markdown 內容
        md_content = self.md_text.get("1.0", tk.END)
        valid, msg = validate_markdown(md_content)
        if not valid:
            errors.append(f"Markdown：{msg}")

        return errors

    def _run_merge(self):
        """背景執行合併流程"""
        try:
            orig_path = self.original_path.get().strip()
            revised_docx = self.revised_path.get().strip()

            # ── 管線 0：兩個 Word 檔案 → 直接 Compare（最高優先）──
            if orig_path and revised_docx:
                self._run_direct_compare(orig_path, revised_docx)
                return

            ref_path = self.reference_path.get() or None
            md_content = self.md_text.get("1.0", tk.END).strip()

            # 偵測 Markdown 是否含 <ins> / <del> 追蹤修訂標記
            has_insdel = "<ins>" in md_content or "<del>" in md_content

            if has_insdel:
                self._run_insdel_pipeline(md_content)
            else:
                has_original = bool(orig_path)
                if has_original:
                    self._run_full_merge(orig_path, ref_path, md_content)
                else:
                    self._run_convert_only(ref_path, md_content)

        except Exception as e:
            self._log(f"✗ 未預期錯誤：{e}")
            self.root.after(0, self._on_merge_failed, str(e))

    def _run_direct_compare(self, orig_path, revised_docx):
        """模式 0：兩個 Word 檔案直接 Compare（不需 Markdown）"""
        self._log("【直接比較模式】兩個 Word 檔案直接 Compare...")
        output_path = suggest_output_path(orig_path)

        success, msg = self.merger.compare_documents(
            original_path=orig_path,
            revised_path=revised_docx,
            output_path=output_path,
        )

        if not success:
            self._log(f"✗ 比較失敗：{msg}")
            self.root.after(0, self._on_merge_failed, msg)
            return

        self._log(f"✓ 追蹤修訂文件已產出：{output_path}")
        self._last_output_dir = os.path.dirname(output_path)
        self.root.after(0, self._on_merge_success, output_path, True)

    def _run_full_merge(self, orig_path, ref_path, md_content):
        """模式 1：有 old.docx → Markdown → DOCX → Word Compare → 追蹤修訂"""
        # ── 步驟 1：Markdown → 新 DOCX ──
        self._log("【步驟 1/2】正在將 Markdown 轉換為 DOCX...")
        temp_docx = get_temp_path(prefix="ai_revision_")

        success, msg, final_temp_docx = self.converter.convert_to_docx(
            markdown_text=md_content,
            output_path=temp_docx,
            reference_docx=ref_path,
        )

        if not success:
            self._log(f"✗ 轉換失敗：{msg}")
            self.root.after(0, self._on_merge_failed, msg)
            return

        temp_docx = final_temp_docx or temp_docx

        self._log("✓ Markdown → DOCX 完成")

        # ── 步驟 2：Word Compare ──
        self._log("【步驟 2/2】正在使用 Word 原生文件比較...")
        output_path = suggest_output_path(orig_path)

        success, msg, final_output_path = self.merger.compare_documents(
            original_path=orig_path,
            revised_path=temp_docx,
            output_path=output_path,
        )

        # 清理暫存檔
        try:
            os.unlink(temp_docx)
        except OSError:
            pass

        if not success:
            self._log(f"✗ 合併失敗：{msg}")
            self.root.after(0, self._on_merge_failed, msg)
            return

        output_path = final_output_path or output_path
        display_ok, display_msg = self.merger.ensure_revision_display_colors()
        if display_ok:
            self._log(f"✓ {display_msg}")
        else:
            self._log(f"⚠ {display_msg}")
        self._log(f"✓ 追蹤修訂文件已產出：{output_path}")
        self._last_output_dir = os.path.dirname(output_path)
        self.root.after(0, self._on_merge_success, output_path, True)

    def _run_convert_only(self, ref_path, md_content):
        """模式 2：只有 Markdown → 純 DOCX 轉換（無追蹤修訂）"""
        self._log("【純轉換模式】正在將 Markdown 轉換為 DOCX...")

        # 讓使用者選擇儲存位置
        output_path = filedialog.asksaveasfilename(
            title="另存 DOCX 為...",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")],
            initialfile="output.docx",
        )

        if not output_path:
            self._log("✗ 使用者取消儲存")
            self.root.after(0, self._set_ui_state, False)
            return

        success, msg, final_output_path = self.converter.convert_to_docx(
            markdown_text=md_content,
            output_path=output_path,
            reference_docx=ref_path,
        )

        if not success:
            self._log(f"✗ 轉換失敗：{msg}")
            self.root.after(0, self._on_merge_failed, msg)
            return

        output_path = final_output_path or output_path
        self._log(f"✓ DOCX 已產出：{output_path}")
        self._last_output_dir = os.path.dirname(output_path)
        self.root.after(0, self._on_merge_success, output_path, False)

    def _run_insdel_pipeline(self, md_content):
        """模式 3：Markdown 含 <ins>/<del> → 直接產出追蹤修訂 DOCX"""
        self._log("【追蹤修訂模式】偵測到 <ins>/<del> 標記，直接轉為 Word 追蹤修訂...")

        # 讓使用者選擇儲存位置
        output_path = filedialog.asksaveasfilename(
            title="另存追蹤修訂 DOCX 為...",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")],
            initialfile="tracked_revision.docx",
        )

        if not output_path:
            self._log("✗ 使用者取消儲存")
            self.root.after(0, self._set_ui_state, False)
            return

        success, msg, final_output_path = self.insdel.convert(
            markdown_text=md_content,
            output_path=output_path,
        )

        if not success:
            self._log(f"✗ 轉換失敗：{msg}")
            self.root.after(0, self._on_merge_failed, msg)
            return

        output_path = final_output_path or output_path
        display_ok, display_msg = self.merger.ensure_revision_display_colors()
        if display_ok:
            self._log(f"✓ {display_msg}")
        else:
            self._log(f"⚠ {display_msg}")
        self._log(f"✓ 追蹤修訂 DOCX 已產出：{output_path}")
        self._last_output_dir = os.path.dirname(output_path)
        self.root.after(0, self._on_merge_success, output_path, True)

    def _on_merge_success(self, output_path: str, has_original: bool = True):
        """合併成功後的 UI 更新"""
        self._set_ui_state(processing=False)
        self.btn_open_dir.config(state=tk.NORMAL)
        self._log("─" * 50)

        if has_original:
            finish_msg = "✅ 處理完成！請用 Microsoft Word 開啟輸出檔案檢視追蹤修訂。"
            ask_msg = f"追蹤修訂文件已產出：\n\n{output_path}\n\n是否立即用 Word 開啟？"
        else:
            finish_msg = "✅ 轉換完成！Markdown 已轉為 Word 文件。"
            ask_msg = f"Word 文件已產出：\n\n{output_path}\n\n是否立即用 Word 開啟？"

        self._log(finish_msg)

        # 詢問是否開啟
        if messagebox.askyesno("處理完成", ask_msg):
            try:
                os.startfile(output_path)
            except Exception as e:
                messagebox.showwarning("無法開啟", f"無法開啟檔案：{e}")

    def _on_merge_failed(self, error_msg: str):
        """合併失敗後的 UI 更新"""
        self._set_ui_state(processing=False)
        messagebox.showerror("處理失敗", f"合併過程中發生錯誤：\n\n{error_msg}")

    # ── UI 輔助方法 ───────────────────────────────────────────

    def _set_ui_state(self, processing: bool):
        """設定 UI 處理狀態"""
        self.is_processing = processing
        if processing:
            self.btn_generate.config(state=tk.DISABLED, text="處理中...")
            self.progress.start(10)
        else:
            self.btn_generate.config(state=tk.NORMAL, text="開始處理")
            self.progress.stop()

    def _log(self, message: str, error: bool = False):
        """寫入狀態日誌（執行緒安全）"""
        def _write():
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, message + "\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)
        self.root.after(0, _write)

    def _clear_log(self):
        """清除狀態日誌"""
        def _clear():
            self.status_text.config(state=tk.NORMAL)
            self.status_text.delete("1.0", tk.END)
            self.status_text.config(state=tk.DISABLED)
        self.root.after(0, _clear)

    def _on_close(self):
        """關閉視窗"""
        if self.is_processing:
            if not messagebox.askyesno("確認", "正在處理中，確定要關閉嗎？"):
                return
        self.root.destroy()


def main():
    """GUI 入口"""
    root = tk.Tk()

    # ── 字型設定（大）──
    BASE_FONT = ("Microsoft JhengHei", 28)
    MONO_FONT = ("Consolas", 26)
    SMALL_FONT = ("Microsoft JhengHei", 24)

    # 方法一：全域 option 覆蓋所有 widget
    root.option_add("*Font", BASE_FONT)
    root.option_add("*Entry*Font", BASE_FONT)
    root.option_add("*Button*Font", BASE_FONT)
    root.option_add("*Label*Font", BASE_FONT)
    root.option_add("*Text*Font", MONO_FONT)

    # 方法二：設定 tk 標準字型名稱
    for name in ("TkDefaultFont", "TkTextFont", "TkFixedFont", "TkMenuFont",
                 "TkHeadingFont", "TkCaptionFont", "TkTooltipFont"):
        try:
            tkfont.nametofont(name).configure(size=26)
        except Exception:
            pass

    # 方法三：ttk 樣式（clam 主題可能無視，但照設）
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", font=BASE_FONT)
    style.configure("TLabelframe.Label", font=BASE_FONT)
    style.configure("TButton", font=BASE_FONT)
    style.configure("TEntry", font=BASE_FONT)
    style.configure("TLabel", font=BASE_FONT)

    # 嘗試設定 Windows DPI 感知（避免模糊）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = RevisionMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
