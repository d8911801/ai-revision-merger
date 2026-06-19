"""
檔案與路徑輔助函式
"""

import os
import tempfile
from pathlib import Path
from datetime import datetime


def ensure_dir(path: str) -> str:
    """確保目錄存在，不存在則建立"""
    os.makedirs(path, exist_ok=True)
    return path


def get_temp_path(prefix: str = "ai_merger_", suffix: str = ".docx") -> str:
    """在系統暫存目錄中產生暫存檔路徑"""
    temp_dir = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"{prefix}{timestamp}{suffix}"
    return os.path.join(temp_dir, filename)


def suggest_output_path(original_path: str) -> str:
    """
    根據原始檔案路徑，建議輸出檔案路徑。
    例如: paper.docx → paper_revised.docx
    """
    path = Path(original_path)
    stem = path.stem
    parent = path.parent
    new_name = f"{stem}_revised.docx"
    output = parent / new_name

    # 避免覆蓋：若已存在，加上數字
    counter = 1
    while output.exists():
        new_name = f"{stem}_revised_{counter}.docx"
        output = parent / new_name
        counter += 1

    return str(output)


def is_file_locked(path: str) -> bool:
    """
    粗略檢查檔案是否正被其他程式占用而無法寫入。
    目前主要針對 Windows / Word 的常見鎖定情境。
    """
    if not os.path.exists(path):
        return False

    try:
        with open(path, "a+b"):
            return False
    except PermissionError:
        return True
    except OSError as e:
        return getattr(e, "winerror", None) in {32, 33}


def resolve_output_conflict(path: str) -> tuple[str, bool]:
    """
    若目標輸出檔正被占用，則自動改存為遞增新檔名。
    回傳：(可用輸出路徑, 是否已改名)
    """
    candidate = Path(path)

    if not candidate.exists() or not is_file_locked(str(candidate)):
        return str(candidate), False

    counter = 1
    while True:
        alt = candidate.with_name(f"{candidate.stem}_{counter}{candidate.suffix}")
        if not alt.exists() or not is_file_locked(str(alt)):
            return str(alt), True
        counter += 1


def validate_docx(path: str) -> tuple[bool, str]:
    """
    驗證是否為有效的 DOCX 檔案。
    不進行深度驗證，只檢查副檔名與檔案存在。
    """
    if not os.path.exists(path):
        return False, f"檔案不存在：{path}"

    if not path.lower().endswith(".docx"):
        return False, f"檔案不是 .docx 格式：{path}"

    file_size = os.path.getsize(path)
    if file_size == 0:
        return False, f"檔案大小為 0：{path}"

    return True, "OK"


def validate_markdown(text: str) -> tuple[bool, str]:
    """
    基本驗證 Markdown 內容是否合理。
    """
    if not text or not text.strip():
        return False, "Markdown 內容為空"

    stripped = text.strip()

    # 最少要有一定長度（避免只是標題）
    if len(stripped) < 10:
        return False, "Markdown 內容過短（少於 10 字元）"

    return True, "OK"
