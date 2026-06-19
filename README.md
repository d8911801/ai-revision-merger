# AI 學術修訂合併器（AI Revision Merger）

將 AI（如千問）產生的 Markdown 修訂稿自動轉為 Word 原生追蹤修訂 DOCX。

支援 `<ins>` / `<del>` 內嵌追蹤標記與 `[^N]` 腳注轉換。

---

## 三種模式

| 模式 | Markdown 內容 | old.docx | 產出 |
|------|-------------|----------|------|
| 追蹤修訂 | 含 `<ins>`/`<del>` |不需要 | 刪除紅字＋插入藍字 DOCX |
| 文檔比較 | 無標記 | 需要 | Word Compare → 追蹤修訂 |
| 純轉換 | 無標記 | 不需要 | 一般 DOCX |

---

## CLI 用法（給 Codex / 自動化工具）

```bash
cd ai_revision_merger

# 模式 1：<ins>/<del> + [^N] → 追蹤修訂 + 腳注
python main.py --cli -m revised.md --output result.docx

# 模式 2：old.docx vs 修訂稿 → Word Compare
python main.py --cli -o old.docx -m revised.md

# 模式 3：純轉換
python main.py --cli -m revised.md --output out.docx
```

### 可選參數

| 參數 | 說明 |
|------|------|
| `--author "千問"` | 修訂作者名稱 |
| `--reference-docx template.docx` | Word 樣式模板 |
| `--output path.docx` | 輸出路徑 |

---

## GUI 用法

雙擊 `啟動_AI合併器.bat`

---

## 需求

- Windows + Microsoft Word 桌面版
- Python 3.11+
- Pandoc（系統安裝）
- `pip install pywin32 python-docx`

---

## Markdown 格式要求

```markdown
<del>被刪除的文字</del><ins>新插入的文字</ins>這是普通內容[^1]。

[^1]: 腳注內容。
```

- `<del>` → Word 追蹤刪除（紅字刪除線）
- `<ins>` → Word 追蹤插入（藍字底線）
- `[^N]` → Word 頁底腳注
- `[^N]: 文字` → 腳注定義（放在文末）
