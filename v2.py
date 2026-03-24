#!/usr/bin/env python3
"""
build_ankiaddon.py  v0.6.2
==========================
執行此腳本，自動下載依賴套件並打包成 Anki 附加元件：

    python build_ankiaddon.py

產出：excel_to_anki.ankiaddon
使用者安裝後無需額外執行 pip install。
"""

import subprocess
import sys
import os
import zipfile
import shutil
import tempfile

OUTPUT_FILE = "excel_to_anki.ankiaddon"

REQUIREMENTS = [
    "genanki==0.13.0",
    "openpyxl>=3.1",
    "requests>=2.31",
    "deep-translator>=1.11",
    "pyyaml",
    "frozendict",
]

# ===========================================================================
# 附加元件原始碼
# ===========================================================================

MANIFEST = """{
    "package": "excel_to_anki",
    "name": "Excel 轉 Anki (Cambridge Dictionary) v0.6.2",
    "version": "0.6.2",
    "homepage": "",
    "conflicts": [],
    "min_point_version": 231000
}
"""

INIT_PY = """\
# Excel 轉 Anki 附加元件 v0.6.1 — 工具選單入口
import sys, os
# 將 vendor 資料夾加入 Python 路徑（Anki 官方建議做法）
# 參考：https://addon-docs.ankiweb.net/python-modules.html
vendor_path = os.path.join(os.path.dirname(__file__), "vendor")
if vendor_path not in sys.path:
    sys.path.insert(0, vendor_path)

from aqt import mw
from aqt.qt import QAction
from .addon import show_converter_dialog

action = QAction("Excel 轉 Anki (劍橋字典)...", mw)
action.triggered.connect(show_converter_dialog)
mw.form.menuTools.addAction(action)
"""

ADDON_PY = '''\
"""
addon.py - Excel 轉 Anki 附加元件 v0.6.2
新功能：多檔案清單、刪除檔案、轉換並合併成單一牌組
依賴套件已隨附於 vendor/ 資料夾，無需手動安裝。
"""

import os
import sys
import json
import tempfile

_vendor = os.path.join(os.path.dirname(__file__), "vendor")
if _vendor not in sys.path:
    sys.path.insert(0, _vendor)


# ---------------------------------------------------------------------------
# AnkiConverter  — 支援多個 excel_paths，全部合入同一個 deck
# ---------------------------------------------------------------------------

class AnkiConverter:
    MODEL_ID = 1607392319
    DECK_ID  = 2059400110

    def __init__(self, excel_paths, output_path, deck_name,
                 include_tts=False, show_translation=True,
                 show_pos=True, show_explanation=True):
        # excel_paths 可以是單一字串或字串清單
        if isinstance(excel_paths, str):
            excel_paths = [excel_paths]
        self.excel_paths      = excel_paths
        self.output_path      = output_path
        self.deck_name        = deck_name
        self.include_tts      = include_tts
        self.show_translation = show_translation
        self.show_pos         = show_pos
        self.show_explanation = show_explanation

        import genanki
        from deep_translator import GoogleTranslator
        self.genanki    = genanki
        self.translator = GoogleTranslator(source="en", target="zh-TW")

        exp_part = (
            \'<div class="explanation">\'\
            \'{{#Explanation}}{{Explanation}}{{/Explanation}} \'\
            \'{{#POS}}(<i>{{POS}}</i>){{/POS}}\'\
            \'</div>\'
        ) if (show_explanation or show_pos) else ""

        trans_part = (
            \'<div class="translation">\'\
            \'{{#Translation}}{{Translation}}{{/Translation}}\'\
            \'</div>\'
        ) if show_translation else ""

        self.model = genanki.Model(
            self.MODEL_ID,
            "Cambridge Dictionary Model v0.6.2",
            fields=[
                {"name": "Word"},
                {"name": "POS"},
                {"name": "Translation"},
                {"name": "Explanation"},
                {"name": "TTS"},
            ],
            templates=[{
                "name": "Card 1",
                "qfmt": f"""
                    <div class="card-content">
                        {exp_part}
                        {trans_part}
                        <div class="type-box">{{{{type:Word}}}}</div>
                        <div class="tts">{{{{TTS}}}}</div>
                    </div>""",
                "afmt": f"""
                    <div class="card-content">
                        {exp_part}
                        {trans_part}
                        <hr id="answer">
                        <div class="word">{{{{Word}}}}</div>
                        <div class="type-box">{{{{type:Word}}}}</div>
                    </div>""",
            }],
            css="""
                .card {
                    font-family: "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                    background-color: #f1f5f9; margin: 0;
                    display: flex; justify-content: center; align-items: center;
                    min-height: 100vh; padding: 20px;
                }
                .card-content {
                    background: white; padding: 40px; border-radius: 20px;
                    box-shadow: 0 10px 25px -5px rgba(0,0,0,.1),
                                0 8px 10px -6px rgba(0,0,0,.1);
                    min-width: 400px; max-width: 90vw;
                    text-align: center; overflow: visible;
                }
                .explanation  { color: #64748b; font-size: 18px; line-height: 1.5; margin-bottom: 15px; }
                .translation  { color: #0d9488; font-size: 28px; font-weight: 700; margin-bottom: 25px; }
                .word         { color: #1e293b; font-size: 36px; font-weight: 800; margin: 20px 0; }
                .type-box     { margin-top: 20px; overflow: visible; display: inline-block; width: 100%; }
                .tts          { margin-top: 25px; }
                #answer       { border: none; border-top: 2px solid #e2e8f0; margin: 30px 0; }
                #typeans {
                    width: 100% !important; box-sizing: border-box; padding: 12px;
                    font-size: 20px; border: 2px solid #cbd5e1; border-radius: 12px;
                    text-align: center; outline: none; display: inline-block;
                }
                table.typeans  { margin: 0 auto; border-collapse: separate; border-spacing: 0 4px; }
                .typeGood   { background-color: #dcfce7 !important; color: #166534 !important; padding: 4px 8px; border-radius: 4px; }
                .typeBad    { background-color: #fee2e2 !important; color: #991b1b !important; padding: 4px 8px; border-radius: 4px; }
                .typeMissed { background-color: #fef9c3 !important; color: #854d0e !important; padding: 4px 8px; border-radius: 4px; }
                .replay-button svg        { width: 45px; height: 45px; }
                .replay-button svg circle { fill: #0d9488; transition: fill .2s, transform .2s; }
                .replay-button svg path   { fill: white; }
                .replay-button:hover svg circle { fill: #0f766e; transform: scale(1.05); }
            """,
        )

    def _fetch_word_data(self, word):
        try:
            import requests
            r = requests.get(
                f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}",
                timeout=5,
            )
            if r.status_code == 200:
                meanings = r.json()[0].get("meanings", [])
                if meanings:
                    pos = meanings[0].get("partOfSpeech", "")
                    exp = meanings[0].get("definitions", [{}])[0].get("definition", "")
                    return pos, exp
        except Exception:
            pass
        return "", ""

    def _translate(self, word):
        try:
            result = self.translator.translate(word)
            return result or ""
        except Exception:
            return ""

    def _read_excel(self, path):
        """回傳該 Excel 的資料列（跳過前兩列標題）。"""
        import openpyxl
        wb   = openpyxl.load_workbook(path, data_only=True)
        rows = list(wb.active.iter_rows(values_only=True))
        return rows[2:] if len(rows) > 2 else []

    def process(self, progress_callback=None):
        genanki = self.genanki

        # 收集所有檔案的資料列
        all_rows = []
        for path in self.excel_paths:
            all_rows.extend(self._read_excel(path))

        deck  = genanki.Deck(self.DECK_ID, self.deck_name)
        total = len(all_rows)
        if total == 0:
            raise ValueError("所有選取的 Excel 檔案均無資料列。")

        for i, row in enumerate(all_rows):
            if not row or not row[0]:
                continue
            word = str(row[0]).strip()
            if not word:
                continue

            def _cell(idx, r=row):
                v = r[idx] if len(r) > idx else None
                return str(v).strip() if v is not None and str(v).lower() != "nan" else ""

            excel_pos, translation, excel_exp = _cell(1), _cell(3), _cell(4)
            api_pos, api_exp = self._fetch_word_data(word)

            final_trans = translation or self._translate(word)
            final_pos   = excel_pos   or api_pos
            final_exp   = excel_exp   or api_exp
            tts = f"[anki:tts lang=en_US]{word}[/anki:tts]" if self.include_tts else ""

            deck.add_note(genanki.Note(
                model=self.model,
                fields=[
                    word,
                    final_pos   if self.show_pos         else "",
                    final_trans if self.show_translation else "",
                    final_exp   if self.show_explanation else "",
                    tts,
                ],
            ))

            if progress_callback:
                progress_callback(int((i + 1) / total * 100))

        genanki.Package(deck).write_to_file(self.output_path)
        return total


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

def _settings_path():
    from aqt import mw
    return os.path.join(mw.pm.profileFolder(), "excel_to_anki_settings.json")

def _load_settings():
    d = {"last_import_dir": "", "last_export_dir": ""}
    p = _settings_path()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                d.update(json.load(f))
        except Exception:
            pass
    return d

def _save_settings(s):
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Background worker
# ---------------------------------------------------------------------------

from aqt.qt import QThread, pyqtSignal

class _Worker(QThread):
    progress = pyqtSignal(int)
    # finished(ok, message, apkg_path, total_cards)
    finished = pyqtSignal(bool, str, str, int)

    def __init__(self, excel_paths, out_path, deck_name,
                 include_tts, show_trans, show_pos, show_exp):
        super().__init__()
        self.excel_paths = excel_paths
        self.out_path    = out_path
        self.deck_name   = deck_name
        self.include_tts = include_tts
        self.show_trans  = show_trans
        self.show_pos    = show_pos
        self.show_exp    = show_exp

    def run(self):
        try:
            conv = AnkiConverter(
                self.excel_paths, self.out_path, self.deck_name,
                self.include_tts, self.show_trans, self.show_pos, self.show_exp,
            )
            total = conv.process(self.progress.emit)
            self.finished.emit(True, "轉換成功！", self.out_path, total)
        except Exception as e:
            self.finished.emit(False, str(e), "", 0)


# ---------------------------------------------------------------------------
# Dialog
# ---------------------------------------------------------------------------

from aqt import mw
from aqt.qt import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QCheckBox, QRadioButton, QGroupBox,
    QProgressBar, QPlainTextEdit, QFileDialog, QSizePolicy,
    QListWidget, QListWidgetItem, QAbstractItemView,
    QApplication,
)
from aqt.utils import showInfo, showWarning, askUser


class ExcelToAnkiDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Excel 轉 Anki (劍橋字典) v0.6.2")
        self.settings = _load_settings()
        self._thread  = None
        self._build_ui()
        self._apply_adaptive_size()

    # ── 自適應尺寸 ───────────────────────────────────────────────────────────

    def _apply_adaptive_size(self):
        """固定視窗大小，置中於螢幕。"""
        W, H = 1000, 620
        self.setMinimumWidth(860)
        self.setMinimumHeight(540)
        self._file_list.setMinimumHeight(180)
        self._go_btn.setMinimumHeight(36)
        for btn in self._extra_btns:
            btn.setMinimumHeight(36)
        self.resize(W, H)
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            self.move(
                geo.x() + (geo.width()  - W) // 2,
                geo.y() + (geo.height() - H) // 2,
            )

    # ── 取得目前清單中所有路徑 ─────────────────────────────────────────────

    def _file_paths(self):
        return [
            self._file_list.item(i).data(32)          # Qt.UserRole = 32
            for i in range(self._file_list.count())
        ]

    def _refresh_btn_state(self):
        has_files = self._file_list.count() > 0
        self._go_btn.setEnabled(has_files)
        self._del_btn.setEnabled(bool(self._file_list.selectedItems()))

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 主佈局：標題 + 橫式兩欄 + 底部
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        # ── 標題 ──────────────────────────────────────────────────────────
        title = QLabel("Excel 轉 Anki 工具 v0.6.2")
        title.setStyleSheet("font-size: 15px; font-weight: bold;")
        root.addWidget(title)

        # ── 橫式兩欄 ──────────────────────────────────────────────────────
        cols = QHBoxLayout()
        cols.setSpacing(18)

        # 左欄：檔案清單 + 牌組名稱
        left = QVBoxLayout()
        left.setSpacing(8)

        file_box = QGroupBox("Excel 檔案清單（可多選）")
        fb = QVBoxLayout(file_box)
        fb.setSpacing(6)
        fb.setContentsMargins(10, 8, 10, 8)

        self._file_list = QListWidget()
        self._file_list.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection)
        self._file_list.itemSelectionChanged.connect(self._refresh_btn_state)
        fb.addWidget(self._file_list)

        file_btn_row = QHBoxLayout()
        self._add_btn = QPushButton("新增檔案")
        self._add_btn.clicked.connect(self._add_files)
        self._del_btn = QPushButton("移除選取")
        self._del_btn.setEnabled(False)
        self._del_btn.clicked.connect(self._remove_selected)
        self._del_btn.setStyleSheet(
            "QPushButton { color: #dc2626; }"
            "QPushButton:disabled { color: #94a3b8; }"
        )
        file_btn_row.addWidget(self._add_btn)
        file_btn_row.addWidget(self._del_btn)
        file_btn_row.addStretch()
        fb.addLayout(file_btn_row)
        left.addWidget(file_box, 1)   # stretch=1，讓清單佔滿左欄高度

        deck_lbl = QLabel("Anki 牌組名稱：")
        left.addWidget(deck_lbl)
        self._deck_input = QLineEdit()
        self._deck_input.setPlaceholderText("單檔自動填入；多檔合併請手動填寫")
        left.addWidget(self._deck_input)

        cols.addLayout(left, 3)   # 左欄比例 3

        # 右欄：輸出方式 + 卡片選項
        right = QVBoxLayout()
        right.setSpacing(8)

        # 輸出方式
        mode_box = QGroupBox("輸出方式")
        ml = QVBoxLayout(mode_box)
        ml.setSpacing(16)
        ml.setContentsMargins(16, 14, 16, 16)
        self._rb_import = QRadioButton("直接匯入 Anki 牌組（預設）")
        self._rb_save   = QRadioButton("下載 .apkg 檔案")
        self._rb_import.setChecked(True)
        ml.addWidget(self._rb_import)
        ml.addWidget(self._rb_save)
        right.addWidget(mode_box)

        # 卡片選項
        opt_box = QGroupBox("卡片顯示選項")
        ol = QVBoxLayout(opt_box)
        ol.setSpacing(16)
        ol.setContentsMargins(16, 14, 16, 16)
        self._cb_trans = QCheckBox("顯示中文翻譯（缺失時自動翻譯）")
        self._cb_pos   = QCheckBox("顯示詞性 (POS)")
        self._cb_exp   = QCheckBox("顯示英文解釋")
        self._cb_tts   = QCheckBox("包含 Anki 原生 TTS 發音")
        for cb in (self._cb_trans, self._cb_pos, self._cb_exp, self._cb_tts):
            cb.setChecked(True)
            ol.addWidget(cb)
        right.addWidget(opt_box)
        right.addStretch()   # 右欄底部留白，防止被拉伸

        cols.addLayout(right, 3)   # 右欄比例 3

        root.addLayout(cols, 1)   # 兩欄區域可伸縮

        # ── 進度條 ────────────────────────────────────────────────────────
        self._progress = QProgressBar()
        self._progress.setValue(0)
        root.addWidget(self._progress)

        # ── 記錄區 ────────────────────────────────────────────────────────
        self._log = QPlainTextEdit()
        self._log.setReadOnly(True)
        self._log.setPlaceholderText("轉換記錄將顯示在此處...")
        self._log.setMinimumHeight(90)
        self._log.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        root.addWidget(self._log)

        # ── 按鈕列 ────────────────────────────────────────────────────────
        br = QHBoxLayout()
        self._go_btn = QPushButton("開始轉換")
        self._go_btn.setEnabled(False)
        self._go_btn.setStyleSheet(
            "QPushButton { background-color: #0d9488; color: white;"
            " padding: 8px 18px; border-radius: 6px; font-weight: bold; }"
            "QPushButton:disabled { background-color: #94a3b8; }"
            "QPushButton:hover:!disabled { background-color: #0f766e; }"
        )
        self._go_btn.clicked.connect(self._start)
        close_btn = QPushButton("關閉")
        close_btn.clicked.connect(self.reject)
        self._extra_btns = [close_btn, self._add_btn, self._del_btn]
        br.addWidget(self._go_btn)
        br.addWidget(close_btn)
        root.addLayout(br)

    # ── 檔案操作 ──────────────────────────────────────────────────────────

    def _add_files(self):
        start = self.settings.get("last_import_dir") or os.path.expanduser("~")
        paths, _ = QFileDialog.getOpenFileNames(
            self, "選擇 Excel 檔案", start,
            "Excel Files (*.xlsx *.xls)")
        if not paths:
            return

        existing = set(self._file_paths())
        added = 0
        for path in paths:
            if path in existing:
                continue
            item = QListWidgetItem(os.path.basename(path))
            item.setData(32, path)           # Qt.UserRole
            item.setToolTip(path)
            self._file_list.addItem(item)
            added += 1

        if added:
            self.settings["last_import_dir"] = os.path.dirname(paths[-1])
            _save_settings(self.settings)
            self._log.appendPlainText(f"新增 {added} 個檔案。")

            # 單一檔案時自動填牌組名
            if self._file_list.count() == 1 and not self._deck_input.text():
                self._deck_input.setText(
                    os.path.splitext(os.path.basename(paths[0]))[0])

        self._refresh_btn_state()

    def _remove_selected(self):
        selected = self._file_list.selectedItems()
        if not selected:
            return
        names = [it.text() for it in selected]
        if not askUser(
            f"確定要從清單中移除以下 {len(selected)} 個檔案？\\n"
            + "\\n".join(f"  - {n}" for n in names),
            parent=self
        ):
            return
        for item in selected:
            self._file_list.takeItem(self._file_list.row(item))
        self._log.appendPlainText(f"已移除 {len(selected)} 個檔案。")
        self._refresh_btn_state()

    # ── 轉換 ──────────────────────────────────────────────────────────────

    def _start(self):
        paths = self._file_paths()
        if not paths:
            return

        deck_name = self._deck_input.text().strip()
        if not deck_name:
            showWarning("請輸入 Anki 牌組名稱！", parent=self)
            return

        multi = len(paths) > 1
        label = f"{len(paths)} 個檔案合併" if multi else os.path.basename(paths[0])

        if self._rb_save.isChecked():
            export_dir = (self.settings.get("last_export_dir")
                          or self.settings.get("last_import_dir")
                          or os.path.expanduser("~"))
            default_name = deck_name if multi else \
                os.path.splitext(os.path.basename(paths[0]))[0]
            out_path, _ = QFileDialog.getSaveFileName(
                self, "儲存 Anki 封裝檔",
                os.path.join(export_dir, f"{default_name}.apkg"),
                "Anki Package (*.apkg)")
            if not out_path:
                return
            if os.path.exists(out_path):
                if not askUser(
                    f"檔案 \'{os.path.basename(out_path)}\' 已存在，是否覆蓋？",
                    parent=self):
                    return
            self.settings["last_export_dir"] = os.path.dirname(out_path)
            _save_settings(self.settings)
            self._pending_mode = "save"
        else:
            tmp = tempfile.NamedTemporaryFile(
                suffix=".apkg", delete=False, prefix="excel_anki_")
            tmp.close()
            out_path = tmp.name
            self._pending_mode = "import"

        self._pending_apkg = out_path
        self._go_btn.setEnabled(False)
        self._progress.setValue(0)
        self._log.appendPlainText(f"正在轉換：{label}...")

        self._thread = _Worker(
            paths, out_path, deck_name,
            self._cb_tts.isChecked(), self._cb_trans.isChecked(),
            self._cb_pos.isChecked(), self._cb_exp.isChecked(),
        )
        self._thread.progress.connect(self._progress.setValue)
        self._thread.finished.connect(self._on_finished)
        self._thread.start()

    def _on_finished(self, success, message, apkg_path, total_cards):
        self._refresh_btn_state()
        if not success:
            self._log.appendPlainText(f"錯誤：{message}")
            showWarning(f"轉換失敗：{message}", parent=self)
            if apkg_path and os.path.exists(apkg_path):
                try:
                    os.remove(apkg_path)
                except Exception:
                    pass
            return

        detail = f"（共 {total_cards} 張卡片）" if total_cards else ""

        if self._pending_mode == "import":
            self._log.appendPlainText("正在匯入牌組至 Anki...")
            try:
                from anki.importing import AnkiPackageImporter
                imp = AnkiPackageImporter(mw.col, apkg_path)
                imp.run()
                mw.col.reset()
                mw.reset()
                self._log.appendPlainText(f"匯入完成 {detail}")
                showInfo(f"匯入成功！{detail}", parent=self)
            except Exception as e:
                self._log.appendPlainText(f"匯入失敗：{e}")
                showWarning(
                    f"APKG 已產生但匯入失敗：{e}\\n\\n"
                    f"請手動匯入暫存檔：\\n{apkg_path}",
                    parent=self)
                return
            finally:
                try:
                    os.remove(apkg_path)
                except Exception:
                    pass
        else:
            self._log.appendPlainText(f"已儲存至：{apkg_path} {detail}")
            showInfo(f"檔案已儲存 {detail}\\n{apkg_path}", parent=self)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def show_converter_dialog():
    dlg = ExcelToAnkiDialog(mw)
    dlg.exec()
'''

# ===========================================================================
# 打包邏輯
# ===========================================================================

def pip_install_to_vendor(vendor_dir: str):
    print("\n[1/3] 下載依賴套件到 vendor/ ...")
    cmd = [
        sys.executable, "-m", "pip", "install",
        "--target", vendor_dir,
        "--no-user",
        "--quiet",
        "--no-compile",
        "--ignore-requires-python",
    ] + REQUIREMENTS

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("  pip 錯誤：")
        print(result.stderr[-2000:])
        sys.exit(1)

    for root, dirs, files in os.walk(vendor_dir, topdown=False):
        for d in dirs:
            full = os.path.join(root, d)
            if d.endswith((".dist-info", ".data", "__pycache__")):
                shutil.rmtree(full, ignore_errors=True)
        for f in files:
            if f.endswith(".pyi"):
                os.remove(os.path.join(root, f))

    n_dirs = sum(1 for _ in os.walk(vendor_dir))
    print(f"    完成（共 {n_dirs} 個目錄）")


def build():
    print("=" * 54)
    print("  Excel 轉 Anki — 附加元件打包工具 v0.6.2")
    print("=" * 54)

    workdir = tempfile.mkdtemp(prefix="excel_anki_build_")
    try:
        vendor_dir = os.path.join(workdir, "vendor")
        os.makedirs(vendor_dir)

        pip_install_to_vendor(vendor_dir)

        print("\n[2/3] 打包成 .ankiaddon ...")

        addon_files = {
            "manifest.json": MANIFEST.encode("utf-8"),
            "__init__.py":   INIT_PY.encode("utf-8"),
            "addon.py":      ADDON_PY.encode("utf-8"),
        }

        with zipfile.ZipFile(OUTPUT_FILE, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in addon_files.items():
                zf.writestr(name, data)
                print(f"    + {name:20s}  ({len(data):,} bytes)")

            vendor_count = 0
            for dirpath, _, filenames in os.walk(vendor_dir):
                for filename in filenames:
                    full = os.path.join(dirpath, filename)
                    arcname = os.path.relpath(full, workdir)
                    zf.write(full, arcname)
                    vendor_count += 1

            print(f"    + vendor/             ({vendor_count} 個檔案)")

        size_kb = os.path.getsize(OUTPUT_FILE) / 1024
        print(f"\n[3/3] 完成！")
        print(f"\n  輸出：{OUTPUT_FILE}  ({size_kb:.1f} KB)")

    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    print("""
安裝方式：
  Anki -> 工具 -> 附加元件 -> 從檔案安裝附加元件
  -> 選擇 excel_to_anki.ankiaddon -> 重啟 Anki

使用者無需安裝任何 pip 套件，依賴已內嵌。
""")


if __name__ == "__main__":
    build()
