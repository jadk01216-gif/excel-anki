import sys
import os
import json

# Force qfluentwidgets to use PyQt6
os.environ['QF_BINDING'] = 'PyQt6'

from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFileDialog
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon

# Create QApplication FIRST - this is critical for Qt to be ready
app = QApplication(sys.argv)

# Now import Fluent Widgets (using the newly installed PyQt6-Fluent-Widgets backend)
from qfluentwidgets import (PushButton, PrimaryPushButton, LineEdit, CheckBox, 
                            ProgressBar, TextEdit, SubtitleLabel, CaptionLabel,
                            FluentIcon, setTheme, Theme, InfoBar, InfoBarPosition, MessageBox)

# Import converter
from converter import AnkiConverter

# File to store persistent settings
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

class ConversionThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, excel_path, output_path, deck_name, include_tts, 
                 show_trans, show_pos, show_exp):
        super().__init__()
        self.excel_path = excel_path
        self.output_path = output_path
        self.deck_name = deck_name
        self.include_tts = include_tts
        self.show_trans = show_trans
        self.show_pos = show_pos
        self.show_exp = show_exp

    def run(self):
        try:
            converter = AnkiConverter(self.excel_path, self.output_path, self.deck_name, 
                                      self.include_tts, self.show_trans, self.show_pos, self.show_exp)
            success = converter.process(self.progress.emit)
            self.finished.emit(success, "轉換成功完成！")
        except Exception as e:
            self.finished.emit(False, str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 設定主題
        setTheme(Theme.LIGHT)
        
        self.setWindowTitle("劍橋字典Excel表格轉apkg檔案 v0.0.3 (PyQt6)")
        self.setMinimumSize(600, 750)
        
        # 設定視窗圖標
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "OIP.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        self.settings = self.load_settings()
        self.init_ui()

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {"last_dir": ""}

    def save_settings(self):
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=4)
        except:
            pass

    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(30, 30, 30, 30)
        self.layout.setSpacing(15)

        # 標題
        self.title_label = SubtitleLabel()
        self.title_label.setText("Excel 轉 Anki 工具 v0.0.3")
        self.layout.addWidget(self.title_label)

        # 檔案選擇
        self.file_group = QVBoxLayout()
        
        self.file_label = CaptionLabel()
        self.file_label.setText("尚未選擇檔案")
        
        self.select_button = PushButton()
        self.select_button.setIcon(FluentIcon.FOLDER)
        self.select_button.setText("選擇 Excel 檔案")
        self.select_button.clicked.connect(self.select_file)
        
        self.file_group.addWidget(self.file_label)
        self.file_group.addWidget(self.select_button)
        self.layout.addLayout(self.file_group)

        # 牌組名稱
        lbl_deck = CaptionLabel()
        lbl_deck.setText("Anki 牌組名稱：")
        self.layout.addWidget(lbl_deck)
        
        self.deck_name_input = LineEdit()
        self.deck_name_input.setPlaceholderText("輸入在 Anki 中顯示的名稱")
        self.layout.addWidget(self.deck_name_input)

        # 顯示選項
        lbl_opts = CaptionLabel()
        lbl_opts.setText("卡片顯示選項：")
        self.layout.addWidget(lbl_opts)
        
        self.show_trans_cb = CheckBox()
        self.show_trans_cb.setText("顯示中文翻譯 (如缺失會自動翻譯)")
        self.show_trans_cb.setChecked(True)
        self.layout.addWidget(self.show_trans_cb)
        
        self.show_pos_cb = CheckBox()
        self.show_pos_cb.setText("顯示詞性 (POS)")
        self.show_pos_cb.setChecked(True)
        self.layout.addWidget(self.show_pos_cb)
        
        self.show_exp_cb = CheckBox()
        self.show_exp_cb.setText("顯示英文解釋")
        self.show_exp_cb.setChecked(True)
        self.layout.addWidget(self.show_exp_cb)
        
        self.tts_checkbox = CheckBox()
        self.tts_checkbox.setText("包含 Anki 原生 TTS")
        self.tts_checkbox.setChecked(True)
        self.layout.addWidget(self.tts_checkbox)

        # 進度條
        self.progress_bar = ProgressBar()
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar)

        # 紀錄區
        self.log_area = TextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setPlaceholderText("轉換記錄將顯示在此處...")
        self.layout.addWidget(self.log_area)

        # 執行按鈕
        self.convert_button = PrimaryPushButton()
        self.convert_button.setText("開始轉換為 APKG")
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setEnabled(False)
        self.layout.addWidget(self.convert_button)

        self.excel_path = None

    def select_file(self):
        last_dir = self.settings.get("last_dir", "")
        file_path, _ = QFileDialog.getOpenFileName(self, "開啟 Excel 檔案", last_dir, "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_path = file_path
            self.file_label.setText(f"已選擇: {os.path.basename(file_path)}")
            self.convert_button.setEnabled(True)
            self.log_area.append(f"已載入檔案: {file_path}")
            
            # 更新上次目錄
            self.settings["last_dir"] = os.path.dirname(file_path)
            self.save_settings()
            
            # 自動填寫牌組名稱
            if not self.deck_name_input.text():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                self.deck_name_input.setText(base_name)

    def start_conversion(self):
        if not self.excel_path:
            return
            
        deck_name = self.deck_name_input.text().strip()
        if not deck_name:
            InfoBar.warning(
                title="提示",
                content="請輸入 Anki 牌組名稱！",
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
            return

        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
        default_output = os.path.join(self.settings.get("last_dir", ""), f"{base_name}.apkg")

        output_path, _ = QFileDialog.getSaveFileName(self, "儲存 Anki 封裝檔", default_output, "Anki Package (*.apkg)")
        if not output_path:
            return
            
        self.settings["last_dir"] = os.path.dirname(output_path)
        self.save_settings()
            
        if os.path.exists(output_path):
            msg_box = MessageBox("覆蓋確認", f"檔案 '{os.path.basename(output_path)}' 已存在。是否要覆蓋它？", self)
            if not msg_box.exec():
                return

        self.convert_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_area.append("狀態: 正在初始化轉換...")

        self.thread = ConversionThread(
            self.excel_path, 
            output_path, 
            deck_name, 
            self.tts_checkbox.isChecked(),
            self.show_trans_cb.isChecked(),
            self.show_pos_cb.isChecked(),
            self.show_exp_cb.isChecked()
        )
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def on_finished(self, success, message):
        self.convert_button.setEnabled(True)
        if success:
            self.log_area.append(f"✅ {message}")
            InfoBar.success(
                title="轉換成功",
                content=message,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
        else:
            self.log_area.append(f"❌ 錯誤: {message}")
            InfoBar.error(
                title="轉換失敗",
                content=message,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=5000,
                parent=self
            )

if __name__ == "__main__":
    # 啟動視窗
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
