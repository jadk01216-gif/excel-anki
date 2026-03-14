import sys
import os
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QCheckBox, QProgressBar, QTextEdit, QMessageBox,
                             QLineEdit)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
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
        self.setWindowTitle("劍橋字典Excel表格轉apkg檔案 v0.0.2")
        self.setMinimumSize(500, 550)
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
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # File Selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("尚未選擇檔案")
        select_button = QPushButton("選擇 Excel 檔案")
        select_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(select_button)
        layout.addLayout(file_layout)

        # Deck Name
        deck_name_layout = QHBoxLayout()
        deck_name_layout.addWidget(QLabel("Anki 牌組名稱："))
        self.deck_name_input = QLineEdit()
        self.deck_name_input.setPlaceholderText("輸入在 Anki 中顯示的名稱")
        deck_name_layout.addWidget(self.deck_name_input)
        layout.addLayout(deck_name_layout)

        # Content Toggles
        toggle_group = QVBoxLayout()
        toggle_group.addWidget(QLabel("卡片顯示選項："))
        
        self.show_trans_cb = QCheckBox("顯示中文翻譯 (如缺失會自動翻譯)")
        self.show_trans_cb.setChecked(True)
        toggle_group.addWidget(self.show_trans_cb)
        
        self.show_pos_cb = QCheckBox("顯示詞性 (POS)")
        self.show_pos_cb.setChecked(True)
        toggle_group.addWidget(self.show_pos_cb)
        
        self.show_exp_cb = QCheckBox("顯示英文解釋")
        self.show_exp_cb.setChecked(True)
        toggle_group.addWidget(self.show_exp_cb)
        
        layout.addLayout(toggle_group)

        # Options
        self.tts_checkbox = QCheckBox("包含 Anki 原生 TTS (英音)")
        self.tts_checkbox.setChecked(True)
        layout.addWidget(self.tts_checkbox)

        # Progress
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Log
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

        # Action
        self.convert_button = QPushButton("轉換為 APKG")
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setEnabled(False)
        layout.addWidget(self.convert_button)

        self.excel_path = None

    def select_file(self):
        last_dir = self.settings.get("last_dir", "")
        file_path, _ = QFileDialog.getOpenFileName(self, "開啟 Excel 檔案", last_dir, "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.convert_button.setEnabled(True)
            self.log_area.append(f"已選擇檔案: {file_path}")
            
            # Update last directory
            self.settings["last_dir"] = os.path.dirname(file_path)
            self.save_settings()
            
            # Auto-fill deck name if empty
            if not self.deck_name_input.text():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                self.deck_name_input.setText(base_name)

    def start_conversion(self):
        if not self.excel_path:
            return
            
        deck_name = self.deck_name_input.text().strip()
        if not deck_name:
            QMessageBox.warning(self, "警告", "請輸入 Anki 牌組名稱！")
            return

        # Suggest default filename based on Excel name
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
        default_output = os.path.join(self.settings.get("last_dir", ""), f"{base_name}.apkg")

        output_path, _ = QFileDialog.getSaveFileName(self, "儲存 Anki 封裝檔", default_output, "Anki Package (*.apkg)")
        if not output_path:
            return
            
        # Update last directory to where user saved
        self.settings["last_dir"] = os.path.dirname(output_path)
        self.save_settings()
            
        # Explicit overwrite check
        if os.path.exists(output_path):
            reply = QMessageBox.question(self, "覆蓋確認", f"檔案 '{os.path.basename(output_path)}' 已存在。是否要覆蓋它？",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                return

        self.convert_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_area.append("開始轉換過程...")

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
            self.log_area.append(message)
            QMessageBox.information(self, "完成", "轉換已成功完成！")
        else:
            self.log_area.append(f"錯誤: {message}")
            QMessageBox.critical(self, "錯誤", f"轉換過程中發生錯誤：\n{message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
