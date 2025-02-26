# gui/main_window.py
import re
from PyQt6.QtWidgets import QMainWindow, QVBoxLayout, QPushButton, QFileDialog, QListWidget, QLabel, QTextEdit, QWidget, QListWidgetItem
from conformity_analysis_module.core.analyzer import Analyzer
from conformity_analysis_module.core.worksheet_updater import WorksheetUpdater
from conformity_analysis_module.core.requirements_loader import RequirementsLoader
from conformity_analysis_module.utils.logger import logger

class ConformityAnalysisWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("符合性分析工具")
        self.setGeometry(100, 100, 800, 600)
        self.requirements = RequirementsLoader.load()
        self.folder_path = None
        self.analyzer = Analyzer()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        self.select_folder_button = QPushButton("選擇資料夾", self)
        self.select_folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.select_folder_button)

        self.folder_label = QLabel("選擇的資料夾: 無", self)
        layout.addWidget(self.folder_label)

        layout.addWidget(QLabel("選擇條款要求 (可多選):", self))
        self.requirements_list = QListWidget(self)
        self.requirements_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        for key in self.requirements.keys():
            self.requirements_list.addItem(QListWidgetItem(key))
        layout.addWidget(self.requirements_list)

        self.analyze_button = QPushButton("開始分析", self)
        self.analyze_button.clicked.connect(self.analyze)
        layout.addWidget(self.analyze_button)

        self.fill_worksheet_button = QPushButton("填入 Worksheet", self)
        self.fill_worksheet_button.clicked.connect(self.fill_worksheet)
        layout.addWidget(self.fill_worksheet_button)

        self.result_text = QTextEdit(self)
        self.result_text.setReadOnly(True)
        layout.addWidget(self.result_text)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇資料夾", "")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"選擇的資料夾: {folder}")

    def analyze(self):
        if not self.folder_path or not self.requirements_list.selectedItems():
            self.result_text.setText("請選擇資料夾和至少一個條款要求")
            return
        
        selected_requirements = {re.sub(r"\s+", "", item.text().upper()): self.requirements[item.text()] 
                                for item in self.requirements_list.selectedItems()}
        logger.info(f"選擇的條款要求: {list(selected_requirements.keys())}")
        results = self.analyzer.analyze(self.folder_path, selected_requirements)
        self.result_text.setText(f"分析完成！共找到 {len(results)} 筆符合結果，結果已儲存至 analysis_results.json")

    def fill_worksheet(self):
        success, message = WorksheetUpdater.update_worksheet()
        self.result_text.setText(message)