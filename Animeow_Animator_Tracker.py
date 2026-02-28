import sys
import os
import threading
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, 
    QTextEdit, QTreeWidget, QTreeWidgetItem, QHeaderView, QFrame,
    QComboBox, QTabWidget
)
from PySide6.QtGui import QIcon, QFont, QColor, QPixmap
from PySide6.QtCore import Qt, QThread, Signal, QMetaObject, Q_ARG

import Test_tool_tracker as TTT

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_folder_suffix(folder_path):
    """
    Finds the name of the folder that contains a 'Raw' subdirectory.
    """
    for root, dirs, files in os.walk(folder_path):
        if 'Raw' in dirs:
            return os.path.basename(root)
    return None

class Worker(QThread):
    progress = Signal(str)
    finished = Signal(bool)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        # We pass a lambda that emits the signal back to the main thread
        def log_callback(msg):
            self.progress.emit(msg)
            
        success = TTT.process_data(self.folder_path, log_callback)
        self.finished.emit(success)

class AnimeowTrackerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Animeow Animator Tracker - V2 (PySide6)")
        self.setMinimumSize(700, 650)
        
        # Thiết lập màu nền tối
        self.setStyleSheet("""
            QMainWindow, QWidget { 
                background-color: #1e1e1e; 
                color: #e0e0e0; 
                font-family: 'Segoe UI', Tahoma, Verdana, sans-serif;
            }
            QPushButton {
                background-color: #2d2d30;
                border: 1px solid #3e3e42;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: bold;
                font-size: 13px;
                color: #e0e0e0;
            }
            QPushButton:hover {
                background-color: #007acc;
                border: 1px solid #007acc;
            }
            QPushButton:disabled {
                background-color: #2d2d30;
                color: #555555;
            }
            QLineEdit {
                background-color: #2d2d30;
                border: 1px solid #3e3e42;
                border-radius: 4px;
                padding: 6px;
                font-size: 13px;
                color: #e0e0e0;
            }
            QTextEdit {
                background-color: #0d0d0d;
                border: 1px solid #3e3e42;
                border-radius: 4px;
                padding: 8px;
                font-family: Consolas, monospace;
                color: #00ff00;
            }
            QTreeWidget {
                background-color: #2d2d30;
                border: 1px solid #3e3e42;
                color: #e0e0e0;
            }
            QHeaderView::section {
                background-color: #1e1e1e;
                color: #007acc;
                font-weight: bold;
                padding: 4px;
                border: 1px solid #3e3e42;
            }
            QScrollBar:vertical {
                border: none;
                background: #1e1e1e;
                width: 14px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #3e3e42;
                min-height: 20px;
                border-radius: 7px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)

        # Main Tabs
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        self.tab_process = QWidget()
        self.tab_animeowee = QWidget()
        self.tab_project = QWidget()
        
        self.tabs.addTab(self.tab_process, "Xử Lý Dữ Liệu")
        self.tabs.addTab(self.tab_animeowee, "Kết Quả Animeowee")
        self.tabs.addTab(self.tab_project, "Kết Quả Dự Án")

        # Tab 1: Processing UI
        main_layout = QVBoxLayout(self.tab_process)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(15)

        # 1. Logo
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        logo_path = resource_path("Animeow_logo.jpg")
            
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            # Scale
            pixmap = pixmap.scaledToWidth(250, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        else:
            logo_label.setText("[Logo Missing]")
        main_layout.addWidget(logo_label)

        # 2. Header
        header_label = QLabel("🚀 ANIMEOW ANIMATOR TRACKER")
        header_font = QFont("Segoe UI", 18, QFont.Bold)
        header_label.setFont(header_font)
        header_label.setStyleSheet("color: #4da6ff;")
        header_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(header_label)

        # 3. Input Section
        input_layout = QHBoxLayout()
        self.path_entry = QLineEdit()
        self.path_entry.setReadOnly(True)
        self.path_entry.setPlaceholderText("Chưa chọn thư mục...")
        input_layout.addWidget(self.path_entry)

        browse_btn = QPushButton("📁 Chọn Thư Mục")
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self.browse_folder)
        input_layout.addWidget(browse_btn)
        main_layout.addLayout(input_layout)

        # 4. Action Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        
        self.process_btn = QPushButton("🔥 BẮT ĐẦU XỬ LÝ")
        self.process_btn.setCursor(Qt.PointingHandCursor)
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.start_processing)
        btn_layout.addWidget(self.process_btn)

        self.view_btn = QPushButton("📊 TRA CỨU KẾT QUẢ")
        self.view_btn.setCursor(Qt.PointingHandCursor)
        self.view_btn.setEnabled(False)
        self.view_btn.clicked.connect(self.view_results)
        btn_layout.addWidget(self.view_btn)
        
        btn_layout.setAlignment(Qt.AlignCenter)
        main_layout.addLayout(btn_layout)

        # 5. Log Console
        log_title = QLabel("📝 Tiến Trình (System Log)")
        log_font = QFont("Segoe UI", 10, QFont.Bold)
        log_title.setFont(log_font)
        main_layout.addWidget(log_title)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        # Footer
        footer_label = QLabel("Developed by Animeow Studio - Qt/PySide6 Edition")
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setStyleSheet("color: #808080; font-size: 11px;")
        main_layout.addWidget(footer_label)

        # Setup Tab 2 & 3: Results UI
        self.setup_tab2_animeowee_ui()
        self.setup_tab3_project_ui()

        # Khởi điểm
        self.write_log("[HỆ THỐNG]: Phiên bản Animeow Animator Tracker V2 đã sẵn sàng!")
        self.write_log("[HỆ THỐNG]: Nhấn 'Chọn Thư Mục' để bắt đầu.")

    def write_log(self, message):
        self.log_text.append(message)
        # Giữ thanh scroll nằm dưới cùng
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def browse_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Chọn Thư Mục Chứa Dữ Liệu Excel")
        if folder_path:
            self.path_entry.setText(folder_path)
            self.process_btn.setEnabled(True)
            self.view_btn.setEnabled(True)
            self.log_text.clear()
            self.write_log(f"📁 Đã chọn đường dẫn:\n => {folder_path}\n")

    def start_processing(self):
        folder_path = self.path_entry.text()
        if not folder_path or not os.path.exists(folder_path):
            QMessageBox.critical(self, "Lỗi", "Đường dẫn không hợp lệ. Vui lòng chọn lại!")
            return

        self.process_btn.setEnabled(False)
        self.view_btn.setEnabled(False)
        self.log_text.clear()
        
        # Chạy logic trong luồng phụ (QThread)
        self.worker = Worker(folder_path)
        self.worker.progress.connect(self.write_log)
        self.worker.finished.connect(self.finish_processing)
        self.worker.start()

    def finish_processing(self, success):
        self.process_btn.setEnabled(True)
        self.view_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "Hoàn Thành", "Quá trình xử lý dữ liệu đã trơn tru thành công!")
            self.refresh_month_selectors()
        else:
            QMessageBox.warning(self, "Cảnh Báo", "Có lỗi xảy ra trong quá trình xử lý.\nVui lòng xem chi tiết trên dòng Log!")

    def find_total_folders(self, root_path):
        """
        Tìm tất cả các thư mục kết quả (Total_*)
        """
        total_folders = []
        for root, dirs, files in os.walk(root_path):
            for d in dirs:
                if d.startswith('Total_'):
                    total_folders.append(os.path.join(root, d))
        return sorted(total_folders, reverse=True)

    def setup_tab2_animeowee_ui(self):
        layout = QVBoxLayout(self.tab_animeowee)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)
        
        control_layout = QHBoxLayout()
        
        month_label = QLabel("📅 Chọn Tháng:")
        month_label.setStyleSheet("color: #e0e0e0; font-weight: bold; font-family: 'Segoe UI'; font-size: 13px;")
        control_layout.addWidget(month_label)
        
        self.combo_month_ani = QComboBox()
        self.combo_month_ani.setStyleSheet("background-color: #2d2d30; color: #4da6ff; padding: 5px; border-radius: 3px; border: 1px solid #3e3e42; font-weight: bold;")
        control_layout.addWidget(self.combo_month_ani)
        
        control_layout.addSpacing(20)
        
        label = QLabel("🔎 Chọn Animeowee:")
        label.setStyleSheet("color: #e0e0e0; font-weight: bold; font-family: 'Segoe UI'; font-size: 13px;")
        control_layout.addWidget(label)
        
        self.combo_animeowee = QComboBox()
        self.combo_animeowee.setStyleSheet("background-color: #2d2d30; color: #e0e0e0; padding: 5px; border-radius: 3px; border: 1px solid #3e3e42;")
        self.combo_animeowee.addItem("Tổng Hợp Animeowee (Total)")
        control_layout.addWidget(self.combo_animeowee)
        
        self.refresh_btn_ani = QPushButton("🔄 Làm Mới")
        self.refresh_btn_ani.setCursor(Qt.PointingHandCursor)
        self.refresh_btn_ani.clicked.connect(self.refresh_animeowee_list)
        control_layout.addWidget(self.refresh_btn_ani)
        
        control_layout.addStretch()
        layout.addLayout(control_layout)
        
        # Thêm nút xem biểu đồ
        self.chart_btn_ani = QPushButton("🖼️ XEM BIỂU ĐỒ TỔNG KẾT")
        self.chart_btn_ani.setStyleSheet("background-color: #4da6ff; color: white;")
        self.chart_btn_ani.setCursor(Qt.PointingHandCursor)
        self.chart_btn_ani.setVisible(False)
        self.chart_btn_ani.clicked.connect(self.show_animeowee_chart)
        layout.addWidget(self.chart_btn_ani)
        
        self.tree_animeowee = QTreeWidget()
        layout.addWidget(self.tree_animeowee)
        
        self.combo_month_ani.currentIndexChanged.connect(self.refresh_animeowee_list)
        self.combo_animeowee.currentIndexChanged.connect(self.update_animeowee_table)

    def setup_tab3_project_ui(self):
        layout = QVBoxLayout(self.tab_project)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)
        
        control_layout = QHBoxLayout()
        
        month_label = QLabel("📅 Chọn Tháng:")
        month_label.setStyleSheet("color: #e0e0e0; font-weight: bold; font-family: 'Segoe UI'; font-size: 13px;")
        control_layout.addWidget(month_label)
        
        self.combo_month_proj = QComboBox()
        self.combo_month_proj.setStyleSheet("background-color: #2d2d30; color: #4da6ff; padding: 5px; border-radius: 3px; border: 1px solid #3e3e42; font-weight: bold;")
        control_layout.addWidget(self.combo_month_proj)
        
        control_layout.addSpacing(20)
        
        label = QLabel("📁 Chọn Dự Án (Project):")
        label.setStyleSheet("color: #e0e0e0; font-weight: bold; font-family: 'Segoe UI'; font-size: 13px;")
        control_layout.addWidget(label)
        
        self.combo_project = QComboBox()
        self.combo_project.setStyleSheet("background-color: #2d2d30; color: #e0e0e0; padding: 5px; border-radius: 3px; border: 1px solid #3e3e42;")
        self.combo_project.addItem("Tổng Hợp Theo Dự Án (Project)")
        control_layout.addWidget(self.combo_project)
        
        self.refresh_btn_proj = QPushButton("🔄 Làm Mới")
        self.refresh_btn_proj.setCursor(Qt.PointingHandCursor)
        self.refresh_btn_proj.clicked.connect(self.refresh_project_list)
        control_layout.addWidget(self.refresh_btn_proj)
        
        control_layout.addStretch()
        layout.addLayout(control_layout)
        
        self.tree_project = QTreeWidget()
        layout.addWidget(self.tree_project)
        
        self.combo_month_proj.currentIndexChanged.connect(self.refresh_project_list)
        self.combo_project.currentIndexChanged.connect(self.update_project_table)

    def find_total_folders(self, root_path):
        """
        Tìm tất cả các thư mục kết quả (Total_*)
        """
        total_folders = []
        for root, dirs, files in os.walk(root_path):
            for d in dirs:
                if d.startswith('Total_'):
                    path = os.path.join(root, d)
                    suffix = d.replace('Total_', '')
                    # Đánh dấu năm (4 chữ số) vs tháng (6 chữ số)
                    display_name = f"📅 NĂM {suffix}" if len(suffix) == 4 else f"🗓️ Tháng {suffix}"
                    total_folders.append((display_name, path))
        
        # Sắp xếp theo tên (Năm lên đầu hoặc theo thứ tự thời gian)
        return sorted(total_folders, key=lambda x: x[0], reverse=True)

    def refresh_month_selectors(self):
        root_path = self.path_entry.text()
        if not root_path: return
        folders_with_names = self.find_total_folders(root_path)
        
        for combo in [self.combo_month_ani, self.combo_month_proj]:
            combo.blockSignals(True)
            old_val = combo.currentData()
            combo.clear()
            for display_name, path in folders_with_names:
                combo.addItem(display_name, path)
            
            idx = combo.findData(old_val)
            if idx >= 0: combo.setCurrentIndex(idx)
            elif combo.count() > 0: combo.setCurrentIndex(0)
            combo.blockSignals(False)

    def refresh_animeowee_list(self):
        total_folder = self.combo_month_ani.currentData()
        if not total_folder or not os.path.exists(total_folder):
            self.combo_animeowee.clear()
            self.combo_animeowee.addItem("Tổng Hợp Animeowee (Total)")
            return
            
        old_selection = self.combo_animeowee.currentText()
        self.combo_animeowee.blockSignals(True)
        self.combo_animeowee.clear()
        self.combo_animeowee.addItem("Tổng Hợp Animeowee (Total)")
        
        anim_folder = os.path.join(total_folder, 'Animeowee_tracker')
        if os.path.exists(anim_folder):
            files = [os.path.splitext(f)[0] for f in os.listdir(anim_folder) if f.endswith('.xlsx')]
            for name in sorted(files):
                self.combo_animeowee.addItem(name)
                
        self.combo_animeowee.blockSignals(False)
        index = self.combo_animeowee.findText(old_selection)
        self.combo_animeowee.setCurrentIndex(index if index >= 0 else 0)
        self.update_animeowee_table()

    def refresh_project_list(self):
        total_folder = self.combo_month_proj.currentData()
        if not total_folder or not os.path.exists(total_folder):
            self.combo_project.clear()
            self.combo_project.addItem("Tổng Hợp Theo Dự Án (Project)")
            return
            
        old_selection = self.combo_project.currentText()
        self.combo_project.blockSignals(True)
        self.combo_project.clear()
        self.combo_project.addItem("Tổng Hợp Theo Dự Án (Project)")
        
        proj_folder = os.path.join(total_folder, 'Project_Animeowee_tracker')
        if os.path.exists(proj_folder):
            files = [os.path.splitext(f)[0] for f in os.listdir(proj_folder) if f.endswith('.xlsx')]
            for name in sorted(files):
                self.combo_project.addItem(name)
                
        self.combo_project.blockSignals(False)
        index = self.combo_project.findText(old_selection)
        self.combo_project.setCurrentIndex(index if index >= 0 else 0)
        self.update_project_table()

    def update_animeowee_table(self):
        import pandas as pd
        self.tree_animeowee.clear()
        total_folder = self.combo_month_ani.currentData()
        if not total_folder: return
        
        total_folder_name = os.path.basename(total_folder)
        suffix = total_folder_name.replace('Total_', '')
        total_file_name = f"Total_{suffix}.xlsx" if suffix != 'Total' else "Total.xlsx"
        
        selection = self.combo_animeowee.currentText()
        
        # Hiện nút biểu đồ nếu chọn "Total"
        if selection == "Tổng Hợp Animeowee (Total)":
            chart_path = os.path.join(total_folder, 'Toan_bo_Animeowee_Chart.png')
            self.chart_btn_ani.setVisible(os.path.exists(chart_path))
            file_to_load = os.path.join(total_folder, total_file_name)
            sheet_name = 'Total_Animeowee'
        else:
            self.chart_btn_ani.setVisible(False)
            file_to_load = os.path.join(total_folder, 'Animeowee_tracker', f"{selection}.xlsx")
            sheet_name = 0
            
        if os.path.exists(file_to_load):
            try:
                df = pd.read_excel(file_to_load, sheet_name=sheet_name)
                self.tree_animeowee.setColumnCount(len(df.columns))
                self.tree_animeowee.setHeaderLabels([str(c) for c in df.columns])
                for idx, row in df.iterrows():
                    self.tree_animeowee.addTopLevelItem(QTreeWidgetItem([str(val) for val in row]))
                for i in range(len(df.columns)):
                    self.tree_animeowee.resizeColumnToContents(i)
            except Exception: pass

    def show_animeowee_chart(self):
        total_folder = self.combo_month_ani.currentData()
        if not total_folder: return
        chart_path = os.path.join(total_folder, 'Toan_bo_Animeowee_Chart.png')
        if os.path.exists(chart_path):
            from PySide6.QtWidgets import QDialog, QLabel, QScrollArea
            dialog = QDialog(self)
            dialog.setWindowTitle("Biểu Đồ Tổng Kết Điểm")
            dialog.setMinimumSize(800, 600)
            dialog_layout = QVBoxLayout(dialog)
            
            scroll = QScrollArea()
            label = QLabel()
            pixmap = QPixmap(chart_path)
            label.setPixmap(pixmap)
            label.setAlignment(Qt.AlignCenter)
            scroll.setWidget(label)
            scroll.setWidgetResizable(True)
            dialog_layout.addWidget(scroll)
            
            close_btn = QPushButton("Đóng")
            close_btn.clicked.connect(dialog.accept)
            dialog_layout.addWidget(close_btn)
            
            dialog.exec()

    def update_project_table(self):
        import pandas as pd
        self.tree_project.clear()
        total_folder = self.combo_month_proj.currentData()
        if not total_folder: return
        
        total_folder_name = os.path.basename(total_folder)
        suffix = total_folder_name.replace('Total_', '')
        total_file_name = f"Total_{suffix}.xlsx" if suffix != 'Total' else "Total.xlsx"
        
        selection = self.combo_project.currentText()
        
        if selection == "Tổng Hợp Theo Dự Án (Project)":
            file_to_load = os.path.join(total_folder, total_file_name)
            sheet_name = 'Total_Project'
        else:
            file_to_load = os.path.join(total_folder, 'Project_Animeowee_tracker', f"{selection}.xlsx")
            sheet_name = 0
            
        if os.path.exists(file_to_load):
            try:
                df = pd.read_excel(file_to_load, sheet_name=sheet_name)
                self.tree_project.setColumnCount(len(df.columns))
                self.tree_project.setHeaderLabels([str(c) for c in df.columns])
                for idx, row in df.iterrows():
                    self.tree_project.addTopLevelItem(QTreeWidgetItem([str(val) for val in row]))
                for i in range(len(df.columns)):
                    self.tree_project.resizeColumnToContents(i)
            except Exception: pass

    def view_results(self):
        if not self.path_entry.text() or not os.path.exists(self.path_entry.text()):
            QMessageBox.warning(self, "Chưa Chọn Thư Mục", "Vui lòng chọn thư mục và xử lý trước!")
            return
            
        self.refresh_month_selectors()
        
        if self.combo_month_ani.count() == 0:
            QMessageBox.information(self, "Thông Báo", "Chưa có dữ liệu thống kê. Vui lòng bấm 'Bắt Đầu Xử Lý' trước!")
            return
            
        # Nhảy qua Tab 2 (Mặc định mở Animeowee)
        self.tabs.setCurrentIndex(1)
        self.refresh_animeowee_list()
        self.refresh_project_list()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Thiết lập icon nếu có
    logo_path = resource_path("Animeow_logo.ico")
    if not os.path.exists(logo_path):
        logo_path = resource_path("Animeow_logo.jpg")
        
    if os.path.exists(logo_path):
        app.setWindowIcon(QIcon(logo_path))
        
    window = AnimeowTrackerApp()
    window.show()
    sys.exit(app.exec())
