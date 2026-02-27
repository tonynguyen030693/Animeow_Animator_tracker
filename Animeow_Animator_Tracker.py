import sys
import os
import threading
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, 
    QTextEdit, QTreeWidget, QTreeWidgetItem, QHeaderView, QFrame,
    QComboBox
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

        # Main Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
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
        footer_label = QLabel("Developed by Tool Team - Qt/PySide6 Edition")
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setStyleSheet("color: #808080; font-size: 11px;")
        main_layout.addWidget(footer_label)

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
        else:
            QMessageBox.warning(self, "Cảnh Báo", "Có lỗi xảy ra trong quá trình xử lý.\nVui lòng xem chi tiết trên dòng Log!")

    def view_results(self):
        import pandas as pd
        folder_path = self.path_entry.text()
        if not folder_path:
            return
            
        total_folder = os.path.join(folder_path, 'Total')
        total_file = os.path.join(total_folder, 'Total.xlsx')
        if not os.path.exists(total_file):
            QMessageBox.information(self, "Thông Báo", "Chưa có dữ liệu thống kê. Vui lòng chạy Xử Lý trước!")
            return
            
        try:
            # Tạo popup window Qt
            self.result_window = QMainWindow(self)
            self.result_window.setWindowTitle("Tra Cứu Kết Quả (Animeow Tracker)")
            self.result_window.setMinimumSize(800, 500)
            
            # Central Widget for popup
            central = QWidget()
            central.setStyleSheet("background-color: #1e1e1e;")
            self.result_window.setCentralWidget(central)
            layout = QVBoxLayout(central)
            
            # Khung control (Combobox chọn Animator)
            control_layout = QHBoxLayout()
            label = QLabel("🔎 Chọn Animator / Chế độ xem:")
            label.setStyleSheet("color: #e0e0e0; font-weight: bold; font-family: 'Segoe UI'; font-size: 13px;")
            control_layout.addWidget(label)
            
            self.combo_box = QComboBox()
            self.combo_box.setStyleSheet("background-color: #2d2d30; color: #e0e0e0; padding: 5px; border-radius: 3px; border: 1px solid #3e3e42;")
            self.combo_box.addItem("Tổng Hợp Tất Cả (Total)")
            
            # Lấy danh sách các file excel trong thư mục Total (loại trừ Total.xlsx và Data_All.xlsx)
            animator_files = [f for f in os.listdir(total_folder) if f.endswith('.xlsx') and f not in ('Total.xlsx', 'Data_All.xlsx')]
            animator_names = [os.path.splitext(f)[0] for f in animator_files]
            
            for name in sorted(animator_names):
                self.combo_box.addItem(name)
                
            control_layout.addWidget(self.combo_box)
            control_layout.addStretch()
            layout.addLayout(control_layout)
            
            # Tạo Treeview
            self.tree = QTreeWidget()
            layout.addWidget(self.tree)
            
            # Hàm cập nhật bảng dữ liệu khi đổi Combobox
            def update_table(index):
                self.tree.clear()
                selection = self.combo_box.currentText()
                
                if selection == "Tổng Hợp Tất Cả (Total)":
                    file_to_load = total_file
                else:
                    file_to_load = os.path.join(total_folder, f"{selection}.xlsx")
                
                if os.path.exists(file_to_load):
                    try:
                        df = pd.read_excel(file_to_load)
                        self.tree.setColumnCount(len(df.columns))
                        self.tree.setHeaderLabels(df.columns)
                        
                        # Đổ dữ liệu
                        for idx, row in df.iterrows():
                            item = QTreeWidgetItem([str(val) for val in row])
                            self.tree.addTopLevelItem(item)
                            
                        # Resize column
                        for i in range(len(df.columns)):
                            self.tree.resizeColumnToContents(i)
                    except Exception as e:
                        QMessageBox.critical(self.result_window, "Lỗi", f"Không thể lấy nội dung bảng:\n{e}")

            # Gắn sự kiện thay đổi lựa chọn
            self.combo_box.currentIndexChanged.connect(update_table)
            
            # Load lần đầu (chọn Total)
            update_table(0)
                        
            self.result_window.show()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở cửa sổ tra cứu:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Thiết lập icon nếu có
    logo_path = resource_path("Animeow_logo.jpg")
    if os.path.exists(logo_path):
        app.setWindowIcon(QIcon(logo_path))
        
    window = AnimeowTrackerApp()
    window.show()
    sys.exit(app.exec())
