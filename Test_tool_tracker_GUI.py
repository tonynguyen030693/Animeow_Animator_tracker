import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import Test_tool_tracker as TTT
import os
try:
    from PIL import Image, ImageTk
except ImportError:
    pass

class ToolTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Test Tool Tracker - Tối Ưu Hóa & Giao Diện Hiện Đại")
        self.root.geometry("650x600") # Tăng chiều cao lên 1 chút để chứa logo
        
        # Configuration cho Dark Mode
        bg_color = '#1e1e1e'       # Nền tối chính
        surface_color = '#2d2d30'  # Nền các khung
        text_color = '#e0e0e0'     # Chữ sáng
        accent_color = '#007acc'   # Màu nhấn (Xanh lá/Xanh dương)
        
        # Style configuration (Using 'clam' theme cho phẳng)
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass # Fallback to default if clam not available
            
        style.configure('TFrame', background=bg_color)
        
        # Định dạng Buttom chung
        style.configure('TButton', 
                        font=('Segoe UI', 10, 'bold'), 
                        padding=8, 
                        background=surface_color, 
                        foreground=text_color,
                        bordercolor=bg_color,
                        lightcolor=surface_color,
                        darkcolor=surface_color)
        style.map('TButton', background=[('active', accent_color)])
        
        # Định dạng Treeview
        style.configure('Treeview', background=surface_color, foreground=text_color, fieldbackground=surface_color, rowheight=25)
        style.configure('Treeview.Heading', background='#1e1e1e', foreground=accent_color, font=('Segoe UI', 10, 'bold'))
        style.map('Treeview', background=[('selected', accent_color)])
        
        # Định dạng Label
        style.configure('TLabel', background=bg_color, foreground=text_color, font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 18, 'bold'), foreground='#4da6ff', background=bg_color)
        style.configure('TLabelframe', background=bg_color, foreground=text_color, font=('Segoe UI', 10, 'bold'))
        style.configure('TLabelframe.Label', background=bg_color, foreground=accent_color)
        
        self.root.configure(bg=bg_color)
        
        # Main Container
        main_frame = ttk.Frame(self.root, padding="25 25 25 25")
        main_frame.pack(fill=tk.BOTH, expand=True)

        def resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            import sys
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)

        # Logo Section
        logo_path = resource_path("Animeow_logo.jpg")
        if os.path.exists(logo_path):
            try:
                # Load và Image Resize logo
                img = Image.open(logo_path)
                # Resize cho vừa giao diện, ví dụ chiều ngang 200px
                width = 200
                ratio = (width / float(img.size[0]))
                height = int((float(img.size[1]) * float(ratio)))
                img = img.resize((width, height), Image.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(img)
                
                logo_label = tk.Label(main_frame, image=self.logo_image, bg=bg_color)
                logo_label.pack(pady=(0, 10))
            except Exception as e:
                print(f"Lỗi khi load logo: {e}")        
        
        # Header
        header = ttk.Label(main_frame, text="✅ CÔNG CỤ TÍNH ĐIỂM DỰ ÁN", style='Header.TLabel')
        header.pack(pady=(0, 25))
        
        # Input Section
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.path_var = tk.StringVar()
        
        # Cấu hình Style cho Entry
        style.configure('Dark.TEntry', fieldbackground=surface_color, foreground=text_color)
        
        self.path_entry = ttk.Entry(
            input_frame, textvariable=self.path_var, font=('Segoe UI', 10), state='readonly', style='Dark.TEntry'
        )
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))
        
        browse_btn = ttk.Button(input_frame, text="📁 Chọn Thư Mục", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)
        
        # Action Buttons
        self.btn_frame = ttk.Frame(main_frame)
        self.btn_frame.pack(pady=(0, 25))
        
        self.process_btn = ttk.Button(
            self.btn_frame, text="BẮT ĐẦU XỬ LÝ 🚀", command=self.start_processing, state=tk.DISABLED
        )
        self.process_btn.pack(side=tk.LEFT, padx=5, ipadx=15)
        
        self.view_btn = ttk.Button(
            self.btn_frame, text="TRA CỨU KẾT QUẢ 📊", command=self.view_results, state=tk.DISABLED
        )
        self.view_btn.pack(side=tk.LEFT, padx=5, ipadx=10)
        
        # Log Section
        log_frame = ttk.LabelFrame(main_frame, text=" 📝 Tiến Trình (Log) ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(
            log_frame, wrap=tk.WORD, height=12, font=('Consolas', 10), 
            bg='#0d0d0d', fg='#00ff00', padx=10, pady=10, insertbackground=text_color
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Footer
        footer = ttk.Label(self.root, text="Developed by Tool Team - Dark Mode Optimized", font=('Segoe UI', 8), foreground='#808080', background=bg_color)
        footer.pack(side=tk.BOTTOM, pady=10)
        
        self.write_log("[HỆ THỐNG]: Sẵn sàng bộ máy xử lý siêu tốc.")
        self.write_log("[HỆ THỐNG]: Vui lòng nhấn 'Chọn Thư Mục' chứa các file Excel của bạn!")
        
    def write_log(self, message):
        """Hàm ghi log, hỗ trợ scroll auto xuống cùng."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def log_callback(self, message):
        """Thread-safe log trigger via tkinter after."""
        self.root.after(0, self.write_log, message)

    def browse_folder(self):
        folder_path = filedialog.askdirectory(title="Chọn Thư Mục Chứa Dữ Liệu Excel")
        if folder_path:
            self.path_var.set(folder_path)
            self.process_btn.config(state=tk.NORMAL)
            self.view_btn.config(state=tk.NORMAL)
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
            self.write_log(f"📁 Đã chọn đường dẫn:\n => {folder_path}\n")

    def start_processing(self):
        folder_path = self.path_var.get()
        if not folder_path or not os.path.exists(folder_path):
            messagebox.showerror("Lỗi", "Đường dẫn không hợp lệ. Vui lòng chọn lại!")
            return
            
        # Khóa nút bấm trong khi chạy
        self.process_btn.config(state=tk.DISABLED)
        self.view_btn.config(state=tk.DISABLED)
        
        # Reset khu vực log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Chạy logic trên Thread phụ dọn dẹp data (Daemon = True để tắt Thread khi đóng Main UI)
        thread = threading.Thread(target=self.run_logic, args=(folder_path,), daemon=True)
        thread.start()
        
    def run_logic(self, folder_path):
        # Trả về bool True/False
        success = TTT.process_data(folder_path, self.log_callback)
        # Báo hiệu UI khi kết thúc
        self.root.after(0, self.finish_processing, success)
        
    def finish_processing(self, success):
        # Mở khóa nút bấm
        self.process_btn.config(state=tk.NORMAL)
        self.view_btn.config(state=tk.NORMAL)
        if success:
            messagebox.showinfo("Hoàn Thành", "Quá trình xử lý dữ liệu đã trơn tru thành công!")
        else:
            messagebox.showwarning("Cảnh Báo", "Có lỗi xảy ra trong quá trình xử lý.\nVui lòng xem chi tiết trên màn hình Log!")

    def view_results(self):
        import pandas as pd
        folder_path = self.path_var.get()
        if not folder_path:
            return
            
        total_file = os.path.join(folder_path, 'Total', 'Total.xlsx')
        if not os.path.exists(total_file):
            messagebox.showinfo("Thông Báo", "Chưa có dữ liệu thống kê. Vui lòng bấm 'Bắt Đầu Xử Lý' trước hoặc chọn thư mục có chứa thư mục Total sẵn!")
            return
            
        try:
            df = pd.read_excel(total_file)
            
            top = tk.Toplevel(self.root)
            top.title("Kết Quả Thống Kê Điểm Animator")
            top.geometry("500x400")
            top.configure(bg='#1e1e1e')
            
            # Khung chứa bảng
            frame = ttk.Frame(top, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Tạo Treeview
            columns = list(df.columns)
            tree = ttk.Treeview(frame, columns=columns, show='headings')
            
            # Thiết lập cột
            for col in columns:
                tree.heading(col, text=col.replace('_', ' '))
                tree.column(col, anchor=tk.CENTER, width=150)
                
            # Đổ dữ liệu
            for index, row in df.iterrows():
                tree.insert("", tk.END, values=list(row))
                
            # Scrollbar
            scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file dữ liệu:\n{e}")

# Chạy Main App
def main():
    root = tk.Tk()
    app = ToolTrackerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
