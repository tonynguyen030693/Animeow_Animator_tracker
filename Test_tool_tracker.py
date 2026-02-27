import pandas as pd
import numpy as np
import os
import re
import traceback
try:
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg') # Không mở cửa sổ UI khi vẽ để tránh lỗi luồng
except ImportError:
    pass

def get_percentages():
    return {
        "P1": 10, "P2": 10, "P3": 10, "P4": 10, "P5": 10,
        "P6": 10, "P7": 10, "P8": 10, "P9": 10, "P10": 90,
        "S1": 10, "S2": 10
    }

def process_data(folder_path, log_callback=None):
    """
    Hàm xử lý chính (thay thế Action cũ).
    Tham số `log_callback` dùng để trả log về GUI.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
            
    try:
        log(f"Đang đọc các file Excel từ: {folder_path} (Bao gồm cả thư mục con)")
        dfs = []
        for root, dirs, files in os.walk(folder_path):
            # Bỏ qua thư mục báo cáo (Total) nếu có để tránh đọc file cũ/gộp
            if 'Total' in dirs:
                dirs.remove('Total')
                
            for filename in files:
                if filename.endswith('.xlsx') and not filename.startswith('~$'): # bỏ qua file tmp
                    file_path = os.path.join(root, filename)
                    try:
                        dfa = pd.read_excel(file_path)
                        if not dfa.empty:
                            dfs.append(dfa)
                    except Exception as e:
                        log(f"[-] Lỗi khi đọc file {filename}: {e}")
        
        if not dfs:
            log("[-] Không tìm thấy dữ liệu hợp lệ trong thư mục.")
            return False
            
        combined_df = pd.concat(dfs, ignore_index=True)
        log(f"[+] Đã gộp {len(dfs)} file, tổng cộng {len(combined_df)} dòng.")
        
        # 1. Tạo thư mục đầu ra
        total_path = os.path.join(folder_path, 'Total')
        os.makedirs(total_path, exist_ok=True)
        
        # 2. Lưu file gộp
        output_file_path = os.path.join(total_path, 'Data_All.xlsx')
        combined_df.to_excel(output_file_path, index=False)
        log("[+] Đã lưu file gộp Data_All.xlsx.")
        
        # 3. Tính điểm
        log("[*] Đang tính toán điểm số và làm sạch dữ liệu...")
        df = combined_df.copy()
        
        def calc_new_point(row):
            point = row.get('Point', 0)
            level = row.get('Level', '')
            if pd.isna(point): point = 0
            if level == 'Hard': return point * 1.1
            elif level == 'Easy': return point * 0.9
            return point

        df['New_Point'] = df.apply(calc_new_point, axis=1)
        df["Index_Point"] = df["New_Point"] / 10
        
        # Lấy danh sách cột Step (P1..S2) nếu có
        possible_steps = list(get_percentages().keys())
        step_cols = [c for c in df.columns if c in possible_steps]
        if not step_cols:
            log("[-] Cảnh báo: Không tìm thấy các cột Step (P1..S2) định cấu hình sẵn, đang cố lấy theo index như cũ...")
            step_cols = df.columns[5:17].tolist()

        # 4. Sử dụng pd.melt thay vì vòng lặp lồng nhau (Tối ưu hóa siêu tốc)
        log("[*] Đang phân bổ dữ liệu cho từng Animator...")
        
        # Đảm bảo các cột ID tồn tại trước khi melt
        id_vars = ['Project', 'Index_Point', 'Level']
        for c in id_vars:
            if c not in df.columns:
                df[c] = ""
                
        melted_df = pd.melt(
            df, 
            id_vars=id_vars, 
            value_vars=step_cols, 
            var_name='Stage', 
            value_name='Animator'
        )
        
        # Xóa các dòng Animator trắng hoặc NaN
        melted_df = melted_df.dropna(subset=['Animator'])
        melted_df['Animator'] = melted_df['Animator'].astype(str).str.strip()
        melted_df = melted_df[~melted_df['Animator'].isin(['nan', 'None', ''])]
        
        # Lưu file New_Data_Clean
        Export_Data_Clean(total_path, melted_df)
        log("[+] Đã xuất data sạch: New_Data_Clean.xlsx")
        
        # 5. Xuất báo cáo cho từng Animator
        animators = melted_df['Animator'].unique()
        log(f"[*] Tìm thấy {len(animators)} Animator. Đang tạo báo cáo cá nhân...")
        
        def remove_duplicates(cell_value):
            if pd.isna(cell_value): return ""
            cell_value = str(cell_value)
            unique_chars = []
            for char in cell_value:
                if char not in unique_chars:
                    unique_chars.append(char)
            return ''.join(unique_chars)
            
        percentages = get_percentages()
        
        def replace_stage(cell):
            if pd.isna(cell): return ""
            cell = str(cell)
            for k, v in percentages.items():
                cell = re.sub(rf'\b{k}\b', f"{v}%", cell)
            return cell
            
        def calculate_total_percentage(cell):
            if pd.isna(cell): return "0%"
            cell = str(cell)
            parts = cell.split('%')
            total = 0
            for p in parts:
                p_clean = ''.join(filter(str.isdigit, p))
                if p_clean:
                    total += float(p_clean)
            return f"{total * 10}%" # Giữ logic nhân 10 từ code cũ
        
        total_points = []
        for index, animator in enumerate(animators):
            df_anim = melted_df[melted_df['Animator'] == animator].copy()
            
            # Group by Project và cộng Index_Point, gộp Stage, Level
            agg_funcs = {'Index_Point': 'sum'}
            if 'Stage' in df_anim.columns:
                agg_funcs['Stage'] = lambda x: ','.join(x.astype(str))
            if 'Level' in df_anim.columns:
                agg_funcs['Level'] = lambda x: ','.join(x.astype(str))
                
            grouped = df_anim.groupby('Project').agg(agg_funcs).reset_index()
            grouped['Animator'] = animator
            
            if 'Level' in grouped.columns:
                grouped['Level'] = grouped['Level'].apply(remove_duplicates)
            
            if 'Stage' in grouped.columns:
                grouped['Stage'] = grouped['Stage'].apply(replace_stage)
                grouped['Stage'] = grouped['Stage'].apply(calculate_total_percentage)
                
            total_point_sum = grouped['Index_Point'].astype(float).sum()
            grouped['Index_Point'] = grouped['Index_Point'].map('{:.2f}'.format)
            
            # Xuất file cá nhân
            out_file = os.path.join(total_path, f"{animator}.xlsx")
            grouped.to_excel(out_file, index=False)
            
            # Lưu tổng point để làm file Total
            total_points.append({
                'Animator': animator,
                'Total_Point': total_point_sum
            })
            
        # 6. Xuất Total và Biểu đồ
        log("[*] Đang xuất file thống kê tổng Total.xlsx...")
        df_total = pd.DataFrame(total_points)
        if not df_total.empty:
            df_total['Total_Point'] = df_total['Total_Point'].map('{:.2f}'.format)
            total_file_path = os.path.join(total_path, "Total.xlsx")
            df_total.to_excel(total_file_path, index=False)
            
            # Xuất biểu đồ nếu có matplotlib
            try:
                if 'plt' in globals():
                    log("[*] Đang tạo biểu đồ tổng kết...")
                    df_chart = pd.DataFrame(total_points)
                    df_chart['Total_Point'] = df_chart['Total_Point'].astype(float)
                    df_chart = df_chart.sort_values(by='Total_Point', ascending=True)
                    
                    plt.figure(figsize=(10, max(6, len(df_chart) * 0.4)))
                    plt.barh(df_chart['Animator'], df_chart['Total_Point'], color='#007acc')
                    plt.xlabel('Tổng Điểm (Total Point)')
                    plt.ylabel('Animator')
                    plt.title('BIỂU ĐỒ TỔNG ĐIỂM CỦA CÁC ANIMATOR')
                    plt.grid(axis='x', linestyle='--', alpha=0.7)
                    
                    # Thêm số liệu lên thanh
                    for index, value in enumerate(df_chart['Total_Point']):
                        plt.text(value, index, f' {value:.2f}', va='center')
                    
                    plt.tight_layout()
                    chart_path = os.path.join(total_path, 'Toan_bo_Animator_Chart.png')
                    plt.savefig(chart_path, dpi=100)
                    plt.close()
                    log(f"[+] Đã lưu biểu đồ: Toan_bo_Animator_Chart.png")
            except Exception as e:
                log(f"[-] Không thể tạo biểu đồ: {e}")
                
        log("[+] Đã xuất Data Total.")
        
        log("\n=> HOÀN THÀNH XUẤT SẮC TRỌN BỘ!")
        return True
        
    except Exception as e:
        log("\n[!!!] LỖI NGHIÊM TRỌNG:")
        log(traceback.format_exc())
        return False

def Export_Data_Clean(output_folder, Animator_DF_Total):
    file_path = os.path.join(output_folder, 'New_Data_Clean.xlsx')
    Animator_DF_Total.to_excel(file_path, index=False)

# Hàm tương thích ngược nếu test file cũ
def Action():
    pass
