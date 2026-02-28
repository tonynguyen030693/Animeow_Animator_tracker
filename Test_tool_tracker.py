import pandas as pd
import numpy as np
import os
import re
import traceback
import sys
try:
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg') # Không mở cửa sổ UI khi vẽ để tránh lỗi luồng
except ImportError:
    pass

def get_percentages():
    return {
        "P1": 10, "P2": 10, "P3": 10, "P4": 10, "P5": 10,
        "P6": 10, "P7": 10, "P8": 10, "P9": 10, "P10": 10,
        "S1": 10, "S2": 10
    }

def get_folder_suffix(folder_path):
    """
    Finds the name of the folder that contains a 'Raw' subdirectory.
    """
    for root, dirs, files in os.walk(folder_path):
        if 'Raw' in dirs:
            return os.path.basename(root)
    return None

def find_month_folders(root_path):
    """
    Tìm tất cả các thư mục chứa thư mục con 'Raw'.
    """
    month_folders = []
    # Nếu bản thân folder_path đã có Raw
    if os.path.exists(os.path.join(root_path, 'Raw')):
        return [root_path]
        
    # Duyệt sâu tối đa 2 cấp để tìm folder tháng (VD: Data_test / 2026 / 202601 / Raw)
    for root, dirs, files in os.walk(root_path):
        if 'Raw' in dirs:
            month_folders.append(root)
            # Không cần duyệt sâu hơn nữa vào folder này vì đã tìm thấy Raw
            dirs.clear() 
    return month_folders

def process_single_month(folder_path, log_callback=None):
    """
    Logic xử lý cho một tháng cụ thể.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            try:
                print(msg)
            except UnicodeEncodeError:
                print(msg.encode(sys.stdout.encoding, errors='backslashreplace').decode(sys.stdout.encoding))
                
    try:
        suffix = get_folder_suffix(folder_path)
        total_folder_name = f'Total_{suffix}' if suffix else 'Total'
        total_file_name = f'Total_{suffix}.xlsx' if suffix else 'Total.xlsx'
        
        log(f"\n--- [ĐANG XỬ LÝ]: {os.path.basename(folder_path)} ---")
        log(f"Đang đọc các file Excel từ: {folder_path}")
        dfs = []
        
        # Chỉ quét trong folder hiện tại và các folder con KHÔNG PHẢI Total_
        for root, dirs, files in os.walk(folder_path):
            dirs[:] = [d for d in dirs if not d.startswith('Total')]
            for filename in files:
                if filename.endswith('.xlsx') and not filename.startswith('~$'):
                    file_path = os.path.join(root, filename)
                    try:
                        dfa = pd.read_excel(file_path)
                        if not dfa.empty:
                            dfa['Source_File'] = filename
                            dfa['Folder_Name'] = os.path.basename(root)
                            dfs.append(dfa)
                    except Exception as e:
                        log(f"[-] Lỗi khi đọc file {filename}: {e}")
        
        if not dfs:
            log(f"[-] Không tìm thấy dữ liệu trong {folder_path}")
            return False
            
        combined_df = pd.concat(dfs, ignore_index=True)
        log(f"[+] Đã gộp {len(dfs)} file, tổng cộng {len(combined_df)} dòng.")
        
        # 1. Tạo thư mục đầu ra
        total_path = os.path.join(folder_path, total_folder_name)
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
        id_vars = ['Project', 'Index_Point', 'Level', 'Source_File', 'Folder_Name']
        for c in id_vars:
            if c not in df.columns:
                df[c] = ""
                
        # Tránh lỗi ValueError của pandas khi columns 'Stage' hoặc 'Animeowee' đã tồn tại trong dữ liệu đầu vào
        for col in ['Stage', 'Animeowee']:
            if col in df.columns:
                df = df.drop(columns=[col])
                
        melted_df = pd.melt(df, id_vars=id_vars, value_vars=step_cols, var_name='Stage', value_name='Animeowee')
        
        # Xóa các dòng Animeowee trắng hoặc NaN
        melted_df = melted_df.dropna(subset=['Animeowee'])
        melted_df['Animeowee'] = melted_df['Animeowee'].astype(str).str.strip()
        melted_df = melted_df[~melted_df['Animeowee'].isin(['nan', 'None', ''])]
        
        # Lưu file New_Data_Clean
        Export_Data_Clean(total_path, melted_df)
        log("[+] Đã xuất data sạch: New_Data_Clean.xlsx")
        
        # 5. Xuất báo cáo cho từng Animeowee
        animators = melted_df['Animeowee'].unique()
        log(f"[*] Tìm thấy {len(animators)} Animeowee. Đang tạo báo cáo cá nhân...")
        
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
            return f"{total}%" # Returns actual value
        
        total_points = []
        for index, animator in enumerate(animators):
            df_anim = melted_df[melted_df['Animeowee'] == animator].copy()
            
            # Group by Project và cộng Index_Point, gộp Stage, Level
            agg_funcs = {'Index_Point': 'sum'}
            if 'Stage' in df_anim.columns:
                agg_funcs['Stage'] = lambda x: ','.join(pd.Series(x).dropna().drop_duplicates().astype(str))
            if 'Level' in df_anim.columns:
                agg_funcs['Level'] = lambda x: ','.join(pd.Series(x).dropna().drop_duplicates().astype(str))
                
            grouped = df_anim.groupby('Project').agg(agg_funcs).reset_index()
            grouped['Animeowee'] = animator
            
            if 'Level' in grouped.columns:
                grouped['Level'] = grouped['Level'].apply(remove_duplicates)
            
            if 'Stage' in grouped.columns:
                grouped['Stage'] = grouped['Stage'].apply(replace_stage)
                grouped['Stage'] = grouped['Stage'].apply(calculate_total_percentage)
                
            total_point_sum = grouped['Index_Point'].astype(float).sum()
            grouped['Index_Point'] = grouped['Index_Point'].map('{:.2f}'.format)
            
            # Xuất file cá nhân
            animator_folder = os.path.join(total_path, 'Animeowee_tracker')
            os.makedirs(animator_folder, exist_ok=True)
            out_file = os.path.join(animator_folder, f"{animator}.xlsx")
            grouped.to_excel(out_file, index=False)
            
            # Lưu tổng point để làm file Total
            total_points.append({
                'Animeowee': animator,
                'Total_Point': total_point_sum
            })
            
        # 6. Xuất Total và Biểu đồ
        log(f"[*] Đang xuất file thống kê tổng {total_file_name}...")
        df_total = pd.DataFrame(total_points)
        # 6.1 Gộp điểm theo Project
        log("[*] Đang tổng hợp điểm theo từng Project...")
        project_points = []
        if not melted_df.empty:
            grouped_projects = melted_df.groupby('Project')['Index_Point'].sum().reset_index()
            for index, row in grouped_projects.iterrows():
                project_points.append({
                    'Project': row['Project'],
                    'Total_Point': row['Index_Point']
                })
        df_project_total = pd.DataFrame(project_points)

        total_file_path = os.path.join(total_path, total_file_name)
        
        # Sử dụng ExcelWriter để lưu nhiều sheet
        with pd.ExcelWriter(total_file_path, engine='openpyxl') as writer:
            # Sheet 1: Animeowee
            if not df_total.empty:
                df_total['Total_Point'] = df_total['Total_Point'].map('{:.2f}'.format)
                df_total.to_excel(writer, sheet_name='Total_Animeowee', index=False)
            
            # Sheet 2: Project
            if not df_project_total.empty:
                df_project_total['Total_Point'] = df_project_total['Total_Point'].map('{:.2f}'.format)
                df_project_total.to_excel(writer, sheet_name='Total_Project', index=False)
                
        # 6.2 Tạo các file _total.xlsx theo TỪNG File Excel Gốc CÙNG VỚI CHI TIẾT TỪNG ANIMEOWEE
        log("[*] Đang xuất báo cáo cho từng file dữ liệu gốc...")
        if 'Source_File' in melted_df.columns:
            # Lấy danh sách các file gốc và folder tương ứng để duyệt
            file_origins = melted_df[['Source_File', 'Folder_Name']].drop_duplicates()
            
            for _, row_origin in file_origins.iterrows():
                src_file = str(row_origin['Source_File'])
                f_name = str(row_origin['Folder_Name'])
                
                # Lọc data chỉ thuộc về file gốc này
                df_single_file = melted_df[melted_df['Source_File'] == src_file]
                
                # Tính tổng điểm cho từng Animeowee trong file này
                animator_group = df_single_file.groupby('Animeowee')['Index_Point'].sum().reset_index()
                
                # Tính tổng điểm của toàn bộ file
                total_file_pt = animator_group['Index_Point'].sum()
                
                # Format cột điểm
                animator_group['Index_Point'] = animator_group['Index_Point'].map('{:.2f}'.format)
                
                # Đổi tên cột cho đẹp trước khi xuất
                animator_group.rename(columns={'Animeowee': 'Animeowee', 'Index_Point': 'Điểm'}, inplace=True)
                
                # Tạo một dòng chứa Tổng Cộng cho file
                total_row = pd.DataFrame([{'Animeowee': '>> TỔNG CỘNG FILE <<', 'Điểm': f"{total_file_pt:.2f}"}])
                final_df_output = pd.concat([animator_group, total_row], ignore_index=True)
                
                base_name, ext = os.path.splitext(src_file)
                new_file_name = f"{base_name}_total{ext}"
                project_folder = os.path.join(total_path, 'Project_Animeowee_tracker')
                os.makedirs(project_folder, exist_ok=True)
                new_file_path = os.path.join(project_folder, new_file_name)
                
                # Lưu ra excel
                try:
                    final_df_output.to_excel(new_file_path, index=False)
                except Exception as e:
                    log(f"[-] Lỗi khi xuất báo cáo file {new_file_name}: {str(e)}")
            # Xuất biểu đồ nếu có matplotlib
            # Sử dụng hàm helper để tạo biểu đồ
            create_summary_chart(df_total, total_path, suffix, log)
                
        log(f"[+] Hoàn thành xử lý tháng cho {os.path.basename(folder_path)}")
        return True
    except Exception as e:
        log(f"[!!!] LỖI tại {folder_path}:")
        log(traceback.format_exc())
        return False

def create_summary_chart(df_total, output_path, title_suffix, log_callback=None):
    """
    Hàm helper để vẽ biểu đồ tổng kết điểm (sắp xếp từ cao xuống thấp)
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else:
            try: print(msg)
            except: pass

    try:
        if 'plt' in globals() and not df_total.empty:
            log(f"[*] Đang tạo biểu đồ tổng kết cho {title_suffix}...")
            # Copy để tránh warning và ép kiểu
            df_chart = df_total.copy()
            df_chart['Total_Point'] = df_chart['Total_Point'].astype(float)
            
            # Sắp xếp ĐỂ VẼ BARH (matplotlib vẽ từ dưới lên nên sort ascending=True để cao nhất nằm trên cùng)
            df_chart = df_chart.sort_values(by='Total_Point', ascending=True)
            
            plt.figure(figsize=(10, max(6, len(df_chart) * 0.4)))
            plt.barh(df_chart['Animeowee'], df_chart['Total_Point'], color='#4da6ff')
            plt.xlabel('Tổng Điểm (Total Point)')
            plt.ylabel('Animeowee')
            plt.title(f'BIỂU ĐỒ TỔNG ĐIỂM {title_suffix}')
            plt.grid(axis='x', linestyle='--', alpha=0.7)
            
            # Thêm số liệu lên đầu mỗi thanh
            for index, value in enumerate(df_chart['Total_Point']):
                plt.text(value, index, f' {value:.2f}', va='center', fontweight='bold')
            
            plt.tight_layout()
            chart_file = os.path.join(output_path, 'Toan_bo_Animeowee_Chart.png')
            plt.savefig(chart_file, dpi=100)
            plt.close()
            log(f"[+] Đã lưu biểu đồ: Toan_bo_Animeowee_Chart.png tại {output_path}")
    except Exception as e:
        log(f"[-] Không thể tạo biểu đồ: {e}")

def generate_yearly_summaries(root_path, month_folders, log_callback=None):
    """
    Gộp các file Total_YYYYMM.xlsx thành Total_YYYY.xlsx
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else:
            try: print(msg)
            except: pass

    # Phân loại month_folders theo năm
    year_to_months = {}
    for mf in month_folders:
        suffix = get_folder_suffix(mf)
        if suffix and len(suffix) >= 4 and suffix[:4].isdigit():
            year = suffix[:4]
            if year not in year_to_months: year_to_months[year] = []
            year_to_months[year].append(mf)

    if not year_to_months: return

    log(f"\n[*] Đang tổng hợp dữ liệu theo năm cho {len(year_to_months)} năm tìm thấy...")

    for year, m_paths in year_to_months.items():
        all_ani_dfs = []
        all_proj_dfs = []

        for mp in m_paths:
            suffix = get_folder_suffix(mp)
            total_file = os.path.join(mp, f"Total_{suffix}", f"Total_{suffix}.xlsx")
            if os.path.exists(total_file):
                try:
                    ani_df = pd.read_excel(total_file, sheet_name='Total_Animeowee')
                    proj_df = pd.read_excel(total_file, sheet_name='Total_Project')
                    all_ani_dfs.append(ani_df)
                    all_proj_dfs.append(proj_df)
                except Exception as e:
                    log(f"[-] Bỏ qua {suffix} do lỗi đọc file: {e}")

        if not all_ani_dfs: continue

        # Gộp và sum điểm
        final_ani = pd.concat(all_ani_dfs).groupby('Animeowee')['Total_Point'].apply(lambda x: x.astype(float).sum()).reset_index()
        final_proj = pd.concat(all_proj_dfs).groupby('Project')['Total_Point'].apply(lambda x: x.astype(float).sum()).reset_index()

        # Tìm folder năm (thường là cha của folder tháng)
        # Ví dụ: Data_test/2026/202601 -> Year folder là Data_test/2026
        # Hoặc nếu chọn ngay 2026 -> root_path chính là year folder
        year_folder = None
        for mp in m_paths:
            parent = os.path.dirname(mp)
            if os.path.basename(parent) == year:
                year_folder = parent
                break
        
        if not year_folder:
            year_folder = root_path # Fallback

        total_year_path = os.path.join(year_folder, f"Total_{year}")
        os.makedirs(total_year_path, exist_ok=True)
        total_year_file = os.path.join(total_year_path, f"Total_{year}.xlsx")

        with pd.ExcelWriter(total_year_file, engine='openpyxl') as writer:
            final_ani['Total_Point'] = final_ani['Total_Point'].map('{:.2f}'.format)
            final_ani.to_excel(writer, sheet_name='Total_Animeowee', index=False)
            final_proj['Total_Point'] = final_proj['Total_Point'].map('{:.2f}'.format)
            final_proj.to_excel(writer, sheet_name='Total_Project', index=False)

        # Tạo biểu đồ cho báo cáo năm
        create_summary_chart(pd.concat(all_ani_dfs).groupby('Animeowee')['Total_Point'].sum().reset_index(), total_year_path, f"NĂM {year}", log)

        log(f"[+] Đã tạo báo cáo năm: Total_{year}.xlsx tại {total_year_path}")

def process_data(folder_path, log_callback=None):
    """
    Hàm entry point, hỗ trợ xử lý hàng loạt.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            try:
                print(msg)
            except UnicodeEncodeError:
                print(msg.encode(sys.stdout.encoding, errors='backslashreplace').decode(sys.stdout.encoding))

    try:
        log(f"[*] Đang tìm kiếm các thư mục dữ liệu trong: {folder_path}")
        month_folders = find_month_folders(folder_path)
        
        if not month_folders:
            log("[-] Không tìm thấy thư mục nào có folder 'Raw'.")
            return False
            
        log(f"[+] Tìm thấy {len(month_folders)} thư mục khả dụng.")
        
        success_count = 0
        processed_successfully = []
        for month_path in month_folders:
            if process_single_month(month_path, log_callback):
                success_count += 1
                processed_successfully.append(month_path)
        
        if processed_successfully:
            generate_yearly_summaries(folder_path, processed_successfully, log_callback)
                
        log(f"\n=> ĐÃ XỬ LÝ XONG: Thành công {success_count}/{len(month_folders)} folder.")
        return success_count > 0
        
    except Exception as e:
        log("\n[!!!] LỖI TRONG QUÁ TRÌNH TÌM KIẾM/XỬ LÝ HÀNG LOẠT:")
        log(traceback.format_exc())
        return False

def Export_Data_Clean(output_folder, Animator_DF_Total):
    file_path = os.path.join(output_folder, 'New_Data_Clean.xlsx')
    Animator_DF_Total.to_excel(file_path, index=False)

# Hàm tương thích ngược nếu test file cũ
def Action():
    pass
