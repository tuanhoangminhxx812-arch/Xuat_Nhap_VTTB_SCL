import streamlit as st
import pandas as pd
import datetime
import os
import sys

# Set up page configurations first
st.set_page_config(
    page_title="TỔNG HỢP VẬT TƯ THIẾT BỊ SCL",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Force system to use UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Import core data processor functions
from data_processor import (
    parse_import,
    parse_export,
    consolidate_data,
    write_to_template,
    generate_voltage_separation_data,
    write_to_voltage_template,
    classify_voltage,
    clean_project_code,
    PROJECT_NAMES,
    write_detailed_scl_classification,
    parse_pm_092
)

# Default File Paths in the Workspace Directory
DEFAULT_IMPORT = "INV-007Atừ 01012026 đến 31052026.xlsx"
DEFAULT_EXPORT = "INV-009 từ 01012026 đến 31052026.xlsx"
DEFAULT_TEMPLATE = "Xuat_Nhap(mau).xlsx"
DEFAULT_OUTPUT = "Xuat_Nhap_TongHop.xlsx"

DEFAULT_TEMPLATE_V = "Tách PP-BL.xlsx"
DEFAULT_OUTPUT_V = "Tach_ChiPhi_PP_BL.xlsx"
DEFAULT_OUTPUT_DETAILED = "TachPP_BL_ChiTiet.xlsx"

# Custom Premium Styling & Outfitted Typography
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Premium KPI Cards */
    .kpi-container {
        display: flex;
        gap: 1.5rem;
        margin-bottom: 2rem;
    }
    
    .kpi-card {
        flex: 1;
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }
    
    .kpi-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.4);
        border-color: rgba(59, 130, 246, 0.4);
    }
    
    .kpi-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 4px;
    }
    
    .kpi-import::before {
        background: linear-gradient(90deg, #10B981, #059669);
    }
    
    .kpi-export::before {
        background: linear-gradient(90deg, #3B82F6, #2563EB);
    }
    
    .kpi-count::before {
        background: linear-gradient(90deg, #F59E0B, #D97706);
    }
    
    .kpi-voltage::before {
        background: linear-gradient(90deg, #8B5CF6, #6D28D9);
    }
    
    .kpi-title {
        font-size: 0.875rem;
        color: #94A3B8;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 0.5rem;
    }
    
    .kpi-value {
        font-size: 1.75rem;
        font-weight: 700;
        color: #F8FAFC;
        line-height: 1.2;
    }
    
    .kpi-sub {
        font-size: 0.75rem;
        color: #64748B;
        margin-top: 0.5rem;
    }
    
    /* Header styling */
    .title-banner {
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 1.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.2);
    }
    
    .title-banner h1 {
        color: white !important;
        margin: 0 !important;
        font-weight: 700;
        letter-spacing: -0.02em;
    }
    
    .title-banner p {
        color: #DBEAFE !important;
        margin: 0.5rem 0 0 0 !important;
        font-size: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to get last modified times of default files on disk
def get_file_mtimes():
    mtimes = {}
    for key, path in [
        ("import", DEFAULT_IMPORT), 
        ("export", DEFAULT_EXPORT), 
        ("template", DEFAULT_TEMPLATE),
        ("template_v", DEFAULT_TEMPLATE_V),
        ("mapping_v", "TachPP_BL mẫu.xlsx")
    ]:
        if os.path.exists(path):
            mtimes[key] = os.path.getmtime(path)
        else:
            mtimes[key] = 0.0
    return mtimes

# Initialize Session States
if "parsed_import" not in st.session_state:
    st.session_state.parsed_import = None
if "parsed_export" not in st.session_state:
    st.session_state.parsed_export = None
if "file_mtimes" not in st.session_state:
    st.session_state.file_mtimes = {}
if "auto_consolidated" not in st.session_state:
    st.session_state.auto_consolidated = False

# Banner Title
st.markdown("""
<div class="title-banner">
    <h1>📊 HỆ THỐNG TỔNG HỢP VẬT TƯ THIẾT BỊ SCL</h1>
    <p>Tự động tổng hợp báo cáo gộp và báo cáo tách chi phí Trung/Hạ áp từ file gốc Nhập (INV-007A) và Xuất (INV-009)</p>
</div>
""", unsafe_allow_html=True)

# ---- SIDEBAR DESIGN ----
st.sidebar.markdown("### ⚙️ Cấu hình dữ liệu")

# Interactive filter controls
scl_only = st.sidebar.toggle("🔍 Chỉ lọc giao dịch SCL (Sửa chữa lớn)", value=True, 
                             help="Nếu bật, chỉ tổng hợp các giao dịch phục vụ sửa chữa lớn (có từ khóa SCL trong nội dung/chứng từ).")


# Manual file overrides (Upload custom files)
st.sidebar.markdown("---")
st.sidebar.markdown("### 📤 Tải lên file tùy chỉnh (Tùy chọn)")
uploaded_import = st.sidebar.file_uploader("Upload file Nhập mới (INV-007A)", type=["xlsx"])
uploaded_export = st.sidebar.file_uploader("Upload file Xuất mới (INV-009)", type=["xlsx"])

# ---- AUTO-CONSOLIDATION LOGIC (FILE WATCHER) ----
current_mtimes = get_file_mtimes()
files_changed = False

# Detect if file timestamps changed on disk compared to cached timestamps
if st.session_state.file_mtimes != current_mtimes:
    files_changed = True
    st.session_state.file_mtimes = current_mtimes

# Trigger parsing if files changed or if not already loaded
if files_changed or st.session_state.parsed_import is None or st.session_state.parsed_export is None:
    # Use uploaded files if present, else use default local files
    import_file = uploaded_import if uploaded_import else DEFAULT_IMPORT
    export_file = uploaded_export if uploaded_export else DEFAULT_EXPORT
    
    if (uploaded_import or os.path.exists(DEFAULT_IMPORT)) and (uploaded_export or os.path.exists(DEFAULT_EXPORT)):
        with st.spinner("🔄 Phát hiện thay đổi hoặc tải mới! Đang xử lý và tổng hợp cả 2 báo cáo..."):
            try:
                # Core parsing
                imp_records = parse_import(import_file)
                exp_records = parse_export(export_file)
                
                # Cache parsed results
                st.session_state.parsed_import = imp_records
                st.session_state.parsed_export = exp_records
                st.session_state.auto_consolidated = True
                
                # 1. Write standard consolidation output file
                if os.path.exists(DEFAULT_TEMPLATE):
                    df_full = consolidate_data(imp_records, exp_records, scl_only=False) # Get all for saving default file
                    write_to_template(df_full, DEFAULT_TEMPLATE, DEFAULT_OUTPUT)
                
                # 2. Write SCL voltage separation output file (Tách PP-BL)
                if os.path.exists(DEFAULT_TEMPLATE_V):
                    df_v = generate_voltage_separation_data(imp_records, exp_records)
                    write_to_voltage_template(df_v, DEFAULT_TEMPLATE_V, DEFAULT_OUTPUT_V)

                # 3. Write detailed SCL classification sheet TachPP_BL_ChiTiet.xlsx
                if os.path.exists(DEFAULT_TEMPLATE):
                    df_scl_only = consolidate_data(imp_records, exp_records, scl_only=True)
                    write_detailed_scl_classification(df_scl_only, DEFAULT_TEMPLATE, DEFAULT_OUTPUT_DETAILED)
                
                st.toast("Tự động tổng hợp cả 2 báo cáo thành công!", icon="🔄")
            except Exception as e:
                st.error(f"Lỗi khi xử lý file Excel: {e}")
    else:
        st.warning("⚠️ Không tìm thấy file nguồn Nhập (INV-007A...) hoặc Xuất (INV-009...) trong thư mục. Vui lòng đặt file vào thư mục hoặc tải lên ở thanh menu bên trái.")

# Use cached data
imp_records = st.session_state.parsed_import if st.session_state.parsed_import else []
exp_records = st.session_state.parsed_export if st.session_state.parsed_export else []

# Setup 2-Tab Navigation
tab1, tab2 = st.tabs(["📊 Tổng Hợp Báo Cáo Gộp", "⚡ Tách Phân Phối - Bán Lẻ (Trung/Hạ Áp)"])

# ==============================================================================
# TAB 1: BÁO CÁO GỘP (Original consolidated sheet)
# ==============================================================================
with tab1:
    st.markdown("### 🔍 Bộ lọc báo cáo gộp")
    
    # Collect all unique warehouses from the parsed voucher codes
    all_whs = set()
    for r in imp_records:
        if r["voucher"] and len(r["voucher"].split('.')) > 1:
            all_whs.add(r["voucher"].split('.')[1])
    for r in exp_records:
        if r["voucher"] and len(r["voucher"].split('.')) > 1:
            all_whs.add(r["voucher"].split('.')[1])
    all_whs = sorted(list(all_whs))

    col_f1, col_f2, col_f3 = st.columns([3, 4, 3])
    with col_f1:
        selected_whs = st.multiselect(
            "🏢 Lọc theo Mã Kho (Warehouse):",
            options=all_whs,
            default=[],
            placeholder="Chọn một hoặc nhiều kho...",
            key="tab1_wh_filter"
        )
    with col_f2:
        keyword = st.text_input(
            "🔍 Tìm kiếm nhanh vật tư:",
            placeholder="Nhập mã vật tư, tên vật tư hoặc nội dung diễn giải...",
            key="tab1_keyword_filter"
        ).strip()
    with col_f3:
        st.markdown("<div style='height: 1.8rem;'></div>", unsafe_allow_html=True)
        re_run = st.button("🔄 Làm mới dữ liệu", use_container_width=True, key="tab1_rerun")
        if re_run:
            # Clear cache to force reload
            st.session_state.parsed_import = None
            st.session_state.parsed_export = None
            st.rerun()

    # Apply Filters dynamically in the view
    wh_filter = selected_whs if len(selected_whs) > 0 else None
    keyword_filter = keyword if keyword != "" else None

    df_filtered = consolidate_data(
        imp_records, 
        exp_records, 
        scl_only=scl_only, 
        warehouse_filter=wh_filter, 
        keyword_filter=keyword_filter
    )

    # Write custom output for download
    temp_download_path = "Xuat_Nhap_Download.xlsx"
    download_ready = False
    if not df_filtered.empty and os.path.exists(DEFAULT_TEMPLATE):
        try:
            write_to_template(df_filtered, DEFAULT_TEMPLATE, temp_download_path)
            download_ready = True
        except Exception as e:
            st.error(f"Lỗi khi viết file download: {e}")

    # Calculate KPI values based on filtered results
    import_df = df_filtered[df_filtered["Nhập - Thành tiền"] > 0] if not df_filtered.empty else pd.DataFrame()
    export_df = df_filtered[df_filtered["Xuất - Thành tiền"] > 0] if not df_filtered.empty else pd.DataFrame()

    total_import_val = import_df["Nhập - Thành tiền"].sum() if not import_df.empty else 0.0
    total_export_val = export_df["Xuất - Thành tiền"].sum() if not export_df.empty else 0.0

    total_import_items = len(import_df)
    total_export_items = len(export_df)

    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card kpi-import">
            <div class="kpi-title">Tổng Giá Trị Nhập</div>
            <div class="kpi-value">{total_import_val:,.0f} <span style='font-size: 1.1rem;'>VNĐ</span></div>
            <div class="kpi-sub">Số lượng vật tư nhập: {total_import_items} dòng giao dịch</div>
        </div>
        <div class="kpi-card kpi-export">
            <div class="kpi-title">Tổng Giá Trị Xuất</div>
            <div class="kpi-value">{total_export_val:,.0f} <span style='font-size: 1.1rem;'>VNĐ</span></div>
            <div class="kpi-sub">Số lượng vật tư xuất: {total_export_items} dòng giao dịch</div>
        </div>
        <div class="kpi-card kpi-count">
            <div class="kpi-title">Tổng Giao Dịch</div>
            <div class="kpi-value">{len(df_filtered):,} <span style='font-size: 1.1rem;'>Giao dịch</span></div>
            <div class="kpi-sub">Chế độ lọc SCL: {"BẬT (Chỉ lọc SCL)" if scl_only else "TẮT (Toàn bộ)"}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.auto_consolidated:
        st.success(f"⚡ **[Hệ thống tự động]**: Đã phát hiện thay đổi và ghi đè file tổng hợp thành công tại các đường dẫn local: `{DEFAULT_OUTPUT}`, `{DEFAULT_OUTPUT_V}`, và `{DEFAULT_OUTPUT_DETAILED}`!")
        st.session_state.auto_consolidated = False

    if not df_filtered.empty:
        st.markdown("### 📈 Phân Tích & Đối Soát Phát Sinh Theo Từng Tháng")
        
        # Calculate Monthly Summary Data
        df_grp = df_filtered.copy()
        df_grp["tháng"] = df_grp["tháng"].fillna(0).astype(int)
        
        months_list = sorted(df_grp["tháng"].unique())
        monthly_data = []
        for m in months_list:
            month_name = f"Tháng {m}" if m > 0 else "Không rõ"
            df_m = df_grp[df_grp["tháng"] == m]
            
            # Nhập
            nhap_df = df_m[df_m["Nhập - Thành tiền"] > 0]
            nhap_val = nhap_df["Nhập - Thành tiền"].sum()
            nhap_count = len(nhap_df)
            
            # Xuất
            xuat_df = df_m[df_m["Xuất - Thành tiền"] > 0]
            xuat_val = xuat_df["Xuất - Thành tiền"].sum()
            xuat_count = len(xuat_df)
            
            monthly_data.append({
                "Tháng": month_name,
                "Tổng Nhập (VNĐ)": nhap_val,
                "Số dòng Nhập": nhap_count,
                "Tổng Xuất (VNĐ)": xuat_val,
                "Số dòng Xuất": xuat_count
            })
            
        df_monthly_summary = pd.DataFrame(monthly_data)
        
        # Render side-by-side: Table on Left, Chart on Right
        col_t1, col_t2 = st.columns([5, 5])
        
        with col_t1:
            st.markdown("##### 📋 Bảng tổng hợp số liệu từng tháng")
            st.dataframe(
                df_monthly_summary,
                column_config={
                    "Tổng Nhập (VNĐ)": st.column_config.NumberColumn("Tổng Nhập (VNĐ)", format="%,.0f"),
                    "Số dòng Nhập": st.column_config.NumberColumn("Số dòng Nhập", format="%d"),
                    "Tổng Xuất (VNĐ)": st.column_config.NumberColumn("Tổng Xuất (VNĐ)", format="%,.0f"),
                    "Số dòng Xuất": st.column_config.NumberColumn("Số dòng Xuất", format="%d")
                },
                use_container_width=True,
                hide_index=True,
                height=300
            )
            
        with col_t2:
            st.markdown("##### 📊 Biểu đồ xu hướng phát sinh")
            df_filtered["tháng_str"] = df_filtered["tháng"].fillna(0).astype(int).apply(lambda x: f"Tháng {x}" if x > 0 else "Không rõ")
            chart_data = df_filtered.groupby("tháng_str")[["Nhập - Thành tiền", "Xuất - Thành tiền"]].sum()
            st.bar_chart(chart_data, height=300)

    st.markdown("### 📋 Xem trước dữ liệu tổng hợp")
    if not df_filtered.empty:
        df_present = df_filtered.copy()
        st.dataframe(
            df_present,
            column_config={
                "Nhập - Số lượng": st.column_config.NumberColumn("Nhập - Số lượng", format="%f"),
                "Nhập - Đơn giá": st.column_config.NumberColumn("Nhập - Đơn giá", format="%,.2f"),
                "Nhập - Thành tiền": st.column_config.NumberColumn("Nhập - Thành tiền", format="%,.0f"),
                "Xuất - Số lượng": st.column_config.NumberColumn("Xuất - Số lượng", format="%f"),
                "Xuất - Đơn giá": st.column_config.NumberColumn("Xuất - Đơn giá", format="%,.2f"),
                "Xuất - Thành tiền": st.column_config.NumberColumn("Xuất - Thành tiền", format="%,.0f"),
                "Ngày viết": st.column_config.DateColumn("Ngày viết", format="YYYY-MM-DD")
            },
            use_container_width=True,
            height=400
        )
        
        if download_ready:
            with open(temp_download_path, "rb") as file:
                st.download_button(
                    label="📥 Tải xuống File Tổng Hợp Báo Cáo Gộp (Đã áp dụng bộ lọc hiện tại)",
                    data=file,
                    file_name="Xuat_Nhap_TongHop_Loc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="tab1_download_btn"
                )
    else:
        st.warning("⚠️ Không có giao dịch nào thỏa mãn bộ lọc hiện tại.")

    # Clean up temp file
    if os.path.exists(temp_download_path) and not download_ready:
        try: os.remove(temp_download_path)
        except: pass

# ==============================================================================
# TAB 2: TÁCH PHÂN PHỐI - BÁN LẺ (TRUNG / HẠ ÁP)
# ==============================================================================
with tab2:
    st.markdown("### ⚡ Tách chi phí Trung áp - Hạ áp theo Khâu Phân phối và Bán lẻ")
    
    # Generate the voltage separation dataset based on cached records
    df_v = generate_voltage_separation_data(imp_records, exp_records)
    
    if not df_v.empty:
        # Collect available months dynamically
        available_months = sorted(list(df_v["tháng"].dropna().unique()))
        
        # Dashboard Filter
        col_t2_1, col_t2_2 = st.columns([3, 7])
        with col_t2_1:
            selected_month_num = st.selectbox(
                "📅 Chọn tháng để xem trước và kiểm tra:",
                options=available_months,
                format_func=lambda x: f"Tháng {int(x)}",
                key="tab2_month_select"
            )
        
        # Calculations for KPIs in selected month
        df_month_v = df_v[df_v["tháng"] == selected_month_num]
        
        # Check if PM_092.xlsx is present
        pm092_file = "PM_092.xlsx"
        pm092_exists = os.path.exists(pm092_file)
        
        if pm092_exists:
            pm_data = parse_pm_092(pm092_file)
            st.markdown("### 📊 ĐỐI SOÁT CHÍNH THỨC VỚI SỔ CHI TIẾT SCL (PM_092)")
            
            recon_rows = []
            all_recon_match = True
            for proj_code in ["VTAD2606001", "VTAD2606002", "VTAD2605001"]:
                our_sum = df_month_v[df_month_v["project_code"] == proj_code]["amount"].sum()
                pm_proj_data = pm_data.get(proj_code, {})
                pm_sum = pm_proj_data.get("net", 0.0) if pm_proj_data.get("month") == selected_month_num else 0.0
                
                diff = our_sum - pm_sum
                is_match = abs(diff) < 1.0 # Float threshold
                
                recon_rows.append({
                    "Mã công trình": proj_code,
                    "Tên công trình": PROJECT_NAMES.get(proj_code, "Dự án Sửa chữa lớn"),
                    "Số liệu Chương trình (A)": our_sum,
                    "Số liệu Sổ chi tiết PM_092 (B)": pm_sum,
                    "Chênh lệch (A - B)": diff,
                    "Trạng thái": "✅ KHỚP CHÍNH XÁC" if is_match else "❌ LỆCH SỐ LIỆU"
                })
                if not is_match:
                    all_recon_match = False
                    
            df_recon = pd.DataFrame(recon_rows)
            
            # Show a nice alert badge
            if all_recon_match:
                st.success("🎉 **[Khớp chính xác tuyệt đối]**: Số liệu tổng hợp từ 2 file gốc INV-007A và INV-009 khớp chính xác 100% từng đồng với Sổ chi tiết đối tượng tài khoản 2413 (PM_092.xlsx)!")
            else:
                st.warning("⚠️ **[Lệch số liệu]**: Có sự khác biệt giữa Số liệu tổng hợp từ file gốc và Sổ chi tiết tài khoản 2413 (PM_092.xlsx). Vui lòng kiểm tra lại!")
                
            # Render a premium table for reconciliation
            st.dataframe(
                df_recon,
                column_config={
                    "Số liệu Chương trình (A)": st.column_config.NumberColumn("Số liệu Chương trình (A)", format="%,.0f"),
                    "Số liệu Sổ chi tiết PM_092 (B)": st.column_config.NumberColumn("Số liệu Sổ chi tiết PM_092 (B)", format="%,.0f"),
                    "Chênh lệch (A - B)": st.column_config.NumberColumn("Chênh lệch (A - B)", format="%,.0f")
                },
                use_container_width=True,
                hide_index=True
            )
            st.markdown("---")
            
        total_scl_cost = df_month_v["amount"].sum()
        total_pp_share = total_scl_cost * 0.8079
        total_bl_share = total_scl_cost * 0.1921
        
        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card kpi-voltage">
                <div class="kpi-title">Tổng Chi Phí SCL (Tháng {int(selected_month_num)})</div>
                <div class="kpi-value">{total_scl_cost:,.0f} <span style='font-size: 1.1rem;'>VNĐ</span></div>
                <div class="kpi-sub">Tổng chi phí phát sinh trong tháng</div>
            </div>
            <div class="kpi-card kpi-import">
                <div class="kpi-title">Khâu Phân Phối (80.79%)</div>
                <div class="kpi-value">{total_pp_share:,.0f} <span style='font-size: 1.1rem;'>VNĐ</span></div>
                <div class="kpi-sub">TK: 627611-1310-610</div>
            </div>
            <div class="kpi-card kpi-export">
                <div class="kpi-title">Khâu Bán Lẻ (19.21%)</div>
                <div class="kpi-value">{total_bl_share:,.0f} <span style='font-size: 1.1rem;'>VNĐ</span></div>
                <div class="kpi-sub">TK: 627611-1320-610</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Construct monthly preview table matching Excel exactly
        st.markdown(f"### 📋 Xem trước Trang Tính 'Tháng {int(selected_month_num)}'")
        
        preview_rows = []
        stt = 1
        projects_m = sorted(df_month_v["project_code"].unique())
        
        for proj_code in projects_m:
            proj_name = PROJECT_NAMES.get(proj_code, f"{proj_code} - Dự án Sửa chữa lớn")
            df_proj = df_month_v[df_month_v["project_code"] == proj_code]
            
            for v_idx, vol in enumerate(["Trung thế", "Hạ thế"]):
                df_vol = df_proj[df_proj["voltage"] == vol]
                amt = df_vol["amount"].sum() if not df_vol.empty else 0.0
                dist = amt * 0.8079
                retail = amt * 0.1921 # equivalent to E - F
                
                preview_rows.append({
                    "STT": stt if v_idx == 0 else "",
                    "Tên công trình (Chi phí Sửa chữa lớn)": proj_name if v_idx == 0 else "",
                    "Mã CT": proj_code,
                    "Trung-Hạ thế": vol,
                    "Tổng CP (E)": amt,
                    "Khâu phân phối (80.79% - F)": dist,
                    "Khâu bán lẻ (19.21% - G)": retail
                })
            stt += 1
            
        df_preview = pd.DataFrame(preview_rows)
        
        if not df_preview.empty:
            # Append Total Sum Row
            tot_row = {
                "STT": "Tổng cộng",
                "Tên công trình (Chi phí Sửa chữa lớn)": "",
                "Mã CT": "",
                "Trung-Hạ thế": "",
                "Tổng CP (E)": total_scl_cost,
                "Khâu phân phối (80.79% - F)": total_pp_share,
                "Khâu bán lẻ (19.21% - G)": total_bl_share
            }
            df_preview = pd.concat([df_preview, pd.DataFrame([tot_row])], ignore_index=True)
            df_preview["STT"] = df_preview["STT"].astype(str)
            
            st.dataframe(
                df_preview,
                column_config={
                    "Tổng CP (E)": st.column_config.NumberColumn("Tổng CP (E)", format="%,.0f"),
                    "Khâu phân phối (80.79% - F)": st.column_config.NumberColumn("Khâu phân phối (80.79% - F)", format="%,.0f"),
                    "Khâu bán lẻ (19.21% - G)": st.column_config.NumberColumn("Khâu bán lẻ (19.21% - G)", format="%,.0f")
                },
                use_container_width=True,
                height=300
            )
            
        # Download Buttons Side-by-Side
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            if os.path.exists(DEFAULT_OUTPUT_V):
                with open(DEFAULT_OUTPUT_V, "rb") as file:
                    st.download_button(
                        label="📥 Tải Báo Cáo Tách PP-BL (Tháng & Công thức)",
                        data=file,
                        file_name="Tach_ChiPhi_PP_BL.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="tab2_download_btn"
                    )
            else:
                st.error("⚠️ Báo cáo tách PP-BL chưa được tạo.")
        with col_dl2:
            if os.path.exists(DEFAULT_OUTPUT_DETAILED):
                with open(DEFAULT_OUTPUT_DETAILED, "rb") as file:
                    st.download_button(
                        label="📥 Tải Chi Tiết Phân Loại SCL (Từng Dòng)",
                        data=file,
                        file_name="TachPP_BL_ChiTiet.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="tab2_download_detailed_btn"
                    )
            else:
                st.error("⚠️ Báo cáo chi tiết SCL chưa được tạo.")
            
    else:
        st.warning("⚠️ Không tìm thấy giao dịch sửa chữa lớn (SCL) nào trong các file nguồn để phân tách Trung/Hạ thế.")
