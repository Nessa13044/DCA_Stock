import os
import sys
import time
import json
import datetime
import urllib.request
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

# ==============================================================================
# CONFIGURATION - VIETNAM STOCK DCA
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_EXCEL = os.path.join(BASE_DIR, "DCA_100_VN_STOCKS_ANALYSIS.xlsx")
os.makedirs(DATA_DIR, exist_ok=True)

WEEKLY_INVESTMENT_VND = 1_000_000  # 1,000,000 VND / week
TOTAL_WEEKS = 56
TOTAL_INVESTMENT_VND_PER_STOCK = WEEKLY_INVESTMENT_VND * TOTAL_WEEKS   # 56,000,000 VND

START_DATE = datetime.date(2025, 5, 5)
START_TIMESTAMP = int(datetime.datetime(2025, 5, 5, 0, 0, tzinfo=datetime.timezone.utc).timestamp())
END_TIMESTAMP = int(datetime.datetime(2026, 6, 1, 0, 0, tzinfo=datetime.timezone.utc).timestamp())

# ==============================================================================
# CANDIDATE POOLS SPECIFICATION (ORDER & COLOR PALETTE)
# ==============================================================================
POOLS_CONFIG = [
    {
        "pool_id": "VN30",
        "pool_name": "VN30 Index (Nhóm Trụ Cột)",
        "badge_bg": "FEF3C7", "badge_fg": "92400E", "row_tint": "FFFBEB", # Warm Amber
        "tickers": [
            ('VCB', 'HOSE'), ('BID', 'HOSE'), ('CTG', 'HOSE'), ('TCB', 'HOSE'), ('MBB', 'HOSE'),
            ('VPB', 'HOSE'), ('ACB', 'HOSE'), ('HDB', 'HOSE'), ('VIB', 'HOSE'), ('TPB', 'HOSE'),
            ('SHB', 'HOSE'), ('SSB', 'HOSE'), ('LPB', 'HOSE'), ('FPT', 'HOSE'), ('HPG', 'HOSE'),
            ('VNM', 'HOSE'), ('MSN', 'HOSE'), ('MWG', 'HOSE'), ('VIC', 'HOSE'), ('VHM', 'HOSE'),
            ('VRE', 'HOSE'), ('GAS', 'HOSE'), ('PLX', 'HOSE'), ('POW', 'HOSE'), ('SAB', 'HOSE'),
            ('VJC', 'HOSE'), ('BCM', 'HOSE'), ('BVH', 'HOSE'), ('GVR', 'HOSE'), ('SSI', 'HOSE')
        ]
    },
    {
        "pool_id": "BANK_FIN",
        "pool_name": "Ngân Hàng & Tài Chính",
        "badge_bg": "DBEAFE", "badge_fg": "1E40AF", "row_tint": "EFF6FF", # Royal Blue
        "tickers": [
            ('MSB', 'HOSE'), ('OCB', 'HOSE'), ('EIB', 'HOSE'), ('NAB', 'HOSE'), ('BVB', 'UPCOM'),
            ('BAB', 'HNX'), ('VND', 'HOSE'), ('VCI', 'HOSE'), ('HCM', 'HOSE'), ('SHS', 'HNX'),
            ('MBS', 'HNX'), ('FTS', 'HOSE'), ('BSI', 'HOSE'), ('CTS', 'HOSE'), ('AGR', 'HOSE'),
            ('VDS', 'HOSE'), ('ORS', 'HOSE'), ('VIX', 'HOSE')
        ]
    },
    {
        "pool_id": "REAL_ESTATE",
        "pool_name": "Bất Động Sản & KCN",
        "badge_bg": "FFE4E6", "badge_fg": "9F1239", "row_tint": "FFF1F2", # Rose
        "tickers": [
            ('KDH', 'HOSE'), ('NLG', 'HOSE'), ('DXG', 'HOSE'), ('PDR', 'HOSE'), ('DIG', 'HOSE'),
            ('KBC', 'HOSE'), ('NVL', 'HOSE'), ('IDC', 'HNX'), ('CEO', 'HNX'), ('TCH', 'HOSE'),
            ('HDG', 'HOSE'), ('CII', 'HOSE'), ('SZC', 'HOSE'), ('NHA', 'HOSE'), ('DXS', 'HOSE'),
            ('HQC', 'HOSE'), ('SCR', 'HOSE'), ('HDC', 'HOSE'), ('KHG', 'HOSE'), ('CRE', 'HOSE')
        ]
    },
    {
        "pool_id": "STEEL_INFRA",
        "pool_name": "Thép, Vật Liệu & Xây Dựng",
        "badge_bg": "E2E8F0", "badge_fg": "1E293B", "row_tint": "F8FAFC", # Slate Gray
        "tickers": [
            ('HSG', 'HOSE'), ('NKG', 'HOSE'), ('VGS', 'HNX'), ('VCS', 'HNX'), ('HT1', 'HOSE'),
            ('BCC', 'HNX'), ('VCG', 'HOSE'), ('HHV', 'HOSE'), ('FCN', 'HOSE'), ('CTD', 'HOSE'),
            ('LCG', 'HOSE'), ('C4G', 'UPCOM'), ('KSB', 'HOSE'), ('PC1', 'HOSE'), ('SMC', 'HOSE')
        ]
    },
    {
        "pool_id": "CHEMICAL_RUBBER",
        "pool_name": "Hóa Chất, Phân Bón & Cao Su",
        "badge_bg": "D1FAE5", "badge_fg": "065F46", "row_tint": "ECFDF5", # Emerald Green
        "tickers": [
            ('DGC', 'HOSE'), ('DCM', 'HOSE'), ('DPM', 'HOSE'), ('CSV', 'HOSE'), ('BFC', 'HOSE'),
            ('DPR', 'HOSE'), ('PHR', 'HOSE')
        ]
    },
    {
        "pool_id": "RETAIL_CONSUMER",
        "pool_name": "Bán Lẻ, Tiêu Dùng & Thực Phẩm",
        "badge_bg": "FFEDD5", "badge_fg": "9A3412", "row_tint": "FFF7ED", # Warm Orange
        "tickers": [
            ('FRT', 'HOSE'), ('DGW', 'HOSE'), ('PNJ', 'HOSE'), ('KDC', 'HOSE'), ('VHC', 'HOSE'),
            ('ANV', 'HOSE'), ('DBC', 'HOSE'), ('BAF', 'HOSE'), ('PAN', 'HOSE'), ('HAG', 'HOSE'),
            ('HNG', 'HOSE'), ('IDI', 'HOSE'), ('ASM', 'HOSE'), ('TNG', 'HNX'), ('TCM', 'HOSE')
        ]
    },
    {
        "pool_id": "TECH_ENERGY",
        "pool_name": "Công Nghệ, Năng Lượng & Tiện Ích",
        "badge_bg": "E0F2FE", "badge_fg": "0369A1", "row_tint": "F0F9FF", # Sky Blue
        "tickers": [
            ('CMG', 'HOSE'), ('CTR', 'HOSE'), ('ELC', 'HOSE'), ('REE', 'HOSE'), ('GEG', 'HOSE'),
            ('NT2', 'HOSE'), ('QTP', 'UPCOM'), ('HND', 'UPCOM')
        ]
    },
    {
        "pool_id": "OIL_LOGISTICS",
        "pool_name": "Dầu Khí, Cảng Biển & Logistics",
        "badge_bg": "EDE9FE", "badge_fg": "5B21B6", "row_tint": "F5F3FF", # Lilac Purple
        "tickers": [
            ('PVD', 'HOSE'), ('PVS', 'HNX'), ('PVT', 'HOSE'), ('BSR', 'UPCOM'), ('OIL', 'UPCOM'),
            ('GMD', 'HOSE'), ('HAH', 'HOSE'), ('VSC', 'HOSE'), ('HVN', 'HOSE'), ('ACV', 'UPCOM'),
            ('PVB', 'HNX'), ('PVC', 'HNX'), ('SGP', 'UPCOM')
        ]
    }
]

# Map ticker to pool configuration
TICKER_TO_POOL = {}
for p in POOLS_CONFIG:
    for t, exch in p['tickers']:
        if t not in TICKER_TO_POOL:
            TICKER_TO_POOL[t] = {
                "pool_id": p['pool_id'],
                "pool_name": p['pool_name'],
                "badge_bg": p['badge_bg'],
                "badge_fg": p['badge_fg'],
                "row_tint": p['row_tint']
            }

def generate_week_labels():
    labels = []
    for i in range(TOTAL_WEEKS):
        w_start = START_DATE + datetime.timedelta(days=i * 7)
        w_end = w_start + datetime.timedelta(days=4)
        label_full = f"Tuần {i+1} ({w_start.strftime('%d/%m/%y')} - {w_end.strftime('%d/%m/%y')})"
        label_short = f"W{i+1:02d} ({w_start.strftime('%d/%m')})"
        labels.append((label_full, label_short, w_start, w_end))
    return labels

def load_cached_or_fetch_stocks():
    cache_file = os.path.join(DATA_DIR, "cached_vn_stocks_raw.json")
    if os.path.exists(cache_file):
        print(f">>> Phát hiện cache tại {cache_file}, đang nạp dữ liệu...")
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)
    print(">>> Không tìm thấy cache, vui lòng kiểm tra lại...")
    return []

def calculate_and_group_metrics(raw_data):
    print(">>> Đang tính toán ma trận DCA và phân loại theo Candidate Pool...")
    
    # 1. Compute metrics for all stocks
    temp_results = []
    for item in raw_data:
        ticker = item['ticker']
        exch = item['exchange']
        prices = item['prices']
        
        shares_bought = [WEEKLY_INVESTMENT_VND / p for p in prices]
        total_shares = sum(shares_bought)
        final_price = prices[-1]
        avg_dca_price = TOTAL_INVESTMENT_VND_PER_STOCK / total_shares
        final_value_vnd = total_shares * final_price
        net_profit_vnd = final_value_vnd - TOTAL_INVESTMENT_VND_PER_STOCK
        roi_pct = (net_profit_vnd / TOTAL_INVESTMENT_VND_PER_STOCK) * 100
        
        pool_info = TICKER_TO_POOL.get(ticker, {
            "pool_id": "OTHER",
            "pool_name": "Nhóm Khác",
            "badge_bg": "F1F5F9",
            "badge_fg": "334155",
            "row_tint": "FFFFFF"
        })
        
        temp_results.append({
            "ticker": ticker,
            "exchange": exch,
            "prices": prices,
            "shares_bought": shares_bought,
            "total_shares": total_shares,
            "avg_dca_price": avg_dca_price,
            "final_price": final_price,
            "total_invested_vnd": TOTAL_INVESTMENT_VND_PER_STOCK,
            "final_value_vnd": final_value_vnd,
            "net_profit_vnd": net_profit_vnd,
            "roi_pct": roi_pct,
            "status": "LÃI" if net_profit_vnd >= 0 else "LỖ",
            "pool_id": pool_info["pool_id"],
            "pool_name": pool_info["pool_name"],
            "badge_bg": pool_info["badge_bg"],
            "badge_fg": pool_info["badge_fg"],
            "row_tint": pool_info["row_tint"]
        })
        
    # 2. Determine Overall Rank (Rank 1 to 100 by ROI)
    temp_results.sort(key=lambda x: x['roi_pct'], reverse=True)
    for overall_idx, r in enumerate(temp_results, 1):
        r['overall_rank'] = overall_idx

    # 3. Sort strictly by Candidate Pool order, and by ROI descending within each pool
    pool_order_map = {p['pool_id']: idx for idx, p in enumerate(POOLS_CONFIG)}
    temp_results.sort(key=lambda x: (pool_order_map.get(x['pool_id'], 999), -x['roi_pct']))
    
    # Assign sequential display order
    for idx, r in enumerate(temp_results, 1):
        r['stt'] = idx
        
    return temp_results

def create_excel_workbook(results):
    print(">>> Đang khởi tạo file Excel chuyên nghiệp (Sắp xếp & Tô màu theo Candidate Pool)...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    week_info = generate_week_labels()
    
    # Typography
    font_family = "Segoe UI"
    font_title = Font(name=font_family, size=16, bold=True, color="FFFFFF")
    font_subtitle = Font(name=font_family, size=10, italic=True, color="E0E0E0")
    font_kpi_num = Font(name=font_family, size=14, bold=True, color="1E293B")
    font_kpi_label = Font(name=font_family, size=9, bold=True, color="64748B")
    font_tbl_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_data = Font(name=font_family, size=10, color="000000")
    font_data_bold = Font(name=font_family, size=10, bold=True, color="000000")
    font_profit = Font(name=font_family, size=10, bold=True, color="155724")
    font_loss = Font(name=font_family, size=10, bold=True, color="721C24")
    
    fill_dark_navy = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid") # Slate 900
    fill_slate = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")     # Slate 800
    fill_kpi_bg = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_green_soft = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
    fill_red_soft = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    fill_summary = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_highlight_shares = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    thick_bottom_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='medium', color='0F172A')
    )
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # ==========================================================================
    # SHEET 1: DCA_Summary_Dashboard
    # ==========================================================================
    ws_dash = wb.create_sheet(title="DCA_Summary_Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True
    
    # Title Banner (A1:M1)
    ws_dash.merge_cells("A1:M1")
    ws_dash["A1"] = "BẢNG PHÂN TÍCH ĐỊNH LƯỢNG HIỆU QUẢ DCA 100 CỔ PHIẾU VIỆT NAM (SẮP XẾP & TÔ MÀU THEO CANDIDATE POOL)"
    ws_dash["A1"].font = font_title
    ws_dash["A1"].fill = fill_dark_navy
    ws_dash["A1"].alignment = align_center
    ws_dash.row_dimensions[1].height = 40
    
    ws_dash.merge_cells("A2:M2")
    ws_dash["A2"] = "Chiến lược: 1.000.000 VNĐ/tuần/mã | Chu kỳ: 56 tuần (05/2025 - 05/2026) | Tổng vốn: 56.000.000 VNĐ/mã | Tổng quy mô rổ: 5,6 tỷ VNĐ | Phân loại & màu sắc theo 8 nhóm ngành"
    ws_dash["A2"].font = font_subtitle
    ws_dash["A2"].fill = fill_slate
    ws_dash["A2"].alignment = align_center
    ws_dash.row_dimensions[2].height = 24

    # Portfolio Aggregates
    total_portfolio_invested = len(results) * TOTAL_INVESTMENT_VND_PER_STOCK
    total_portfolio_final = sum(r['final_value_vnd'] for r in results)
    total_portfolio_profit = total_portfolio_final - total_portfolio_invested
    portfolio_roi = (total_portfolio_profit / total_portfolio_invested) * 100
    win_count = sum(1 for r in results if r['net_profit_vnd'] >= 0)
    loss_count = len(results) - win_count
    best_performer = min(results, key=lambda x: x['overall_rank']) # Rank 1 overall

    # KPI Cards (Row 4 & 5 across A to M)
    kpi_cards = [
        ("A4:B4", "A5:B5", "TỔNG VỐN (100 CỔ PHIẾU)", f"{total_portfolio_invested:,.0f} ₫", "A4", "A5"),
        ("C4:D4", "C5:D5", "TỔNG GIÁ TRỊ CUỐI KỲ", f"{total_portfolio_final:,.0f} ₫", "C4", "C5"),
        ("E4:F4", "E5:F5", "LỢI NHUẬN RÒNG DANH MỤC", f"{total_portfolio_profit:+,.0f} ₫", "E4", "E5"),
        ("G4:H4", "G5:H5", "ROI TOÀN DANH MỤC", f"{portfolio_roi:+.2f}%", "G4", "G5"),
        ("I4:J4", "I5:J5", "TỶ LỆ LÃI / LỖ", f"{win_count} Lãi / {loss_count} Lỗ", "I4", "I5"),
        ("K4:M4", "K5:M5", "QUÁN QUÂN TĂNG TRƯỞNG", f"{best_performer['ticker']} ({best_performer['roi_pct']:+.1f}%)", "K4", "K5"),
    ]
    
    ws_dash.row_dimensions[4].height = 18
    ws_dash.row_dimensions[5].height = 26
    
    for title_range, val_range, title, val, top_left_t, top_left_v in kpi_cards:
        ws_dash.merge_cells(title_range)
        ws_dash.merge_cells(val_range)
        cell_t = ws_dash[top_left_t]
        cell_t.value = title
        cell_t.font = font_kpi_label
        cell_t.fill = fill_kpi_bg
        cell_t.alignment = align_center
        
        cell_v = ws_dash[top_left_v]
        cell_v.value = val
        cell_v.font = font_kpi_num
        cell_v.fill = fill_kpi_bg
        cell_v.alignment = align_center
        
        start_col_str = title_range.split(":")[0][0]
        end_col_str = title_range.split(":")[1][0] if ":" in title_range else start_col_str
        start_num = column_index_from_string(start_col_str)
        end_num = column_index_from_string(end_col_str)
        for r_idx in (4, 5):
            for c_i in range(start_num, end_num + 1):
                ws_dash.cell(row=r_idx, column=c_i).border = thin_border

    # Headers for Summary Table (Row 7)
    dash_headers = [
        ("STT", 8),
        ("Mã Cổ Phiếu", 14),
        ("Nhóm Ngành (Candidate Pool)", 30),
        ("Sàn GD", 12),
        ("Hạng ROI Tổng", 15),
        ("Tổng Vốn (VNĐ)", 18),
        ("Giá DCA TB (VNĐ)", 18),
        ("Giá Tuần 56 (VNĐ)", 18),
        ("Tổng CP Tích Lũy", 18),
        ("Giá Trị Tuần 56 (VNĐ)", 22),
        ("Lãi / Lỗ Ròng (VNĐ)", 22),
        ("Tỷ Suất ROI (%)", 16),
        ("Trạng Thái", 12)
    ]
    
    ws_dash.row_dimensions[7].height = 28
    for col_idx, (h_name, col_w) in enumerate(dash_headers, 1):
        cell = ws_dash.cell(row=7, column=col_idx, value=h_name)
        cell.font = font_tbl_header
        cell.fill = fill_slate
        cell.alignment = align_center
        cell.border = thick_bottom_border
        col_letter = get_column_letter(col_idx)
        ws_dash.column_dimensions[col_letter].width = col_w

    start_row = 8
    for idx, r in enumerate(results):
        row_num = start_row + idx
        ws_dash.row_dimensions[row_num].height = 21
        
        # Pool styles
        pool_row_fill = PatternFill(start_color=r['row_tint'], end_color=r['row_tint'], fill_type="solid")
        pool_badge_fill = PatternFill(start_color=r['badge_bg'], end_color=r['badge_bg'], fill_type="solid")
        pool_badge_font = Font(name=font_family, size=10, bold=True, color=r['badge_fg'])
        
        is_profit = r['net_profit_vnd'] >= 0
        
        c1 = ws_dash.cell(row=row_num, column=1, value=r['stt'])
        c1.alignment = align_center
        
        c2 = ws_dash.cell(row=row_num, column=2, value=r['ticker'])
        c2.alignment = align_center
        c2.font = font_data_bold
        
        c3 = ws_dash.cell(row=row_num, column=3, value=r['pool_name'])
        c3.alignment = align_left
        c3.font = pool_badge_font
        c3.fill = pool_badge_fill
        
        c4 = ws_dash.cell(row=row_num, column=4, value=r['exchange'])
        c4.alignment = align_center
        
        c5 = ws_dash.cell(row=row_num, column=5, value=f"#{r['overall_rank']}")
        c5.alignment = align_center
        c5.font = font_data_bold
        
        c6 = ws_dash.cell(row=row_num, column=6, value=r['total_invested_vnd'])
        c6.alignment = align_right
        c6.number_format = "#,##0 ₫"
        
        c7 = ws_dash.cell(row=row_num, column=7, value=r['avg_dca_price'])
        c7.alignment = align_right
        c7.number_format = "#,##0 ₫"
        
        c8 = ws_dash.cell(row=row_num, column=8, value=r['final_price'])
        c8.alignment = align_right
        c8.number_format = "#,##0 ₫"
        
        c9 = ws_dash.cell(row=row_num, column=9, value=r['total_shares'])
        c9.alignment = align_right
        c9.number_format = "#,##0.00"
        
        c10 = ws_dash.cell(row=row_num, column=10, value=r['final_value_vnd'])
        c10.alignment = align_right
        c10.number_format = "#,##0 ₫"
        
        c11 = ws_dash.cell(row=row_num, column=11, value=r['net_profit_vnd'])
        c11.alignment = align_right
        c11.number_format = "+#,##0 ₫;-#,##0 ₫;0 ₫"
        c11.font = font_profit if is_profit else font_loss
        c11.fill = fill_green_soft if is_profit else fill_red_soft
        
        c12 = ws_dash.cell(row=row_num, column=12, value=r['roi_pct'] / 100.0)
        c12.alignment = align_right
        c12.number_format = "+0.00%;-0.00%;0.00%"
        c12.font = font_profit if is_profit else font_loss
        c12.fill = fill_green_soft if is_profit else fill_red_soft
        
        c13 = ws_dash.cell(row=row_num, column=13, value=r['status'])
        c13.alignment = align_center
        c13.font = font_profit if is_profit else font_loss
        c13.fill = fill_green_soft if is_profit else fill_red_soft
        
        for c in (c1, c2, c4, c5, c6, c7, c8, c9, c10):
            c.font = font_data_bold if c in (c2, c5, c10) else font_data
            c.fill = pool_row_fill
            c.border = thin_border
            
        c3.border = thin_border
        for c in (c11, c12, c13):
            c.border = thin_border

    # Bottom Total Row
    tot_row = start_row + len(results)
    ws_dash.row_dimensions[tot_row].height = 25
    ws_dash.cell(row=tot_row, column=1, value="").fill = fill_summary
    ws_dash.cell(row=tot_row, column=2, value="TỔNG CỘNG").fill = fill_summary
    ws_dash.cell(row=tot_row, column=2).font = font_data_bold
    ws_dash.cell(row=tot_row, column=2).alignment = align_center
    ws_dash.cell(row=tot_row, column=3, value=f"{len(POOLS_CONFIG)} Nhóm Ngành").fill = fill_summary
    ws_dash.cell(row=tot_row, column=3).alignment = align_center
    ws_dash.cell(row=tot_row, column=3).font = font_data_bold
    ws_dash.cell(row=tot_row, column=4, value="100 Mã").fill = fill_summary
    ws_dash.cell(row=tot_row, column=4).alignment = align_center
    ws_dash.cell(row=tot_row, column=5, value="-").fill = fill_summary
    ws_dash.cell(row=tot_row, column=5).alignment = align_center
    
    c_tot_inv = ws_dash.cell(row=tot_row, column=6, value=f"=SUM(F{start_row}:F{tot_row-1})")
    c_tot_inv.number_format = "#,##0 ₫"
    c_tot_inv.font = font_data_bold
    c_tot_inv.fill = fill_summary
    c_tot_inv.alignment = align_right
    
    ws_dash.cell(row=tot_row, column=7, value="-").fill = fill_summary
    ws_dash.cell(row=tot_row, column=7).alignment = align_center
    ws_dash.cell(row=tot_row, column=8, value="-").fill = fill_summary
    ws_dash.cell(row=tot_row, column=8).alignment = align_center
    ws_dash.cell(row=tot_row, column=9, value="-").fill = fill_summary
    ws_dash.cell(row=tot_row, column=9).alignment = align_center
    
    c_tot_val = ws_dash.cell(row=tot_row, column=10, value=f"=SUM(J{start_row}:J{tot_row-1})")
    c_tot_val.number_format = "#,##0 ₫"
    c_tot_val.font = font_data_bold
    c_tot_val.fill = fill_summary
    c_tot_val.alignment = align_right
    
    c_tot_pnl = ws_dash.cell(row=tot_row, column=11, value=f"=SUM(K{start_row}:K{tot_row-1})")
    c_tot_pnl.number_format = "+#,##0 ₫;-#,##0 ₫;0 ₫"
    c_tot_pnl.font = font_profit if total_portfolio_profit >= 0 else font_loss
    c_tot_pnl.fill = fill_green_soft if total_portfolio_profit >= 0 else fill_red_soft
    c_tot_pnl.alignment = align_right
    
    c_tot_roi = ws_dash.cell(row=tot_row, column=12, value=f"=K{tot_row}/F{tot_row}")
    c_tot_roi.number_format = "+0.00%;-0.00%;0.00%"
    c_tot_roi.font = font_profit if total_portfolio_profit >= 0 else font_loss
    c_tot_roi.fill = fill_green_soft if total_portfolio_profit >= 0 else fill_red_soft
    c_tot_roi.alignment = align_right
    
    c_tot_st = ws_dash.cell(row=tot_row, column=13, value="DANH MỤC")
    c_tot_st.alignment = align_center
    c_tot_st.fill = fill_summary
    c_tot_st.font = font_data_bold

    for col_c in range(1, 14):
        ws_dash.cell(row=tot_row, column=col_c).border = thick_bottom_border

    # ==========================================================================
    # SHEET 2: Weekly_Prices_VND (100 mã x 56 tuần - Sắp xếp theo Pool)
    # ==========================================================================
    ws_price = wb.create_sheet(title="Weekly_Prices_VND")
    ws_price.views.sheetView[0].showGridLines = True
    
    price_static_headers = [("STT", 8), ("Mã Cổ Phiếu", 14), ("Nhóm Ngành", 28), ("Sàn GD", 12)]
    ws_price.row_dimensions[1].height = 28
    
    for c_i, (h_txt, w_val) in enumerate(price_static_headers, 1):
        c = ws_price.cell(row=1, column=c_i, value=h_txt)
        c.font = font_tbl_header
        c.fill = fill_slate
        c.alignment = align_center
        c.border = thick_bottom_border
        ws_price.column_dimensions[get_column_letter(c_i)].width = w_val
        
    for w_i, (_, label_short, _, _) in enumerate(week_info, 1):
        c_idx = 4 + w_i
        c = ws_price.cell(row=1, column=c_idx, value=label_short)
        c.font = font_tbl_header
        c.fill = fill_dark_navy
        c.alignment = align_center
        c.border = thick_bottom_border
        ws_price.column_dimensions[get_column_letter(c_idx)].width = 14
        
    c_min = ws_price.cell(row=1, column=61, value="Giá Thấp Nhất")
    c_max = ws_price.cell(row=1, column=62, value="Giá Cao Nhất")
    c_avg = ws_price.cell(row=1, column=63, value="Giá TB (VNĐ)")
    for c in (c_min, c_max, c_avg):
        c.font = font_tbl_header
        c.fill = fill_slate
        c.alignment = align_center
        c.border = thick_bottom_border
    ws_price.column_dimensions["BI"].width = 16
    ws_price.column_dimensions["BJ"].width = 16
    ws_price.column_dimensions["BK"].width = 16

    for idx, r in enumerate(results):
        r_num = 2 + idx
        ws_price.row_dimensions[r_num].height = 19
        pool_row_fill = PatternFill(start_color=r['row_tint'], end_color=r['row_tint'], fill_type="solid")
        pool_badge_fill = PatternFill(start_color=r['badge_bg'], end_color=r['badge_bg'], fill_type="solid")
        pool_badge_font = Font(name=font_family, size=10, bold=True, color=r['badge_fg'])
        
        c1 = ws_price.cell(row=r_num, column=1, value=r['stt'])
        c2 = ws_price.cell(row=r_num, column=2, value=r['ticker'])
        c3 = ws_price.cell(row=r_num, column=3, value=r['pool_name'])
        c4 = ws_price.cell(row=r_num, column=4, value=r['exchange'])
        
        c1.alignment = align_center
        c2.alignment = align_center
        c2.font = font_data_bold
        c3.alignment = align_left
        c3.font = pool_badge_font
        c3.fill = pool_badge_fill
        c4.alignment = align_center
        
        for c in (c1, c2, c4):
            c.fill = pool_row_fill
            c.border = thin_border
        c3.border = thin_border
            
        for w_i, p_val in enumerate(r['prices'], 1):
            c_p = ws_price.cell(row=r_num, column=4 + w_i, value=p_val)
            c_p.alignment = align_right
            c_p.font = font_data
            c_p.number_format = "#,##0 ₫"
            c_p.fill = pool_row_fill
            c_p.border = thin_border
            
        first_col_let = get_column_letter(5)
        last_col_let = get_column_letter(60)
        c_min = ws_price.cell(row=r_num, column=61, value=f"=MIN({first_col_let}{r_num}:{last_col_let}{r_num})")
        c_max = ws_price.cell(row=r_num, column=62, value=f"=MAX({first_col_let}{r_num}:{last_col_let}{r_num})")
        c_avg = ws_price.cell(row=r_num, column=63, value=f"=AVERAGE({first_col_let}{r_num}:{last_col_let}{r_num})")
        for c in (c_min, c_max, c_avg):
            c.alignment = align_right
            c.font = font_data_bold
            c.number_format = "#,##0 ₫"
            c.fill = pool_row_fill
            c.border = thin_border

    # ==========================================================================
    # SHEET 3: Weekly_Shares_Bought (Số cổ phiếu mua mỗi tuần với 1 triệu VNĐ)
    # ==========================================================================
    ws_shares = wb.create_sheet(title="Weekly_Shares_Bought")
    ws_shares.views.sheetView[0].showGridLines = True
    
    ws_shares.row_dimensions[1].height = 28
    for c_i, (h_txt, w_val) in enumerate(price_static_headers, 1):
        c = ws_shares.cell(row=1, column=c_i, value=h_txt)
        c.font = font_tbl_header
        c.fill = fill_slate
        c.alignment = align_center
        c.border = thick_bottom_border
        ws_shares.column_dimensions[get_column_letter(c_i)].width = w_val
        
    for w_i, (_, label_short, _, _) in enumerate(week_info, 1):
        c_idx = 4 + w_i
        c = ws_shares.cell(row=1, column=c_idx, value=label_short)
        c.font = font_tbl_header
        c.fill = fill_dark_navy
        c.alignment = align_center
        c.border = thick_bottom_border
        ws_shares.column_dimensions[get_column_letter(c_idx)].width = 14
        
    # Cột 61 (Col BI): TỔNG CỔ PHIẾU TÍCH LŨY
    c_tot_header = ws_shares.cell(row=1, column=61, value="TỔNG CP TÍCH LŨY")
    c_tot_header.font = font_tbl_header
    c_tot_header.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    c_tot_header.alignment = align_center
    c_tot_header.border = thick_bottom_border
    ws_shares.column_dimensions["BI"].width = 20
    
    # Cột 62 (Col BJ): GIÁ TRỊ TUẦN 56 (VNĐ)
    c_val_header = ws_shares.cell(row=1, column=62, value="GIÁ TRỊ TUẦN 56 (VNĐ)")
    c_val_header.font = font_tbl_header
    c_val_header.fill = PatternFill(start_color="064E3B", end_color="064E3B", fill_type="solid")
    c_val_header.alignment = align_center
    c_val_header.border = thick_bottom_border
    ws_shares.column_dimensions["BJ"].width = 22

    for idx, r in enumerate(results):
        r_num = 2 + idx
        ws_shares.row_dimensions[r_num].height = 19
        pool_row_fill = PatternFill(start_color=r['row_tint'], end_color=r['row_tint'], fill_type="solid")
        pool_badge_fill = PatternFill(start_color=r['badge_bg'], end_color=r['badge_bg'], fill_type="solid")
        pool_badge_font = Font(name=font_family, size=10, bold=True, color=r['badge_fg'])
        
        c1 = ws_shares.cell(row=r_num, column=1, value=r['stt'])
        c2 = ws_shares.cell(row=r_num, column=2, value=r['ticker'])
        c3 = ws_shares.cell(row=r_num, column=3, value=r['pool_name'])
        c4 = ws_shares.cell(row=r_num, column=4, value=r['exchange'])
        
        c1.alignment = align_center
        c2.alignment = align_center
        c2.font = font_data_bold
        c3.alignment = align_left
        c3.font = pool_badge_font
        c3.fill = pool_badge_fill
        c4.alignment = align_center
        
        for c in (c1, c2, c4):
            c.fill = pool_row_fill
            c.border = thin_border
        c3.border = thin_border
            
        for w_i, s_val in enumerate(r['shares_bought'], 1):
            c_s = ws_shares.cell(row=r_num, column=4 + w_i, value=s_val)
            c_s.alignment = align_right
            c_s.font = font_data
            c_s.number_format = "#,##0.00"
            c_s.fill = pool_row_fill
            c_s.border = thin_border
            
        first_s_col = get_column_letter(5)
        last_s_col = get_column_letter(60)
        c_tot = ws_shares.cell(row=r_num, column=61, value=f"=SUM({first_s_col}{r_num}:{last_s_col}{r_num})")
        c_tot.alignment = align_right
        c_tot.font = font_data_bold
        c_tot.number_format = "#,##0.00"
        c_tot.fill = fill_highlight_shares
        c_tot.border = thin_border
        
        # Formula: Total Shares (BI) * Week 56 price from Weekly_Prices_VND!BH{r}
        c_val = ws_shares.cell(row=r_num, column=62, value=f"=BI{r_num}*Weekly_Prices_VND!BH{r_num}")
        c_val.alignment = align_right
        c_val.font = font_data_bold
        c_val.number_format = "#,##0 ₫"
        c_val.fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
        c_val.border = thin_border

    wb.save(OUTPUT_EXCEL)
    print(f"\n>>> Đã xuất thành công file Excel chất lượng cao: {OUTPUT_EXCEL}")

def export_csv_summary(results):
    summary_df = pd.DataFrame([{
        "STT": r['stt'],
        "Ticker": r['ticker'],
        "Pool_ID": r['pool_id'],
        "Candidate_Pool": r['pool_name'],
        "Exchange": r['exchange'],
        "Overall_Rank": r['overall_rank'],
        "Total_Invested_VND": r['total_invested_vnd'],
        "Avg_DCA_Price_VND": round(r['avg_dca_price'], 2),
        "Final_Price_VND": round(r['final_price'], 2),
        "Total_Shares_Accumulated": round(r['total_shares'], 2),
        "Final_Value_VND": round(r['final_value_vnd'], 0),
        "Net_Profit_VND": round(r['net_profit_vnd'], 0),
        "ROI_Pct": round(r['roi_pct'], 2),
        "Status": r['status']
    } for r in results])
    
    csv_path = os.path.join(DATA_DIR, "dca_100_vn_stocks_summary.csv")
    summary_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f">>> Đã xuất file CSV tổng kết: {csv_path}")

def main():
    start_time = time.time()
    print("==================================================================================")
    print("CHƯƠNG TRÌNH PHÂN TÍCH ĐỊNH LƯỢNG DCA 100 CỔ PHIẾU VN (SẮP XẾP & TÔ MÀU THEO POOL)")
    print("==================================================================================")
    raw_data = load_cached_or_fetch_stocks()
    results = calculate_and_group_metrics(raw_data)
    create_excel_workbook(results)
    export_csv_summary(results)
    elapsed = time.time() - start_time
    print(f">>> Hoàn thành cập nhật toàn bộ trang tính trong {elapsed:.1f} giây!")

if __name__ == "__main__":
    main()
