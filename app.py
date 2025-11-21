import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os
import urllib.request
import re
from datetime import datetime, timedelta
import io

# ==========================================
# 0. 全域設定：AI 顧問指令 (完整修復版)
# ==========================================
AI_CONSULTANT_PROMPT = """
# ⚠️ SYSTEM OVERRIDE: DATA LOADING INSTRUCTION
**To the AI Analysis Model:**
The user has uploaded a **Single-Sheet Excel File**.
1. **ALL DATA** is contained in the **First Sheet** named '📘_完整分析報告'.
2. The content is organized as **Vertically Stacked Tables**.
3. The structure is:
   - **[Top Section]**: This Instruction (Prompt).
   - **[Middle Section]**: Q13_Trend Data (Daily Trend).
   - **[Bottom Section]**: Consolidated Data Tables for P7D, PP7D, and P30D (Campaign/AdSet/Ad levels).
4. **ACTION**: Please read the entire sheet. Scan for headers like "Table: ..." to identify different datasets.

---

# Role
你是一位擁有 10 年經驗的資深成效廣告分析師。請根據本頁面中的所有數據進行帳戶健檢。

# Data Structure & Sorting Logic
- **Q13_Trend**: 依日期排序的每日趨勢。
- **Consolidated Tables (P7D/PP7D/P30D)**:
    - 這些表格預設 **「依花費金額 (Spend) 由高到低排名」**。
    - **分析重點**: 請優先關注排名前 3-5 名的「高花費項目」，它們對整體帳戶影響最大。
    - 表格最後一列通常是 **「全帳戶平均 (Account Average)」**，請以此作為基準線 (Benchmark)。

# Analysis Requirements

## 1. 波動偵測 (Fluctuation Analysis)
- **全站體檢**: 優先查看上方 `Q13_Trend` 表格中的 **「🏆 整體帳戶」** 趨勢線，判斷整體 CVR 與 CPA 走勢。
- **細項對比**: 往下捲動，找到 **P7D (本週)** 與 **PP7D (上週)** 的表格進行環比分析。
- 找出 CPA 暴漲或 CVR 驟降的「警示區」。

## 2. 擴量機會 (Scaling)
- 找出 **CPA 低且穩定** 的行銷活動/廣告組合 -> 建議加碼。
- 找出 **High CTR / Low Spend** 的潛力素材 -> 建議給予獨立預算。
- 找出 **High CTR / Low CVR** 的項目 -> 建議優化落地頁。

## 3. 止損建議 (Cost Cutting)
- 找出 **高花費 but 0 轉換** 的項目。
- 找出 **CPA 過高且 CTR 低落** 的無效廣告。

## 4. 綜合戰術行動清單 (Action Plan)
請列出具體的：
- **🔴 應關閉**: 具體列出該關閉的素材/受眾名稱。
- **🟢 應加強**: 具體列出該加碼的項目。
- **💰 預算調整**: 具體的預算增減建議。
- **🎨 素材/網頁優化**: 下一步該做什麼圖？該改什麼文案？

# Output Format
請輸出專業分析報告，並確保「戰術行動清單」清晰可執行。
"""

# ==========================================
# 1. 基礎設定與字型處理
# ==========================================
st.set_page_config(page_title="廣告成效全能分析 v4.1", layout="wide")

@st.cache_resource
def get_chinese_font():
    font_path = "NotoSansCJKtc-Regular.otf"
    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            with st.spinner('正在下載中文字型 (首次執行需時較久)...'):
                urllib.request.urlretrieve(url, font_path)
        except:
            return None
    return fm.FontProperties(fname=font_path)

font_prop = get_chinese_font()

# ==========================================
# 2. 資料處理核心函數
# ==========================================

def clean_ad_name(name):
    return re.sub(r' - 複本.*$', '', str(name)).strip()

def create_summary_row(df, metric_cols):
    summary_dict = {}
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        summary_dict[col] = df[col].sum()
        
    for metric, (num, denom, is_pct) in metric_cols.items():
        total_num = summary_dict.get(num, 0)
        total_denom = summary_dict.get(denom, 0)
        if total_denom > 0:
            val = (total_num / total_denom)
            if is_pct: val *= 100
            summary_dict[metric] = round(val, 2)
        else:
            summary_dict[metric] = 0

    non_numeric_cols = df.select_dtypes(exclude=[np.number]).columns
    if len(non_numeric_cols) > 0:
        summary_dict[non_numeric_cols[0]] = '全帳戶平均'
        for col in non_numeric_cols[1:]:
            summary_dict[col] = '-'
    return pd.DataFrame([summary_dict])

def calculate_metrics_consolidated(df_group, conv_col):
    # 聚合
    df_metrics = df_group.agg({
        '花費金額 (TWD)': 'sum',
        conv_col: 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    }).reset_index()
    
    df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]
    
    # 計算指標
    df_metrics['CPA'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0, axis=1)
    df_metrics['CTR (%)'] = df_metrics.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    df_metrics['CVR (%)'] = df_metrics.apply(lambda x: (x[conv_col] / x['連結點擊次數']) * 100 if x['連結點擊次數'] > 0 else 0, axis=1)
    
    df_metrics = df_metrics.round(2).sort_values(by='花費金額 (TWD)', ascending=False)
    
    # 平均列
    metric_config = {
        'CPA': ('花費金額 (TWD)', conv_col, False),
        'CTR (%)': ('連結點擊次數', '曝光次數', True),
        'CVR (%)': (conv_col, '連結點擊次數', True)
    }
    summary_row = create_summary_row(df_metrics, metric_config)
    
    if not df_metrics.empty:
        return pd.concat([df_metrics, summary_row], ignore_index=True)
    return df_metrics

def prepare_excel_data(df, period_name_short, conv_col):
    """準備 Excel 用的數據塊"""
    df['廣告名稱_clean'] = df['廣告名稱'].apply(clean_ad_name)
    cols = [conv_col, '花費金額 (TWD)', '連結點擊次數', '曝光次數']
    df[cols] = df[cols].fillna(0)
    
    results = []
    results.append((f'{period_name_short}_Ad_廣告', calculate_metrics_consolidated(df.groupby('廣告名稱_clean'), conv_col)))
    results.append((f'{period_name_short}_AdSet_廣告組合', calculate_metrics_consolidated(df.groupby(['行銷活動名稱', '廣告組合名稱']), conv_col)))
    results.append((f'{period_name_short}_Campaign_行銷活動', calculate_metrics_consolidated(df.groupby('行銷活動名稱'), conv_col)))
    return results

def get_trend_data_excel(df_p30d, conv_col):
    """Excel 用的趨勢數據 (含行銷活動細分)"""
    trend_df = df_p30d.copy()
    
    camp_daily = trend_df.groupby(['天數', '行銷活動名稱']).agg({
        '花費金額 (TWD)': 'sum', conv_col: 'sum', '連結點擊次數': 'sum', '曝光次數': 'sum'
    }).reset_index()
    
    acc_daily = trend_df.groupby(['天數']).agg({
        '花費金額 (TWD)': 'sum', conv_col: 'sum', '連結點擊次數': 'sum', '曝光次數': 'sum'
    }).reset_index()
    acc_daily['行銷活動名稱'] = '🏆 整體帳戶'
    
    final = pd.concat([acc_daily, camp_daily], ignore_index=True)
    final = final[final['花費金額 (TWD)'] > 0]
    
    final['CPA'] = final.apply(lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0, axis=1)
    final['CVR (%)'] = final.apply(lambda x: (x[conv_col] / x['連結點擊次數']) * 100 if x['連結點擊次數'] > 0 else 0, axis=1)
    
    final['天數'] = final['天數'].dt.strftime('%Y-%m-%d')
    return final.sort_values(by=['天數', '行銷活動名稱'])

def generate_single_sheet_excel(dfs_list, prompt_text):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet('📘_完整分析報告')
        writer.sheets['📘_完整分析報告'] = ws
        
        fmt_header = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#0068C9'})
        fmt_prompt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'bg_color': '#F0F2F6'})
        fmt_th = workbook.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1})
        
        row = 0
        # 寫入 Prompt
        ws.merge_range('A1:H1', "🤖 AI 分析顧問指令", fmt_header)
        row += 1
        # 自動計算 Prompt 行數
        prompt_lines = prompt_text.count('\n') + 5
        ws.merge_range(row, 0, row + prompt_lines, 10, prompt_text, fmt_prompt)
        row += prompt_lines + 2
        
        ws.write(row, 0, "--- 📊 DATA SECTION START ---", fmt_header)
        row += 2
        
        for title, df in dfs_list:
            ws.write(row, 0, f"📌 Table: {title}", fmt_header)
            row += 1
            df.to_excel(writer, sheet_name='📘_完整分析報告', startrow=row, index=False)
            for i, col in enumerate(df.columns):
                ws.write(row, i, col, fmt_th)
            row += len(df) + 4
            
        ws.set_column('A:A', 40)
        ws.set_column('B:Z', 15)
    output.seek(0)
    return output.getvalue()

# ==========================================
# 3. 主程式 UI 邏輯
# ==========================================
st.title("📊 廣告成效全能分析儀表板 (v4.1 完整指令版)")
st.caption("整合功能：每日趨勢可視化 + 詳細數據表格 + AI 專用單頁報表匯出")

uploaded_file = st.file_uploader("請上傳 CSV 報表檔案", type=['csv'])

if uploaded_file is not None:
    try:
        # 1. 讀取與欄位偵測
        df = pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip()
        all_columns = df.columns.tolist()
        
        # 側邊欄：設定區
        with st.sidebar:
            st.header("⚙️ 分析設定")
            
            # 智慧偵測轉換欄位
            suggested_idx = 0
            for idx, col in enumerate(all_columns):
                c_low = col.lower()
                if '成本' in col or 'cost' in c_low: continue
                if ('free' in c_low and 'course' in c_low): suggested_idx = idx; break
                if '購買' in col or 'purchase' in c_low: suggested_idx = idx; break
                if '轉換' in col: suggested_idx = idx; break
                if '成果' in col: suggested_idx = idx; break
                
            conversion_col = st.selectbox(
                "🎯 目標轉換欄位:",
                options=all_columns,
                index=suggested_idx
            )
            
            # 嘗試找標準欄位
            def find_col(opts, default):
                for opt in opts:
                    for col in all_columns:
                        if opt in col: return col
                return default

            spend_col = find_col(['花費金額 (TWD)', '花費', '金額'], '花費金額 (TWD)')
            clicks_col = find_col(['連結點擊次數', '連結點擊'], '連結點擊次數')
            impressions_col = find_col(['曝光次數', '曝光'], '曝光次數')
            
            st.info(f"已鎖定：\n💰 花費: {spend_col}\n🖱️ 點擊: {clicks_col}\n👀 曝光: {impressions_col}")

        # 2. 資料清洗
        req_cols = [spend_col, clicks_col, impressions_col, conversion_col]
        df[req_cols] = df[req_cols].fillna(0)
        df['天數'] = pd.to_datetime(df['天數'])
        
        # 標準化欄位名稱 (方便後續處理，除了 conversion_col)
        df_std = df.rename(columns={
            spend_col: '花費金額 (TWD)',
            clicks_col: '連結點擊次數',
            impressions_col: '曝光次數'
        })
        
        # 日期區間設定
        max_date = df_std['天數'].max().normalize()
        today = max_date + timedelta(days=1)
        p7d_start = today - timedelta(days=7)
        p30d_start = today - timedelta(days=30)
        
        df_p7d = df_std[(df_std['天數'] >= p7d_start) & (df_std['天數'] < today)].copy()
        df_p30d = df_std[(df_std['天數'] >= p30d_start) & (df_std['天數'] < today)].copy()
        
        # --- 分頁內容 ---
        tab1, tab2 = st.tabs(["📈 視覺化儀表板 (Dashboard)", "📑 詳細數據列表 (Details)"])
        
        # === TAB 1: 視覺化圖表 ===
        with tab1:
            # 1. 數據摘要
            total_spend = df_p30d['花費金額 (TWD)'].sum()
            total_conv = df_p30d[conversion_col].sum()
            cpa_30d = total_spend / total_conv if total_conv > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("近30日總花費", f"${total_spend:,.0f}")
            c2.metric(f"近30日總轉換 ({conversion_col})", f"{total_conv:,.0f}")
            c3.metric("近30日平均 CPA", f"${cpa_30d:,.0f}")
            
            st.divider()
            
            # 2. 每日趨勢圖 (使用 Matplotlib)
            daily = df_p30d.groupby('天數')[['花費金額 (TWD)', conversion_col, '連結點擊次數', '曝光次數']].sum().reset_index()
            
            # 計算每日指標
            daily['CPM'] = daily.apply(lambda x: x['花費金額 (TWD)']/x['曝光次數']*1000 if x['曝光次數']>0 else 0, axis=1)
            daily['CTR'] = daily.apply(lambda x: x['連結點擊次數']/x['曝光次數'] if x['曝光次數']>0 else 0, axis=1)
            daily['CVR'] = daily.apply(lambda x: x[conversion_col]/x['連結點擊次數'] if x['連結點擊次數']>0 else 0, axis=1)
            
            plot_data = daily[daily['花費金額 (TWD)'] > 0].copy()
            plot_data['日期str'] = plot_data['天數'].dt.strftime('%m-%d')
            
            metrics_cfg = [
                ('花費金額 (TWD)', '每日花費 (Spend)', 'red'),
                (conversion_col, f'每日轉換 ({conversion_col})', 'brown'),
                ('CVR', '轉換率 (CVR)', 'magenta'),
                ('CPA', '轉換成本 (CPA)', 'purple', lambda x: x['花費金額 (TWD)']/x[conversion_col] if x[conversion_col]>0 else 0)
            ]
            
            # 繪圖
            fig, axes = plt.subplots(2, 2, figsize=(12, 10))
            axes = axes.flatten()
            
            for i, cfg in enumerate(metrics_cfg):
                col_name, title, color = cfg[0], cfg[1], cfg[2]
                ax = axes[i]
                
                # 特別處理 CPA 計算
                if len(cfg) > 3: # 自定義計算函數
                    y_vals = plot_data.apply(cfg[3], axis=1)
                    label_fmt = "{:.0f}"
                else:
                    y_vals = plot_data[col_name]
                    label_fmt = "{:.1%}" if col_name in ['CTR', 'CVR'] else "{:.0f}"
                
                ax.plot(plot_data['日期str'], y_vals, marker='o', color=color, linewidth=2)
                
                if font_prop:
                    ax.set_title(title, fontproperties=font_prop, fontsize=14)
                    ax.set_xlabel('日期', fontproperties=font_prop)
                    for label in ax.get_xticklabels() + ax.get_yticklabels():
                        label.set_fontproperties(font_prop)
                else:
                    ax.set_title(title) # Fallback
                
                ax.grid(True, linestyle='--', alpha=0.7)
                
                # 標籤
                for x, y in zip(plot_data['日期str'], y_vals):
                    ax.annotate(label_fmt.format(y), (x, y), textcoords="offset points", xytext=(0,8), ha='center', fontsize=9)
            
            plt.tight_layout()
            st.pyplot(fig)

        # === TAB 2: 詳細數據 ===
        with tab2:
            st.markdown("#### 📊 各維度數據總覽")
            t_p7, t_p30 = st.tabs(["P7D (近7天)", "P30D (近30天)"])
            
            with t_p7:
                res_p7 = prepare_excel_data(df_p7d, 'P7D', conversion_col)
                st.dataframe(res_p7[2][1], use_container_width=True) # 行銷活動
                with st.expander("查看廣告組合與廣告細節"):
                    st.write("廣告組合 (AdSet):")
                    st.dataframe(res_p7[1][1], use_container_width=True)
                    st.write("廣告 (Ad):")
                    st.dataframe(res_p7[0][1], use_container_width=True)
            
            with t_p30:
                res_p30 = prepare_excel_data(df_p30d, 'P30D', conversion_col)
                st.dataframe(res_p30[2][1], use_container_width=True)

        # === 側邊欄下載區 (最後執行) ===
        with st.sidebar:
            st.divider()
            st.header("📥 報告下載")
            
            # 準備 Excel 數據
            excel_stack = []
            # 1. Trend
            excel_stack.append(('Q13_Trend', get_trend_data_excel(df_p30d, conversion_col)))
            # 2. Periods
            excel_stack.extend(prepare_excel_data(df_p7d, 'P7D', conversion_col))
            excel_stack.extend(prepare_excel_data(df_p30d, 'P30D', conversion_col))
            
            excel_bytes = generate_single_sheet_excel(excel_stack, AI_CONSULTANT_PROMPT)
            
            st.download_button(
                label="下載 AI 專用單頁報表 (.xlsx)",
                data=excel_bytes,
                file_name=f"Full_Report_{conversion_col}_{max_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="包含完整的 System Prompt、每日趨勢與各層級數據，可直接上傳給 AI 進行分析。"
            )

    except Exception as e:
        st.error(f"發生錯誤: {e}")
        st.info("請確認 CSV 檔案格式正確，且包含花費、點擊與轉換數據。")
