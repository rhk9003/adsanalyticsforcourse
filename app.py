import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import io

# ==========================================
# 0. 全域設定：AI 顧問指令 (針對單頁堆疊版優化)
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
# 1. 輔助函數 (資料處理核心)
# ==========================================

def clean_ad_name(name):
    """移除廣告名稱中的 ' - 複本' 及後續所有內容。"""
    return re.sub(r' - 複本.*$', '', str(name)).strip()

def create_summary_row(df, metric_cols):
    """計算加總平均列的輔助函數 (支援多欄位)。"""
    summary_dict = {}
    
    # 先計算所有數值欄位的總和
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        summary_dict[col] = df[col].sum()
        
    # 重新計算衍生指標
    for metric, (num, denom, is_pct) in metric_cols.items():
        total_num = summary_dict.get(num, 0)
        total_denom = summary_dict.get(denom, 0)
        
        if total_denom > 0:
            val = (total_num / total_denom)
            if is_pct: val *= 100
            summary_dict[metric] = round(val, 2)
        else:
            summary_dict[metric] = 0

    # 處理非數值欄位
    non_numeric_cols = df.select_dtypes(exclude=[np.number]).columns
    if len(non_numeric_cols) > 0:
        summary_dict[non_numeric_cols[0]] = '全帳戶平均'
        for col in non_numeric_cols[1:]:
            summary_dict[col] = '-'
            
    return pd.DataFrame([summary_dict])

def calculate_consolidated_metrics(df_group):
    """核心函數：一次計算所有指標並合併。"""
    # 1. 聚合
    df_metrics = df_group.agg({
        '花費金額 (TWD)': 'sum',
        'free-course': 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    }).reset_index()

    # 2. 過濾
    df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]

    # 3. 計算指標
    df_metrics['CPA (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x['free-course'] if x['free-course'] > 0 else 0, axis=1)
    df_metrics['CTR (%)'] = df_metrics.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    df_metrics['CVR (%)'] = df_metrics.apply(lambda x: (x['free-course'] / x['連結點擊次數']) * 100 if x['連結點擊次數'] > 0 else 0, axis=1)
    df_metrics['CPC (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else 0, axis=1)

    # 4. 數值修整與排序
    df_metrics = df_metrics.round(2)
    df_metrics = df_metrics.sort_values(by='花費金額 (TWD)', ascending=False)

    # 5. 平均列
    metric_config = {
        'CPA (TWD)': ('花費金額 (TWD)', 'free-course', False),
        'CTR (%)': ('連結點擊次數', '曝光次數', True),
        'CVR (%)': ('free-course', '連結點擊次數', True),
        'CPC (TWD)': ('花費金額 (TWD)', '連結點擊次數', False)
    }
    summary_row = create_summary_row(df_metrics, metric_config)
    
    if not df_metrics.empty:
        return pd.concat([df_metrics, summary_row], ignore_index=True)
    else:
        return df_metrics

def collect_all_results_consolidated(df, period_name_short):
    """產生整合版的數據列表"""
    # 預處理
    df['廣告名稱_clean'] = df['廣告名稱'].apply(clean_ad_name)
    cols_to_fill = ['free-course', '花費金額 (TWD)', '連結點擊次數', '曝光次數']
    df[cols_to_fill] = df[cols_to_fill].fillna(0)
    
    results = []
    results.append((f'{period_name_short}_Ad_廣告', calculate_consolidated_metrics(df.groupby('廣告名稱_clean'))))
    results.append((f'{period_name_short}_AdSet_廣告組合', calculate_consolidated_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']))))
    results.append((f'{period_name_short}_Campaign_行銷活動', calculate_consolidated_metrics(df.groupby('行銷活動名稱'))))
    return results

def get_trend_data(df_p30d):
    """計算每日趨勢"""
    trend_df = df_p30d.copy()
    
    campaign_daily = trend_df.groupby(['天數', '行銷活動名稱']).agg({
        '花費金額 (TWD)': 'sum', 'free-course': 'sum', '連結點擊次數': 'sum', '曝光次數': 'sum'
    }).reset_index()
    
    account_daily = trend_df.groupby(['天數']).agg({
        '花費金額 (TWD)': 'sum', 'free-course': 'sum', '連結點擊次數': 'sum', '曝光次數': 'sum'
    }).reset_index()
    account_daily['行銷活動名稱'] = '🏆 整體帳戶 (Account Overall)'
    
    final_trend = pd.concat([account_daily, campaign_daily], ignore_index=True)
    final_trend = final_trend[final_trend['花費金額 (TWD)'] > 0]
    
    final_trend['CPA (TWD)'] = final_trend.apply(lambda x: x['花費金額 (TWD)'] / x['free-course'] if x['free-course'] > 0 else 0, axis=1)
    final_trend['CTR (%)'] = final_trend.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    final_trend['CVR (%)'] = final_trend.apply(lambda x: (x['free-course'] / x['連結點擊次數']) * 100 if x['連結點擊次數'] > 0 else 0, axis=1)
    
    final_trend['天數'] = final_trend['天數'].dt.strftime('%Y-%m-%d')
    return final_trend.round(2).sort_values(by=['天數', '行銷活動名稱'])

def to_excel_single_sheet(dfs_list, prompt_text):
    """
    將所有數據垂直堆疊在同一個 Excel 分頁中。
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # 建立唯一的分頁
        sheet_name = '📘_完整分析報告'
        ws = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws
        
        # 格式設定
        fmt_prompt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_size': 11, 'bg_color': '#F0F2F6'})
        fmt_header = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#0068C9'})
        fmt_note = workbook.add_format({'italic': True, 'font_size': 10, 'font_color': '#555555'}) # [NEW] 註解格式
        fmt_table_header = workbook.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1})
        
        current_row = 0
        
        # 1. 寫入 AI 指令 (Prompt)
        ws.merge_range('A1:H1', "🤖 AI 分析顧問指令 (SYSTEM PROMPT)", fmt_header)
        current_row += 1
        
        # 估算 Prompt 行數 (概略)
        prompt_lines = prompt_text.count('\n') + 5
        ws.merge_range(current_row, 0, current_row + prompt_lines, 10, prompt_text, fmt_prompt)
        current_row += prompt_lines + 2
        
        ws.write(current_row, 0, "--- 📊 DATA SECTION START (Below are Stacked Tables) ---", fmt_header)
        current_row += 2
        
        # 2. 迴圈寫入所有 DataFrame
        for title, df in dfs_list:
            # 寫標題
            ws.write(current_row, 0, f"📌 Table: {title}", fmt_header)
            current_row += 1
            
            # [NEW] 新增排序說明註解 (Trend 表格除外，因為 Trend 是依日期排序)
            if "Trend" not in title:
                ws.write(current_row, 0, "   ℹ️ Ranking: Sorted by Spend (High to Low). Last row is Account Average.", fmt_note)
                current_row += 1
            
            # 寫入 DataFrame
            # 使用 pandas to_excel 寫入數據，不包含 index
            df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
            
            # 簡單的 Header 樣式覆蓋 (為了美觀，可選)
            for col_num, value in enumerate(df.columns.values):
                ws.write(current_row, col_num, value, fmt_table_header)
            
            # 更新 current_row (數據行數 + Header + 間距)
            current_row += len(df) + 4 # 留 3 行空白
            
        # 設定欄寬 (概略)
        ws.set_column('A:A', 40) # 名稱欄寬一點
        ws.set_column('B:J', 15) # 數值欄
            
    output.seek(0)
    return output.getvalue()

# ==========================================
# 2. Streamlit 顯示組件
# ==========================================

def display_consolidated_block(df, period_name, period_name_short):
    """顯示整合版數據預覽"""
    st.markdown(f"### 🎯 {period_name} 綜合數據概覽")
    results = collect_all_results_consolidated(df, period_name_short)
    
    st.caption("1. 廣告層級 (Ad Level) - 含所有指標")
    st.dataframe(results[0][1], use_container_width=True, hide_index=True)
    st.caption("2. 廣告組合層級 (AdSet Level)")
    st.dataframe(results[1][1], use_container_width=True, hide_index=True)
    st.caption("3. 行銷活動層級 (Campaign Level)")
    st.dataframe(results[2][1], use_container_width=True, hide_index=True)

# ==========================================
# 3. Streamlit 主程式
# ==========================================

def marketing_analysis_app():
    st.set_page_config(layout="wide", page_title="廣告成效智能分析工具")
    
    st.title("📊 廣告成效多週期分析工具 (AI Ready)")
    st.markdown("### 🚀 最終進化版：單頁報告模式")
    st.info("已將所有指令與數據合併為 **單一 Excel 分頁 (Single Sheet)**，採用垂直堆疊格式。這能確保 AI 能夠一次性讀取所有內容，不再發生「讀不到分頁」的問題。")
    
    uploaded_file = st.file_uploader("上傳 CSV 檔案", type=["csv"])

    if uploaded_file is not None:
        try:
            # 讀取與清洗
            df = pd.read_csv(uploaded_file)
            df.columns = df.columns.str.strip()
            
            col_map = {
                'free course': 'free-course', 'Free course': 'free-course',
                'Free Course': 'free-course', '花費金額': '花費金額 (TWD)',
                '金額': '花費金額 (TWD)'
            }
            df.rename(columns=col_map, inplace=True)
            
            # 檢查
            req_cols = ['天數', '行銷活動名稱', 'free-course', '花費金額 (TWD)', '連結點擊次數', '曝光次數']
            missing = [c for c in req_cols if c not in df.columns]
            if missing:
                st.error(f"❌ 缺少欄位: {missing}")
                st.stop()

            # 日期處理
            df['天數'] = pd.to_datetime(df['天數'])
            max_date = df['天數'].max().normalize()
            today = max_date + timedelta(days=1)
            
            st.success(f"資料最新日期：{max_date.strftime('%Y-%m-%d')}")

            # 定義區間
            p7d_start = today - timedelta(days=7)
            p7d_end = today - timedelta(days=1)
            pp7d_start = p7d_start - timedelta(days=7)
            pp7d_end = p7d_start - timedelta(days=1)
            p30d_start = today - timedelta(days=30)
            p30d_end = today - timedelta(days=1) # 確保變數存在
            
            df_p7d = df[(df['天數'] >= p7d_start) & (df['天數'] <= p7d_end)].copy()
            df_pp7d = df[(df['天數'] >= pp7d_start) & (df['天數'] <= pp7d_end)].copy()
            df_p30d = df[(df['天數'] >= p30d_start) & (df['天數'] <= p30d_end)].copy()

            # 執行分析與收集 (準備堆疊的數據)
            stacked_data = []
            
            # 1. Trend
            q13_df = get_trend_data(df_p30d)
            stacked_data.append(('Q13_P30D_Trend (含整體帳戶)', q13_df))
            
            # 2. Periods Data
            stacked_data.extend(collect_all_results_consolidated(df_p7d, 'P7D'))
            stacked_data.extend(collect_all_results_consolidated(df_pp7d, 'PP7D'))
            stacked_data.extend(collect_all_results_consolidated(df_p30d, 'P30D'))

            # UI 顯示 (保持分頁瀏覽以便人類閱讀)
            t1, t2, t3, t4 = st.tabs(["📈 趨勢", "P7D (本週)", "PP7D (上週)", "P30D (月報)"])
            with t1: st.dataframe(q13_df, use_container_width=True)
            with t2: display_consolidated_block(df_p7d, "P7D", "P7D")
            with t3: display_consolidated_block(df_pp7d, "PP7D", "PP7D")
            with t4: display_consolidated_block(df_p30d, "P30D", "P30D")

            # 下載 (單頁版)
            excel_data = to_excel_single_sheet(stacked_data, AI_CONSULTANT_PROMPT)
            
            st.markdown("### 📥 下載 AI 專用報表")
            st.download_button(
                label="下載單頁式完整分析報表 (.xlsx)",
                data=excel_data,
                file_name=f"Ad_Analysis_SingleSheet_{max_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="所有數據與指令都在同一個分頁中，直接上傳給 AI 即可，保證讀取成功。"
            )

        except Exception as e:
            st.error(f"發生錯誤: {e}")

if __name__ == "__main__":
    marketing_analysis_app()
