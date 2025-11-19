import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import io

# --- 1. 輔助函數 ---

def clean_ad_name(name):
    """移除廣告名稱中的 ' - 複本' 及後續所有內容，以便將相同創意合併。"""
    return re.sub(r' - 複本.*$', '', str(name)).strip()

def calculate_and_rank_metrics(df_group, metric_type, sort_ascending):
    """計算 CPA/CPC/CTR 指標並排名。"""
    
    if metric_type == 'CPA':
        # Q1, Q2, Q3 metrics
        df_metrics = df_group.agg({
            '花費金額 (TWD)': 'sum',
            'free-course': 'sum'
        }).reset_index()
        df_metrics['CPA (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x['free-course'] if x['free-course'] > 0 else np.nan, axis=1)
        df_metrics.replace([np.inf, -np.inf], np.nan, inplace=True)
        return df_metrics.sort_values(by='CPA (TWD)', ascending=sort_ascending).round(2)

    elif metric_type == 'CPC':
        # Q4, Q5, Q6 metrics
        df_metrics = df_group.agg({
            '花費金額 (TWD)': 'sum',
            '連結點擊次數': 'sum'
        }).reset_index()
        df_metrics['CPC (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else np.nan, axis=1)
        df_metrics.replace([np.inf, -np.inf], np.nan, inplace=True)
        return df_metrics.sort_values(by='CPC (TWD)', ascending=sort_ascending).round(2)

    elif metric_type == 'CTR':
        # Q7, Q8, Q9 metrics
        df_metrics = df_group.agg({
            '連結點擊次數': 'sum',
            '曝光次數': 'sum'
        }).reset_index()
        df_metrics['CTR (%)'] = df_metrics.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
        return df_metrics.sort_values(by='CTR (%)', ascending=sort_ascending).round(2)

def collect_all_results(df, period_name_short):
    """執行 Q1-Q9 分析並收集結果為 (Sheet Name, DataFrame) 列表。"""
    
    # 預處理當前 DF
    df['廣告名稱_clean'] = df['廣告名稱'].apply(clean_ad_name)
    df['free-course'] = df['free-course'].fillna(0)
    df['花費金額 (TWD)'] = df['花費金額 (TWD)'].fillna(0)
    df['連結點擊次數'] = df['連結點擊次數'].fillna(0)
    df['曝光次數'] = df['曝光次數'].fillna(0)
    
    results = []
    
    # CPA (Q1-Q3)
    results.append((f'{period_name_short}_Q1_Ad_CPA', calculate_and_rank_metrics(df.groupby('廣告名稱_clean'), 'CPA', True)))
    results.append((f'{period_name_short}_Q2_AdSet_CPA', calculate_and_rank_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), 'CPA', True)))
    results.append((f'{period_name_short}_Q3_Campaign_CPA', calculate_and_rank_metrics(df.groupby('行銷活動名稱'), 'CPA', True)))

    # CPC (Q4-Q6)
    results.append((f'{period_name_short}_Q4_Ad_CPC', calculate_and_rank_metrics(df.groupby('廣告名稱_clean'), 'CPC', True)))
    results.append((f'{period_name_short}_Q5_AdSet_CPC', calculate_and_rank_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), 'CPC', True)))
    results.append((f'{period_name_short}_Q6_Campaign_CPC', calculate_and_rank_metrics(df.groupby('行銷活動名稱'), 'CPC', True)))

    # CTR (Q7-Q9)
    results.append((f'{period_name_short}_Q7_Ad_CTR', calculate_and_rank_metrics(df.groupby('廣告名稱_clean'), 'CTR', False)))
    results.append((f'{period_name_short}_Q8_AdSet_CTR', calculate_and_rank_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), 'CTR', False)))
    results.append((f'{period_name_short}_Q9_Campaign_CTR', calculate_and_rank_metrics(df.groupby('行銷活動名稱'), 'CTR', False)))
    
    return results

def to_excel_bytes(dfs_to_export):
    """將列表中的 (sheet_name, DataFrame) 寫入 Excel 文件的 BytesIO。"""
    output = io.BytesIO()
    # 使用 xlsxwriter 引擎
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dfs_to_export:
            # 確保 sheet name 不超過 Excel 限制 (31字元)
            safe_sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            
    # 將指標移到開頭，準備下載
    output.seek(0)
    return output.getvalue()

# --- 2. Streamlit 顯示函數 ---

def display_analysis_block(df, period_name, period_name_short):
    """在 Streamlit 中顯示單一時間區間的 Q1-Q9 分析結果。"""
    
    st.markdown(f"### 🎯 {period_name} 成效指標排名")
    
    # 重新運行計算以便顯示，這裡只需要顯示，數據已經被 collect_all_results 函數處理
    # 這裡的 df 已經是經過預處理的副本
    
    # 方便地獲取所有結果來顯示
    all_results = collect_all_results(df, period_name_short)
    
    # 顯示 CPA
    st.subheader("📊 每次成果成本 (CPA) 排名 - 低到高")
    st.caption("1. 廣告 CPA")
    st.dataframe(all_results[0][1].rename(columns={'廣告名稱_clean': '廣告名稱'}), use_container_width=True, hide_index=True)
    st.caption("2. 廣告組合 CPA")
    st.dataframe(all_results[1][1], use_container_width=True, hide_index=True)
    st.caption("3. 行銷活動 CPA")
    st.dataframe(all_results[2][1], use_container_width=True, hide_index=True)
    
    # 顯示 CPC
    st.subheader("💰 每次連結點擊成本 (CPC) 排名 - 低到高")
    st.caption("4. 廣告 CPC")
    st.dataframe(all_results[3][1].rename(columns={'廣告名稱_clean': '廣告名稱'}), use_container_width=True, hide_index=True)
    st.caption("5. 廣告組合 CPC")
    st.dataframe(all_results[4][1], use_container_width=True, hide_index=True)
    st.caption("6. 行銷活動 CPC")
    st.dataframe(all_results[5][1], use_container_width=True, hide_index=True)

    # 顯示 CTR
    st.subheader("⚡ 連結點閱率 (CTR) 排名 - 高到低")
    st.caption("7. 廣告 CTR")
    st.dataframe(all_results[6][1].rename(columns={'廣告名稱_clean': '廣告名稱'}), use_container_width=True, hide_index=True)
    st.caption("8. 廣告組合 CTR")
    st.dataframe(all_results[7][1], use_container_width=True, hide_index=True)
    st.caption("9. 行銷活動 CTR")
    st.dataframe(all_results[8][1], use_container_width=True, hide_index=True)


def display_trend_analysis(df_p30d):
    """顯示 Q10 每日趨勢波動分析並返回其 DataFrame。"""
    
    st.header("📈 趨勢與波動檢視 (Q10) - 過去 30 天")
    st.markdown("以**每日**的**行銷活動**為基礎，檢視 CPA 與 CTR 的波動情況，以幫助判斷趨勢變化。")
    
    trend_df = df_p30d.copy()
    trend_df['廣告名稱_clean'] = trend_df['廣告名稱'].apply(clean_ad_name)

    campaign_daily_trend = trend_df.groupby(['天數', '行銷活動名稱']).agg({
        '花費金額 (TWD)': 'sum',
        'free-course': 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    }).reset_index()

    campaign_daily_trend['CPA (TWD)'] = campaign_daily_trend.apply(lambda x: x['花費金額 (TWD)'] / x['free-course'] if x['free-course'] > 0 else np.nan, axis=1)
    campaign_daily_trend['CTR (%)'] = campaign_daily_trend.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    
    # 格式化輸出
    campaign_daily_trend['天數'] = campaign_daily_trend['天數'].dt.strftime('%Y-%m-%d')
    campaign_daily_trend.replace([np.inf, -np.inf], np.nan, inplace=True)
    
    trend_output_df = campaign_daily_trend[['天數', '行銷活動名稱', '花費金額 (TWD)', 'free-course', 'CPA (TWD)', 'CTR (%)']].round(2)
    
    st.dataframe(trend_output_df, use_container_width=True, hide_index=True)
    
    return trend_output_df


# --- 3. Streamlit 主程式 ---

def marketing_analysis_app():
    st.set_page_config(layout="wide")
    st.title("📊 廣告成效多週期分析工具")
    st.markdown("請上傳您的廣告數據 CSV 檔案。系統將自動依據檔案中**最新日期**，計算三個時間區間的指標排名與趨勢分析。")

    uploaded_file = st.file_uploader("上傳 CSV 檔案", type=["csv"])

    if uploaded_file is not None:
        try:
            # 讀取檔案
            df = pd.read_csv(uploaded_file)
            
            # 初始預處理
            df['天數'] = pd.to_datetime(df['天數'])
            
            # 確認日期區間
            max_date = df['天數'].max().normalize()
            today = max_date + timedelta(days=1)
            
            st.success(f"檔案讀取成功！資料集最新日期為：**{max_date.strftime('%Y-%m-%d')}**")

            # --- 定義時間區間 ---
            
            # 1. 過去七天 (P7D)
            p7d_end = today - timedelta(days=1)
            p7d_start = today - timedelta(days=7)
            p7d_period = f'過去七天 ({p7d_start.strftime("%Y-%m-%d")} ~ {p7d_end.strftime("%Y-%m-%d")})'
            df_p7d = df[(df['天數'] >= p7d_start) & (df['天數'] <= p7d_end)].copy()
            P7D_SHORT = 'P7D'

            # 2. 過去七天的前七天 (PP7D)
            pp7d_end = p7d_start - timedelta(days=1)
            pp7d_start = p7d_start - timedelta(days=7)
            pp7d_period = f'前七天 ({pp7d_start.strftime("%Y-%m-%d")} ~ {pp7d_end.strftime("%Y-%m-%d")})'
            df_pp7d = df[(df['天數'] >= pp7d_start) & (df['天數'] <= pp7d_end)].copy()
            PP7D_SHORT = 'PP7D'

            # 3. 過去三十天 (P30D)
            p30d_end = today - timedelta(days=1)
            p30d_start = today - timedelta(days=30)
            p30d_period = f'過去三十天 ({p30d_start.strftime("%Y-%m-%d")} ~ {p30d_end.strftime("%Y-%m-%d")})'
            df_p30d = df[(df['天數'] >= p30d_start) & (df['天數'] <= p30d_end)].copy()
            P30D_SHORT = 'P30D'
            
            # --- 執行分析並收集所有結果 ---
            
            all_dfs_for_excel = []
            
            # Q1-Q9: 排名數據
            all_dfs_for_excel.extend(collect_all_results(df_p7d.copy(), P7D_SHORT))
            all_dfs_for_excel.extend(collect_all_results(df_pp7d.copy(), PP7D_SHORT))
            all_dfs_for_excel.extend(collect_all_results(df_p30d.copy(), P30D_SHORT))

            # --- 顯示 Tabs 輸出 ---

            tab1, tab2, tab3 = st.tabs([p7d_period, pp7d_period, p30d_period])

            with tab1:
                display_analysis_block(df_p7d, p7d_period, P7D_SHORT)

            with tab2:
                display_analysis_block(df_pp7d, pp7d_period, PP7D_SHORT)

            with tab3:
                display_analysis_block(df_p30d, p30d_period, P30D_SHORT)

            # --- Q10 趨勢分析單獨顯示 (使用 P30D 資料) ---
            st.markdown("---")
            q10_df = display_trend_analysis(df_p30d)
            
            # Q10: 趨勢數據加入 Excel 輸出列表
            all_dfs_for_excel.append(('Q10_P30D_Trend', q10_df))

            # --- 創建 Excel 下載按鈕 ---
            excel_data = to_excel_bytes(all_dfs_for_excel)
            
            st.download_button(
                label="📥 下載所有分析結果 (.xlsx)",
                data=excel_data,
                file_name=f"Ad_Analysis_Report_{max_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="點擊下載包含所有週期和指標的 Excel 報表。"
            )


        except Exception as e:
            st.error(f"資料處理發生錯誤，請檢查您的 CSV 檔案格式，特別是日期欄位（'天數'）和數字欄位：{e}")
            st.code(str(e))

if __name__ == "__main__":
    marketing_analysis_app()
