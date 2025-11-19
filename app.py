import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta

# --- 輔助函數 ---

def clean_ad_name(name):
    """移除廣告名稱中的 ' - 複本' 及後續所有內容，以便將相同創意合併。"""
    # 專業處理：確保將所有「基本版1 內向I人_圖檔1 - 複本」等歸為「基本版1 內向I人_圖檔1」
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

def display_analysis_block(df, period_name):
    """在 Streamlit 中顯示單一時間區間的 Q1-Q9 分析結果。"""
    
    st.markdown(f"### 🎯 {period_name} 成效指標排名 (Q1-Q9)")
    
    # 在執行分析前，先對當前 DF 進行必要的數據清理
    df['廣告名稱_clean'] = df['廣告名稱'].apply(clean_ad_name)
    df['free-course'] = df['free-course'].fillna(0)
    df['花費金額 (TWD)'] = df['花費金額 (TWD)'].fillna(0)
    df['連結點擊次數'] = df['連結點擊次數'].fillna(0)
    df['曝光次數'] = df['曝光次數'].fillna(0)
    
    # --- 1-3. CPA (低到高) ---
    st.subheader("📊 每次成果成本 (CPA) 排名 - 低到高")
    
    # Q1. 廣告 CPA
    st.caption("1. 廣告 CPA：將相同創意合併 (排除複本)")
    ad_cpa = calculate_and_rank_metrics(df.groupby('廣告名稱_clean'), 'CPA', True)
    st.dataframe(ad_cpa[['廣告名稱_clean', 'free-course', '花費金額 (TWD)', 'CPA (TWD)']], use_container_width=True, hide_index=True)
    
    # Q2. 廣告組合 CPA
    st.caption("2. 行銷活動 + 廣告組合 CPA")
    adset_cpa = calculate_and_rank_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), 'CPA', True)
    st.dataframe(adset_cpa[['行銷活動名稱', '廣告組合名稱', 'free-course', '花費金額 (TWD)', 'CPA (TWD)']], use_container_width=True, hide_index=True)

    # Q3. 行銷活動 CPA
    st.caption("3. 行銷活動 CPA")
    campaign_cpa = calculate_and_rank_metrics(df.groupby('行銷活動名稱'), 'CPA', True)
    st.dataframe(campaign_cpa[['行銷活動名稱', 'free-course', '花費金額 (TWD)', 'CPA (TWD)']], use_container_width=True, hide_index=True)

    st.markdown("---")
    
    # --- 4-6. CPC (低到高) ---
    st.subheader("💰 每次連結點擊成本 (CPC) 排名 - 低到高")

    # Q4. 廣告 CPC
    st.caption("4. 廣告 CPC：將相同創意合併 (排除複本)")
    ad_cpc = calculate_and_rank_metrics(df.groupby('廣告名稱_clean'), 'CPC', True)
    st.dataframe(ad_cpc[['廣告名稱_clean', '連結點擊次數', '花費金額 (TWD)', 'CPC (TWD)']], use_container_width=True, hide_index=True)
    
    # Q5. 廣告組合 CPC
    st.caption("5. 行銷活動 + 廣告組合 CPC")
    adset_cpc = calculate_and_rank_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), 'CPC', True)
    st.dataframe(adset_cpc[['行銷活動名稱', '廣告組合名稱', '連結點擊次數', '花費金額 (TWD)', 'CPC (TWD)']], use_container_width=True, hide_index=True)

    # Q6. 行銷活動 CPC
    st.caption("6. 行銷活動 CPC")
    campaign_cpc = calculate_and_rank_metrics(df.groupby('行銷活動名稱'), 'CPC', True)
    st.dataframe(campaign_cpc[['行銷活動名稱', '連結點擊次數', '花費金額 (TWD)', 'CPC (TWD)']], use_container_width=True, hide_index=True)

    st.markdown("---")

    # --- 7-9. CTR (高到低) ---
    st.subheader("⚡ 連結點閱率 (CTR) 排名 - 高到低")

    # Q7. 廣告 CTR
    st.caption("7. 廣告 CTR：將相同創意合併 (排除複本)")
    ad_ctr = calculate_and_rank_metrics(df.groupby('廣告名稱_clean'), 'CTR', False)
    st.dataframe(ad_ctr[['廣告名稱_clean', '連結點擊次數', '曝光次數', 'CTR (%)']], use_container_width=True, hide_index=True)
    
    # Q8. 廣告組合 CTR
    st.caption("8. 行銷活動 + 廣告組合 CTR")
    adset_ctr = calculate_and_rank_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), 'CTR', False)
    st.dataframe(adset_ctr[['行銷活動名稱', '廣告組合名稱', '連結點擊次數', '曝光次數', 'CTR (%)']], use_container_width=True, hide_index=True)

    # Q9. 行銷活動 CTR
    st.caption("9. 行銷活動 CTR")
    campaign_ctr = calculate_and_rank_metrics(df.groupby('行銷活動名稱'), 'CTR', False)
    st.dataframe(campaign_ctr[['行銷活動名稱', '連結點擊次數', '曝光次數', 'CTR (%)']], use_container_width=True, hide_index=True)
    
# --- Q10 趨勢分析 ---
def display_trend_analysis(df_p30d):
    """顯示 Q10 每日趨勢波動分析。"""
    
    st.header("📈 趨勢與波動檢視 (Q10) - 過去 30 天")
    st.markdown("以**每日**的**行銷活動**為基礎，檢視 CPA 與 CTR 的波動情況，以幫助判斷趨勢變化。")
    
    # 預處理 Q10 數據
    trend_df = df_p30d.copy()
    
    # 每日行銷活動趨勢
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
    
    st.dataframe(
        campaign_daily_trend[['天數', '行銷活動名稱', '花費金額 (TWD)', 'free-course', 'CPA (TWD)', 'CTR (%)']].round(2), 
        use_container_width=True,
        hide_index=True
    )
    st.markdown("---")

# --- 3. Streamlit 主程式 ---

def marketing_analysis_app():
    st.set_page_config(layout="wide")
    st.title("📊 廣告成效多週期分析工具")
    st.markdown("請上傳您的廣告數據 CSV 檔案。系統將自動依據檔案中**最新日期**，計算過去七天 (P7D)、前七天 (PP7D) 及過去三十天 (P30D) 的所有指標排名。")

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

            # 2. 過去七天的前七天 (PP7D)
            pp7d_end = p7d_start - timedelta(days=1)
            pp7d_start = p7d_start - timedelta(days=7)
            pp7d_period = f'前七天 ({pp7d_start.strftime("%Y-%m-%d")} ~ {pp7d_end.strftime("%Y-%m-%d")})'
            df_pp7d = df[(df['天數'] >= pp7d_start) & (df['天數'] <= pp7d_end)].copy()

            # 3. 過去三十天 (P30D)
            p30d_end = today - timedelta(days=1)
            p30d_start = today - timedelta(days=30)
            p30d_period = f'過去三十天 ({p30d_start.strftime("%Y-%m-%d")} ~ {p30d_end.strftime("%Y-%m-%d")})'
            df_p30d = df[(df['天數'] >= p30d_start) & (df['天數'] <= p30d_end)].copy()
            
            # --- 運行分析並使用 Tabs 輸出 ---

            tab1, tab2, tab3 = st.tabs([p7d_period, pp7d_period, p30d_period])

            with tab1:
                display_analysis_block(df_p7d, p7d_period)

            with tab2:
                display_analysis_block(df_pp7d, pp7d_period)

            with tab3:
                display_analysis_block(df_p30d, p30d_period)

            # --- Q10 趨勢分析單獨顯示 (使用 P30D 資料) ---
            st.markdown("---")
            display_trend_analysis(df_p30d)


        except Exception as e:
            st.error(f"資料處理發生錯誤，請檢查您的 CSV 檔案格式，特別是日期欄位（'天數'）和數字欄位：{e}")
            st.code(str(e))

if __name__ == "__main__":
    marketing_analysis_app()
