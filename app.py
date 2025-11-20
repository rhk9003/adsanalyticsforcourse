import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import io

# ==========================================
# 1. 輔助函數 (資料處理核心)
# ==========================================

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

# ==========================================
# 2. Streamlit 顯示組件
# ==========================================

def display_analysis_block(df, period_name, period_name_short):
    """在 Streamlit 中顯示單一時間區間的 Q1-Q9 分析結果。"""
    
    st.markdown(f"### 🎯 {period_name} 成效指標排名")
    
    # 獲取所有結果來顯示
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


# ==========================================
# 3. Streamlit 主程式 (包含 AI Prompt 功能)
# ==========================================

def marketing_analysis_app():
    st.set_page_config(layout="wide", page_title="廣告成效智能分析工具")
    
    st.title("📊 廣告成效多週期分析工具 (AI Ready)")
    
    # ------------------------------------------
    # 新增功能：AI 顧問指令生成區
    # ------------------------------------------
    with st.expander("🤖 步驟 1：獲取 AI 深度診斷指令 (Prompt)", expanded=True):
        st.info("💡 使用說明：請點擊右上角「複製」按鈕，將此指令連同下方下載的 **Excel 報表** 一起貼給 ChatGPT/Claude/Gemini，即可獲得專業分析。")
        
        ai_consultant_prompt = """
# Role
你是一位擁有 10 年經驗的資深成效廣告分析師，擅長數據解讀、商業策略推演與消費者心理分析。請根據我上傳的廣告數據 Excel 檔案（涵蓋 Campaign, AdSet, Ad 三個層級，以及 P7D, PP7D, P30D 不同時間區間），進行深度的廣告帳戶健檢。

# Data Context & File Naming Logic
- **P7D**: 過去 7 天數據（近期表現）。
- **PP7D**: 上一個 7 天數據（用於做 WoW 環比比較）。
- **P30D**: 過去 30 天數據（用於看長期趨勢與累積數據）。
- **Q10_Trend**: 每日趨勢數據。
- **關鍵指標**: CPA (Cost Per Action), CTR (點擊率), CPC (點擊成本), Spend (花費), Conversions (free-course/成果)。

# Analysis Requirements (請依序執行以下任務)

## 1. 波動偵測 (Fluctuation Analysis)
- **目標**: 找出近期表現劇烈變化的項目。
- **執行動作**:
    - 對比 Campaign 與 AdSet 層級的 **P7D vs. PP7D** 數據。
    - 找出 CPA 暴漲（>30%）或 轉單量驟跌的「警示區」。
    - 找出 CPA 顯著下降或 轉單量激增的「機會區」。
- **輸出重點**: 不要只列數字，請告訴我「哪裡變好了？哪裡變壞了？」。

## 2. 擴量機會掃描 (Scaling Opportunities)
- **目標**: 找出值得加碼預算的「明星項目」。
- **篩選標準**:
    - **高效率**: P7D CPA 低於帳戶平均值，且具備一定轉單量。
    - **高潛力**: CTR 顯著高於平均（代表受眾對素材有高興趣），但目前預算/曝光不足（Impression 較低）的項目。
    - **受眾紅利**: 在 AdSet 層級，找出那些「花費少但 CPA 極低」的受眾（例如特定興趣或版位）。
- **輸出重點**: 明確指出哪一個 Campaign/AdSet/Ad 應該增加預算？建議加碼的理由是什麼？

## 3. 止損與縮編建議 (Cost Cutting)
- **目標**: 揪出浪費預算的「黑洞」。
- **篩選標準**:
    - **無效花費**: P7D/P30D 花費高昂但 0 轉單的項目。
    - **低效能**: CPA 遠高於平均（>1.5倍），且 CTR 低落（表示受眾不買單）的項目。
    - **素材疲勞**: P30D 表現尚可，但 P7D CPA 飆升且 CTR 下滑的素材。
- **輸出重點**: 明確列出哪些應該「立即關閉」？哪些應該「縮減預算」？

## 4. 受眾動機與素材洞察 (Audience & Creative Strategy)
- **目標**: 從數據反推「為什麼這群人會買單？」。
- **執行動作**:
    - 分析表現最好的前 3-5 名素材名稱（Ad Name）與視覺/文案標籤（如：I人、媽媽、創業、上班族...）。
    - 結合 CTR 數據，解讀哪
