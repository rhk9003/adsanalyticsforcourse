import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import io

# ==========================================
# 0. 全域設定：AI 顧問指令 (將內建於 Excel 中)
# ==========================================

AI_CONSULTANT_PROMPT = """
# Role
你是一位擁有 10 年經驗的資深成效廣告分析師，擅長數據解讀、商業策略推演與消費者心理分析。
請讀取本 Excel 檔案中的所有數據分頁（涵蓋 Campaign, AdSet, Ad 三個層級，以及 P7D, PP7D, P30D 不同時間區間），進行深度的廣告帳戶健檢。

# Data Context & File Naming Logic
- **P7D**: 過去 7 天數據（近期表現）。
- **PP7D**: 上一個 7 天數據（用於做 WoW 環比比較）。
- **P30D**: 過去 30 天數據（用於看長期趨勢與累積數據）。
- **Q10_Trend**: 每日趨勢數據。
- **關鍵指標**: CPA (Cost Per Action), CTR (點擊率), CPC (點擊成本), Spend (花費), Conversions (free-course/成果)。
- **全帳戶平均**: 每個表格的最下方有一列「全帳戶平均」，請以此作為基準線來判斷優劣。

# Analysis Requirements (請依序執行以下任務)

## 1. 波動偵測 (Fluctuation Analysis)
- **目標**: 找出近期表現劇烈變化的項目。
- **執行動作**:
    - 對比 Campaign 與 AdSet 層級的 **P7D vs. PP7D** 數據。
    - 找出 CPA 暴漲（>30%）或 轉單量驟跌的「警示區」。
    - 找出 CPA 顯著下降或 轉單量激增的「機會區」。
- **輸出重點**: 不要只列數字，請告訴我「哪裡變好了？哪裡變壞了？」。

## 2. 擴量機會與潛力黑馬掃描 (Scaling & Hidden Gems)
- **目標**: 找出「明星項目」以及「被系統低估的潛力股」。
- **篩選標準**:
    - **明星與穩健號 (Proven Winners)**: P7D CPA 低於帳戶平均，且具備穩定轉單量。 -> 建議：加碼擴量。
    - **預算受限的潛力股 (Budget Constrained Potential)**: **CTR 顯著高於平均 (例如 > 2.5%~3%)**，但因為花費過低 (Low Spend) 導致尚未累積足夠轉換或 0 轉換的項目。這代表素材吸睛，但沒機會表現。 -> 建議：給予獨立預算測試。
    - **高流量低轉化 (High Interest, Low Conv)**: **CTR 極高**，但 CPA 偏高或轉換率低。這代表受眾對素材極有興趣，問題可能出在「登陸頁面」或「價格」。 -> 建議：不要急著關閉，應優先優化落地頁體驗。
- **輸出重點**: 明確指出哪些是「該加預算的贏家」，哪些是「值得再給機會的潛力股」。

## 3. 止損與縮編建議 (Cost Cutting)
- **目標**: 揪出浪費預算的「黑洞」。
- **篩選標準**:
    - **無效花費**: P7D/P30D 花費高昂但 0 轉單的項目。
    - **低效能**: CPA 遠高於平均（>1.5倍），且 **CTR 低落**（表示受眾根本不買單，這才是真正的爛廣告）的項目。
    - **素材疲勞**: P30D 表現尚可，但 P7D CPA 飆升且 CTR 下滑的素材。
- **輸出重點**: 明確列出哪些應該「立即關閉」？哪些應該「縮減預算」？

## 4. 受眾動機與素材洞察 (Audience & Creative Strategy)
- **目標**: 從數據反推「為什麼這群人會買單？」。
- **執行動作**:
    - 分析表現最好的前 3-5 名素材名稱（Ad Name）與視覺/文案標籤（如：I人、媽媽、創業、上班族...）。
    - 結合 CTR 數據，解讀哪種「溝通切角（Hook）」最能打動受眾？
    - 對比不同受眾（AdSet）對同一類素材的反應差異。
- **輸出重點**: 總結出一個「受眾偏好框架」，並具體建議下一波素材該怎麼做。

## 5. 綜合戰術行動清單 (Consolidated Action Plan) - 最重要的一步
請將上述所有分析收斂為一份 **「可直接執行的操作清單」**，請務必以表格或列點方式呈現，包含以下三個維度：

### A. 開關操作 (On/Off Decisions)
- **🔴 應關閉/暫停 (Turn Off)**:
    - [素材層級]: 具體列出該關閉的爛素材名稱。
    - [架構層級]: 該暫停的受眾(AdSet)或行銷活動(Campaign)。
- **🟢 應開啟/加強 (Turn On/Scale)**:
    - [潛力股]: 建議重新啟動或給予更多機會的項目。

### B. 預算調控 (Budget Allocation)
- **💰 預算加碼**: 具體建議哪個 Campaign/AdSet 預算應該增加？增加幅度建議？
- **💸 預算縮減**: 哪個 Campaign/AdSet 預算應該砍半或縮編？

### C. 製作與優化 (Creation & Optimization)
- **🎨 素材補量**: 根據贏家素材，設計師下一波該做什麼圖？(例如：請多做幾張「I人」切角的圖、多做幾張黃色背景的圖)。
- **🎯 受眾測試**: 建議測試什麼新興趣、新版位或新方向？

# Output Format
請以專業顧問報告的形式輸出，使用粗體標示關鍵數據，並確保最後的「綜合戰術行動清單」清晰易讀，讓投手可以按表操課。
"""

# ==========================================
# 1. 輔助函數 (資料處理核心)
# ==========================================

def clean_ad_name(name):
    """移除廣告名稱中的 ' - 複本' 及後續所有內容，以便將相同創意合併。"""
    return re.sub(r' - 複本.*$', '', str(name)).strip()

def create_summary_row(df, metric_col, numerator_col, denominator_col, is_percentage=False):
    """計算加總平均列的輔助函數"""
    total_num = df[numerator_col].sum()
    total_denom = df[denominator_col].sum()
    
    if is_percentage:
        # CTR = Clicks / Impressions * 100
        avg_metric = (total_num / total_denom * 100) if total_denom > 0 else 0
    else:
        # CPA = Spend / Conv, CPC = Spend / Clicks
        avg_metric = (total_num / total_denom) if total_denom > 0 else 0
        
    summary_dict = {
        numerator_col: total_num,
        denominator_col: total_denom,
        metric_col: round(avg_metric, 2)
    }
    
    # 處理分組欄位 (非數值欄位)
    # 找出不在 [分子, 分母, 指標] 中的欄位名稱
    group_cols = [c for c in df.columns if c not in [numerator_col, denominator_col, metric_col]]
    
    # 設定第一欄為 "全帳戶平均"，其他為 "-"
    if group_cols:
        summary_dict[group_cols[0]] = '全帳戶平均'
        for col in group_cols[1:]:
            summary_dict[col] = '-'
            
    return pd.DataFrame([summary_dict])

def calculate_and_rank_metrics(df_group, metric_type, sort_ascending):
    """計算 CPA/CPC/CTR 指標並排名，並在最後加入全帳戶平均列。"""
    
    df_metrics = None
    summary_row = None

    if metric_type == 'CPA':
        # 1. 聚合
        df_metrics = df_group.agg({
            '花費金額 (TWD)': 'sum',
            'free-course': 'sum'
        }).reset_index()
        
        # 2. 過濾：排除花費為 0 的項目
        df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]

        # 3. 計算個別指標
        df_metrics['CPA (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x['free-course'] if x['free-course'] > 0 else np.nan, axis=1)
        df_metrics.replace([np.inf, -np.inf], np.nan, inplace=True)
        
        # 4. 排序 (先排序再加平均)
        df_metrics = df_metrics.sort_values(by='CPA (TWD)', ascending=sort_ascending).round(2)
        
        # 5. 計算全帳戶平均列 (僅包含有花費的項目)
        summary_row = create_summary_row(df_metrics, 'CPA (TWD)', '花費金額 (TWD)', 'free-course')

    elif metric_type == 'CPC':
        df_metrics = df_group.agg({
            '花費金額 (TWD)': 'sum',
            '連結點擊次數': 'sum'
        }).reset_index()
        
        # 2. 過濾：排除花費為 0 的項目
        df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]
        
        df_metrics['CPC (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else np.nan, axis=1)
        df_metrics.replace([np.inf, -np.inf], np.nan, inplace=True)
        
        df_metrics = df_metrics.sort_values(by='CPC (TWD)', ascending=sort_ascending).round(2)
        
        summary_row = create_summary_row(df_metrics, 'CPC (TWD)', '花費金額 (TWD)', '連結點擊次數')

    elif metric_type == 'CTR':
        # 注意：CTR 原本不需要花費，但為了過濾，必須把 '花費金額 (TWD)' 加進來聚合
        df_metrics = df_group.agg({
            '連結點擊次數': 'sum',
            '曝光次數': 'sum',
            '花費金額 (TWD)': 'sum' 
        }).reset_index()
        
        # 2. 過濾：排除花費為 0 的項目
        df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]

        df_metrics['CTR (%)'] = df_metrics.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
        
        df_metrics = df_metrics.sort_values(by='CTR (%)', ascending=sort_ascending).round(2)
        
        summary_row = create_summary_row(df_metrics, 'CTR (%)', '連結點擊次數', '曝光次數', is_percentage=True)
        
        # 在輸出前移除花費欄位 (因為是 CTR 表)
        df_metrics = df_metrics.drop(columns=['花費金額 (TWD)'])
    
    # 5. 合併：將平均列放到最下方
    if df_metrics is not None and summary_row is not None:
        return pd.concat([df_metrics, summary_row], ignore_index=True)
    
    return df_metrics

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

def to_excel_bytes(dfs_to_export, prompt_text):
    """
    將列表中的 (sheet_name, DataFrame) 寫入 Excel 文件的 BytesIO。
    同時將 Prompt 寫入第一個 '📘_AI_指令說明書' 分頁。
    """
    output = io.BytesIO()
    # 使用 xlsxwriter 引擎以支援格式設定
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- 1. 建立並寫入 AI 指令分頁 (最優先) ---
        instruction_sheet_name = '📘_AI_指令說明書'
        worksheet = workbook.add_worksheet(instruction_sheet_name)
        writer.sheets[instruction_sheet_name] = worksheet  # 註冊分頁
        
        # 設定格式：自動換行、頂部對齊
        text_format = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_size': 11})
        
        # 設定欄寬 (A欄寬一點以便閱讀)
        worksheet.set_column('A:A', 100)
        
        # 寫入指令內容
        worksheet.write('A1', prompt_text, text_format)
        
        # --- 2. 寫入其餘數據分頁 ---
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

    # 過濾掉花費為 0 的天/行銷活動
    campaign_daily_trend = campaign_daily_trend[campaign_daily_trend['花費金額 (TWD)'] > 0]

    campaign_daily_trend['CPA (TWD)'] = campaign_daily_trend.apply(lambda x: x['花費金額 (TWD)'] / x['free-course'] if x['free-course'] > 0 else np.nan, axis=1)
    campaign_daily_trend['CTR (%)'] = campaign_daily_trend.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    
    # 格式化輸出
    campaign_daily_trend['天數'] = campaign_daily_trend['天數'].dt.strftime('%Y-%m-%d')
    campaign_daily_trend.replace([np.inf, -np.inf], np.nan, inplace=True)
    
    trend_output_df = campaign_daily_trend[['天數', '行銷活動名稱', '花費金額 (TWD)', 'free-course', 'CPA (TWD)', 'CTR (%)']].round(2)
    
    st.dataframe(trend_output_df, use_container_width=True, hide_index=True)
    
    return trend_output_df


# ==========================================
# 3. Streamlit 主程式
# ==========================================

def marketing_analysis_app():
    # Page Config
    st.set_page_config(layout="wide", page_title="廣告成效智能分析工具")
    
    st.title("📊 廣告成效多週期分析工具 (AI Ready)")
    st.markdown("### 🚀 流程簡化：")
    st.info("現在，您只需下載 Excel 檔，直接上傳給 ChatGPT/Claude。**AI 分析指令已自動內建在 Excel 的第一個分頁「📘_AI_指令說明書」中**，無需再手動複製貼上。")
    
    st.markdown("---")
    st.markdown("### 步驟 1：上傳原始 CSV 進行資料處理")

    uploaded_file = st.file_uploader("上傳 CSV 檔案", type=["csv"])

    if uploaded_file is not None:
        try:
            # 讀取檔案
            df = pd.read_csv(uploaded_file)

            # --- [FIXED] 智慧欄位名稱標準化 ---
            df.columns = df.columns.str.strip()
            
            column_mapping = {
                'free course': 'free-course',
                'Free course': 'free-course',
                'Free Course': 'free-course',
                '花費金額': '花費金額 (TWD)',
                '金額': '花費金額 (TWD)'
            }
            df.rename(columns=column_mapping, inplace=True)
            
            # 檢查關鍵欄位是否存在
            required_cols = ['天數', '行銷活動名稱', 'free-course', '花費金額 (TWD)', '連結點擊次數', '曝光次數']
            missing_cols = [c for c in required_cols if c not in df.columns]
            
            if missing_cols:
                st.error(f"❌ 檔案欄位對應失敗。")
                st.error(f"找不到以下欄位: {missing_cols}")
                st.warning(f"系統目前偵測到的欄位: {list(df.columns)}")
                st.info("提示：請檢查您的 CSV 檔是否包含 'free-course' (或 'free course') 以及 '花費金額 (TWD)'。")
                st.stop()
            # --------------------------------

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

            # --- 創建 Excel 下載按鈕 (包含指令分頁) ---
            # 將 AI_CONSULTANT_PROMPT 傳入函數
            excel_data = to_excel_bytes(all_dfs_for_excel, AI_CONSULTANT_PROMPT)
            
            st.markdown("### 步驟 2：下載分析報表")
            st.download_button(
                label="📥 下載完整分析報表 (已內建 AI 指令).xlsx",
                data=excel_data,
                file_name=f"Ad_Analysis_Report_{max_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="此 Excel 檔已包含 AI 分析指令與所有週期的指標數據，直接上傳給 ChatGPT 即可。"
            )

        except Exception as e:
            st.error(f"資料處理發生錯誤，請檢查您的 CSV 檔案格式：{e}")
            st.code(str(e))

if __name__ == "__main__":
    marketing_analysis_app()
