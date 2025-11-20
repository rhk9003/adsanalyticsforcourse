# ... (上面原本的 import 和輔助函數保持不變: clean_ad_name, calculate_and_rank_metrics, collect_all_results, to_excel_bytes, display_analysis_block, display_trend_analysis) ...

# --- 3. Streamlit 主程式 (已更新) ---

def marketing_analysis_app():
    st.set_page_config(layout="wide", page_title="廣告成效智能分析工具")
    
    st.title("📊 廣告成效多週期分析工具 (AI Ready)")
    
    # ==========================================
    # 新增功能：AI 顧問指令生成區
    # ==========================================
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
    - 結合 CTR 數據，解讀哪種「溝通切角（Hook）」最能打動受眾？
    - 對比不同受眾（AdSet）對同一類素材的反應差異。
- **輸出重點**: 總結出一個「受眾偏好框架」，並具體建議下一波素材該怎麼做。

# Output Format
請以專業顧問報告的形式輸出，使用粗體標示關鍵數據，並在每個分析段落後提供具體的 **「Next Step 行動建議」**。語氣保持客觀、直指核心。
"""
        st.code(ai_consultant_prompt, language='markdown')
    
    st.markdown("---")
    st.markdown("### 步驟 2：上傳原始 CSV 進行資料處理")
    st.markdown("系統將自動依據檔案中**最新日期**，計算三個時間區間 (P7D/PP7D/P30D) 的指標排名與趨勢分析，並生成可供 AI 讀取的 Excel 報表。")

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
                label="📥 下載完整分析報表 (.xlsx)",
                data=excel_data,
                file_name=f"Ad_Analysis_Report_{max_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="包含所有週期的 CPA/CPC/CTR 排名與趨勢數據，請將此檔案提供給 AI。"
            )

        except Exception as e:
            st.error(f"資料處理發生錯誤，請檢查您的 CSV 檔案格式：{e}")
            st.code(str(e))

if __name__ == "__main__":
    marketing_analysis_app()
