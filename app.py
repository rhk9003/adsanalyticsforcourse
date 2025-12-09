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
import requests  # 新增：用於 REST API 兼容模式
import json      # 新增：用於處理 API 回傳格式

# --- 核心修正：安全引入套件以防止 App 閃退 ---
try:
    import google.generativeai as genai
    HAS_GENAI = True
except ModuleNotFoundError:
    HAS_GENAI = False

# 檢查 xlsxwriter 是否存在 (Excel 匯出需要)
try:
    import xlsxwriter
    HAS_XLSXWRITER = True
except ModuleNotFoundError:
    HAS_XLSXWRITER = False
# -------------------------------------------

# ==========================================
# 0. 全域設定：AI 顧問指令
# ==========================================
AI_CONSULTANT_PROMPT = """
# ⚠️ SYSTEM OVERRIDE: DATA LOADING INSTRUCTION
**To the AI Analysis Model:**
The user has provided summary data tables from an advertising account.
1. **Data Source**: Summarized metrics from P1D (Yesterday), P7D (Last 7 Days), and comparison vs PP7D (Previous Period).
2. **Key Comparisons**: 
   - **P1D vs P7D**: Immediate daily alerts.
   - **P7D vs PP7D**: Week-over-Week trend analysis.

# Role
你是一位資深成效廣告分析師。請使用繁體中文回答。

# Analysis Requirements
## 1. 🚨 P1D 緊急異常 (Daily Alert)
- 檢查 **P1D (昨日)** 相較於 **P7D (均值)** 是否有 CPA 暴漲 (>30%) 或 CTR 驟降 (>20%)。
- 這是「救火」層級，請優先指出需要立即關閉或檢查的廣告。

## 2. 📉 P7D vs PP7D 週環比分析 (WoW Trend)
- 對比 **P7D (本週)** 與 **PP7D (上週)**。
- 找出 CPA 變高、CVR 變低的「衰退行銷活動」。
- 若本週花費增加但 ROAS/CPA 變差，請標記為「擴量失敗 (Inefficient Scaling)」。
- 若本週 CTR 提升但 CVR 下降，請標記為「流量品質變差 (Traffic Quality Drop)」。

## 3. 綜合優化建議
- 針對衰退項目提出具體假設（素材疲乏？競價激烈？受眾飽和？）。
- 請條列式給出具體的調整建議（例如：暫停廣告、更換受眾、優化落地頁）。
"""

# ==========================================
# 1. 基礎設定與字型處理
# ==========================================
st.set_page_config(page_title="廣告成效全能分析 v6.2 (Gemini 2.5 Pro)", layout="wide")

@st.cache_resource
def get_chinese_font():
    font_path = "NotoSansCJKtc-Regular.otf"
    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            with st.spinner('正在下載中文字型 (首次執行需時較久)...'):
                urllib.request.urlretrieve(url, font_path)
        except Exception as e:
            return None
    return fm.FontProperties(fname=font_path)

font_prop = get_chinese_font()

# ==========================================
# 2. 核心計算邏輯
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

def calculate_consolidated_metrics(df_group, conv_col):
    df_metrics = df_group.agg({
        '花費金額 (TWD)': 'sum',
        conv_col: 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    }).reset_index()

    df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]

    df_metrics['CPA (TWD)'] = df_metrics.apply(lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0, axis=1)
    df_metrics['CTR (%)'] = df_metrics.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    df_metrics['CVR (%)'] = df_metrics.apply(lambda x: (x[conv_col] / x['連結點擊次數']) * 100 if x['連結點擊次數'] > 0 else 0, axis=1)
    
    df_metrics = df_metrics.round(2).sort_values(by='花費金額 (TWD)', ascending=False)

    metric_config = {
        'CPA (TWD)': ('花費金額 (TWD)', conv_col, False),
        'CTR (%)': ('連結點擊次數', '曝光次數', True),
        'CVR (%)': (conv_col, '連結點擊次數', True)
    }
    summary_row = create_summary_row(df_metrics, metric_config)
    
    if not df_metrics.empty:
        return pd.concat([df_metrics, summary_row], ignore_index=True)
    else:
        return df_metrics

def collect_period_results(df, period_name_short, conv_col):
    df['廣告名稱_clean'] = df['廣告名稱'].apply(clean_ad_name)
    results = []
    
    # 0. 詳細層級
    results.append((
        f'{period_name_short}_Detail_詳細(組合+廣告)', 
        calculate_consolidated_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱', '廣告名稱']), conv_col)
    ))
    # 1. 廣告層級
    results.append((f'{period_name_short}_Ad_廣告', calculate_consolidated_metrics(df.groupby('廣告名稱_clean'), conv_col)))
    # 2. 廣告組合層級
    results.append((f'{period_name_short}_AdSet_廣告組合', calculate_consolidated_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), conv_col)))
    # 3. 行銷活動層級
    results.append((f'{period_name_short}_Campaign_行銷活動', calculate_consolidated_metrics(df.groupby('行銷活動名稱'), conv_col)))
    
    return results

# ==========================================
# 3. 異常偵測與趨勢分析邏輯
# ==========================================
def check_daily_anomalies(df_p1, df_p7, level_name='行銷活動名稱'):
    p1 = df_p1[df_p1[level_name] != '全帳戶平均'].copy()
    p7 = df_p7[df_p7[level_name] != '全帳戶平均'].copy()
    
    if p1.empty or p7.empty: return pd.DataFrame()

    merged = pd.merge(p1, p7, on=level_name, suffixes=('_P1', '_P7'), how='inner')
    alerts = []
    
    for _, row in merged.iterrows():
        if row['花費金額 (TWD)_P1'] < 200: continue 

        name = row[level_name]
        cpa_p1, cpa_p7 = row['CPA (TWD)_P1'], row['CPA (TWD)_P7']
        ctr_p1, ctr_p7 = row['CTR (%)_P1'], row['CTR (%)_P7']
        spend_p1 = row['花費金額 (TWD)_P1']

        if cpa_p7 > 0 and cpa_p1 > cpa_p7 * 1.3:
            diff = int(((cpa_p1 - cpa_p7) / cpa_p7) * 100)
            alerts.append({'層級': level_name, '名稱': name, '類型': '🔴 CPA 暴漲', 
                           '數據對比': f"昨${cpa_p1:.0f} vs 均${cpa_p7:.0f} (🔺{diff}%)", '建議': '檢查競價或受眾'})
            
        if ctr_p7 > 0 and ctr_p1 < ctr_p7 * 0.8:
            diff = int(((ctr_p7 - ctr_p1) / ctr_p7) * 100)
            alerts.append({'層級': level_name, '名稱': name, '類型': '📉 CTR 驟降', 
                           '數據對比': f"昨{ctr_p1}% vs 均{ctr_p7}% (🔻{diff}%)", '建議': '素材疲乏/更換素材'})
            
        if cpa_p1 == 0 and spend_p1 > 500:
             alerts.append({'層級': level_name, '名稱': name, '類型': '🛑 高花費0轉換', 
                            '數據對比': f"昨花費 ${spend_p1:.0f}", '建議': '檢查落地頁/設定'})

    return pd.DataFrame(alerts)

def check_weekly_trends(df_p7, df_pp7, level_name='行銷活動名稱'):
    curr = df_p7[df_p7[level_name] != '全帳戶平均'].copy()
    prev = df_pp7[df_pp7[level_name] != '全帳戶平均'].copy()
    
    if curr.empty or prev.empty: return pd.DataFrame()
    
    merged = pd.merge(curr, prev, on=level_name, suffixes=('_This', '_Last'), how='inner')
    trends = []
    
    for _, row in merged.iterrows():
        if row['花費金額 (TWD)_This'] < 1000: continue
        
        name = row[level_name]
        cpa_this, cpa_last = row['CPA (TWD)_This'], row['CPA (TWD)_Last']
        ctr_this, ctr_last = row['CTR (%)_This'], row['CTR (%)_Last']
        spend_this, spend_last = row['花費金額 (TWD)_This'], row['花費金額 (TWD)_Last']
        
        if cpa_last > 0 and cpa_this > cpa_last * 1.2:
            diff = int(((cpa_this - cpa_last) / cpa_last) * 100)
            trends.append({
                '層級': level_name, '名稱': name, '狀態': '⚠️ 成本惡化',
                '數據變化': f"${cpa_this:.0f} (vs ${cpa_last:.0f})",
                '變化幅度': f"🔺 +{diff}%",
                '診斷': '競爭加劇或轉換率下降'
            })
            
        if ctr_last > 0 and ctr_this < ctr_last * 0.85:
            diff = int(((ctr_last - ctr_this) / ctr_last) * 100)
            trends.append({
                '層級': level_name, '名稱': name, '狀態': '📉 CTR 衰退',
                '數據變化': f"{ctr_this}% (vs {ctr_last}%)",
                '變化幅度': f"🔻 -{diff}%",
                '診斷': '素材開始老化'
            })

        if spend_last > 0 and spend_this > spend_last * 1.2:
            if cpa_last > 0 and cpa_this > cpa_last * 1.1:
                trends.append({
                    '層級': level_name, '名稱': name, '狀態': '💸 擴量效率差',
                    '數據變化': f"花費增至 ${spend_this:,.0f}",
                    '變化幅度': f"CPA 亦漲",
                    '診斷': '邊際效應遞減，建議暫停加碼'
                })

    return pd.DataFrame(trends)

def get_trend_data_excel(df_p30d, conv_col):
    trend_df = df_p30d.copy()
    acc_daily = trend_df.groupby(['天數']).agg({
        '花費金額 (TWD)': 'sum', conv_col: 'sum', '連結點擊次數': 'sum', '曝光次數': 'sum'
    }).reset_index()
    acc_daily['行銷活動名稱'] = '🏆 整體帳戶 (Account Overall)'
    final_trend = acc_daily[acc_daily['花費金額 (TWD)'] > 0]
    final_trend['CPA (TWD)'] = final_trend.apply(lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0, axis=1)
    final_trend['天數'] = final_trend['天數'].dt.strftime('%Y-%m-%d')
    return final_trend.round(2)

# 修改：Excel 匯出函數增加 ai_response 參數
def to_excel_single_sheet_stacked(dfs_list, prompt_text, ai_response=None):
    # 檢查 xlsxwriter 引擎是否可用
    engine = 'xlsxwriter' if HAS_XLSXWRITER else None
    if not engine:
        # 如果沒有 xlsxwriter，回退到預設或拋出警告
        # 這裡為了簡單，我們假設使用者會安裝。如果真的沒有，pandas 可能會報錯或使用 openpyxl
        pass

    output = io.BytesIO()
    # 使用 engine 參數
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            sheet_name = '📘_完整分析報告'
            ws = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws
            
            fmt_prompt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_size': 10, 'bg_color': '#F0F2F6'})
            fmt_ai_response = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_size': 11, 'bg_color': '#FFF8DC', 'border': 1})
            fmt_header = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#0068C9'})
            fmt_table_header = workbook.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1})
            
            current_row = 0
            
            # 1. 寫入 AI 分析結果 (如果有的話)
            if ai_response:
                ws.merge_range('A1:K1', "🤖 Gemini AI 廣告診斷報告 (AI Analysis Report)", fmt_header)
                current_row += 1
                # 估算行數 (粗略估計每行 50 字)
                ai_lines = ai_response.count('\n') + (len(ai_response) // 50) + 2
                ws.merge_range(current_row, 0, current_row + ai_lines, 10, ai_response, fmt_ai_response)
                current_row += ai_lines + 2
            
            # 2. 寫入 System Prompt (留底用)
            ws.merge_range(current_row, 0, current_row, 8, "🛠️ 系統分析指令 (System Prompt Log)", fmt_header)
            current_row += 1
            prompt_lines = prompt_text.count('\n') + 3
            ws.merge_range(current_row, 0, current_row + prompt_lines, 10, prompt_text, fmt_prompt)
            current_row += prompt_lines + 2
            
            # 3. 寫入所有數據表
            for title, df in dfs_list:
                ws.write(current_row, 0, f"📌 Table: {title}", fmt_header)
                current_row += 1
                df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
                for col_num, value in enumerate(df.columns.values):
                    ws.write(current_row, col_num, value, fmt_table_header)
                current_row += len(df) + 4
                
            ws.set_column('A:A', 40)
            ws.set_column('B:Z', 15)
    except Exception as e:
        # 如果 Excel 寫入失敗 (例如缺少 xlsxwriter)，回傳空 byte 或錯誤提示
        return None
            
    output.seek(0)
    return output.getvalue()

# ==========================================
# 4. 新增功能：Gemini AI 分析串接 (雙模式：SDK / REST API)
# ==========================================

# 新增輔助函數：安全地將 DataFrame 轉換為文字格式，避免缺少 tabulate 報錯
def safe_to_markdown(df):
    """
    嘗試使用 markdown 格式，如果缺少 tabulate 套件，則回退到 Pipe 分隔的 CSV 格式。
    LLM 都能理解這兩種格式。
    """
    try:
        return df.to_markdown(index=False)
    except ImportError:
        # 如果沒有 tabulate，手動轉為類似 Markdown 的格式 (Pipe 分隔)
        # 這裡使用 to_csv 並用 '|' 分隔，效果跟 Markdown 很像
        return df.to_csv(sep='|', index=False)
    except Exception:
        # 最後的防線：直接轉字串
        return df.to_string(index=False)

def call_gemini_analysis(api_key, alerts_daily, alerts_weekly, campaign_summary):
    # 準備 Prompt (兩種模式共用)
    data_context = "\n\n# 📊 Account Data Summary\n"
    data_context += "## 1. Daily Alerts (P1D vs P7D Anomalies)\n"
    if not alerts_daily.empty:
        # 使用安全的轉換函數
        data_context += safe_to_markdown(alerts_daily)
    else:
        data_context += "No critical daily anomalies detected."
        
    data_context += "\n\n## 2. Weekly Trends (P7D vs PP7D Decline)\n"
    if not alerts_weekly.empty:
        # 使用安全的轉換函數
        data_context += safe_to_markdown(alerts_weekly)
    else:
        data_context += "No significant weekly decline trends detected."
        
    data_context += "\n\n## 3. Current Week Campaign Performance (P7D)\n"
    # 使用安全的轉換函數
    data_context += safe_to_markdown(campaign_summary.head(10))
    
    full_prompt = AI_CONSULTANT_PROMPT + data_context + "\n\n# User Request: 請根據上述數據，產生一份廣告優化診斷報告。"

    with st.spinner('🤖 AI 正在分析數據中... (這可能需要 10-20 秒)'):
        try:
            # 模式 A: 使用官方 SDK (如果已安裝)
            if HAS_GENAI:
                genai.configure(api_key=api_key)
                # 修改點：更換模型為 gemini-2.5-pro
                model = genai.GenerativeModel('gemini-2.5-pro')
                response = model.generate_content(full_prompt)
                return response.text
            
            # 模式 B: 使用 REST API (Fallback 模式)
            else:
                # 修改點：更換模型為 gemini-2.5-pro
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key={api_key}"
                headers = {'Content-Type': 'application/json'}
                data = {
                    "contents": [{
                        "parts": [{"text": full_prompt}]
                    }]
                }
                
                response = requests.post(url, headers=headers, json=data)
                
                if response.status_code == 200:
                    result_json = response.json()
                    # 安全地解析 JSON 回傳結構
                    try:
                        return result_json['candidates'][0]['content']['parts'][0]['text']
                    except (KeyError, IndexError):
                        return f"⚠️ API 回傳格式不如預期: {str(result_json)}"
                else:
                    return f"⚠️ API 連線錯誤 ({response.status_code}): {response.text}"
                
        except Exception as e:
            return f"❌ 系統發生錯誤: {str(e)}\n請檢查 API Key 是否正確，或該 Key 是否有權限存取 2.5 Pro 模型。"

# ==========================================
# 5. 主程式 UI
# ==========================================
st.title("📊 廣告成效全能分析 v6.2 (Gemini 2.5 Pro)")

# 顯示環境警告 (如果缺少關鍵套件)
if not HAS_GENAI:
    st.warning("ℹ️ 提示：未偵測到 `google-generativeai` 套件。系統將自動切換為 **REST API 兼容模式** (只需 API Key 即可運作)。")
if not HAS_XLSXWRITER:
    st.warning("⚠️ 警告：未偵測到 `xlsxwriter` 套件。Excel 匯出功能可能會失效。")

# 初始化 Session State
if 'gemini_result' not in st.session_state:
    st.session_state['gemini_result'] = None

uploaded_file = st.file_uploader("請上傳 CSV 報表檔案", type=['csv'])

if uploaded_file is not None:
    try:
        # 1. 讀取與欄位偵測
        try:
            df = pd.read_csv(uploaded_file, encoding='utf-8')
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding='cp950')
        except Exception as e:
            st.error(f"檔案讀取未知的錯誤: {e}")
            st.stop()

        df.columns = df.columns.str.strip()
        all_columns = df.columns.tolist()
        
        with st.sidebar:
            st.header("⚙️ 分析設定")
            
            st.subheader("🤖 AI 分析設定")
            gemini_api_key = st.text_input("Gemini API Key", type="password", placeholder="輸入 Key 以啟用 AI 分析")
            st.caption("[取得 Google AI Studio Key](https://aistudio.google.com/app/apikey)")
            st.divider()
            
            suggested_idx = 0
            for idx, col in enumerate(all_columns):
                c_low = col.lower()
                if '成本' in col or 'cost' in c_low: continue
                if ('free' in c_low and 'course' in c_low): suggested_idx = idx; break
                if '購買' in col or 'purchase' in c_low: suggested_idx = idx; break
                if '轉換' in col: suggested_idx = idx; break
                
            conversion_col = st.selectbox("🎯 目標轉換欄位:", options=all_columns, index=suggested_idx)
            
            def find_col(opts, default):
                for opt in opts:
                    for col in all_columns:
                        if opt in col: return col
                return default

            spend_col = find_col(['花費金額 (TWD)', '花費', '金額'], '花費金額 (TWD)')
            clicks_col = find_col(['連結點擊次數', '連結點擊'], '連結點擊次數')
            impressions_col = find_col(['曝光次數', '曝光'], '曝光次數')

        # 2. 數據清洗
        cols_to_numeric = [spend_col, clicks_col, impressions_col, conversion_col]
        for col in cols_to_numeric:
            if col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).str.replace(',', '', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        if '天數' not in df.columns:
             st.error("錯誤：CSV 檔案中找不到「天數」欄位，請檢查檔案格式。")
             st.stop()

        df['天數'] = pd.to_datetime(df['天數'], errors='coerce')
        df = df.dropna(subset=['天數']) 

        df_std = df.rename(columns={
            spend_col: '花費金額 (TWD)',
            clicks_col: '連結點擊次數',
            impressions_col: '曝光次數'
        })
        
        # 3. 日期區間與資料分組
        if df_std.empty:
            st.error("錯誤：資料經過清洗後為空，請檢查原始檔案是否包含有效的日期與數據。")
            st.stop()

        max_date = df_std['天數'].max().normalize()
        today = max_date + timedelta(days=1)
        
        # P1D / P7D / PP7D / P30D
        p1d_start = max_date
        df_p1d = df_std[df_std['天數'] == p1d_start].copy()
        
        p7d_start = today - timedelta(days=7)
        p7d_end = today - timedelta(days=1)
        pp7d_start = p7d_start - timedelta(days=7)
        pp7d_end = p7d_start - timedelta(days=1)
        p30d_start = today - timedelta(days=30)
        p30d_end = today - timedelta(days=1)
        
        df_p7d = df_std[(df_std['天數'] >= p7d_start) & (df_std['天數'] <= p7d_end)].copy()
        df_pp7d = df_std[(df_std['天數'] >= pp7d_start) & (df_std['天數'] <= pp7d_end)].copy()
        df_p30d = df_std[(df_std['天數'] >= p30d_start) & (df_std['天數'] <= p30d_end)].copy()
        
        res_p1d_camp = calculate_consolidated_metrics(df_p1d.groupby('行銷活動名稱'), conversion_col)
        res_p7d_camp = calculate_consolidated_metrics(df_p7d.groupby('行銷活動名稱'), conversion_col)
        res_pp7d_camp = calculate_consolidated_metrics(df_pp7d.groupby('行銷活動名稱'), conversion_col)
        
        alerts_daily = check_daily_anomalies(res_p1d_camp, res_p7d_camp, '行銷活動名稱')
        alerts_weekly = check_weekly_trends(res_p7d_camp, res_pp7d_camp, '行銷活動名稱')

        # --- UI 呈現 ---
        tab1, tab2, tab3 = st.tabs(["📈 戰情室 & 雙重監控", "📑 詳細數據表 (AdSet+Ad)", "🤖 AI 深度診斷 (Gemini)"])
        
        with tab1:
            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("🚨 P1D 緊急警示 (昨日 vs 均值)")
                if not alerts_daily.empty:
                    st.dataframe(alerts_daily, hide_index=True, use_container_width=True)
                else:
                    st.success("昨日表現平穩 (無 CPA暴漲 / CTR驟降)")
            
            with col_b:
                st.subheader("📉 P7D 週環比衰退 (本週 vs 上週)")
                if not alerts_weekly.empty:
                    st.dataframe(alerts_weekly, hide_index=True, use_container_width=True)
                else:
                    st.info("本週無顯著衰退項目 (CPA與CTR皆穩定)")

            st.divider()
            # 30日概況
            total_spend = df_p30d['花費金額 (TWD)'].sum()
            total_conv = df_p30d[conversion_col].sum()
            cpa_30d = total_spend / total_conv if total_conv > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("近30日總花費", f"${total_spend:,.0f}")
            c2.metric(f"近30日總轉換", f"{total_conv:,.0f}")
            c3.metric("近30日平均 CPA", f"${cpa_30d:,.0f}")
            
            # 趨勢圖
            daily = df_p30d.groupby('天數')[['花費金額 (TWD)', conversion_col, '連結點擊次數', '曝光次數']].sum().reset_index()
            daily['日期str'] = daily['天數'].dt.strftime('%m-%d')
            
            fig, ax1 = plt.subplots(figsize=(12, 5))
            ax2 = ax1.twinx()
            ax1.bar(daily['日期str'], daily['花費金額 (TWD)'], color='#ddd', label='花費', alpha=0.6)
            ax2.plot(daily['日期str'], daily[conversion_col], color='red', marker='o', label='轉換數', linewidth=2)
            ax1.set_xlabel('日期', fontproperties=font_prop)
            ax1.set_ylabel('花費 (TWD)', fontproperties=font_prop)
            ax2.set_ylabel('轉換數', fontproperties=font_prop)
            if font_prop:
                for label in ax1.get_xticklabels(): label.set_fontproperties(font_prop)
            st.pyplot(fig)

        with tab2:
            st.markdown("### 🔍 各區間詳細數據 (行銷活動 > 廣告組合 > 廣告)")
            t_p1, t_p7, t_pp7, t_p30 = st.tabs(["P1D (昨日)", "P7D (本週)", "PP7D (上週)", "P30D (月報)"])
            
            res_p1 = collect_period_results(df_p1d, 'P1D', conversion_col)
            res_p7 = collect_period_results(df_p7d, 'P7D', conversion_col)
            res_pp7 = collect_period_results(df_pp7d, 'PP7D', conversion_col)
            res_p30 = collect_period_results(df_p30d, 'P30D', conversion_col)
            
            def render_data_tab(results_list, unique_key):
                st.info("💡 下表已展開為「詳細層級」，您可看到每個行銷活動 > 廣告組合 下的各別廣告表現。")
                st.dataframe(results_list[0][1], use_container_width=True)
                
                with st.expander("查看其他匯總層級 (行銷活動 / 廣告組合 / 廣告整體)"):
                    view_mode = st.radio(
                        "選擇其他檢視層級:", 
                        ["行銷活動 (Campaign)", "廣告組合 (AdSet)", "廣告 (Ad)"],
                        horizontal=True,
                        key=unique_key
                    )
                    if view_mode == "行銷活動 (Campaign)":
                        st.dataframe(results_list[3][1], use_container_width=True)
                    elif view_mode == "廣告組合 (AdSet)":
                        st.dataframe(results_list[2][1], use_container_width=True)
                    elif view_mode == "廣告 (Ad)":
                        st.dataframe(results_list[1][1], use_container_width=True)

            with t_p1: render_data_tab(res_p1, "radio_p1")
            with t_p7: render_data_tab(res_p7, "radio_p7")
            with t_pp7: render_data_tab(res_pp7, "radio_pp7")
            with t_p30: render_data_tab(res_p30, "radio_p30")

        # === Tab 3: AI 分析區塊 ===
        with tab3:
            st.header("🤖 Gemini AI 廣告成效診斷")
            st.markdown("""
            AI 將根據 **每日警示 (Daily Alerts)**、**週趨勢 (Weekly Trends)** 與 **本週行銷活動 (P7D Campaign)** 數據，
            自動依照左側設定的「AI 顧問指令」進行診斷並提供優化建議。
            """)
            
            col_ai_btn, col_ai_warn = st.columns([1, 2])
            with col_ai_btn:
                # 即使沒安裝套件，現在也允許按下按鈕（會使用 REST API Fallback）
                run_ai = st.button("🚀 開始 AI 智能分析", type="primary")
            
            if run_ai:
                if not gemini_api_key:
                    st.warning("⚠️ 請先於左側側邊欄輸入 Gemini API Key")
                else:
                    analysis_result = call_gemini_analysis(
                        gemini_api_key, 
                        alerts_daily, 
                        alerts_weekly, 
                        res_p7d_camp
                    )
                    # 關鍵：將結果存入 Session State，確保切換 Tab 或點擊下載時內容不消失
                    st.session_state['gemini_result'] = analysis_result
            
            # 顯示分析結果 (如果存在)
            if st.session_state['gemini_result']:
                 st.markdown("### 📝 AI 診斷報告")
                 st.markdown("---")
                 st.markdown(st.session_state['gemini_result'])

        # 下載區 (維持並增強功能)
        with st.sidebar:
            st.divider()
            excel_stack = []
            excel_stack.append(('Trend_Daily', get_trend_data_excel(df_p30d, conversion_col)))
            excel_stack.extend(res_p1)
            excel_stack.extend(res_p7)
            excel_stack.extend(res_pp7)
            excel_stack.extend(res_p30)
            
            # 從 Session State 獲取最新的 AI 分析結果 (如果有的話)
            current_ai_result = st.session_state.get('gemini_result', None)
            
            # 傳入 AI 結果到 Excel 生成函數
            excel_bytes = to_excel_single_sheet_stacked(excel_stack, AI_CONSULTANT_PROMPT, current_ai_result)
            
            if excel_bytes:
                button_label = "📥 下載完整分析報表"
                if current_ai_result:
                    button_label += " (已包含 AI 診斷)"
                
                st.download_button(
                    label=button_label,
                    data=excel_bytes,
                    file_name=f"Full_Report_{max_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Excel 產生失敗，請檢查 xlsxwriter 套件是否安裝。")

    except Exception as e:
        st.error(f"系統發生未預期的錯誤: {e}")
        st.write("建議檢查：1. CSV格式是否正確 2. 是否包含轉換/花費欄位")
