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
import requests  # 用於 REST API 兼容模式
import json      # 用於處理 API 回傳格式

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
# ==========================================
# 0. 全域設定：AI 顧問指令（v4.0 深度細節+高階邏輯完全體）
# ==========================================
AI_CONSULTANT_PROMPT = """
# Role｜你的身份不是分析師，是「媒體採買裁判」
你是一位資深成效廣告顧問，但此任務中你**不是負責解釋數據**，
而是負責在資訊不完美的情況下，做出「可執行的媒體採買裁決」。

你的任務不是給可能性，而是：
- 判斷哪個方向是對的
- 哪些素材 / 組合該被保留、關閉、拆分或獨立
- 明確告訴我「現在該動誰、不該動誰」

請使用 **繁體中文**，語氣務實、精準、偏決策而非教學。

---

# 資料說明
系統會提供以下資料表（不一定全部齊全）：
- Daily Alerts（P1D vs P7D）
- Weekly Trends（P7D vs PP7D）
- P7D / PP7D / P30D Campaign / AdSet / Ad 表
- CPM Change Table（P7D vs PP7D vs P30D）

請在「資料可能不完整」的前提下仍做出判斷，必要時標註不確定性來源。

---

# 🔴 核心規則（非常重要）
你**不可只做分析說明**，必須完成「裁決」。
每一則廣告、每一個廣告組合，**必須被歸類到下列六種決策類型之一，而且只能選一種**。

---

## 🧭 強制決策分類（不得新增或合併類別）

### A. ✅ 方向正確的代表（Direction Proof）
定義：
- 整體 CPA 明顯優於帳戶平均或同層級中位數
- CTR / CVR 至少一項具備說服力
- 即使 CPM 偏高，仍能轉換，代表「方向是對的」

👉 意義：這是「訊息 × 受眾 × 素材」正確性的證據

---

### B. 🧩 組合表現良好（Good Combo）
定義：
- 在「目前 AdSet 結構」中相對其他素材表現穩定
- 不一定是帳戶最佳，但是該組合的健康成員

👉 意義：這個組合內部邏輯成立，可維持

---

### C. ❌ 在此組合應被關閉（Kill in This Combo）
定義：
- 在此 AdSet 中 CTR / CVR 明顯落後
- 持續吸收預算卻無法帶來對等轉換
- 拖累該組合整體 CPA

👉 注意：這代表「在這個組合該關」，**不等於素材永久報廢**

---

### D. 🕳️ 被組合掩埋的潛力素材（Buried Potential）
定義：
- CTR / CVR 不差，甚至優於平均
- 但曝光或花費明顯過低
- 同組存在歷史王者或高 CTR 吸血素材

👉 意義：素材可能好，但被系統偏食或歷史數據壓制

---

### E. 🚀 值得獨立給預算（Spin-off Candidate）
定義：
- 在有限預算或不利環境下仍能維持好 CPA
- 表現穩定，方向明確
- 具備「如果給乾淨環境可能擴量」的特徵

👉 意義：值得獨立成立新 AdSet / Campaign 測試或擴量

---

### F. 🛑 維持不動（Do Nothing / Protect）
定義：
- 表現穩定但不特別亮眼
- 屬於帳戶的安全基本盤
- 改動風險高於潛在收益

👉 意義：不要為了優化而破壞穩定現金流

---

# 📌 輸出要求（不可省略）

## 1️⃣ 帳戶層級裁決摘要
- 目前帳戶整體狀態（穩定 / 有結構問題 / 方向正確但配置錯）
- 是否存在：
  - 預算吸血鬼
  - 系統偏食（新素材被壓制）
  - 組合內部互相拖累

---

## 2️⃣ 強制決策清單（核心）
請依序列出 A → F 六類，每一類至少包含：
- 廣告 / 廣告組合名稱
- 關鍵數據（CPA / CTR / CVR / CPM）
- 為何「相對於誰」而做此判斷
- 明確動作指令（關閉 / 移出 / 獨立 / 保留）

---

## 3️⃣ 行動版待辦清單（給人直接照做）
請輸出一份可直接執行的清單，格式如下：

- [暫停] Ad X（原因：C 類，在此組合拖累 CPA）
- [拆分] Ad Y → 新 AdSet（原因：E 類，具獨立擴量潛力）
- [保留不動] AdSet Z（原因：F 類，穩定基本盤）

---

# ⚠️ 重要提醒
- 若資料不足，請說明「哪一段判斷風險較高」
- 若某素材不是爛，而是「放錯地方」，請明確指出
- 請避免模糊建議（如：可考慮、也許、可能）

你現在是裁判，不是旁白。

"""

# ==========================================
# 1. 基礎設定與字型處理
# ==========================================
st.set_page_config(page_title="廣告成效全能分析 v6.3 (Gemini 2.5 Pro + CPM)", layout="wide")

@st.cache_resource
def get_chinese_font():
    font_path = "NotoSansCJKtc-Regular.otf"
    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
    if not os.path.exists(font_path):
        try:
            with st.spinner('正在下載中文字型 (首次執行需時較久)...'):
                urllib.request.urlretrieve(url, font_path)
        except Exception:
            return None
    return fm.FontProperties(fname=font_path)

font_prop = get_chinese_font()

# ==========================================
# 2. 核心計算邏輯
# ==========================================

def clean_ad_name(name):
    return re.sub(r' - 複本.*$', '', str(name)).strip()

def create_summary_row(df, metric_cols):
    """
    metric_cols: dict
      key: 指標名稱，如 'CPA (TWD)'
      val: (numerator_col, denominator_col, multiplier)
      multiplier: 1 (純比值), 100 (百分比), 1000 (每千次，如 CPM)
    """
    summary_dict = {}
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        summary_dict[col] = df[col].sum()
        
    for metric, (num, denom, multiplier) in metric_cols.items():
        total_num = summary_dict.get(num, 0)
        total_denom = summary_dict.get(denom, 0)
        if total_denom > 0:
            val = (total_num / total_denom) * multiplier
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
    """
    對任一層級（Campaign / AdSet / Ad / Detail）：
    - 先 sum 花費 / 曝光 / 點擊 / 轉換
    - 再用 aggregated 數字算 CPA / CTR / CVR / CPM
    """
    df_metrics = df_group.agg({
        '花費金額 (TWD)': 'sum',
        conv_col: 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    }).reset_index()

    df_metrics = df_metrics[df_metrics['花費金額 (TWD)'] > 0]

    # CPA / CTR / CVR / CPM
    df_metrics['CPA (TWD)'] = df_metrics.apply(
        lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0, axis=1
    )
    df_metrics['CTR (%)'] = df_metrics.apply(
        lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1
    )
    df_metrics['CVR (%)'] = df_metrics.apply(
        lambda x: (x[conv_col] / x['連結點擊次數']) * 100 if x['連結點擊次數'] > 0 else 0, axis=1
    )
    df_metrics['CPM (TWD)'] = df_metrics.apply(
        lambda x: (x['花費金額 (TWD)'] / x['曝光次數']) * 1000 if x['曝光次數'] > 0 else 0, axis=1
    )
    
    df_metrics = df_metrics.round(2).sort_values(by='花費金額 (TWD)', ascending=False)

    metric_config = {
        'CPA (TWD)': ('花費金額 (TWD)', conv_col, 1),
        'CTR (%)': ('連結點擊次數', '曝光次數', 100),
        'CVR (%)': (conv_col, '連結點擊次數', 100),
        'CPM (TWD)': ('花費金額 (TWD)', '曝光次數', 1000)
    }
    summary_row = create_summary_row(df_metrics, metric_config)
    
    if not df_metrics.empty:
        return pd.concat([df_metrics, summary_row], ignore_index=True)
    else:
        return df_metrics

def collect_period_results(df, period_name_short, conv_col):
    df['廣告名稱_clean'] = df['廣告名稱'].apply(clean_ad_name)
    results = []
    
    # 0. 詳細層級：活動 + 組合 + 廣告
    results.append((
        f'{period_name_short}_Detail_詳細(組合+廣告)', 
        calculate_consolidated_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱', '廣告名稱']), conv_col)
    ))
    # 1. 廣告層級
    results.append(
        (f'{period_name_short}_Ad_廣告',
         calculate_consolidated_metrics(df.groupby('廣告名稱_clean'), conv_col))
    )
    # 2. 廣告組合層級（這裡也會有 CPM）
    results.append(
        (f'{period_name_short}_AdSet_廣告組合',
         calculate_consolidated_metrics(df.groupby(['行銷活動名稱', '廣告組合名稱']), conv_col))
    )
    # 3. 行銷活動層級
    results.append(
        (f'{period_name_short}_Campaign_行銷活動',
         calculate_consolidated_metrics(df.groupby('行銷活動名稱'), conv_col))
    )
    
    return results

# ==========================================
# 3. 異常偵測與趨勢分析邏輯
# ==========================================
def check_daily_anomalies(df_p1, df_p7, level_name='行銷活動名稱'):
    p1 = df_p1[df_p1[level_name] != '全帳戶平均'].copy()
    p7 = df_p7[df_p7[level_name] != '全帳戶平均'].copy()
    
    if p1.empty or p7.empty:
        return pd.DataFrame()

    merged = pd.merge(p1, p7, on=level_name, suffixes=('_P1', '_P7'), how='inner')
    alerts = []
    
    for _, row in merged.iterrows():
        if row['花費金額 (TWD)_P1'] < 200: 
            continue 

        name = row[level_name]
        cpa_p1, cpa_p7 = row['CPA (TWD)_P1'], row['CPA (TWD)_P7']
        ctr_p1, ctr_p7 = row['CTR (%)_P1'], row['CTR (%)_P7']
        spend_p1 = row['花費金額 (TWD)_P1']

        if cpa_p7 > 0 and cpa_p1 > cpa_p7 * 1.3:
            diff = int(((cpa_p1 - cpa_p7) / cpa_p7) * 100)
            alerts.append({
                '層級': level_name,
                '名稱': name,
                '類型': '🔴 CPA 暴漲', 
                '數據對比': f"昨${cpa_p1:.0f} vs 均${cpa_p7:.0f} (🔺{diff}%)",
                '建議': '檢查競價或受眾'
            })
            
        if ctr_p7 > 0 and ctr_p1 < ctr_p7 * 0.8:
            diff = int(((ctr_p7 - ctr_p1) / ctr_p7) * 100)
            alerts.append({
                '層級': level_name,
                '名稱': name,
                '類型': '📉 CTR 驟降', 
                '數據對比': f"昨{ctr_p1}% vs 均{ctr_p7}% (🔻{diff}%)",
                '建議': '素材疲乏/更換素材'
            })
            
        if cpa_p1 == 0 and spend_p1 > 500:
             alerts.append({
                 '層級': level_name,
                 '名稱': name,
                 '類型': '🛑 高花費0轉換', 
                 '數據對比': f"昨花費 ${spend_p1:.0f}",
                 '建議': '檢查落地頁/設定'
             })

    return pd.DataFrame(alerts)

def check_weekly_trends(df_p7, df_pp7, level_name='行銷活動名稱'):
    curr = df_p7[df_p7[level_name] != '全帳戶平均'].copy()
    prev = df_pp7[df_pp7[level_name] != '全帳戶平均'].copy()
    
    if curr.empty or prev.empty:
        return pd.DataFrame()
    
    merged = pd.merge(curr, prev, on=level_name, suffixes=('_This', '_Last'), how='inner')
    trends = []
    
    for _, row in merged.iterrows():
        if row['花費金額 (TWD)_This'] < 1000: 
            continue
        
        name = row[level_name]
        cpa_this, cpa_last = row['CPA (TWD)_This'], row['CPA (TWD)_Last']
        ctr_this, ctr_last = row['CTR (%)_This'], row['CTR (%)_Last']
        spend_this, spend_last = row['花費金額 (TWD)_This'], row['花費金額 (TWD)_Last']
        
        if cpa_last > 0 and cpa_this > cpa_last * 1.2:
            diff = int(((cpa_this - cpa_last) / cpa_last) * 100)
            trends.append({
                '層級': level_name,
                '名稱': name,
                '狀態': '⚠️ 成本惡化',
                '數據變化': f"${cpa_this:.0f} (vs ${cpa_last:.0f})",
                '變化幅度': f"🔺 +{diff}%",
                '診斷': '競爭加劇或轉換率下降'
            })
            
        if ctr_last > 0 and ctr_this < ctr_last * 0.85:
            diff = int(((ctr_last - ctr_this) / ctr_this) * 100) if ctr_this > 0 else 100
            trends.append({
                '層級': level_name,
                '名稱': name,
                '狀態': '📉 CTR 衰退',
                '數據變化': f"{ctr_this}% (vs {ctr_last}%)",
                '變化幅度': f"🔻 -{diff}%",
                '診斷': '素材開始老化'
            })

        if spend_last > 0 and spend_this > spend_last * 1.2:
            if cpa_last > 0 and cpa_this > cpa_last * 1.1:
                trends.append({
                    '層級': level_name,
                    '名稱': name,
                    '狀態': '💸 擴量效率差',
                    '數據變化': f"花費增至 ${spend_this:,.0f}",
                    '變化幅度': f"CPA 亦漲",
                    '診斷': '邊際效應遞減，建議暫停加碼'
                })

    return pd.DataFrame(trends)

def get_trend_data_excel(df_p30d, conv_col):
    trend_df = df_p30d.copy()
    acc_daily = trend_df.groupby(['天數']).agg({
        '花費金額 (TWD)': 'sum',
        conv_col: 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    }).reset_index()
    acc_daily['行銷活動名稱'] = '🏆 整體帳戶 (Account Overall)'
    final_trend = acc_daily[acc_daily['花費金額 (TWD)'] > 0]
    final_trend['CPA (TWD)'] = final_trend.apply(
        lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0,
        axis=1
    )
    final_trend['CPM (TWD)'] = final_trend.apply(
        lambda x: (x['花費金額 (TWD)'] / x['曝光次數']) * 1000 if x['曝光次數'] > 0 else 0,
        axis=1
    )
    final_trend['天數'] = final_trend['天數'].dt.strftime('%Y-%m-%d')
    return final_trend.round(2)

def build_cpm_change_table(p7_camp_df, pp7_camp_df, p30_camp_df):
    """
    建立行銷活動層級的 CPM 變化表：P7D / PP7D / P30D
    """
    def prep(df, suffix):
        if df is None or df.empty:
            return pd.DataFrame(columns=['行銷活動名稱', f'CPM_{suffix}', f'花費金額_{suffix}', f'曝光次數_{suffix}'])
        tmp = df.copy()
        cols_keep = ['行銷活動名稱', 'CPM (TWD)', '花費金額 (TWD)', '曝光次數']
        cols_exist = [c for c in cols_keep if c in tmp.columns]
        tmp = tmp[cols_exist]
        tmp = tmp[tmp['行銷活動名稱'].notna()]
        tmp = tmp.rename(columns={
            'CPM (TWD)': f'CPM_{suffix}',
            '花費金額 (TWD)': f'花費金額_{suffix}',
            '曝光次數': f'曝光次數_{suffix}'
        })
        return tmp

    p7 = prep(p7_camp_df, 'P7D')
    pp7 = prep(pp7_camp_df, 'PP7D')
    p30 = prep(p30_camp_df, 'P30D')

    merged = p7.merge(pp7, on='行銷活動名稱', how='outer').merge(p30, on='行銷活動名稱', how='outer')
    if merged.empty:
        return merged

    for c in ['CPM_P7D', 'CPM_PP7D', 'CPM_P30D',
              '花費金額_P7D', '花費金額_PP7D', '花費金額_P30D',
              '曝光次數_P7D', '曝光次數_PP7D', '曝光次數_P30D']:
        if c in merged.columns:
            merged[c] = merged[c].fillna(0)

    def pct_change(new, old):
        if old == 0:
            return None
        return round((new - old) / old * 100, 2)

    merged['CPM_週環比變化_vs_PP7D_(%)'] = merged.apply(
        lambda x: pct_change(x['CPM_P7D'], x['CPM_PP7D']), axis=1
    )
    merged['CPM_月度對比_vs_P30D_(%)'] = merged.apply(
        lambda x: pct_change(x['CPM_P7D'], x['CPM_P30D']), axis=1
    )

    if '花費金額_P7D' in merged.columns:
        merged = merged.sort_values('花費金額_P7D', ascending=False)

    return merged

# ==========================================
# 4. Excel 匯出函數（含 AI 回覆）
# ==========================================
def to_excel_single_sheet_stacked(dfs_list, prompt_text, ai_response=None):
    engine = 'xlsxwriter' if HAS_XLSXWRITER else None
    if not engine:
        pass

    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            sheet_name = '📘_完整分析報告'
            ws = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws
            
            fmt_prompt = workbook.add_format({
                'text_wrap': True, 'valign': 'top',
                'font_size': 10, 'bg_color': '#F0F2F6'
            })
            fmt_ai_response = workbook.add_format({
                'text_wrap': True, 'valign': 'top',
                'font_size': 11, 'bg_color': '#FFF8DC',
                'border': 1
            })
            fmt_header = workbook.add_format({
                'bold': True, 'font_size': 14,
                'font_color': '#0068C9'
            })
            fmt_table_header = workbook.add_format({
                'bold': True, 'bg_color': '#E6E6E6', 'border': 1
            })
            
            current_row = 0
            
            # 1. AI 分析結果
            if ai_response:
                ws.merge_range('A1:K1', "🤖 Gemini AI 廣告診斷報告 (AI Analysis Report)", fmt_header)
                current_row += 1
                ai_lines = ai_response.count('\n') + (len(ai_response) // 50) + 2
                ws.merge_range(current_row, 0, current_row + ai_lines, 10, ai_response, fmt_ai_response)
                current_row += ai_lines + 2
            
            # 2. System Prompt
            ws.merge_range(current_row, 0, current_row, 8, "🛠️ 系統分析指令 (System Prompt Log)", fmt_header)
            current_row += 1
            prompt_lines = prompt_text.count('\n') + 3
            ws.merge_range(current_row, 0, current_row + prompt_lines, 10, prompt_text, fmt_prompt)
            current_row += prompt_lines + 2
            
            # 3. 數據表
            for title, df in dfs_list:
                ws.write(current_row, 0, f"📌 Table: {title}", fmt_header)
                current_row += 1
                df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
                for col_num, value in enumerate(df.columns.values):
                    ws.write(current_row, col_num, value, fmt_table_header)
                current_row += len(df) + 4
                
            ws.set_column('A:A', 40)
            ws.set_column('B:Z', 15)
    except Exception:
        return None
            
    output.seek(0)
    return output.getvalue()

# ==========================================
# 5. AI 分析串接：輔助函式（多層級餵入）
# ==========================================
def safe_to_markdown(df):
    try:
        return df.to_markdown(index=False)
    except ImportError:
        return df.to_csv(sep='|', index=False)
    except Exception:
        return df.to_string(index=False)

def get_top_by_spend(df, n=20, min_spend=0):
    if df is None or df.empty:
        return df

    tmp = df.copy()

    for col in ['行銷活動名稱', '廣告名稱_clean']:
        if col in tmp.columns:
            tmp = tmp[tmp[col] != '全帳戶平均']

    if '花費金額 (TWD)' in tmp.columns:
        tmp = tmp[tmp['花費金額 (TWD)'] >= min_spend]
        tmp = tmp.sort_values('花費金額 (TWD)', ascending=False).head(n)

    return tmp

def call_gemini_analysis(
    api_key,
    alerts_daily,
    alerts_weekly,
    campaign_summary,
    adset_p7=None,
    ad_p7=None,
    trend_30d=None,
    cpm_change_table=None
):
    data_context = "\n\n# 📊 Account Data Summary（多層級視角）\n"

    data_context += "\n## 1. Daily Alerts (P1D vs P7D Anomalies)\n"
    if alerts_daily is not None and not alerts_daily.empty:
        data_context += safe_to_markdown(alerts_daily)
    else:
        data_context += "No critical daily anomalies detected."

    data_context += "\n\n## 2. Weekly Trends (P7D vs PP7D Decline)\n"
    if alerts_weekly is not None and not alerts_weekly.empty:
        data_context += safe_to_markdown(alerts_weekly)
    else:
        data_context += "No significant weekly decline trends detected."

    data_context += "\n\n## 3. Current Week Campaign Performance (P7D)\n"
    if campaign_summary is not None and not campaign_summary.empty:
        top_campaigns = get_top_by_spend(campaign_summary, n=20, min_spend=0)
        data_context += safe_to_markdown(top_campaigns)
    else:
        data_context += "No campaign-level data available."

    if adset_p7 is not None and not adset_p7.empty:
        data_context += "\n\n## 4. P7D AdSet Performance (Top by Spend)\n"
        top_adsets = get_top_by_spend(adset_p7, n=30, min_spend=500)
        if top_adsets is not None and not top_adsets.empty:
            data_context += safe_to_markdown(top_adsets)

    if ad_p7 is not None and not ad_p7.empty:
        data_context += "\n\n## 5. P7D Ad Performance (Top by Spend)\n"
        top_ads = get_top_by_spend(ad_p7, n=50, min_spend=300)
        if top_ads is not None and not top_ads.empty:
            data_context += safe_to_markdown(top_ads)

    if trend_30d is not None and not trend_30d.empty:
        data_context += "\n\n## 6. 30D Account Daily Trend (Account Overall)\n"
        data_context += safe_to_markdown(trend_30d)

    if cpm_change_table is not None and not cpm_change_table.empty:
        data_context += "\n\n## 7. CPM Change Table (P7D vs PP7D vs P30D, Campaign Level)\n"
        data_context += safe_to_markdown(cpm_change_table)

    full_prompt = (
        AI_CONSULTANT_PROMPT
        + data_context
        + "\n\n# User Request: 請根據上述多層級數據，產生一份廣告優化診斷報告，並明確指出：活動 / AdSet / 廣告層級的調整建議，特別說明 CPM 變化如何影響 CPA 與 CPC。"
    )

    with st.spinner('🤖 AI 正在分析數據中... (這可能需要 10–20 秒)'):
        try:
            if HAS_GENAI:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-2.5-pro')
                response = model.generate_content(full_prompt)
                return response.text if hasattr(response, "text") else str(response)

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
                try:
                    return result_json['candidates'][0]['content']['parts'][0]['text']
                except (KeyError, IndexError):
                    return f"⚠️ API 回傳格式不如預期: {str(result_json)}"
            else:
                return f"⚠️ API 連線錯誤 ({response.status_code}): {response.text}"

        except Exception as e:
            return f"❌ 系統發生錯誤: {str(e)}\n請檢查 API Key 是否正確，或該 Key 是否有權限存取 2.5 Pro 模型。"

# ==========================================
# 6. 主程式 UI
# ==========================================
st.title("📊 廣告成效全能分析 v6.3 (Gemini 2.5 Pro + CPM)")

if not HAS_GENAI:
    st.warning("ℹ️ 提示：未偵測到 `google-generativeai` 套件。系統將自動切換為 **REST API 兼容模式** (只需 API Key 即可運作)。")
if not HAS_XLSXWRITER:
    st.warning("⚠️ 警告：未偵測到 `xlsxwriter` 套件。Excel 匯出功能可能會失效。")

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
        
        # 側邊欄設定
        with st.sidebar:
            st.header("⚙️ 分析設定")
            
            st.subheader("🤖 AI 分析設定")
            gemini_api_key = st.text_input("Gemini API Key", type="password", placeholder="輸入 Key 以啟用 AI 分析")
            st.caption("[取得 Google AI Studio Key](https://aistudio.google.com/app/apikey)")
            st.divider()
            
            suggested_idx = 0
            for idx, col in enumerate(all_columns):
                c_low = col.lower()
                if '成本' in col or 'cost' in c_low: 
                    continue
                if ('free' in c_low and 'course' in c_low):
                    suggested_idx = idx
                    break
                if '購買' in col or 'purchase' in c_low:
                    suggested_idx = idx
                    break
                if '轉換' in col:
                    suggested_idx = idx
                    break
                
            conversion_col = st.selectbox("🎯 目標轉換欄位:", options=all_columns, index=suggested_idx)
            
            def find_col(opts, default):
                for opt in opts:
                    for col in all_columns:
                        if opt in col:
                            return col
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
        
        # 各區間 Campaign 層級
        res_p1d_camp = calculate_consolidated_metrics(df_p1d.groupby('行銷活動名稱'), conversion_col)
        res_p7d_camp = calculate_consolidated_metrics(df_p7d.groupby('行銷活動名稱'), conversion_col)
        res_pp7d_camp = calculate_consolidated_metrics(df_pp7d.groupby('行銷活動名稱'), conversion_col)
        
        # 警示與週趨勢
        alerts_daily = check_daily_anomalies(res_p1d_camp, res_p7d_camp, '行銷活動名稱')
        alerts_weekly = check_weekly_trends(res_p7d_camp, res_pp7d_camp, '行銷活動名稱')

        # 各區間多層級匯總
        res_p1 = collect_period_results(df_p1d, 'P1D', conversion_col)
        res_p7 = collect_period_results(df_p7d, 'P7D', conversion_col)
        res_pp7 = collect_period_results(df_pp7d, 'PP7D', conversion_col)
        res_p30 = collect_period_results(df_p30d, 'P30D', conversion_col)

        # P7D 多層級 DataFrame 給 AI 用
        p7_detail_df = res_p7[0][1]
        p7_ad_df     = res_p7[1][1]
        p7_adset_df  = res_p7[2][1]
        p7_camp_df   = res_p7[3][1]

        # P30D 行銷活動層級，用於 CPM 變化表
        p30_camp_df = res_p30[3][1] if len(res_p30) >= 4 else None

        # 30 日帳戶趨勢 DataFrame
        trend_30d_df = get_trend_data_excel(df_p30d, conversion_col)

        # CPM 變化表
        cpm_change_df = build_cpm_change_table(
            p7_camp_df,
            res_pp7d_camp,
            p30_camp_df
        )

        # --- UI Tabs ---
        tab1, tab2, tab3 = st.tabs(["📈 戰情室 & 雙重監控", "📑 詳細數據表 (AdSet+Ad)", "🤖 AI 深度診斷 (Gemini)"])
        
        # ========== Tab 1：戰情室 ==========
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
            total_impr = df_p30d['曝光次數'].sum()
            cpa_30d = total_spend / total_conv if total_conv > 0 else 0
            cpm_30d = (total_spend / total_impr * 1000) if total_impr > 0 else 0
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("近30日總花費", f"${total_spend:,.0f}")
            c2.metric("近30日總轉換", f"{total_conv:,.0f}")
            c3.metric("近30日平均 CPA", f"${cpa_30d:,.0f}")
            c4.metric("近30日平均 CPM", f"${cpm_30d:,.0f}")

            # 趨勢圖：花費 vs 轉換
            daily = df_p30d.groupby('天數')[['花費金額 (TWD)', conversion_col, '連結點擊次數', '曝光次數']].sum().reset_index()
            daily['日期str'] = daily['天數'].dt.strftime('%m-%d')
            
            fig, ax1 = plt.subplots(figsize=(12, 5))
            ax2 = ax1.twinx()
            ax1.bar(daily['日期str'], daily['花費金額 (TWD)'], alpha=0.6, label='花費')
            ax2.plot(daily['日期str'], daily[conversion_col], marker='o', label='轉換數', linewidth=2)
            ax1.set_xlabel('日期', fontproperties=font_prop)
            ax1.set_ylabel('花費 (TWD)', fontproperties=font_prop)
            ax2.set_ylabel('轉換數', fontproperties=font_prop)
            if font_prop:
                for label in ax1.get_xticklabels():
                    label.set_fontproperties(font_prop)
            st.pyplot(fig)

            st.divider()
            st.subheader("💰 CPM 變化概況（行銷活動層級：P7D / PP7D / P30D）")
            if cpm_change_df is not None and not cpm_change_df.empty:
                st.dataframe(cpm_change_df, use_container_width=True)
            else:
                st.info("目前無法產生 CPM 變化表（可能是資料不足或欄位不完整）。")

        # ========== Tab 2：詳細數據表 ==========
        with tab2:
            st.markdown("### 🔍 各區間詳細數據 (行銷活動 > 廣告組合 > 廣告)")
            t_p1, t_p7, t_pp7, t_p30 = st.tabs(["P1D (昨日)", "P7D (本週)", "PP7D (上週)", "P30D (月報)"])
            
            def render_data_tab(results_list, unique_key):
                st.info("💡 下表為「詳細層級」，可看到每個 行銷活動 > 廣告組合 > 廣告 的表現（含 CPA / CTR / CVR / CPM）。")
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

            with t_p1:
                render_data_tab(res_p1, "radio_p1")
            with t_p7:
                render_data_tab(res_p7, "radio_p7")
            with t_pp7:
                render_data_tab(res_pp7, "radio_pp7")
            with t_p30:
                render_data_tab(res_p30, "radio_p30")

        # ========== Tab 3：AI 深度診斷 ==========
        with tab3:
            st.header("🤖 Gemini AI 廣告成效診斷")
            st.markdown("""
AI 將依照「帳戶層級 → 行銷活動 → AdSet → 廣告 → 30 日趨勢 → CPM 變化」的多層級數據，
自動產生優化診斷報告與可執行建議，並特別說明 CPM 變化對 CPA / CPC 的影響。
            """)
            
            col_ai_btn, _ = st.columns([1, 2])
            with col_ai_btn:
                run_ai = st.button("🚀 開始 AI 智能分析", type="primary")
            
            if run_ai:
                if not gemini_api_key:
                    st.warning("⚠️ 請先於左側側邊欄輸入 Gemini API Key")
                else:
                    analysis_result = call_gemini_analysis(
                        api_key=gemini_api_key,
                        alerts_daily=alerts_daily,
                        alerts_weekly=alerts_weekly,
                        campaign_summary=p7_camp_df,
                        adset_p7=p7_adset_df,
                        ad_p7=p7_ad_df,
                        trend_30d=trend_30d_df,
                        cpm_change_table=cpm_change_df
                    )
                    st.session_state['gemini_result'] = analysis_result
            
            if st.session_state['gemini_result']:
                st.markdown("### 📝 AI 診斷報告")
                st.markdown("---")
                st.markdown(st.session_state['gemini_result'])

        # ========== 側邊欄：下載 Excel ==========
        with st.sidebar:
            st.divider()
            excel_stack = []
            excel_stack.append(('Trend_Daily_30D', trend_30d_df))
            if cpm_change_df is not None and not cpm_change_df.empty:
                excel_stack.append(('CPM_Change_P7D_PP7D_P30D', cpm_change_df))
            excel_stack.extend(res_p1)
            excel_stack.extend(res_p7)
            excel_stack.extend(res_pp7)
            excel_stack.extend(res_p30)
            
            current_ai_result = st.session_state.get('gemini_result', None)
            
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
        st.write("建議檢查：1. CSV格式是否正確 2. 是否包含轉換/花費/曝光欄位")
