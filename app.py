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
# 0. 全域設定：AI 顧問指令（含 CPM 分析）
# ==========================================
AI_CONSULTANT_PROMPT = """
# Role
你是一位資深成效廣告分析師，同時也是「媒體採買決策顧問」。
請使用繁體中文回答，語氣專業精準、條列清楚、直接給可執行決策。

# 你會拿到的資料視角
系統會依序提供數個表格，分別來自：

1. **Daily Alerts Table：P1D vs P7D**
   - 內容：昨日本帳戶各行銷活動的異常警示。
   - 功能：判斷是否有需要立刻處理／暫停／降出價的項目。

2. **Weekly Trends Table：P7D vs PP7D**
   - 內容：本週 (P7D) 相較上週 (PP7D) 的趨勢變化。
   - 功能：判斷是否有結構性變壞、擴量後效率變差。

3. **P7D Campaign Summary（行銷活動層級）**
   - 內容：本週各行銷活動的整體成效（CPA / CTR / CVR / 花費 / 轉換 / CPM）。
   - 功能：判斷誰是主力活動、誰佔用大量預算但效率不佳。

4. **P7D AdSet Performance（廣告組合層級，依花費篩選 Top N）**
   - 功能：在同一行銷活動內，判斷是否只有少數 AdSet 拖累整體成效。
   - 用途：找出應該被減碼或停掉的 AdSet、以及可以保留的穩定 AdSet。

5. **P7D Ad Performance（廣告層級，依花費篩選 Top N）**
   - 功能：判斷是否只有某幾支素材的 CTR / CPA 出問題。
   - 用途：找出素材疲乏、點擊高但不轉換的廣告、應該優先調整的廣告。

6. **30D Account Daily Trend（帳戶近 30 日日別趨勢）**
   - 功能：判斷衰退是短期波動還是已形成週期性／長期趨勢。

7. **CPM 變化表：P7D / PP7D / P30D（行銷活動層級）**
   - 內容：每個行銷活動在不同觀察期間的 CPM (TWD) 以及變化幅度。
   - 功能：判斷出價與競價壓力是否提升、哪些活動 CPM 明顯變貴但成效未同步改善。

> 所有匯總表會同時計算 CPM (每千次曝光成本)，請將 CPA / CPC / CPM 視為成本結構的一體三面來看。

---

# 分析任務要求（請務必依序完成）

## 1. 帳戶整體快速總結（3–5 行）
- 描述帳戶目前整體狀態：
  - 「偏穩定 / 輕微惡化 / 明顯惡化 / 有成長空間」。
  - 近 7 日整體 CPA 與轉換量大致狀況。
  - 若有明顯 CPM 變貴或變便宜，可簡要註記（如：整體 CPM 上升但 CTR/CVR 也有明顯改變）。
- 若樣本數偏低或資料不完整，請明講「樣本不足風險」。

---

## 2. 🚨 昨日救火清單（使用 Daily Alerts）
- 僅針對 **Daily Alerts Table** 中有異常的活動。
- 產出「救火清單」，格式示意：

  - 【層級：行銷活動】〈活動名稱〉  
    - 問題來源：Daily Alert（例如：CPA 暴漲 / CTR 驟降 / 高花費 0 轉換）
    - 關鍵數字：簡要列出昨日 vs 均值對比（CPA / CTR / 花費）
    - 建議動作（1–2 個）：
      - 例如：暫停該活動、降低預算 X%、限縮出價、暫停表現最差的廣告組合／素材

- 若沒有任何 Daily Alert，請明確寫出：「昨日沒有需要即刻救火的活動」。

---

## 3. 📉 週環比衰退診斷（使用 Weekly Trends）
- 僅針對 **Weekly Trends Table** 中「明顯惡化」的活動。
- 將活動分類（可複選）：
  1. 「擴量效率差」：花費大幅增加，CPA 變差
  2. 「素材疲乏 / CTR 衰退」：CTR 明顯下降
  3. 「轉換效率下降」：CVR 下降 / CPA 上漲

- 每個惡化活動請列出：

  - 【層級：行銷活動】〈活動名稱〉  
    - 問題來源：Weekly Trend（例如：CPA +X%，CTR -Y%，花費 +Z%）
    - 可能原因假設（2–3 點）：
      - 例如：受眾飽和、素材看膩、競價加劇、落地頁無法承接新增流量
    - 建議策略：
      - 減碼：預算縮減多少成數 / 暫停擴量
      - 重構：重切受眾、調整投放區間、只保留表現最好的一兩個 AdSet
      - 素材：新增何種類型素材（更強 CTA、強調差異化、補社會證據等）

- 若可能，請嘗試往 AdSet / Ad 層級對應，找出「最可能拖累」的組合或廣告。
- 必要時補充該活動的 CPM 變化（例如：CPM 上漲 +30%，但 CTR 沒有同步上升）。

---

## 3.5 💰 CPM 變化與成本結構連動（使用 CPM 變化表 + P7D/PP7D/P30D）
- 專門針對 CPM 做一段獨立分析，內容請包含：

  1. **CPM 變化總覽**
     - 說明：哪些活動的 CPM 在 P7D 相較於 PP7D / P30D 明顯上升或下降？
     - 可列出 3–5 個代表性活動。

  2. **對 CPA 與 CPC 的連動推論**（請分情境明講）：
     - CPM 上升 + CPA 也上升：
       - 多半是「每千次曝光變貴，且轉換效率沒有跟上」，整體成本結構惡化。
     - CPM 上升 + CPA 大致持平：
       - 代表在更貴的競價環境中，帳戶只是勉強守住，不算真正優化，長期壓力偏高。
     - CPM 上升 + CPA 反而下降：
       - 代表雖然每千次曝光變貴，但 CTR / CVR 有明顯提升，流量品質改善，是值得優先保留與觀察的區塊。
     - CPM 下降 + CPA 沒明顯改善或變差：
       - 可能只是買到更便宜但較不精準的曝光，流量品質不足。

  3. **具體建議**
     - 請點名 2–3 個「CPM 明顯變貴且 CPA 沒有改善（持平或變差）」的活動，建議：
       - 減碼預算 / 限縮受眾 / 優先調整出價策略。
     - 同時點名 2–3 個「CPM 變貴但 CPA 更好」的活動，建議：
       - 視為高品質流量來源，可作為優先保留與適度加碼的對象。

---

## 4. 🔎 AdSet / 廣告層級的「元兇定位」
- 利用 **P7D AdSet Performance** 與 **P7D Ad Performance**，針對上一步標記「有問題」的行銷活動，嘗試回答：

  - 哪些 AdSet 是主要拖累來源？（高花費 + 高 CPA / 低 CTR）
  - 哪些 AdSet 表現穩定，可保留甚至加碼？
  - 哪些廣告素材疑似疲乏（CTR 下滑）？
  - 是否出現「點擊高但不轉換」的廣告（CTR 高、CVR 低）？

- 請分段列出：

  - 【問題 AdSet / 廣告】〈名稱〉  
    - 所屬行銷活動（若能對應）
    - 關鍵指標：花費、CPA、CTR、CVR、轉換、CPM
    - 判斷：是「素材問題」、「受眾問題」或「出價／預算配置問題」的可能性較高
    - 建議動作：暫停／減碼／更換素材／改受眾／調整出價

---

## 5. 📈 擴量與加碼機會（使用 P7D Campaign + AdSet/Ad）
- 找出兩類目標：

  1. 「可加碼活動」：
     - CPA 明顯低於帳戶平均，且轉換量有一定基礎。
     - CPM 與 CPC 處於合理或偏低水準（代表買到便宜且有效的流量）。

  2. 「穩定基本盤」：
     - CPA 接近帳戶平均但轉換量穩定、波動不大。
     - CPM 波動不大，代表成本結構穩定。

- 每個候選對象請列出：

  - 【行銷活動 / AdSet】〈名稱〉  
    - 關鍵數字：CPA、CTR、CVR、CPM、花費、轉換數
    - 理由：為何認定適合加碼或當基本盤？
    - 建議加碼／調整策略：
      - 如：預算上調 20–30% 觀察 3 天、複製活動到新受眾、沿用既有素材測試其他出價策略

---

## 6. 📆 30 日趨勢觀察（使用 30D Trend）
- 利用近 30 日日別趨勢，說明：

  - 近期問題是：
    - 過去幾天才出現的短期波動？
    - 還是已連續數週的趨勢變壞？
  - 同時說明 CPA / CPM 在 30 日內的大致走勢：
    - 若 CPM 長期上升且 CPA 也上升：代表整體競價環境變貴且策略未跟上。
    - 若 CPM 長期上升但 CPA 大致持平：代表策略勉強維持，風險在累積。
    - 若 CPM 長期上升但 CPA 下降：代表流量品質提升，值得保留與加碼。
  - 對「要馬上砍」 vs 「先調整觀察」的判斷有何影響？

---

## 7. ✅ 優先級待辦清單（整合所有視角）
請用「行動優先順序」收斂為三段清單：

1. **Priority A：立即執行（今天就要動）**
   - 例如：暫停明顯虧損活動、停掉高花費 0 轉換組合、強烈建議降預算。
   - 每點請註明依據（來自：Daily / Weekly / AdSet / Ad / CPM 變化）。

2. **Priority B：本週內調整與觀察**
   - 例如：週環比惡化但尚有潛力的活動。
   - 用「測試假設 + 觀察期」寫法（先調整 3–5 天，再決定去留）。

3. **Priority C：實驗與 A/B Test 題目**
   - 例如：針對成效好活動的擴量測試、針對低 CVR 活動的落地頁優化、針對 CTR 下滑活動的素材重製。

---

# 回覆格式要求
- 使用標題與條列明確分段（例如：`## 帳戶整體狀態`、`## 昨日救火清單`、`## CPM 變化分析`）。
- 每當引用特定活動／AdSet／廣告的建議時，若能，請標註資料主要依據（Daily / Weekly / P7D Campaign / AdSet / Ad / 30D Trend / CPM 變化表）。
- 每段分析都要附帶「具體可執行動作」，避免只有描述沒有決策建議。
- 當提到成本時，請刻意區分 CPA（每次轉換成本）、CPC（每次點擊成本）、CPM（每千次曝光成本）的角色與關聯。
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
