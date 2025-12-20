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
# Role
你是一位資深成效廣告分析師，同時也是「媒體採買決策顧問」。
你的分析風格必須兼具**「數據顆粒度的細膩拆解」**與**「預算配置的戰略判斷」**。
請使用繁體中文回答，語氣專業精準、條列清楚、直接給可執行決策。

# 資料來源說明
系統會提供多個表格（Daily Alerts, Weekly Trends, P7D Campaign/AdSet/Ad, 30D Trend, CPM Change）。
請綜合這些數據進行分析。

---

# 分析任務要求（請務必依序完成，不可省略細節）

## 1. 帳戶整體快速總結 & 風險預警
- **整體狀態**：描述帳戶目前是「偏穩定 / 輕微惡化 / 明顯惡化 / 有成長空間」。
- **數據概覽**：近 7 日整體 CPA 與轉換量的大致水位。
- **【關鍵偵測】**：請直接點出帳戶中是否存在**「預算吸血鬼」**（高花費、高 CTR 但低 CVR 的素材）或**「新舊素材預算排擠」**現象？這是否為當前成效受阻的主因？
- 若樣本數偏低，請標註「樣本不足風險」。

---

## 2. 🚨 昨日救火清單 (Daily Alerts)
- 僅針對 **Daily Alerts Table** 中有異常的活動。
- 格式：
  - 【層級：行銷活動】〈名稱〉
    - 問題來源：Daily Alert（CPA 暴漲 / CTR 驟降 / 高花費 0 轉換）
    - 關鍵數字：昨日 vs 均值對比
    - **急救指令**：暫停 / 降預算 / 檢查設定（請給出明確動作）

---

## 3. 📉 週環比衰退診斷 (Weekly Trends)
- 針對 **Weekly Trends Table** 中「明顯惡化」的活動，請依據數據特徵分類（可複選）：
  1. **「擴量效率差」**：花費大幅增加，CPA 同步變差（邊際效益遞減）。
  2. **「素材疲乏 / CTR 衰退」**：CTR 明顯下降，導致 CPC 變貴。
  3. **「轉換效率下降」**：CTR 持平，但 CVR 下降（落地頁或受眾意圖問題）。
- 每個惡化活動請給出具體建議（減碼 / 重構 / 換素材）。

---

## 3.5 💰 CPM 變化與成本結構連動（核心洞察）
- 結合 **CPM 變化表** 與 **P7D/30D 數據**，分析競價環境對 CPA 的影響。請依照以下情境邏輯進行推論：

  1. **CPM 上升 + CPA 也上升**：
     - 診斷：競價變貴且轉化未跟上，成本結構惡化。建議檢查是否受眾過窄或競爭加劇。
  2. **CPM 上升 + CPA 持平/下降**：
     - 診斷：**高品質流量**。雖然貴但受眾精準（CVR 高），是值得保護的黃金區塊。
  3. **CPM 下降 + CPA 沒改善/變差**：
     - 診斷：**劣質流量陷阱**。買到了便宜曝光，但受眾不買單（CVR 低）。建議排除特定版位或緊縮受眾。
  4. **CPM 下降 + CPA 改善**：
     - 診斷：市場紅利或素材中了，應考慮擴量。

---

## 4. 🩸 深度診斷：預算效率與元兇定位 (AdSet & Ad Level)
**這是最重要的段落。請利用 P7D AdSet/Ad 表格，執行「微觀偵測」：**

1.  **偵測「預算吸血鬼」(Vampire Creatives)**：
    - 找出花費排名前 20% 的素材中，是否有 **「CTR 高 (吸睛) 但 CVR 顯著低於平均」** 的廣告？
    - **診斷**：它造成了「高點擊假象」，騙取了系統預算。**建議動作：立即暫停。**

2.  **偵測「系統偏食症」(System Bias / Cannibalization)**：
    - 檢查同一 AdSet 內，是否有 **「新素材 (如 202512xx)」CPA 優於「舊素材」，但花費卻遠低於舊素材**？
    - **診斷**：舊素材憑藉歷史數據霸佔預算，導致新素材無法發揮。**建議動作：暫停同組內的舊素材，強迫預算流向新素材。**

3.  **One Bad Apple (害群之馬) 理論**：
    - 當某個 AdSet CPA 過高時，檢查是否 **「只有一支爛廣告在拖累」**？
    - **診斷**：若是，**建議「關閉該廣告」而非「關閉整個 AdSet」**；若全體廣告都差，才建議關閉 AdSet。

---

## 5. 📈 擴量與加碼機會 (Scaling)
- 找出兩類目標：
  1. **「可加碼潛力股」**：CPA 低於帳戶平均，且預算佔比尚低（通常是被埋沒的新素材或新受眾）。
  2. **「穩定基本盤」**：CPA 穩定、量體大的舊活動。
- 建議：明確指出哪個 AdSet/廣告 值得加碼，以及加碼的方式（直接加預算 / 獨立出來開新活動）。

---

## 6. ✅ 優先級待辦清單 (Action Plan)
請將所有分析收斂為三類具體指令，並**註明判斷依據**：

1.  **Priority A：止血與清創（立即執行）**
    - 針對「預算吸血鬼」、「高花費 0 轉換」與「CPA 嚴重超標」項目的處決指令。
    - **指令格式**：`[暫停]` 廣告 X（依據：吸血鬼素材，高點擊低轉換）

2.  **Priority B：導流與優化（資源重分配）**
    - 針對「資源錯置」與「系統偏食」的修正。
    - **指令格式**：`[暫停]` AdSet Y 中的舊素材 A，`[保留]` 新素材 B（依據：CPA B < A，強迫導流測試新素材）

3.  **Priority C：保護基本盤（請勿更動）**
    - 點名那些「雖然舊但很穩」的黃金素材/受眾。
    - **指令格式**：`[維持]` AdSet Z（依據：穩定獲利來源，勿因擴量測試而干擾）

---


## 7. 🧪 新項目帶動判斷（必填）
請根據系統提供的「New Creatives Summary」與「New AdSets Summary」回答兩題，必須引用數字：

1) 新素材是否有帶動整體成長？
- 結論：有 / 沒有 / 不確定（資料不足）
- 依據：新素材的「轉換占比(%)、花費占比(%)、CPA vs 全帳戶 P7D CPA」並引用數字
- 動作：加碼 / 保留觀察 / 淘汰 / 拆分獨立

2) 新廣告組合是否有帶動成長？
- 結論：有 / 沒有 / 不確定
- 依據：PP7D→P7D 花費變化（新組合判定）、轉換占比(%)、CPA 表現（引用數字）
- 動作：擴量 / 拆分獨立 / 停止測試

---

# 回覆格式要求
- 必須使用標題與條列明確分段。
- 每一項建議都必須有**數據支持**（例如引用 CPA / CTR / CVR 數值）。
- 在提到的成本時，請明確區分是 CPA (轉換成本) 還是 CPM (曝光成本)。
"""

# ==========================================
# 1. 基礎設定與字型處理
# ==========================================
st.set_page_config(page_title="廣告成效全能分析 v6.4 (Dashboard + Instant DL)", layout="wide")

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


# --- 新素材/新組合判定（低 token：程式先聚合，AI 只判讀） ---
DATE_RE = re.compile(r'(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])')  # YYYYMMDD

def extract_yyyymmdd(s: str):
    """從字串中抓出第一個 YYYYMMDD，回傳 date 或 None"""
    m = DATE_RE.search(str(s))
    if not m:
        return None
    try:
        return datetime.strptime(m.group(0), "%Y%m%d").date()
    except Exception:
        return None

def is_recent_date(d, anchor_date, days=14):
    """以 anchor_date（資料 max_date）為基準，判斷 d 是否在最近 N 天內"""
    if not d:
        return False
    if isinstance(anchor_date, pd.Timestamp):
        anchor_date = anchor_date.date()
    return (anchor_date - d).days >= 0 and (anchor_date - d).days <= days

def build_new_creatives_summary(df_p7d, conv_col, anchor_date, recent_days=14, top_n=15, min_spend=300):
    """新素材（廣告）摘要：依名稱中的 YYYYMMDD 判定「近期新素材」"""
    if df_p7d is None or df_p7d.empty:
        return pd.DataFrame()

    tmp = df_p7d.copy()
    tmp['廣告名稱_clean'] = tmp['廣告名稱'].apply(clean_ad_name)
    tmp['creative_date'] = tmp['廣告名稱'].apply(extract_yyyymmdd)
    tmp['is_new_creative'] = tmp['creative_date'].apply(lambda d: is_recent_date(d, anchor_date, days=recent_days))

    agg = tmp.groupby(['廣告名稱_clean', 'is_new_creative'], as_index=False).agg({
        '花費金額 (TWD)': 'sum',
        conv_col: 'sum',
        '連結點擊次數': 'sum',
        '曝光次數': 'sum'
    })

    agg['CPA (TWD)'] = agg.apply(lambda x: x['花費金額 (TWD)'] / x[conv_col] if x[conv_col] > 0 else 0, axis=1)
    agg['CTR (%)'] = agg.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    agg['CPC (TWD)'] = agg.apply(lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else 0, axis=1)

    total_spend = agg['花費金額 (TWD)'].sum()
    total_conv = agg[conv_col].sum()
    agg['花費占比(%)'] = agg['花費金額 (TWD)'].apply(lambda v: (v / total_spend * 100) if total_spend > 0 else 0)
    agg['轉換占比(%)'] = agg[conv_col].apply(lambda v: (v / total_conv * 100) if total_conv > 0 else 0)

    agg = agg[agg['花費金額 (TWD)'] >= min_spend].copy()
    agg = agg.sort_values(['is_new_creative', '花費金額 (TWD)'], ascending=[False, False]).head(top_n)

    return agg.round(2)

def build_new_adsets_summary(df_p7d, df_pp7d, conv_col, top_n=15, min_spend_p7=500, old_spend_threshold=200):
    """新廣告組合判定：PP7D 花費很低但 P7D 有明顯花費"""
    if df_p7d is None or df_p7d.empty:
        return pd.DataFrame()

    def agg_adset(df):
        if df is None or df.empty:
            return pd.DataFrame(columns=['行銷活動名稱', '廣告組合名稱', '花費金額 (TWD)', '轉換', '連結點擊次數', '曝光次數'])
        tmp = df.groupby(['行銷活動名稱', '廣告組合名稱'], as_index=False).agg({
            '花費金額 (TWD)': 'sum',
            conv_col: 'sum',
            '連結點擊次數': 'sum',
            '曝光次數': 'sum'
        })
        tmp = tmp.rename(columns={conv_col: '轉換'})
        return tmp

    p7 = agg_adset(df_p7d)
    pp7 = agg_adset(df_pp7d)[['行銷活動名稱', '廣告組合名稱', '花費金額 (TWD)']].rename(columns={'花費金額 (TWD)': '花費金額_PP7D'})

    merged = p7.merge(pp7, on=['行銷活動名稱', '廣告組合名稱'], how='left')
    merged['花費金額_PP7D'] = merged['花費金額_PP7D'].fillna(0)

    merged['is_new_adset'] = (merged['花費金額_PP7D'] < old_spend_threshold) & (merged['花費金額 (TWD)'] >= min_spend_p7)

    merged['CPA (TWD)'] = merged.apply(lambda x: x['花費金額 (TWD)'] / x['轉換'] if x['轉換'] > 0 else 0, axis=1)
    merged['CTR (%)'] = merged.apply(lambda x: (x['連結點擊次數'] / x['曝光次數']) * 100 if x['曝光次數'] > 0 else 0, axis=1)
    merged['CPC (TWD)'] = merged.apply(lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else 0, axis=1)

    total_spend = merged['花費金額 (TWD)'].sum()
    total_conv = merged['轉換'].sum()
    merged['花費占比(%)'] = merged['花費金額 (TWD)'].apply(lambda v: (v / total_spend * 100) if total_spend > 0 else 0)
    merged['轉換占比(%)'] = merged['轉換'].apply(lambda v: (v / total_conv * 100) if total_conv > 0 else 0)

    merged = merged.sort_values(['is_new_adset', '花費金額 (TWD)'], ascending=[False, False]).head(top_n)

    return merged.round(2)
# --- end ---
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

    df_metrics['CPC (TWD)'] = df_metrics.apply(
        lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else 0, axis=1
    )

    df_metrics = df_metrics.round(2).sort_values(by='花費金額 (TWD)', ascending=False)

    metric_config = {
        'CPA (TWD)': ('花費金額 (TWD)', conv_col, 1),
        'CTR (%)': ('連結點擊次數', '曝光次數', 100),
        'CVR (%)': (conv_col, '連結點擊次數', 100),
        'CPM (TWD)': ('花費金額 (TWD)', '曝光次數', 1000),
        'CPC (TWD)': ('花費金額 (TWD)', '連結點擊次數', 1)
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


def calc_period_overall(df_period, conv_col):
    """
    計算期間整體（帳戶層級）指標：花費 / 轉換 / CPA / CTR / CPC
    - 使用原始明細 df_period 聚合，避免受中間匯總表結構影響
    """
    spend = float(df_period['花費金額 (TWD)'].sum()) if '花費金額 (TWD)' in df_period.columns else 0.0
    conv = float(df_period[conv_col].sum()) if conv_col in df_period.columns else 0.0
    clicks = float(df_period['連結點擊次數'].sum()) if '連結點擊次數' in df_period.columns else 0.0
    impr = float(df_period['曝光次數'].sum()) if '曝光次數' in df_period.columns else 0.0

    cpa = (spend / conv) if conv > 0 else 0.0
    ctr = (clicks / impr * 100) if impr > 0 else 0.0
    cpc = (spend / clicks) if clicks > 0 else 0.0

    return {
        'spend': round(spend, 0),
        'conv': round(conv, 0),
        'cpa': round(cpa, 2),
        'ctr': round(ctr, 2),
        'cpc': round(cpc, 2),
    }

def call_gemini_analysis(
    api_key,
    alerts_daily,
    alerts_weekly,
    campaign_summary,
    adset_p7=None,
    ad_p7=None,
    trend_30d=None,
    cpm_change_table=None,
    new_creatives=None,
    new_adsets=None
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

    # 新素材 / 新組合摘要（低 token）
    if new_creatives is not None and not new_creatives.empty:
        data_context += "\n\n## 8. New Creatives Summary (Recent Creatives, P7D Top)\n"
        data_context += safe_to_markdown(new_creatives)

    if new_adsets is not None and not new_adsets.empty:
        data_context += "\n\n## 9. New AdSets Summary (New AdSets by Spend Shift, P7D Top)\n"
        data_context += safe_to_markdown(new_adsets)

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
st.title("📊 廣告成效全能分析 v6.4 (Dashboard + Instant DL)")

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

        # 新素材 / 新廣告組合摘要（供 AI 判讀：避免丟全量表造成 token 壓力）
        new_creatives_df = build_new_creatives_summary(
            df_p7d=df_p7d,
            conv_col=conversion_col,
            anchor_date=max_date,
            recent_days=14,
            top_n=15,
            min_spend=300
        )

        new_adsets_df = build_new_adsets_summary(
            df_p7d=df_p7d,
            df_pp7d=df_pp7d,
            conv_col=conversion_col,
            top_n=15,
            min_spend_p7=500,
            old_spend_threshold=200
        )
        
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

        # ==========================================
        # [NEW] 調整 1：將下載邏輯提前至此（確保沒做 AI 也能下載）
        # ==========================================
        excel_stack = []
        excel_stack.append(('Trend_Daily_30D', trend_30d_df))
        if cpm_change_df is not None and not cpm_change_df.empty:
            excel_stack.append(('CPM_Change_P7D_PP7D_P30D', cpm_change_df))
        excel_stack.extend(res_p1)
        excel_stack.extend(res_p7)
        excel_stack.extend(res_pp7)
        excel_stack.extend(res_p30)
        
        # 取得目前 session state 的結果 (可能是 None，也可能是跑完後的文字)
        current_ai_result = st.session_state.get('gemini_result', None)
        
        # 產生 Excel Bytes
        excel_bytes = to_excel_single_sheet_stacked(excel_stack, AI_CONSULTANT_PROMPT, current_ai_result)
        
        with st.sidebar:
            st.divider()
            if excel_bytes:
                dl_label = "📥 下載完整分析報表"
                if current_ai_result:
                    dl_label += " (含 AI 分析)"
                
                st.download_button(
                    label=dl_label,
                    data=excel_bytes,
                    file_name=f"Full_Report_{max_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Excel 產生失敗 (xlsxwriter 未安裝)")

        # ==========================================
        # [NEW] 調整 2：新增 Dashboard 分頁 (Tab 0)
        # ==========================================
        tab_dash, tab1, tab2, tab3, tab4 = st.tabs(["📊 自訂儀表板", "📈 戰情室 & 雙重監控", "📑 詳細數據表 (AdSet+Ad)", "🤖 AI 深度診斷 (Gemini)", "🧾 週報產生器 (LINE Markdown)"])
        
        # ========== Tab 0：自訂儀表板 ==========
        with tab_dash:
            st.subheader("📈 30天趨勢比較儀表板")
            st.caption("勾選不同對象，比較其在指定指標上的每日變化趨勢。")
            
            # 1. 選擇層級
            dash_level = st.radio("1. 選擇分析層級", ["全帳戶 (Account)", "行銷活動 (Campaign)", "廣告組合 (AdSet)", "廣告 (Ad)"], horizontal=True)
            
            # 2. 準備篩選資料
            df_dash = df_p30d.copy()
            level_col_map = {
                "行銷活動 (Campaign)": "行銷活動名稱",
                "廣告組合 (AdSet)": "廣告組合名稱",
                "廣告 (Ad)": "廣告名稱"
            }
            
            selected_entities = []
            if dash_level == "全帳戶 (Account)":
                df_dash['分析對象'] = '全帳戶'
                selected_entities = ['全帳戶']
            else:
                target_col = level_col_map[dash_level]
                # 過濾掉 '全帳戶平均' 這種統計行
                unique_items = sorted([x for x in df_dash[target_col].dropna().unique() if '平均' not in str(x)])
                selected_entities = st.multiselect(f"2. 選擇 {dash_level} (可多選比對)", unique_items)
                
                if not selected_entities:
                    st.info("👆 請從上方選單選擇至少一個項目來顯示圖表")
                else:
                    df_dash = df_dash[df_dash[target_col].isin(selected_entities)].copy()
                    df_dash['分析對象'] = df_dash[target_col]

            # 3. 選擇指標
            metric_options = ["花費金額", "轉換數", "CPA", "CTR", "CVR", "CPC", "CPM", "曝光次數", "連結點擊次數"]
            selected_metric = st.selectbox("3. 選擇指標 (Y軸)", metric_options, index=2) # 預設 CPA

            if selected_entities:
                # 4. 計算每日數據
                # 先依 日期 + 分析對象 Groupby Sum
                daily_agg = df_dash.groupby(['天數', '分析對象']).agg({
                    '花費金額 (TWD)': 'sum',
                    conversion_col: 'sum',
                    '連結點擊次數': 'sum',
                    '曝光次數': 'sum'
                }).reset_index()
                
                # 計算衍生指標
                daily_agg['CPA'] = daily_agg.apply(lambda x: x['花費金額 (TWD)'] / x[conversion_col] if x[conversion_col] > 0 else 0, axis=1)
                daily_agg['CTR'] = daily_agg.apply(lambda x: x['連結點擊次數'] / x['曝光次數'] * 100 if x['曝光次數'] > 0 else 0, axis=1)
                daily_agg['CVR'] = daily_agg.apply(lambda x: x[conversion_col] / x['連結點擊次數'] * 100 if x['連結點擊次數'] > 0 else 0, axis=1)
                daily_agg['CPC'] = daily_agg.apply(lambda x: x['花費金額 (TWD)'] / x['連結點擊次數'] if x['連結點擊次數'] > 0 else 0, axis=1)
                daily_agg['CPM'] = daily_agg.apply(lambda x: x['花費金額 (TWD)'] / x['曝光次數'] * 1000 if x['曝光次數'] > 0 else 0, axis=1)
                
                # 對應中文欄位到 DataFrame 欄位
                metric_map = {
                    "花費金額": "花費金額 (TWD)",
                    "轉換數": conversion_col,
                    "CPA": "CPA",
                    "CTR": "CTR",
                    "CVR": "CVR",
                    "CPC": "CPC",
                    "CPM": "CPM",
                    "曝光次數": "曝光次數",
                    "連結點擊次數": "連結點擊次數"
                }
                
                plot_col = metric_map[selected_metric]
                
                # Pivot 轉換成 st.line_chart 需要的格式 (Index=Date, Columns=Entities, Values=Metric)
                chart_data = daily_agg.pivot(index='天數', columns='分析對象', values=plot_col)
                chart_data = chart_data.fillna(0)
                
                st.markdown(f"#### 📊 {selected_metric} 每日變化趨勢")
                st.line_chart(chart_data)

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
                        cpm_change_table=cpm_change_df,
                        new_creatives=new_creatives_df,
                        new_adsets=new_adsets_df
                    )
                    st.session_state['gemini_result'] = analysis_result
                    # 強制重新執行一次以刷新側邊欄下載按鈕的內容
                    st.rerun()
            
            if st.session_state['gemini_result']:
                st.markdown("### 📝 AI 診斷報告")
                st.markdown("---")
                st.markdown(st.session_state['gemini_result'])


        # ========== Tab 4：週報產生器（LINE Markdown） ==========
        PLAN_TYPES = [
            "1. 做簡易的開關、預算調配即可",
            "2. 補素材",
            "3. 補受眾",
            "4. 進行到達頁面優化",
            "5. 預算縮減、提高",
            "6. 維持即可",
        ]

        def _fmt_pct(x):
            return f"{x:.2f}%"

        def _fmt_money(x):
            return f"${x:,.0f}"

        def _weekly_report_ai_prompt(p7_overall, pp7_overall, top_adsets_p7, top_ads_p7):
            return f"""
你是一位成效廣告代操顧問。請用「可直接貼給客戶的週報語氣」輸出繁體中文，保持簡潔、可執行。

【本週 P7D 概況】
- 花費：{p7_overall['spend']}
- 轉換：{p7_overall['conv']}
- CPA：{p7_overall['cpa']}
- CTR：{p7_overall['ctr']}%
- CPC：{p7_overall['cpc']}

【上週 PP7D 概況】
- 花費：{pp7_overall['spend']}
- 轉換：{pp7_overall['conv']}
- CPA：{pp7_overall['cpa']}
- CTR：{pp7_overall['ctr']}%
- CPC：{pp7_overall['cpc']}

【AdSet（視為受眾單位）P7D Top】
{safe_to_markdown(top_adsets_p7)}

【Ad（視為素材單位）P7D Top】
{safe_to_markdown(top_ads_p7)}

請輸出 JSON（務必是 JSON，不能有多餘文字），格式如下：
{{
  "status_summary": "一段 2~4 句的現況描述（包含：哪些受眾有效/無效、哪些素材有效/無效）",
  "audience_effective": ["受眾/AdSet A（理由）", "..."],
  "audience_ineffective": ["受眾/AdSet B（理由）", "..."],
  "creative_effective": ["素材/Ad X（理由）", "..."],
  "creative_ineffective": ["素材/Ad Y（理由）", "..."],
  "next_week_plan_reco": [
    {{
      "type": "1. 做簡易的開關、預算調配即可",
      "recommend": true,
      "reason": "為何建議/不建議",
      "actions": ["具體動作 1", "具體動作 2"]
    }}
  ]
}}
"""

        def _try_parse_json(s):
            try:
                return json.loads(s)
            except Exception:
                s2 = re.sub(r"^```json\s*|\s*```$", "", str(s).strip(), flags=re.IGNORECASE)
                try:
                    return json.loads(s2)
                except Exception:
                    return None

        with tab4:
            st.subheader("🧾 每週周報（可貼 LINE）")
            st.caption("流程：AI 先產草案 → 你勾選/編輯 → 產出 Markdown")

            # 1) P7D / PP7D 帳戶概況
            p7_overall = calc_period_overall(df_p7d, conversion_col)
            pp7_overall = calc_period_overall(df_pp7d, conversion_col)

            # 2) 取受眾/素材 Top（避免把整張表丟給 AI 太長）
            top_adsets = get_top_by_spend(p7_adset_df, n=12, min_spend=500)
            top_ads = get_top_by_spend(p7_ad_df, n=12, min_spend=300)

            # 3) 顯示概況
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**P7D 概況**")
                st.write({
                    "花費": _fmt_money(p7_overall["spend"]),
                    "轉換": int(p7_overall["conv"]),
                    "CPA": _fmt_money(p7_overall["cpa"]),
                    "CTR": _fmt_pct(p7_overall["ctr"]),
                    "CPC": _fmt_money(p7_overall["cpc"]),
                })
            with c2:
                st.markdown("**PP7D 概況**")
                st.write({
                    "花費": _fmt_money(pp7_overall["spend"]),
                    "轉換": int(pp7_overall["conv"]),
                    "CPA": _fmt_money(pp7_overall["cpa"]),
                    "CTR": _fmt_pct(pp7_overall["ctr"]),
                    "CPC": _fmt_money(pp7_overall["cpc"]),
                })

            st.divider()

            # 4) 生成週報草案（AI）
            if "weekly_draft" not in st.session_state:
                st.session_state["weekly_draft"] = None

            col_btn, col_hint = st.columns([1, 3])
            with col_btn:
                gen_weekly = st.button("🤖 生成週報草案", type="primary")
            with col_hint:
                st.info("會輸出：現況描述 / 有效無效受眾與素材 / 下週計畫（6 類）→ 你再勾選與改字")

            if gen_weekly:
                if not gemini_api_key:
                    st.warning("⚠️ 請先於左側側邊欄輸入 Gemini API Key")
                else:
                    prompt = _weekly_report_ai_prompt(p7_overall, pp7_overall, top_adsets, top_ads)
                    with st.spinner("AI 週報草案生成中..."):
                        try:
                            if HAS_GENAI:
                                genai.configure(api_key=gemini_api_key)
                                model = genai.GenerativeModel("gemini-2.5-pro")
                                resp = model.generate_content(prompt)
                                raw_text = resp.text if hasattr(resp, "text") else str(resp)
                            else:
                                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key={gemini_api_key}"
                                headers = {"Content-Type": "application/json"}
                                data = {"contents": [{"parts": [{"text": prompt}]}]}
                                r = requests.post(url, headers=headers, json=data)
                                if r.status_code == 200:
                                    j = r.json()
                                    raw_text = j["candidates"][0]["content"]["parts"][0]["text"]
                                else:
                                    raw_text = ""
                        except Exception as e:
                            raw_text = ""

                    parsed = _try_parse_json(raw_text)
                    if not parsed:
                        st.error("AI 回傳不是可解析 JSON（可能混入其它文字）。你可以把回傳貼到下方手動修正。")
                        st.text_area("AI 原始回傳", value=str(raw_text), height=220)
                    else:
                        st.session_state["weekly_draft"] = parsed

            draft = st.session_state.get("weekly_draft")
            if not draft:
                st.stop()

            st.divider()
            st.subheader("✍️ 你可勾選、編輯、補充")

            # 5) 現況描述（可編輯）
            status_summary = st.text_area(
                "現況描述（可改）",
                value=str(draft.get("status_summary", "")),
                height=120
            )

            # 6) 有效/無效受眾與素材：勾選 + 可編輯
            def editable_checklist(title, items, key_prefix):
                st.markdown(f"### {title}")
                chosen = []

                items = items or []
                for i, it in enumerate(items):
                    k_chk = f"{key_prefix}_chk_{i}"
                    k_txt = f"{key_prefix}_txt_{i}"

                    # 只在 widget 建立前初始化預設值（一次）
                    if k_chk not in st.session_state:
                        st.session_state[k_chk] = True
                    if k_txt not in st.session_state:
                        st.session_state[k_txt] = str(it)

                    # 由 widget 自行更新 session_state，避免重複賦值造成錯誤
                    st.checkbox("採用", key=k_chk)
                    st.text_input("內容", key=k_txt)

                    if st.session_state.get(k_chk) and str(st.session_state.get(k_txt, "")).strip():
                        chosen.append(str(st.session_state.get(k_txt, "")).strip())

                    st.divider()

                return chosen

            colL, colR = st.columns(2)
            with colL:
                aud_eff = editable_checklist("✅ 有效受眾（AdSet）", draft.get("audience_effective", []), "aud_eff")
                aud_bad = editable_checklist("❌ 無效受眾（AdSet）", draft.get("audience_ineffective", []), "aud_bad")
            with colR:
                cre_eff = editable_checklist("✅ 有效素材（Ad）", draft.get("creative_effective", []), "cre_eff")
                cre_bad = editable_checklist("❌ 無效素材（Ad）", draft.get("creative_ineffective", []), "cre_bad")

            st.divider()

            # 7) 下週計畫：6 類型逐一顯示（勾選採用 + 編輯理由 + 編輯 actions）
            st.markdown("### 📌 下週計畫（你決定採用哪些）")
            plan_recos = draft.get("next_week_plan_reco", [])

            reco_map = {p.get("type"): p for p in plan_recos if isinstance(p, dict) and p.get("type")}
            merged_plans = []
            for t in PLAN_TYPES:
                p = reco_map.get(t, {"type": t, "recommend": False, "reason": "", "actions": []})
                merged_plans.append(p)

            selected_plans = []
            for idx, p in enumerate(merged_plans):
                t = p.get("type", "")
                default_on = bool(p.get("recommend", False))

                k_on = f"plan_on_{idx}"
                k_reason = f"plan_reason_{idx}"
                k_actions = f"plan_actions_{idx}"

                if k_on not in st.session_state:
                    st.session_state[k_on] = default_on
                if k_reason not in st.session_state:
                    st.session_state[k_reason] = str(p.get("reason", ""))
                if k_actions not in st.session_state:
                    st.session_state[k_actions] = "\n".join(p.get("actions", []) or [])

                st.markdown(f"**{t}**")
                st.checkbox("採用此計畫", key=k_on)
                st.text_area("理由（可改）", height=80, key=k_reason)
                st.text_area("具體動作（每行一條，可改）", height=100, key=k_actions)

                if st.session_state[k_on]:
                    actions_list = [x.strip() for x in st.session_state[k_actions].splitlines() if x.strip()]
                    selected_plans.append({
                        "type": t,
                        "reason": st.session_state[k_reason].strip(),
                        "actions": actions_list
                    })
                st.divider()

            # 8) 補充輸入框
            client_note = st.text_area("補充說明（可選）", value="", height=120)

            # 9) 拼 Markdown（LINE 可貼）
            def build_markdown():
                lines = []
                lines.append("## 📊 本週廣告週報")
                lines.append("")
                lines.append("### 1) 簡要概況")
                lines.append(f"- **P7D** 花費 {_fmt_money(p7_overall['spend'])}｜轉換 {int(p7_overall['conv'])}｜CPA {_fmt_money(p7_overall['cpa'])}｜CTR {_fmt_pct(p7_overall['ctr'])}｜CPC {_fmt_money(p7_overall['cpc'])}")
                lines.append(f"- **PP7D** 花費 {_fmt_money(pp7_overall['spend'])}｜轉換 {int(pp7_overall['conv'])}｜CPA {_fmt_money(pp7_overall['cpa'])}｜CTR {_fmt_pct(pp7_overall['ctr'])}｜CPC {_fmt_money(pp7_overall['cpc'])}")
                lines.append("")
                lines.append("### 2) 現況描述")
                if status_summary.strip():
                    lines.append(status_summary.strip())
                lines.append("")
                lines.append("### 3) 受眾與素材表現")
                if aud_eff:
                    lines.append("**✅ 有效受眾（AdSet）**")
                    lines += [f"- {x}" for x in aud_eff]
                if aud_bad:
                    lines.append("**❌ 無效受眾（AdSet）**")
                    lines += [f"- {x}" for x in aud_bad]
                if cre_eff:
                    lines.append("**✅ 有效素材（Ad）**")
                    lines += [f"- {x}" for x in cre_eff]
                if cre_bad:
                    lines.append("**❌ 無效素材（Ad）**")
                    lines += [f"- {x}" for x in cre_bad]
                lines.append("")
                lines.append("### 4) 下週計畫")
                if selected_plans:
                    for p in selected_plans:
                        lines.append(f"**{p['type']}**")
                        if p.get("reason"):
                            lines.append(f"- 理由：{p['reason']}")
                        if p.get("actions"):
                            lines.append("- 動作：")
                            lines += [f"  - {a}" for a in p["actions"]]
                else:
                    lines.append("- 本週建議維持為主，先觀察數據穩定性。")
                if client_note.strip():
                    lines.append("")
                    lines.append("### 5) 補充")
                    lines.append(client_note.strip())
                return "\n".join(lines)

            md = build_markdown()

            st.subheader("📋 可複製 Markdown（貼給客戶）")
            st.code(md, language="markdown")

    except Exception as e:
        st.error(f"系統發生未預期的錯誤: {e}")
        st.write("建議檢查：1. CSV格式是否正確 2. 是否包含轉換/花費/曝光欄位")
