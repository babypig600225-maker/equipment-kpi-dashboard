import streamlit as st
import streamlit.components.v1 as components
import openpyxl
import io, json, os, re

st.set_page_config(
    page_title="富強醫材 設備維修 KPI 儀表板",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 隱藏 Streamlit 預設 UI 元素，讓 HTML 完全佔滿
st.markdown("""
<style>
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding: 0 !important; margin: 0 !important; max-width: 100% !important; }
    section[data-testid="stSidebar"] { display: none; }
    [data-testid="stAppViewContainer"] > div:first-child { padding: 0; }
    .upload-wrapper {
        max-width: 640px; margin: 80px auto 0;
        background: #fff; border: 2px dashed #b4b2a9;
        border-radius: 16px; padding: 48px 32px; text-align: center;
    }
    .upload-wrapper h2 { font-size: 20px; color: #1a1a18; margin-bottom: 8px; }
    .upload-wrapper p  { font-size: 13px; color: #888780; margin-bottom: 24px; }
    [data-testid="stFileUploaderDropzone"] { border: none !important; background: #f5f4f0!important; }
</style>
""", unsafe_allow_html=True)

# ── 讀取 HTML 模板 ──
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'dashboard_template.html')

@st.cache_data
def load_template():
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        return f.read()

# ── Excel 解析 ──
def pn(v):
    if v is None or str(v).strip() == '無故障': return 0
    try: return float(str(v).replace(',', ''))
    except: return 0

def r2(v): return round(v, 2)

@st.cache_data
def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    kpi_name = next((s for s in wb.sheetnames if '關鍵設備績效指標' in s), None)
    if not kpi_name:
        raise ValueError("找不到「關鍵設備績效指標」工作表")

    ws = wb[kpi_name]
    kpi_rows = list(ws.iter_rows(values_only=True))

    # 偵測月份列
    month_row_idx, month_cols = -1, []
    for ri, row in enumerate(kpi_rows[:20]):
        found = [(ci, str(v).strip()) for ci, v in enumerate(row or [])
                 if v and str(v).strip().endswith('月') and str(v).strip()[:-1].isdigit()]
        if found:
            month_row_idx = ri
            month_cols = found
            break
    if month_row_idx == -1:
        raise ValueError("無法偵測月份列")

    month_row = kpi_rows[month_row_idx]
    year_label = next((str(v).strip() for v in (month_row or []) if v and '年' in str(v)), '2026年')
    std_row = kpi_rows[month_row_idx + 1] or []

    months = []
    for ci, label in month_cols:
        std_h = 0
        try: std_h = float(std_row[ci]) if std_row[ci] else 0
        except: pass
        if std_h > 0:
            months.append({'label': year_label + label, 'short': label, 'offset': ci, 'stdHrs': std_h})

    # 機台資料起始列
    machine_start = month_row_idx + 3
    for ri in range(month_row_idx + 1, min(len(kpi_rows), month_row_idx + 10)):
        row = kpi_rows[ri] or []
        if any(str(c).startswith('M-') for c in row if c):
            machine_start = ri
            break

    machines = []
    for row in kpi_rows[machine_start:]:
        if not row or not row[0] or not row[1]: continue
        name = str(row[0]).strip()
        code = str(row[1]).strip()
        if not code.startswith('M-'): continue
        mdata = []
        for m in months:
            mc = m['offset']
            f = int(pn(row[mc])) if len(row) > mc else 0
            r = r2(pn(row[mc+1])) if len(row) > mc+1 else 0
            d = r2(pn(row[mc+2])) if len(row) > mc+2 else 0
            braw = row[mc+3] if len(row) > mc+3 else None
            mtbf = None if (braw is None or str(braw).strip()=='無故障' or f==0) else r2(float(braw))
            mttr = r2(pn(row[mc+4])) if (len(row) > mc+4 and f > 0) else 0
            mdata.append({'f': f, 'r': r, 'd': d, 'mtbf': mtbf, 'mttr': mttr})
        machines.append({'name': name, 'code': code, 'months': mdata})

    # 月統計
    SRS = ['押出機','單機型壓鑄成型機','自動射出成型機','液態矽膠射出成型機','無廢料射出成型機','後射式矽膠射出成型機']

    def get_series(name):
        prefixes = {'押出機':'押出機','壓鑄':'單機型壓鑄成型機','自動射出':'自動射出成型機',
                    '液態矽膠':'液態矽膠射出成型機','無廢料':'無廢料射出成型機','後射':'後射式矽膠射出成型機'}
        for k, v in prefixes.items():
            if k in name: return v
        return None

    # 月別統計
    monthly_stats = []
    for m in months:
        sname = next((s for s in wb.sheetnames if m['short'] in s), None)
        total, comp = 0, 0
        if sname:
            ws2 = wb[sname]
            for ri2, row2 in enumerate(ws2.iter_rows(min_row=1, max_row=6, values_only=True)):
                row2 = list(row2) if row2 else []
                if ri2 == 2 and len(row2) > 2:
                    v = pn(row2[2])
                    if 30 <= v <= 500: total = int(v)
                if ri2 == 3 and len(row2) > 5:
                    v = pn(row2[5])
                    if 0.5 < v < 1.0: comp = v
                if comp == 0 or total == 0:
                    for cell in row2:
                        v = pn(cell)
                        if 0.5 < v < 1.0 and comp == 0: comp = v
                        if 30 <= v <= 500 and v == int(v) and total == 0: total = int(v)
        monthly_stats.append({'totalRepairs': total, 'completionRate': comp})

    # 系列彙總 → SERIES_DATA 格式
    STATUS_MAP = {
        lambda b, t, f: b is None: '表現卓越',
    }

    def compute_st(by_month_slice):
        valid_b = [d['mtbf'] for d in by_month_slice if d['mtbf'] is not None]
        valid_t = [d['mttr'] for d in by_month_slice if d['f'] > 0]
        if not valid_b: return '表現卓越'
        avg_b = sum(valid_b) / len(valid_b)
        avg_t = sum(valid_t) / len(valid_t) if valid_t else 0
        improving = len(valid_b) >= 2 and valid_b[-1] > valid_b[0]
        if avg_b >= 400: return '可靠度高'
        if avg_b >= 200: return '趨於穩定' if improving else '待觀察'
        if avg_t <= 10 and avg_t > 0: return '維修高效'
        if improving: return '持續改善中'
        return '需重點改善' if avg_b < 50 else '待觀察'

    series_data = []
    for sname in SRS:
        ms = [m for m in machines if m['name'] == sname]
        by_month = []
        for mi, month in enumerate(months):
            f = sum(m['months'][mi]['f'] for m in ms)
            r = r2(sum(m['months'][mi]['r'] for m in ms))
            d = r2(sum(m['months'][mi]['d'] for m in ms))
            std = month['stdHrs']
            mtbf = r2((std - d) / f) if f > 0 else None
            mttr = r2(r / f) if f > 0 else 0
            by_month.append({'f': f, 'r': r, 'd': d, 'mtbf': mtbf, 'mttr': mttr})
        st_label = compute_st(by_month)
        series_data.append([{**d, 'st': st_label} for d in by_month])

    return months, machines, monthly_stats, series_data

# ── 產生 JS DATA 區段 ──
def generate_data_block(months, machines, monthly_stats, series_data):
    month_labels  = json.dumps([m['label'] for m in months], ensure_ascii=False)
    month_shorts  = json.dumps([m['short'] for m in months], ensure_ascii=False)
    total_repairs = json.dumps([s['totalRepairs'] for s in monthly_stats])
    completion    = json.dumps([round(s['completionRate'], 4) for s in monthly_stats])

    # SERIES_DATA
    sd_js = '[\n'
    for row in series_data:
        sd_js += '  ['
        items = []
        for d in row:
            mtbf = 'null' if d['mtbf'] is None else d['mtbf']
            items.append(f"{{f:{d['f']},r:{d['r']},d:{d['d']},mtbf:{mtbf},mttr:{d['mttr']},st:'{d['st']}'}}")
        sd_js += ','.join(items) + '],\n'
    sd_js += ']'

    # MACHINES
    mach_js = '[\n'
    for m in machines:
        months_js = '['
        for md in m['months']:
            mtbf = 'null' if md['mtbf'] is None else md['mtbf']
            months_js += f"{{f:{md['f']},r:{md['r']},d:{md['d']},mtbf:{mtbf},mttr:{md['mttr']}}},"
        months_js += ']'
        name_js = m['name'].replace("'", "\\'")
        code_js = m['code'].replace("'", "\\'")
        mach_js += f"  {{name:'{name_js}',code:'{code_js}',months:{months_js}}},\n"
    mach_js += ']'

    return f"""/* ══ DATA ══ */
const MONTHS={month_labels};
const MO={month_shorts};
const TOTAL_REPAIRS={total_repairs};
const COMPLETION={completion};

const SRS=['押出機','單機型壓鑄成型機','自動射出成型機','液態矽膠射出成型機','無廢料射出成型機','後射式矽膠射出成型機'];
const SRS_COLORS=['#378ADD','#1D9E75','#BA7517','#D85A30','#7F77DD','#D4537E'];

const SERIES_DATA={sd_js};

const MACHINES={mach_js};"""

# ══ 主程式 ══
def main():
    if 'parsed' not in st.session_state:
        # 上傳畫面
        st.markdown("""
        <div class="upload-wrapper">
            <div style="font-size:52px;margin-bottom:16px">📊</div>
            <h2>富強醫材 — 設備維修 KPI 儀表板</h2>
            <p>請上傳「設備維修統計.xlsx」<br>需包含「關鍵設備績效指標」工作表</p>
        </div>
        """, unsafe_allow_html=True)

        uploaded = st.file_uploader("", type=['xlsx'], label_visibility='collapsed')
        if uploaded:
            with st.spinner("解析資料中..."):
                try:
                    result = parse_excel(uploaded.read())
                    st.session_state['parsed'] = result
                    st.session_state['filename'] = uploaded.name
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析失敗：{e}")
        return

    # 已載入資料 → 渲染 HTML
    months, machines, monthly_stats, series_data = st.session_state['parsed']
    fname = st.session_state.get('filename', '')

    # 重新上傳按鈕（浮在右上）
    col1, col2 = st.columns([6, 1])
    with col2:
        if st.button("↩ 重新上傳", use_container_width=True):
            del st.session_state['parsed']
            st.rerun()

    # 產生資料並注入 HTML
    data_block = generate_data_block(months, machines, monthly_stats, series_data)
    template   = load_template()
    html_out   = template.replace('/* ══ INJECT_DATA ══ */', data_block)

    # 更新標頭資訊（檔名與日期範圍）
    date_range = f"{months[0]['label']} – {months[-1]['label']}"
    html_out = html_out.replace(
        '資料來源：設備維修統計.xlsx',
        f'資料來源：{fname}'
    ).replace(
        '2026年 1–7月',
        date_range
    )

    # 顯示
    components.html(html_out, height=2400, scrolling=True)

main()
