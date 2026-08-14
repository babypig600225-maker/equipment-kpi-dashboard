import streamlit as st
import streamlit.components.v1 as components
import openpyxl, io, json
import plotly.graph_objects as go

st.set_page_config(
    page_title="富強醫材 設備維修 KPI 儀表板",
    page_icon="🏭", layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown("""
<style>
    #MainMenu, footer { visibility: hidden; }
    .block-container { padding-top:0.5rem !important; }
    [data-testid="stAppViewContainer"]>[data-testid="stVerticalBlock"] { gap:0; }
</style>
""", unsafe_allow_html=True)

# ── 常數 ──
SRS  = ['押出機','單機型壓鑄成型機','自動射出成型機','液態矽膠射出成型機','無廢料射出成型機','後射式矽膠射出成型機']
SRSN = ['押出機','壓鑄','自動射出','液態矽膠','無廢料','後射式矽膠']
COLS = ['#378ADD','#1D9E75','#BA7517','#D85A30','#7F77DD','#D4537E']
CMAP = dict(zip(SRS, COLS))
CSHORT = dict(zip(SRS, SRSN))

# ── 解析 ──
def pn(v):
    if v is None or str(v).strip()=='無故障': return 0
    try: return float(str(v).replace(',',''))
    except: return 0

@st.cache_data(show_spinner=False)
def parse(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    kpi_name = next((s for s in wb.sheetnames if '關鍵設備績效指標' in s), None)
    if not kpi_name: raise ValueError("找不到「關鍵設備績效指標」工作表")
    ws = wb[kpi_name]
    rows = list(ws.iter_rows(values_only=True))

    # 月份
    mrow_idx, mcols = -1, []
    for ri, row in enumerate(rows[:20]):
        found = [(ci,str(v).strip()) for ci,v in enumerate(row or []) if v and str(v).strip().endswith('月') and str(v).strip()[:-1].isdigit()]
        if found: mrow_idx,mcols = ri,found; break

    yr = next((str(v).strip() for v in (rows[mrow_idx] or []) if v and '年' in str(v)), '2026年')
    std_row = rows[mrow_idx+1] or []
    months = []
    for ci,lbl in mcols:
        try: sh = float(std_row[ci]) if std_row[ci] else 0
        except: sh=0
        if sh>0: months.append({'label':yr+lbl,'short':lbl,'offset':ci,'std':sh})

    ms = mrow_idx+3
    for ri in range(mrow_idx+1, min(len(rows), mrow_idx+10)):
        if any(str(c).startswith('M-') for c in (rows[ri] or []) if c): ms=ri; break

    machines=[]
    for row in rows[ms:]:
        if not row or not row[0] or not row[1]: continue
        nm,cd = str(row[0]).strip(), str(row[1]).strip()
        if not cd.startswith('M-'): continue
        md=[]
        for m in months:
            mc=m['offset']
            f=int(pn(row[mc])) if len(row)>mc else 0
            r=round(pn(row[mc+1]),2) if len(row)>mc+1 else 0
            d=round(pn(row[mc+2]),2) if len(row)>mc+2 else 0
            braw=row[mc+3] if len(row)>mc+3 else None
            mtbf=None if(braw is None or str(braw).strip()=='無故障' or f==0) else round(float(braw),2)
            mttr=round(pn(row[mc+4]),2) if(len(row)>mc+4 and f>0) else 0
            md.append({'f':f,'r':r,'d':d,'mtbf':mtbf,'mttr':mttr})
        machines.append({'name':nm,'code':cd,'months':md})

    stats=[]
    for m in months:
        sn=next((s for s in wb.sheetnames if m['short'] in s), None)
        tot,comp=0,0
        if sn:
            for ri2,r2 in enumerate(wb[sn].iter_rows(min_row=1,max_row=6,values_only=True)):
                r2=list(r2) if r2 else []
                if ri2==2 and len(r2)>2:
                    v=pn(r2[2]);
                    if 30<=v<=500: tot=int(v)
                if ri2==3 and len(r2)>5:
                    v=pn(r2[5]);
                    if 0.5<v<1.0: comp=v
                if not comp or not tot:
                    for c in r2:
                        v=pn(c)
                        if 0.5<v<1.0 and not comp: comp=v
                        if 30<=v<=500 and v==int(v) and not tot: tot=int(v)
        stats.append({'rep':tot,'comp':comp})
    return months, machines, stats

def agg_series(machines, months):
    res={}
    for nm in SRS:
        ms=[m for m in machines if m['name']==nm]
        by=[]
        for mi,mo in enumerate(months):
            f=sum(m['months'][mi]['f'] for m in ms)
            r=round(sum(m['months'][mi]['r'] for m in ms),2)
            d=round(sum(m['months'][mi]['d'] for m in ms),2)
            mtbf=round((mo['std']-d)/f,2) if f>0 else None
            mttr=round(r/f,2) if f>0 else 0
            by.append({'f':f,'r':r,'d':d,'mtbf':mtbf,'mttr':mttr})
        res[nm]=by
    return res

# ── 圖表 helper ──
LAYOUT = dict(plot_bgcolor='white', paper_bgcolor='white', margin=dict(t=40,b=60,l=50,r=20),
              height=280, font=dict(family='微軟正黑體,Noto Sans TC,sans-serif', size=11),
              xaxis=dict(gridcolor='#eae8e1',showgrid=True),
              yaxis=dict(gridcolor='#eae8e1',showgrid=True))
LEG = dict(orientation='h', y=-0.28, x=0.5, xanchor='center', font_size=10)

def line_chart(title, x, series_data, y_key, y_title, show_none=True):
    fig=go.Figure()
    for nm in series_data:
        y=[series_data[nm][i][y_key] for i in range(len(x))]
        if show_none: y=[v if v else None for v in y]
        fig.add_trace(go.Scatter(x=x,y=y,name=CSHORT[nm],mode='lines+markers',
            line=dict(color=CMAP[nm],width=2),marker=dict(size=6),connectgaps=True))
    fig.update_layout(title=title,yaxis_title=y_title,legend=LEG,**LAYOUT)
    return fig

def bar_chart(title, x, series_data, y_key, y_title):
    fig=go.Figure()
    for nm in series_data:
        y=[series_data[nm][i][y_key] for i in range(len(x))]
        fig.add_trace(go.Bar(x=x,y=y,name=CSHORT[nm],marker_color=CMAP[nm],opacity=0.85))
    fig.update_layout(title=title,barmode='group',yaxis_title=y_title,legend=LEG,**LAYOUT)
    return fig

def status_badge(v):
    cfg={'需重點改善':('🔴','#fcebeb','#a32d2d'),'可靠度高':('🟢','#e2efda','#1f7a4d'),
         '表現卓越':('✨','#e1f5ee','#0f6e56'),'待觀察':('🟡','#fff2cc','#7b4f00'),
         '維修高效':('🔵','#e6f1fb','#185fa5'),'持續改善中':('🟠','#fff2cc','#7b4f00'),
         '趨於穩定':('🔵','#e6f1fb','#185fa5')}
    ico,bg,fg=cfg.get(v,('⚪','#f1f0ea','#5f5e5a'))
    return f'<span style="background:{bg};color:{fg};padding:2px 10px;border-radius:4px;font-weight:600;font-size:12px">{ico} {v}</span>'

def compute_status(data):
    vb=[d['mtbf'] for d in data if d['mtbf'] is not None]
    vt=[d['mttr'] for d in data if d['f']>0]
    if not vb: return '表現卓越'
    ab=sum(vb)/len(vb); at=sum(vt)/len(vt) if vt else 0
    imp=len(vb)>=2 and vb[-1]>vb[0]
    if ab>=400: return '可靠度高'
    if ab>=200: return '趨於穩定' if imp else '待觀察'
    if 0<at<=10: return '維修高效'
    return '持續改善中' if imp else ('需重點改善' if ab<50 else '待觀察')

# ══ 主程式 ══
def main():
    if 'data' not in st.session_state:
        st.markdown("""
        <div style="max-width:580px;margin:80px auto;background:#fff;border:2px dashed #b4b2a9;
                    border-radius:16px;padding:48px 32px;text-align:center">
            <div style="font-size:52px;margin-bottom:16px">📊</div>
            <div style="font-size:20px;font-weight:600;color:#1a1a18;margin-bottom:8px">富強醫材 — 設備維修 KPI 儀表板</div>
            <div style="font-size:13px;color:#888780;margin-bottom:8px">請上傳「設備維修統計.xlsx」</div>
            <div style="font-size:12px;color:#b4b2a9">需包含「關鍵設備績效指標」工作表</div>
        </div>
        """, unsafe_allow_html=True)
        up=st.file_uploader("",type=['xlsx'],label_visibility='collapsed')
        if up:
            with st.spinner("解析資料中..."):
                try:
                    st.session_state['data']=parse(up.read())
                    st.session_state['fn']=up.name
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")
        return

    months,machines,stats=st.session_state['data']
    fn=st.session_state.get('fn','')
    sagg=agg_series(machines,months)
    mlabels=[m['label'] for m in months]
    mshorts=[m['short'] for m in months]

    # ── 標頭 ──
    c1,c2=st.columns([5,1])
    with c1:
        st.markdown(f"""
        <div style="padding:12px 0 0">
        <div style="font-size:20px;font-weight:700;color:#1A3A5C">🏭 富強醫材 — 設備維修 KPI 儀表板</div>
        <div style="font-size:12px;color:#888780;margin-top:4px">
            📁 {fn}&nbsp;&nbsp;&nbsp;📅 {mlabels[0]} – {mlabels[-1]}&nbsp;&nbsp;&nbsp;富強醫材股份有限公司
        </div></div>""", unsafe_allow_html=True)
    with c2:
        if st.button("↩ 重新上傳",use_container_width=True):
            del st.session_state['data']; st.rerun()

    st.divider()
    tab1,tab2=st.tabs(["📊 機台系列","🔍 個別機台"])

    # ══ 機台系列 ══
    with tab1:
        c1,c2=st.columns(2)
        with c1: sel_m=st.selectbox("月份",["全部月份"]+mlabels,key='sm')
        with c2: sel_s=st.selectbox("機台系列",["全部機台系列"]+SRS,key='ss')

        midxs=list(range(len(months))) if sel_m=="全部月份" else [mlabels.index(sel_m)]
        fsrs=[sel_s] if sel_s!="全部機台系列" else SRS
        xlabels=[mshorts[i] for i in midxs]
        fdata={nm:sagg[nm] for nm in fsrs}

        # KPI
        tf=sum(fdata[nm][i]['f'] for nm in fdata for i in midxs)
        tr=round(sum(fdata[nm][i]['r'] for nm in fdata for i in midxs),1)
        td=round(sum(fdata[nm][i]['d'] for nm in fdata for i in midxs),1)
        trep=sum(stats[i]['rep'] for i in midxs) if sel_s=="全部機台系列" else tf
        ac=round(sum(stats[i]['comp'] for i in midxs)/len(midxs)*100,1)

        k1,k2,k3,k4,k5=st.columns(5)
        k1.metric("設備報修總數",f"{trep} 件",f"{len(midxs)}個月")
        k2.metric("故障次數",f"{tf} 次",f"{len(fsrs)}個系列")
        k3.metric("維修工時",f"{tr} h","含等待+作業時間")
        k4.metric("停工工時",f"{td} h",f"停機 {td}h",delta_color="inverse")
        k5.metric("平均完成率",f"{ac}%","✓ 達標" if ac>=95 else "△ 需改善")

        # 圖表 Row1: MTBF + MTTR
        c1,c2=st.columns(2)
        with c1:
            fig=line_chart("MTBF 趨勢（hr / 次）",xlabels,fdata,'mtbf',"MTBF (h)")
            st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
            st.caption("＊ MTBF（平均無故障間隔時間）— 數值越高代表機台越可靠")
        with c2:
            fig2=go.Figure()
            for nm in fsrs:
                y=[fdata[nm][i]['mttr'] if fdata[nm][i]['f']>0 else None for i in midxs]
                fig2.add_trace(go.Scatter(x=xlabels,y=y,name=CSHORT[nm],mode='lines+markers',
                    line=dict(color=CMAP[nm],width=2),marker=dict(size=6),connectgaps=True))
            fig2.update_layout(title="MTTR 趨勢（hr / 次）",yaxis_title="MTTR (h)",legend=LEG,**LAYOUT)
            st.plotly_chart(fig2,use_container_width=True,config={'displayModeBar':False})
            st.caption("＊ MTTR（平均修復時間）— 數值越低代表維修效率越高")

        # 圖表 Row2: 故障+維修+停工
        c1,c2,c3=st.columns(3)
        with c1:
            st.plotly_chart(bar_chart("故障次數（月別）",xlabels,fdata,'f',"次"),
                use_container_width=True,config={'displayModeBar':False})
        with c2:
            st.plotly_chart(line_chart("維修工時（hr，月別）",xlabels,fdata,'r',"h",False),
                use_container_width=True,config={'displayModeBar':False})
        with c3:
            st.plotly_chart(bar_chart("停工工時（hr，月別）",xlabels,fdata,'d',"h"),
                use_container_width=True,config={'displayModeBar':False})

        # 績效表
        st.markdown("#### 系列績效總覽")
        rows=[]
        for nm in fsrs:
            data=[fdata[nm][i] for i in midxs]
            f=sum(d['f'] for d in data); r=round(sum(d['r'] for d in data),1); d=round(sum(d['d'] for d in data),1)
            vb=[x['mtbf'] for x in data if x['mtbf']]; vt=[x['mttr'] for x in data if x['f']>0]
            ab=f"{sum(vb)/len(vb):.1f}h" if vb else "無故障"
            at=f"{sum(vt)/len(vt):.1f}h" if vt else "—"
            sl=compute_status(data)
            rows.append({'機台系列':nm,'故障次數':f,'維修工時(h)':r,'停工工時(h)':d,'平均MTBF':ab,'平均MTTR':at,'狀態':status_badge(sl)})

        # 用 HTML 表格保留狀態標籤顏色
        th_style="background:#1A3A5C;color:white;padding:8px 12px;text-align:center;font-weight:600;font-size:13px"
        td_style="padding:7px 12px;text-align:center;font-size:13px;border-bottom:1px solid #eae8e1"
        td_l="padding:7px 12px;font-size:13px;border-bottom:1px solid #eae8e1"
        html_tbl = f'<table style="width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.06)">'
        html_tbl += f'<tr><th style="{th_style};text-align:left">機台系列</th><th style="{th_style}">故障次數</th><th style="{th_style}">維修工時(h)</th><th style="{th_style}">停工工時(h)</th><th style="{th_style}">平均MTBF</th><th style="{th_style}">平均MTTR</th><th style="{th_style}">狀態</th></tr>'
        for i,row in enumerate(rows):
            bg="#f9f8f6" if i%2 else "white"
            nm_color = CMAP[row["機台系列"]]
            nm_name  = row["機台系列"]
            html_tbl += (f'<tr style="background:{bg}">'
                f'<td style="{td_l}"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:{nm_color};margin-right:6px"></span>{nm_name}</td>'
                f'<td style="{td_style}">{row["故障次數"]}</td>'
                f'<td style="{td_style}">{row["維修工時(h)"]}</td>'
                f'<td style="{td_style}">{row["停工工時(h)"]}</td>'
                f'<td style="{td_style};color:#185fa5;font-weight:600">{row["平均MTBF"]}</td>'
                f'<td style="{td_style}">{row["平均MTTR"]}</td>'
                f'<td style="{td_style}">{row["狀態"]}</td></tr>')
        html_tbl += '</table>'
        st.markdown(html_tbl, unsafe_allow_html=True)

    # ══ 個別機台 ══
    with tab2:
        sub=st.radio("檢視角度",["工時 & 故障","MTBF & MTTR"],horizontal=True,key='msub')
        c1,c2,c3=st.columns(3)
        with c1: sel_mm=st.selectbox("月份",["全部月份"]+mlabels,key='mm')
        with c2: sel_ms=st.selectbox("機台系列",["全部機台系列"]+SRS,key='ms')
        with c3: hz=st.checkbox("隱藏零故障機台")

        midxs_m=list(range(len(months))) if sel_mm=="全部月份" else [mlabels.index(sel_mm)]
        mlist=[m for m in machines if sel_ms=="全部機台系列" or m['name']==sel_ms]
        total_all=len(mlist)
        if hz: mlist=[m for m in mlist if any(m['months'][i]['f']>0 for i in midxs_m)]

        tf=sum(m['months'][i]['f'] for m in mlist for i in midxs_m)
        tr=round(sum(m['months'][i]['r'] for m in mlist for i in midxs_m),1)
        td=round(sum(m['months'][i]['d'] for m in mlist for i in midxs_m),1)
        ac_m=round(sum(stats[i]['comp'] for i in midxs_m)/len(midxs_m)*100,1)

        k1,k2,k3,k4,k5=st.columns(5)
        k1.metric("顯示機台數",f"{len(mlist)} 台",f"共{total_all}台")
        k2.metric("故障次數",f"{tf} 次",f"{len(midxs_m)}個月")
        k3.metric("維修工時",f"{tr} h")
        k4.metric("停工工時",f"{td} h",delta_color="inverse")
        k5.metric("平均完成率",f"{ac_m}%","✓ 達標" if ac_m>=95 else "△ 需改善")

        ranked=sorted([{'code':m['code'],'name':m['name'],
            'tF':sum(m['months'][i]['f'] for i in midxs_m),
            'tR':round(sum(m['months'][i]['r'] for i in midxs_m),1),
            'tD':round(sum(m['months'][i]['d'] for i in midxs_m),1),
            'vB':[m['months'][i]['mtbf'] for i in midxs_m if m['months'][i]['mtbf']],
            'vT':[m['months'][i]['mttr'] for i in midxs_m if m['months'][i]['f']>0]}
            for m in mlist],key=lambda x:-x['tF'])[:15]
        ranked=[r for r in ranked if r['tF']>0 or r['tR']>0]

        if sub=="工時 & 故障":
            if ranked:
                c1,c2=st.columns(2)
                with c1:
                    fig6=go.Figure(go.Bar(x=[r['code'] for r in ranked],y=[r['tF'] for r in ranked],
                        marker_color=[CMAP.get(r['name'],'#888') for r in ranked],
                        text=[r['tF'] for r in ranked],textposition='outside'))
                    fig6.update_layout(title="累計故障次數 TOP 15",height=350,
                        margin=dict(t=40,b=90,l=40,r=10),plot_bgcolor='white',paper_bgcolor='white',
                        xaxis=dict(tickangle=45,gridcolor='#eae8e1'),yaxis=dict(gridcolor='#eae8e1'))
                    st.plotly_chart(fig6,use_container_width=True,config={'displayModeBar':False})
                with c2:
                    fig7=go.Figure()
                    fig7.add_trace(go.Bar(x=[r['code'] for r in ranked],y=[r['tR'] for r in ranked],name='維修工時',marker_color='#378ADD',opacity=0.85))
                    fig7.add_trace(go.Bar(x=[r['code'] for r in ranked],y=[r['tD'] for r in ranked],name='停工工時',marker_color='#E24B4A',opacity=0.85))
                    fig7.update_layout(title="維修 vs 停工工時 TOP 15",height=350,barmode='stack',
                        margin=dict(t=40,b=90,l=40,r=10),plot_bgcolor='white',paper_bgcolor='white',
                        legend=dict(orientation='h',y=1.05,font_size=11),
                        xaxis=dict(tickangle=45,gridcolor='#eae8e1'),yaxis=dict(gridcolor='#eae8e1',title='h'))
                    st.plotly_chart(fig7,use_container_width=True,config={'displayModeBar':False})

            st.markdown("#### 個別機台明細")
            xlm=[mshorts[i] for i in midxs_m]
            tbl=[]
            for m in mlist:
                row={'機台名稱':m['name'],'機台編號':m['code']}
                for mi in midxs_m:
                    d=m['months'][mi]; lbl=mshorts[mi]
                    row[f'{lbl} 故障']=d['f']; row[f'{lbl} 維修(h)']=d['r'] or ''; row[f'{lbl} 停工(h)']=d['d'] or ''
                if len(midxs_m)>1:
                    row['累計故障']=sum(m['months'][i]['f'] for i in midxs_m)
                    row['維修工時(h)']=round(sum(m['months'][i]['r'] for i in midxs_m),1)
                    row['停工工時(h)']=round(sum(m['months'][i]['d'] for i in midxs_m),1)
                tbl.append(row)
            import pandas as pd
            st.dataframe(pd.DataFrame(tbl),use_container_width=True,hide_index=True)

        else:
            hf=[r for r in ranked if r['tF']>0]
            if hf:
                c1,c2,c3=st.columns(3)
                bc=[CMAP.get(r['name'],'#888') for r in hf]
                with c1:
                    fig8=go.Figure(go.Bar(x=[r['code'] for r in hf],
                        y=[round(sum(r['vB'])/len(r['vB']),1) if r['vB'] else 0 for r in hf],
                        marker_color=bc,opacity=0.75,marker_line_color=bc,marker_line_width=1.5))
                    fig8.update_layout(title="MTBF 分佈",height=320,margin=dict(t=40,b=90,l=40,r=10),
                        plot_bgcolor='white',paper_bgcolor='white',
                        xaxis=dict(tickangle=45),yaxis=dict(gridcolor='#eae8e1',title='h'))
                    st.plotly_chart(fig8,use_container_width=True,config={'displayModeBar':False})
                with c2:
                    mt_vals=[round(sum(r['vT'])/len(r['vT']),1) if r['vT'] else 0 for r in hf]
                    mtc=['#1f7a4d' if v<=5 else '#7b4f00' if v<=15 else '#a32d2d' for v in mt_vals]
                    fig9=go.Figure(go.Bar(x=[r['code'] for r in hf],y=mt_vals,
                        marker_color=mtc,opacity=0.85,marker_line_color=mtc,marker_line_width=1.5))
                    fig9.update_layout(title="MTTR 比較",height=320,margin=dict(t=40,b=90,l=40,r=10),
                        plot_bgcolor='white',paper_bgcolor='white',
                        xaxis=dict(tickangle=45),yaxis=dict(gridcolor='#eae8e1',title='h'))
                    st.plotly_chart(fig9,use_container_width=True,config={'displayModeBar':False})
                with c3:
                    fig10=go.Figure()
                    for r in hf:
                        ab=round(sum(r['vB'])/len(r['vB']),1) if r['vB'] else 0
                        at=round(sum(r['vT'])/len(r['vT']),1) if r['vT'] else 0
                        fig10.add_trace(go.Scatter(x=[ab],y=[at],mode='markers+text',
                            text=[r['code']],textposition='top center',showlegend=False,
                            marker=dict(size=10,color=CMAP.get(r['name'],'#888'))))
                    fig10.update_layout(title="MTBF vs MTTR 散佈圖",height=320,
                        margin=dict(t=40,b=30,l=50,r=10),plot_bgcolor='white',paper_bgcolor='white',
                        xaxis=dict(title="MTBF (h) →",gridcolor='#eae8e1'),
                        yaxis=dict(title="MTTR (h) ↑",gridcolor='#eae8e1'))
                    st.plotly_chart(fig10,use_container_width=True,config={'displayModeBar':False})

            st.markdown("#### MTBF / MTTR 月別明細")
            import pandas as pd
            tbl2=[]
            for m in mlist:
                row={'機台名稱':m['name'],'機台編號':m['code']}
                for mi in midxs_m:
                    d=m['months'][mi]; lbl=mshorts[mi]
                    row[f'{lbl} 故障']=d['f']
                    row[f'{lbl} MTBF(h)']=d['mtbf'] if d['mtbf'] else '—'
                    row[f'{lbl} MTTR(h)']=d['mttr'] if d['f']>0 else '—'
                if len(midxs_m)>1:
                    vb=[m['months'][i]['mtbf'] for i in midxs_m if m['months'][i]['mtbf']]
                    vt=[m['months'][i]['mttr'] for i in midxs_m if m['months'][i]['f']>0]
                    row['平均MTBF(h)']=round(sum(vb)/len(vb),1) if vb else '—'
                    row['平均MTTR(h)']=round(sum(vt)/len(vt),1) if vt else '—'
                tbl2.append(row)
            st.dataframe(pd.DataFrame(tbl2),use_container_width=True,hide_index=True)

    st.caption("富強醫材股份有限公司　設備課　© 2026")

main()
