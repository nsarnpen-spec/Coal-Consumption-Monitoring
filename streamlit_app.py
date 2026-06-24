"""
Coal Consumption Dashboard – HONGSA Power Plant
CHS Operation Team

Deploy: Streamlit Community Cloud
Local:  streamlit run streamlit_app.py
"""

import math
import io
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime

# ─────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Coal Consumption Dashboard – HONGSA",
    page_icon="⚡",
    layout="wide",
)

# ─────────────────────────────────────────────────────────
#  CUSTOM CSS
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #0f1117; }
[data-testid="stHeader"]           { background: #0f1117; }
[data-testid="stSidebar"]          { background: #1a1d27; }
[data-testid="metric-container"] {
  background: #1a1d27; border: 1px solid #2a2d3e;
  border-radius: 10px; padding: 10px 14px;
}
.banner {
  border-radius: 8px; padding: 10px 14px;
  font-size: 0.83rem; margin-bottom: 6px;
  display: flex; align-items: flex-start; gap: 10px;
}
.banner-ok  { background:rgba(34,197,94,0.08);  border:1px solid rgba(34,197,94,0.25);  color:#86efac; }
.banner-c1  { background:rgba(239,68,68,0.10);  border:1px solid rgba(239,68,68,0.35);  color:#fca5a5; }
.banner-c2  { background:rgba(249,115,22,0.10); border:1px solid rgba(249,115,22,0.35); color:#fdba74; }
.banner-tot { background:rgba(249,115,22,0.08); border:1px solid rgba(249,115,22,0.30); color:#fdba74; }
.alarm-row {
  display:flex; gap:10px; font-size:0.75rem;
  padding:5px 10px; border-radius:5px; margin-bottom:3px;
  background:rgba(249,115,22,0.07);
  border-left:3px solid #f97316; color:#fdba74;
}
.alarm-time { color:#e2e8f0; font-weight:600; min-width:140px; }
.section-h  { font-size:0.9rem; font-weight:600; color:#e2e8f0; margin-bottom:4px; }
.thr-note   {
  font-size:0.72rem; color:#8b95a3; padding:6px 10px;
  background:rgba(255,255,255,0.03); border-radius:6px;
  border-left:3px solid #2a2d3e; line-height:1.7;
  margin-top:4px; margin-bottom:10px;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
#  PARSER
# ─────────────────────────────────────────────────────────
def safe_float(v, ndigits=2):
    try:
        f = float(v)
        return 0.0 if (math.isnan(f) or math.isinf(f)) else round(f, ndigits)
    except (ValueError, TypeError):
        return 0.0


def process_excel(file_bytes: bytes) -> list[dict]:
    xf = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=xf.sheet_names[0], header=None)

    data_start = None
    for i in range(len(df)):
        try:
            t = pd.to_datetime(df.iloc[i, 1], dayfirst=True)
            if not pd.isna(t):
                data_start = i; break
        except Exception:
            continue
    if data_start is None:
        raise ValueError("ไม่พบ Timestamp ในคอลัมน์ B")

    rows = []
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        try:
            t = pd.to_datetime(row.iloc[1], dayfirst=True)
            if pd.isna(t): break
        except Exception:
            break
        u1mw = safe_float(row.iloc[3]);  u2mw = safe_float(row.iloc[4]);  u3mw = safe_float(row.iloc[5])
        u1c  = safe_float(row.iloc[6]);  u2c  = safe_float(row.iloc[7]);  u3c  = safe_float(row.iloc[8])
        u1r  = safe_float(row.iloc[9], 4); u2r = safe_float(row.iloc[10], 4); u3r = safe_float(row.iloc[11], 4)
        rows.append({"t": t, "u1c": u1c, "u2c": u2c, "u3c": u3c,
                     "u1mw": u1mw, "u2mw": u2mw, "u3mw": u3mw,
                     "u1r": u1r, "u2r": u2r, "u3r": u3r,
                     "tot": round(u1c + u2c + u3c, 2)})
    if not rows:
        raise ValueError("ไม่พบข้อมูลที่ valid ในไฟล์")

    # Cond.1: ≥2 units > 650 t/h ต่อเนื่อง > 120 นาที
    c1t = [sum([r["u1c"] > 650, r["u2c"] > 650, r["u3c"] > 650]) >= 2 for r in rows]
    rl  = [0] * len(rows)
    for i in range(len(rows)):
        rl[i] = (rl[i-1] + 1) if (i > 0 and c1t[i]) else (1 if c1t[i] else 0)
    a1 = [False] * len(rows)
    for i in range(len(rows) - 1, -1, -1):
        if rl[i] >= 121:
            j = i
            while j >= 0 and c1t[j]:
                a1[j] = True; j -= 1

    for i, r in enumerate(rows):
        r["a1"] = a1[i]
        r["a2"] = r["u1c"] > 670 or r["u2c"] > 670 or r["u3c"] > 670
        r["a3"] = r["u1r"] > 1.08 or r["u2r"] > 1.08 or r["u3r"] > 1.08
        r["a4"] = r["tot"] > 1950
    return rows


# ─────────────────────────────────────────────────────────
#  CHART HELPERS
# ─────────────────────────────────────────────────────────
# Base layout WITHOUT yaxis (to avoid duplicate keyword error)
def base_layout(shapes=None):
    return dict(
        paper_bgcolor="#1a1d27",
        plot_bgcolor="#1a1d27",
        font=dict(color="#e2e8f0", size=11),
        margin=dict(l=55, r=20, t=20, b=40),
        xaxis=dict(
            gridcolor="#2a2d3e", linecolor="#2a2d3e",
            tickfont=dict(color="#6b7280", size=9),
        ),
        legend=dict(
            bgcolor="rgba(0,0,0,0)", borderwidth=0,
            font=dict(size=10, color="#8b95a3"),
        ),
        hovermode="x unified",
        hoverlabel=dict(bgcolor="#1a1d27", bordercolor="#2a2d3e",
                        font=dict(color="#e2e8f0")),
        shapes=shapes or [],
    )


# Common yaxis style
def y_style(y_min, y_max, title):
    return dict(
        gridcolor="#2a2d3e", linecolor="#2a2d3e",
        tickfont=dict(color="#6b7280", size=9),
        title_text=title, title_font=dict(size=10, color="#8b95a3"),
        range=[y_min, y_max],
    )


def alarm_shapes(data, field, color):
    shapes = []
    in_alarm = False
    start_t  = None
    for r in data:
        if r[field] and not in_alarm:
            in_alarm = True; start_t = r["t"]
        elif not r[field] and in_alarm:
            shapes.append(dict(type="rect", xref="x", yref="paper",
                               x0=start_t, x1=r["t"], y0=0, y1=1,
                               fillcolor=color, opacity=1, line_width=0, layer="below"))
            in_alarm = False
    if in_alarm and start_t:
        shapes.append(dict(type="rect", xref="x", yref="paper",
                           x0=start_t, x1=data[-1]["t"], y0=0, y1=1,
                           fillcolor=color, opacity=1, line_width=0, layer="below"))
    return shapes


# ─────────────────────────────────────────────────────────
#  CHART BUILDERS
# ─────────────────────────────────────────────────────────
def make_coal_chart(data):
    ts  = [r["t"] for r in data]
    fig = go.Figure()
    for key, name, color in [("u1c","Unit 1","#60a5fa"),
                              ("u2c","Unit 2","#34d399"),
                              ("u3c","Unit 3","#f472b6")]:
        fig.add_trace(go.Scatter(x=ts, y=[r[key] for r in data], name=name,
                                 line=dict(color=color, width=1.5), mode="lines"))
    fig.add_hline(y=650, line_dash="dash", line_color="rgba(249,115,22,0.7)",
                  line_width=1, annotation_text="650 t/h",
                  annotation_font_color="#f97316", annotation_font_size=9)
    fig.add_hline(y=670, line_dash="dot", line_color="rgba(239,68,68,0.85)",
                  line_width=1.5, annotation_text="670 t/h",
                  annotation_font_color="#ef4444", annotation_font_size=9)
    fig.update_layout(**base_layout(alarm_shapes(data, "a2", "rgba(239,68,68,0.10)")))
    fig.update_yaxes(**y_style(560, 730, "t/h"))
    return fig


def make_total_chart(data):
    ts  = [r["t"] for r in data]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=ts, y=[r["tot"] for r in data], name="Total Coal",
                             line=dict(color="#facc15", width=1.8), mode="lines",
                             fill="tozeroy", fillcolor="rgba(250,204,21,0.06)"))
    fig.add_hline(y=1950, line_dash="dash", line_color="rgba(239,68,68,0.85)",
                  line_width=1.5, annotation_text="1,950 t/h",
                  annotation_font_color="#ef4444", annotation_font_size=9)
    fig.update_layout(**base_layout(alarm_shapes(data, "a4", "rgba(249,115,22,0.12)")))
    fig.update_yaxes(**y_style(1700, 2100, "t/h"))
    return fig


def make_ratio_chart(data):
    ts  = [r["t"] for r in data]
    fig = go.Figure()
    for key, name, color in [("u1r","Unit 1","#60a5fa"),
                              ("u2r","Unit 2","#34d399"),
                              ("u3r","Unit 3","#f472b6")]:
        fig.add_trace(go.Scatter(x=ts, y=[r[key] for r in data], name=name,
                                 line=dict(color=color, width=1.5), mode="lines"))
    fig.add_hline(y=1.08, line_dash="dash", line_color="rgba(239,68,68,0.85)",
                  line_width=1.5, annotation_text="1.08 limit",
                  annotation_font_color="#ef4444", annotation_font_size=9)
    fig.update_layout(**base_layout(alarm_shapes(data, "a3", "rgba(234,179,8,0.12)")))
    fig.update_yaxes(**y_style(0.95, 1.20, "t/MWh"))
    return fig


# ─────────────────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; margin-bottom:20px;">
  <div style="font-size:1.6rem; font-weight:700; color:#3b82f6;">⚡ Coal Consumption Dashboard</div>
  <div style="font-size:0.82rem; color:#8b95a3; margin-top:4px;">HONGSA Power Plant – CHS Operation Team</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
#  FILE UPLOAD
# ─────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "📂 อัปโหลดไฟล์ Coal Consumption จาก DCS/PI",
    type=["xlsx", "xls"],
    help="รองรับไฟล์ .xlsx / .xls จาก DCS หรือ PI System",
)

if uploaded is None:
    st.info("👆 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มแสดงผล Dashboard")
    st.stop()

# ─────────────────────────────────────────────────────────
#  PROCESS DATA
# ─────────────────────────────────────────────────────────
with st.spinner("⏳ กำลังประมวลผลข้อมูล…"):
    try:
        data = process_excel(uploaded.read())
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.stop()

df = pd.DataFrame(data)
st.success(
    f"✅ โหลดสำเร็จ — {len(data):,} records  |  "
    f"{df['t'].min().strftime('%d/%m/%y %H:%M')} – {df['t'].max().strftime('%d/%m/%y %H:%M')}"
)

# ─────────────────────────────────────────────────────────
#  SIDEBAR — FILTER
# ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Filter")
    zoom = st.radio("ช่วงเวลา", ["ทั้งหมด", "6 ชม. แรก", "6 ชม. สุดท้าย", "เฉพาะ Alarm Zone"],
                    index=0)
    if zoom == "6 ชม. แรก":
        data = data[:360]
    elif zoom == "6 ชม. สุดท้าย":
        data = data[-360:]
    elif zoom == "เฉพาะ Alarm Zone":
        alarm_idx = [i for i, r in enumerate(data) if r["a2"]]
        if alarm_idx:
            lo = max(0, alarm_idx[0] - 30)
            hi = min(len(data) - 1, alarm_idx[-1] + 30)
            data = data[lo:hi+1]

    st.markdown("---")
    st.markdown("**Alarm Thresholds**")
    st.markdown("🔴 Coal/unit > **650 t/h** (≥2 units, >2h)")
    st.markdown("🔴 Coal/unit > **670 t/h** (instant)")
    st.markdown("🟠 Total > **1,950 t/h**")
    st.markdown("🟡 Ratio > **1.08 t/MWh**")

# ─────────────────────────────────────────────────────────
#  KPI METRICS
# ─────────────────────────────────────────────────────────
u1c  = [r["u1c"] for r in data]
u2c  = [r["u2c"] for r in data]
u3c  = [r["u3c"] for r in data]
tots = [r["tot"] for r in data]
c1 = sum(1 for r in data if r["a1"])
c2 = sum(1 for r in data if r["a2"])
c3 = sum(1 for r in data if r["a3"])
c4 = sum(1 for r in data if r["a4"])

cols = st.columns(6)
cols[0].metric("U1 Avg Coal",  f"{sum(u1c)/len(u1c):.1f} t/h")
cols[1].metric("U2 Avg Coal",  f"{sum(u2c)/len(u2c):.1f} t/h")
cols[2].metric("U3 Avg Coal",  f"{sum(u3c)/len(u3c):.1f} t/h")
cols[3].metric("U1 Peak", f"{max(u1c):.1f} t/h",
               delta="⚠️ >670" if max(u1c) > 670 else "✅ ปกติ",
               delta_color="inverse" if max(u1c) > 670 else "off")
cols[4].metric("U2 Peak", f"{max(u2c):.1f} t/h",
               delta="⚠️ >670" if max(u2c) > 670 else "✅ ปกติ",
               delta_color="inverse" if max(u2c) > 670 else "off")
cols[5].metric("U3 Peak", f"{max(u3c):.1f} t/h",
               delta="⚠️ >670" if max(u3c) > 670 else "✅ ปกติ",
               delta_color="inverse" if max(u3c) > 670 else "off")

cols2 = st.columns(6)
cols2[0].metric("Total Avg",  f"{sum(tots)/len(tots):.0f} t/h")
cols2[1].metric("Total Peak", f"{max(tots):.0f} t/h",
                delta="⚠️ >1950" if max(tots) > 1950 else "✅ ปกติ",
                delta_color="inverse" if max(tots) > 1950 else "off")
cols2[2].metric("Cond.1 Alarm", f"{c1} นาที",
                delta="ALARM" if c1 > 0 else "ปกติ",
                delta_color="inverse" if c1 > 0 else "off")
cols2[3].metric("Cond.2 Alarm", f"{c2} นาที",
                delta="ALARM" if c2 > 0 else "ปกติ",
                delta_color="inverse" if c2 > 0 else "off")
cols2[4].metric("Ratio >1.08",  f"{c3} นาที",
                delta="ALARM" if c3 > 0 else "ปกติ",
                delta_color="inverse" if c3 > 0 else "off")
cols2[5].metric("Total >1950",  f"{c4} นาที",
                delta="ALARM" if c4 > 0 else "ปกติ",
                delta_color="inverse" if c4 > 0 else "off")

# ─────────────────────────────────────────────────────────
#  ALARM BANNERS
# ─────────────────────────────────────────────────────────
def banner(cls, icon, text):
    st.markdown(
        f'<div class="banner banner-{cls}"><span>{icon}</span><span>{text}</span></div>',
        unsafe_allow_html=True
    )

banner("ok" if c1 == 0 else "c1",
       "✅" if c1 == 0 else "🚨",
       "Cond.1 ปกติ — ไม่มี ≥2 units เกิน 650 t/h เกิน 2 ชม." if c1 == 0
       else f"<b>Cond.1 ALARM</b> — ≥2 units เกิน 650 t/h ต่อเนื่อง &gt;2 ชม. ({c1} นาที)")

banner("ok" if c2 == 0 else "c2",
       "✅" if c2 == 0 else "⚠️",
       "Cond.2 ปกติ — ไม่มี unit เกิน 670 t/h" if c2 == 0
       else f"<b>Cond.2 ALARM</b> — มี unit เกิน 670 t/h รวม {c2} นาที → เตรียม Start Coal Feeder Mill #7")

banner("ok" if c3 == 0 else "c2",
       "✅" if c3 == 0 else "📊",
       "Ratio ปกติ — Ratio ≤ 1.08 ตลอดช่วง" if c3 == 0
       else f"<b>Ratio &gt; 1.08</b> — Ratio Coal/Gross MW เกิน 1.08 รวม {c3} นาที")

banner("ok" if c4 == 0 else "tot",
       "✅" if c4 == 0 else "⚖️",
       "Total Coal ปกติ — Total ≤ 1,950 t/h ตลอดช่วง" if c4 == 0
       else f"<b>Total Coal &gt; 1,950 t/h</b> — เกิน Limit รวม {c4} นาที")

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
#  CHARTS
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-h">🔥 Coal Flow Rate per Unit (t/h)</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="thr-note">'
    '<b style="color:#ef4444">Cond.1:</b> ≥ 2 จาก 3 units &gt; 650 t/h ต่อเนื่อง &gt; 2 ชม. → เปลี่ยน Coal Stockpile'
    ' &nbsp;|&nbsp; '
    '<b style="color:#ef4444">Cond.2:</b> unit ใดเกิน 670 t/h → เตรียม Start Coal Feeder Mill #7'
    '</div>',
    unsafe_allow_html=True
)
st.plotly_chart(make_coal_chart(data), use_container_width=True)

st.markdown('<div class="section-h">⚖️ Total Coal Flow Rate – All 3 Units (t/h)</div>', unsafe_allow_html=True)
st.plotly_chart(make_total_chart(data), use_container_width=True)

st.markdown('<div class="section-h">📊 Ratio Coal / Gross MW (t/MWh)</div>', unsafe_allow_html=True)
st.plotly_chart(make_ratio_chart(data), use_container_width=True)

# ─────────────────────────────────────────────────────────
#  ALARM EVENT LIST
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-h">🚨 Alarm Events – Cond.2 (unit &gt; 670 t/h)</div>',
            unsafe_allow_html=True)
alarm_rows = [r for r in data if r["a2"]]
if not alarm_rows:
    st.markdown('<p style="color:#6b7280;font-size:0.82rem">ไม่มี Cond.2 Alarm ในช่วงนี้</p>',
                unsafe_allow_html=True)
else:
    rows_html = ""
    for r in alarm_rows:
        units = []
        if r["u1c"] > 670: units.append(f"U1: {r['u1c']:.1f}")
        if r["u2c"] > 670: units.append(f"U2: {r['u2c']:.1f}")
        if r["u3c"] > 670: units.append(f"U3: {r['u3c']:.1f}")
        ts_str = r["t"].strftime("%d/%m/%y %H:%M") if hasattr(r["t"], "strftime") else str(r["t"])
        rows_html += (f'<div class="alarm-row">'
                      f'<span class="alarm-time">{ts_str}</span>'
                      f'<span>{", ".join(units)} t/h</span></div>')
    st.markdown(
        f'<p style="font-size:0.75rem;color:#8b95a3;margin-bottom:6px">'
        f'Cond.2: unit เกิน 670 t/h ({len(alarm_rows)} จุด)</p>' + rows_html,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────────────────
st.markdown(
    f'<div style="text-align:center;font-size:0.7rem;color:#4b5563;margin-top:24px;padding-bottom:16px;">'
    f'Generated: {datetime.now().strftime("%d/%m/%Y %H:%M")} &nbsp;|&nbsp;'
    f' Source: DCS/PI Minute Snapshot &nbsp;|&nbsp; HONGSA CHS Operation'
    f'</div>',
    unsafe_allow_html=True
)
