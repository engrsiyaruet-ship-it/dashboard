import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Student Research Progress", page_icon="🎓", layout="wide")

DATA = Path("data.xlsx")
ADMIN_PASSWORD = os.environ.get("DASH_PASSWORD", "")

# ------------------------------------------------------------------ styling
ACCENT = ["#22d3ee", "#a78bfa", "#f472b6", "#fbbf24", "#34d399", "#60a5fa", "#fb923c"]
STATUS_COLORS = {
    "Drafting": "#94a3b8", "Submitted": "#60a5fa", "Review Pending": "#22d3ee",
    "Under Review": "#a78bfa", "Revision in Progress": "#fbbf24",
    "Accepted": "#34d399", "Published": "#10b981", "Rejected": "#f87171",
}
PRIORITY_COLORS = {"Critical": "#f87171", "High": "#fb923c", "Medium": "#fbbf24", "Low": "#34d399"}

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px;}
#MainMenu, footer, header[data-testid="stHeader"] {visibility: hidden; height: 0;}
.hero {background: linear-gradient(120deg,#0ea5e9 0%,#6366f1 50%,#a855f7 100%);
       border-radius: 20px; padding: 26px 32px; margin-bottom: 18px;
       box-shadow: 0 10px 30px rgba(99,102,241,.35);}
.hero h1 {color:#fff; margin:0; font-size: 2rem; font-weight: 800; letter-spacing:.3px;}
.hero p {color:rgba(255,255,255,.88); margin:4px 0 0 0; font-size:.95rem;}
.kpi {background: linear-gradient(145deg,#1e293b,#0f172a); border:1px solid rgba(148,163,184,.18);
      border-radius:16px; padding:16px 18px; border-top:4px solid var(--c); height:100%;
      box-shadow: 0 6px 18px rgba(0,0,0,.25);}
.kpi .lbl {color:#94a3b8; font-size:.72rem; text-transform:uppercase; letter-spacing:1.2px; font-weight:600;}
.kpi .val {color:#f8fafc; font-size:2rem; font-weight:800; line-height:1.15;}
.kpi .sub {color:var(--c); font-size:.78rem; font-weight:600;}
.scard {background: linear-gradient(145deg,#1e293b,#111827); border:1px solid rgba(148,163,184,.18);
        border-radius:18px; padding:18px 20px; box-shadow: 0 6px 18px rgba(0,0,0,.25);}
.scard h3 {margin:0; color:#f8fafc; font-size:1.25rem;}
.scard .meta {color:#94a3b8; font-size:.82rem; margin-bottom:10px;}
.bar {background:#0b1220; border-radius:99px; height:12px; overflow:hidden; margin:4px 0 10px 0;}
.bar > div {height:100%; border-radius:99px;}
.pill {display:inline-block; padding:2px 10px; border-radius:99px; font-size:.72rem; font-weight:700; color:#0b1220;}
.feed {border-left:3px solid var(--c); background:#111827; border-radius:0 12px 12px 0;
       padding:10px 14px; margin-bottom:10px;}
.feed b {color:#f8fafc;} .feed span {color:#94a3b8; font-size:.82rem;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {background:#1e293b; border-radius:10px 10px 0 0; padding:8px 18px;}
</style>
""", unsafe_allow_html=True)


def style(fig, h=380, legend=True):
    fig.update_layout(
        template="plotly_dark", height=h, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=50, b=10), showlegend=legend,
        font=dict(family="Inter, Segoe UI, sans-serif", size=13),
        title=dict(font=dict(size=16)), legend=dict(orientation="h", y=-0.15),
    )
    fig.update_xaxes(gridcolor="rgba(148,163,184,.12)")
    fig.update_yaxes(gridcolor="rgba(148,163,184,.12)")
    return fig


def kpi(label, value, sub, color):
    return f'<div class="kpi" style="--c:{color}"><div class="lbl">{label}</div><div class="val">{value}</div><div class="sub">{sub}</div></div>'


# ------------------------------------------------------------------ data
@st.cache_data(show_spinner="Loading data...")
def load(path: str, mtime: float):
    raw = pd.read_excel(path, sheet_name=None, header=None)

    def table(name):
        d = raw[name]
        hdr = next(i for i in range(len(d)) if str(d.iloc[i, 0]).endswith("_ID") and d.iloc[i].notna().sum() >= 4)
        out = d.iloc[hdr + 1:].copy()
        out.columns = [str(c).strip() for c in d.iloc[hdr]]
        out = out.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        for c in out.columns:
            if any(k in c for k in ("Date", "Deadline", "Expected_Completion", "Updated")):
                out[c] = pd.to_datetime(out[c], errors="coerce")
            elif c.endswith("_pct") or c in ("Credits", "Current_Year"):
                out[c] = pd.to_numeric(out[c], errors="coerce")
        return out

    s = table("Students")
    names = s.set_index("Student_ID")["Student_Name"]
    tabs = {"students": s}
    for key, sheet in [("courses", "Coursework"), ("ms", "Manuscripts"),
                       ("dl", "Deadlines"), ("log", "Progress Log")]:
        t = table(sheet)
        t.insert(1, "Student", t["Student_ID"].map(names))
        tabs[key] = t
    return tabs


# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.markdown("### ⚙️ Controls")
    asof = pd.Timestamp(st.date_input("Dashboard date", value=pd.Timestamp.today().date()))

    with st.expander("🔒 Admin: update data"):
        pw = st.text_input("Password", type="password")
        if ADMIN_PASSWORD and pw == ADMIN_PASSWORD:
            up = st.file_uploader("Upload new Excel file", type=["xlsx"])
            if up is not None and (not DATA.exists() or DATA.read_bytes() != up.getvalue()):
                DATA.write_bytes(up.getvalue())
                st.cache_data.clear()
                st.success("Data updated ✔")
        elif pw:
            st.error("Wrong password")

T = load(str(DATA), DATA.stat().st_mtime)
stu, cw, ms, dl, log = T["students"], T["courses"], T["ms"], T["dl"], T["log"]

with st.sidebar:
    picked = st.multiselect("Students", stu["Student_Name"].tolist(), default=stu["Student_Name"].tolist())
    degrees = st.multiselect("Degree", sorted(stu["Degree"].dropna().unique()),
                             default=sorted(stu["Degree"].dropna().unique()))
    st.caption(f"Data last updated: {stu['Last_Updated'].max():%d %b %Y}")

stu = stu[stu["Student_Name"].isin(picked) & stu["Degree"].isin(degrees)]
names_ok = stu["Student_Name"].tolist()
cw, ms, dl, log = [d[d["Student"].isin(names_ok)] for d in (cw, ms, dl, log)]
color_map = {n: ACCENT[i % len(ACCENT)] for i, n in enumerate(T["students"]["Student_Name"])}

# derived metrics
dl = dl.copy()
dl["Days_Left"] = (dl["Due_Date"] - asof).dt.days
done_words = ("complete", "done", "submitted")
dl_open = dl[~dl["Status"].str.lower().isin(done_words)]
upcoming = dl_open[(dl_open["Days_Left"] >= 0) & (dl_open["Days_Left"] <= 30)]
overdue = dl_open[dl_open["Days_Left"] < 0]

total_span = (stu["Expected_Completion"] - stu["Start_Date"]).dt.days
stu = stu.assign(
    Time_Elapsed_pct=((asof - stu["Start_Date"]).dt.days / total_span * 100).clip(0, 100).round(0),
)
stu["Gap"] = stu["Overall_Progress_pct"] - stu["Time_Elapsed_pct"]
stu["Health"] = np.where(stu["Gap"] >= -10, "On track", "Needs attention")

# ------------------------------------------------------------------ hero
st.markdown(f"""
<div class="hero"><h1>🎓 Student Research Progress Dashboard</h1>
<p>Live overview of coursework, manuscripts, deadlines and weekly progress • as of {asof:%d %B %Y}</p></div>
""", unsafe_allow_html=True)

if stu.empty:
    st.warning("No students selected.")
    st.stop()

pub_done = ms["Manuscript_Status"].isin(["Accepted", "Published"]).sum()
in_rev = ms["Manuscript_Status"].isin(["Under Review", "Review Pending", "Revision in Progress"]).sum()
cols = st.columns(6)
cards = [
    ("Students", len(stu), f"{(stu['Degree'] == 'PhD').sum()} PhD • {(stu['Degree'] == 'MS').sum()} MS", ACCENT[0]),
    ("Avg. progress", f"{stu['Overall_Progress_pct'].mean():.0f}%", "across selected students", ACCENT[1]),
    ("Manuscripts", len(ms), f"{pub_done} accepted / published", ACCENT[2]),
    ("In peer review", in_rev, "under review or revision", ACCENT[3]),
    ("Due in 30 days", len(upcoming), f"{(upcoming['Priority'].isin(['Critical', 'High'])).sum()} high priority", ACCENT[4]),
    ("Overdue", len(overdue), "open items past due date", "#f87171"),
]
for c, (a, b, s_, col) in zip(cols, cards):
    c.markdown(kpi(a, b, s_, col), unsafe_allow_html=True)
st.write("")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📊 Overview", "👩‍🎓 Students", "📚 Coursework", "📝 Manuscripts", "⏰ Deadlines", "📈 Progress log"])

# ------------------------------------------------------------------ overview
with tab1:
    n = len(stu)
    g = go.Figure()
    for i, (_, r) in enumerate(stu.iterrows()):
        g.add_trace(go.Indicator(
            mode="gauge+number", value=r["Overall_Progress_pct"], number={"suffix": "%", "font": {"size": 34}},
            title={"text": f"<b>{r['Student_Name']}</b><br><span style='font-size:12px;color:#94a3b8'>{r['Degree']} • {r['Research_Area'][:34]}</span>"},
            domain={"x": [i / n + 0.01, (i + 1) / n - 0.01], "y": [0, 1]},
            gauge={"axis": {"range": [0, 100], "tickcolor": "#475569"}, "bar": {"color": color_map[r["Student_Name"]], "thickness": 0.35},
                   "bgcolor": "#0b1220", "borderwidth": 0,
                   "threshold": {"line": {"color": "#f8fafc", "width": 3}, "thickness": .8, "value": r["Time_Elapsed_pct"]}}))
    style(g, 320, False)
    g.update_layout(title="Overall research progress (white marker = share of program time elapsed)")
    st.plotly_chart(g, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        f = go.Figure()
        f.add_bar(x=stu["Student_Name"], y=stu["Time_Elapsed_pct"], name="Time elapsed", marker_color="#475569")
        f.add_bar(x=stu["Student_Name"], y=stu["Overall_Progress_pct"], name="Progress",
                  marker_color=[color_map[x] for x in stu["Student_Name"]], text=stu["Overall_Progress_pct"].astype(int).astype(str) + "%",
                  textposition="outside")
        f.update_layout(barmode="group", title="Progress vs. time elapsed", yaxis_range=[0, 110])
        st.plotly_chart(style(f), width="stretch")
    with c2:
        cnt = ms["Manuscript_Status"].value_counts().reset_index()
        cnt.columns = ["Status", "Count"]
        f = px.pie(cnt, names="Status", values="Count", hole=.62, color="Status", color_discrete_map=STATUS_COLORS,
                   title="Manuscript pipeline")
        f.update_traces(textinfo="value", textfont_size=15)
        f.add_annotation(text=f"<b>{len(ms)}</b><br>papers", showarrow=False, font=dict(size=22))
        st.plotly_chart(style(f), width="stretch")

    st.markdown("#### 🔔 Next 30 days")
    show = upcoming.sort_values("Due_Date")[["Student", "Task_or_Milestone", "Category", "Priority", "Due_Date", "Days_Left"]]
    st.dataframe(show, hide_index=True, width="stretch", column_config={
        "Due_Date": st.column_config.DateColumn("Due", format="DD MMM YYYY"),
        "Days_Left": st.column_config.ProgressColumn("Days left", min_value=0, max_value=30, format="%d d"),
        "Task_or_Milestone": "Task"})

# ------------------------------------------------------------------ students
with tab2:
    cc = st.columns(min(len(stu), 3))
    for i, (_, r) in enumerate(stu.iterrows()):
        col = color_map[r["Student_Name"]]
        health_col = "#34d399" if r["Health"] == "On track" else "#fb923c"
        n_ms = (ms["Student"] == r["Student_Name"]).sum()
        n_cw = (cw["Student"] == r["Student_Name"]).sum()
        with cc[i % len(cc)]:
            st.markdown(f"""
            <div class="scard">
              <h3>{r['Student_Name']} <span class="pill" style="background:{health_col};float:right">{r['Health']}</span></h3>
              <div class="meta">{r['Degree']} • Year {int(r['Current_Year'])} • {r['Supervisor']}<br>{r['Research_Area']}</div>
              <div style="color:#94a3b8;font-size:.78rem">Overall progress — {r['Overall_Progress_pct']:.0f}%</div>
              <div class="bar"><div style="width:{r['Overall_Progress_pct']}%;background:{col}"></div></div>
              <div style="color:#94a3b8;font-size:.78rem">Time elapsed — {r['Time_Elapsed_pct']:.0f}%</div>
              <div class="bar"><div style="width:{r['Time_Elapsed_pct']}%;background:#475569"></div></div>
              <div style="color:#cbd5e1;font-size:.85rem">📅 {r['Start_Date']:%b %Y} → {r['Expected_Completion']:%b %Y}<br>
              📚 {n_cw} courses &nbsp; 📝 {n_ms} manuscripts</div>
              <div style="color:#94a3b8;font-size:.8rem;margin-top:8px"><i>{r.get('Notes', '')}</i></div>
            </div><br>""", unsafe_allow_html=True)

# ------------------------------------------------------------------ coursework
with tab3:
    c1, c2 = st.columns([3, 2])
    with c1:
        d = cw.sort_values(["Student", "Course_Deadline"])
        f = px.bar(d, y="Course_Name", x="Completion_pct", color="Student", orientation="h", color_discrete_map=color_map,
                   text="Completion_pct", title="Course completion (%)", hover_data=["Semester", "Course_Status", "Grade"])
        f.update_traces(texttemplate="%{text}%", textposition="inside")
        f.update_layout(xaxis_range=[0, 100], yaxis=dict(autorange="reversed", title=""))
        st.plotly_chart(style(f, 120 + 34 * len(d)), width="stretch")
    with c2:
        cr = cw.groupby(["Student", "Course_Status"])["Credits"].sum().reset_index()
        f = px.bar(cr, x="Student", y="Credits", color="Course_Status", title="Credits by status",
                   color_discrete_sequence=["#34d399", "#fbbf24", "#94a3b8", "#60a5fa"])
        st.plotly_chart(style(f, 360), width="stretch")
        done = cw.loc[cw["Course_Status"] == "Completed", "Credits"].sum()
        st.metric("Credits completed", f"{done:.0f} / {cw['Credits'].sum():.0f}")
    st.dataframe(cw.drop(columns=["Course_ID", "Student_ID"]), hide_index=True, width="stretch", column_config={
        "Completion_pct": st.column_config.ProgressColumn("Completion", min_value=0, max_value=100, format="%d%%"),
        "Course_Deadline": st.column_config.DateColumn("Deadline", format="DD MMM YYYY")})

# ------------------------------------------------------------------ manuscripts
with tab4:
    c1, c2 = st.columns(2)
    with c1:
        d = ms.groupby(["Student", "Manuscript_Status"]).size().reset_index(name="Count")
        f = px.bar(d, x="Student", y="Count", color="Manuscript_Status", color_discrete_map=STATUS_COLORS,
                   title="Manuscripts per student by status")
        st.plotly_chart(style(f), width="stretch")
    with c2:
        d = ms.dropna(subset=["Submission_Date"])
        f = px.scatter(d, x="Submission_Date", y="Student", color="Manuscript_Status", color_discrete_map=STATUS_COLORS,
                       hover_data=["Manuscript_Title", "Target_Journal"], title="Submission timeline")
        f.update_traces(marker=dict(size=16, line=dict(width=1, color="white")))
        st.plotly_chart(style(f), width="stretch")
    st.dataframe(ms[["Student", "Manuscript_Title", "Target_Journal", "Manuscript_Status", "Current_Stage",
                     "Next_Action", "Action_Deadline", "DOI_or_ID"]], hide_index=True, width="stretch",
                 column_config={"Action_Deadline": st.column_config.DateColumn("Action due", format="DD MMM YYYY")})

# ------------------------------------------------------------------ deadlines
with tab5:
    d = dl_open.sort_values("Due_Date")
    f = px.scatter(d, x="Due_Date", y="Student", color="Priority", color_discrete_map=PRIORITY_COLORS,
                   hover_data=["Task_or_Milestone", "Category", "Status"], title="Deadline radar")
    f.update_traces(marker=dict(size=18, line=dict(width=1, color="white"), symbol="diamond"))
    f.add_shape(type="line", x0=asof, x1=asof, y0=0, y1=1, yref="paper", line=dict(color="#f8fafc", dash="dot"))
    f.add_annotation(x=asof, y=1.05, yref="paper", text="Today", showarrow=False)
    st.plotly_chart(style(f, 340), width="stretch")

    d = d.assign(Alert=np.select([d["Days_Left"] < 0, d["Days_Left"] <= 7, d["Days_Left"] <= 14], ["🔴 Overdue", "🔴 ≤ 7 days", "🟠 ≤ 14 days"], "🟢 OK"))
    st.dataframe(d[["Alert", "Student", "Task_or_Milestone", "Category", "Priority", "Status", "Due_Date", "Days_Left", "Required_Action"]],
                 hide_index=True, width="stretch",
                 column_config={"Due_Date": st.column_config.DateColumn("Due", format="DD MMM YYYY")})

# ------------------------------------------------------------------ progress log
with tab6:
    d = log.sort_values("Update_Date")
    f = px.line(d, x="Update_Date", y="Progress_pct", color="Student", markers=True, color_discrete_map=color_map,
                title="Progress trend from weekly updates", hover_data=["Activity_Area"])
    f.update_traces(line=dict(width=4), marker=dict(size=11))
    f.update_layout(yaxis_range=[0, 100])
    st.plotly_chart(style(f, 400), width="stretch")
    st.markdown("#### 🗒️ Latest updates")
    for _, r in d.sort_values("Update_Date", ascending=False).head(8).iterrows():
        st.markdown(f"""<div class="feed" style="--c:{color_map.get(r['Student'], '#22d3ee')}">
        <b>{r['Student']}</b> • <span>{r['Update_Date']:%d %b %Y} • {r['Activity_Area']} • {r['Progress_pct']:.0f}%</span><br>
        ✅ {r['Achievement_or_Update']}<br>⚠️ <span>{r['Blocker_or_Risk']}</span><br>➡️ <span>{r['Next_Step']}</span></div>""",
                    unsafe_allow_html=True)

st.caption("Built with Streamlit & Plotly • demo data are fictional")
