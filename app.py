from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

# === SETTINGS ===
WORK_START = datetime.strptime("09:00", "%H:%M").time()
WORK_END = datetime.strptime("18:00", "%H:%M").time()
WORK_HOURS = 8


# --- LOGO ---
from pathlib import Path

logo_path_svg = Path("assets/logo.svg")

if logo_path_svg.exists():
    st.image(str(logo_path_svg), width=160)

st.set_page_config(page_title="Ish vaqti hisoblagich", page_icon="⏱️", layout="centered")
st.title("📊 Ish vaqti va yo‘q vaqtni hisoblash")
st.caption(
    "Excel faylni yuklang (attendance.xlsx). Faylda ustunlar: Дата, приход, уход bo‘lishi kerak."
)

uploaded_file = st.file_uploader("Excel faylni yuklang:", type=["xls", "xlsx"])

# --- Downloadable template for users ---
template_df = pd.DataFrame(
    [
        {"Дата": "2025-10-01", "приход": "09:00", "уход": "18:00"},
        {"Дата": "2025-10-02", "приход": "09:15", "уход": "17:30"},
        {"Дата": "2025-10-03", "приход": "(нет)", "уход": "(нет)"},
    ]
)
buf = BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    template_df.to_excel(writer, index=False, sheet_name="attendance_template")
template_bytes = buf.getvalue()
st.sidebar.download_button(
    label="Shablonni yuklab olish — Excel",
    data=template_bytes,
    file_name="attendance_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# === HELPERS ===
def detect_header_row(raw_df, token="Дата"):
    """Find the row index where the real header starts (contains 'Дата')."""
    matches = raw_df.apply(
        lambda r: r.astype(str).str.contains(token, case=False, na=False).any(), axis=1
    )
    return int(matches[matches].index[0]) if matches.any() else None


def find_column(df, token_list):
    """Find column name matching one of the given tokens."""
    for t in token_list:
        for c in df.columns:
            if t.lower() in str(c).lower():
                return c
    return None


def parse_time(date_val, time_val):
    """Convert time strings like '13:01 (1)' into datetime, else None."""
    if pd.isna(time_val):
        return None
    s = str(time_val).strip()
    if "(нет)" in s or s == "":
        return None
    s = s.split()[0]
    try:
        date_str = (
            date_val.strftime("%Y-%m-%d")
            if hasattr(date_val, "strftime")
            else str(date_val)
        )
        return datetime.strptime(f"{date_str} {s}", "%Y-%m-%d %H:%M")
    except Exception:
        return None


# === MAIN LOGIC ===
if uploaded_file:
    # 1️⃣ Read raw data
    raw = pd.read_excel(
        uploaded_file,
        header=None,
        engine="xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl",
    )
    header_row_idx = detect_header_row(raw, token="Дата")
    if header_row_idx is None:
        st.error("❌ Faylda 'Дата' sarlavhasi topilmadi.")
        st.stop()

    # 2️⃣ Load actual table
    df = pd.read_excel(
        uploaded_file,
        header=header_row_idx,
        engine="xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl",
    )

    # 3️⃣ Map columns
    col_date = find_column(df, ["дата", "date"])
    col_prihod = find_column(df, ["приход", "in", "entry"])
    col_uhod = find_column(df, ["уход", "out", "exit"])

    if not col_date or not col_prihod or not col_uhod:
        st.error("Kerakli ustunlar topilmadi. Fayl ustunlarini tekshiring.")
        st.write(df.head())
        st.stop()

    df = df[[col_date, col_prihod, col_uhod]].copy()
    df.rename(
        columns={col_date: "Дата", col_prihod: "приход", col_uhod: "уход"}, inplace=True
    )
    df["Дата"] = df["Дата"].ffill()

    # 4️⃣ Parse datetimes
    df["приход_time"] = df.apply(lambda r: parse_time(r["Дата"], r["приход"]), axis=1)
    df["уход_time"] = df.apply(lambda r: parse_time(r["Дата"], r["уход"]), axis=1)

    # 5️⃣ Calculate per-day
    per_day = []
    for date, group in df.groupby("Дата"):
        total = timedelta(0)
        first_in = None
        last_out = None
        for _, row in group.iterrows():
            if row["приход_time"] and row["уход_time"]:
                total += row["уход_time"] - row["приход_time"]
            if row["приход_time"] and (
                first_in is None or row["приход_time"] < first_in
            ):
                first_in = row["приход_time"]
            if row["уход_time"] and (last_out is None or row["уход_time"] > last_out):
                last_out = row["уход_time"]

        worked_hours = round(total.total_seconds() / 3600, 2)
        missing = round(max(WORK_HOURS - worked_hours, 0), 2)

        # safe time checks
        late_min = 0
        early_min = 0
        if first_in and pd.notna(first_in):
            late_delta = (first_in.hour * 60 + first_in.minute) - (
                WORK_START.hour * 60 + WORK_START.minute
            )
            late_min = max(late_delta, 0)
        if last_out and pd.notna(last_out):
            early_delta = (WORK_END.hour * 60 + WORK_END.minute) - (
                last_out.hour * 60 + last_out.minute
            )
            early_min = max(early_delta, 0)

        per_day.append(
            {
                "Дата": date,
                "Ish vaqti (soat)": worked_hours,
                "Ishxonada bo‘lmagan (soat)": missing,
                "Birinchi kirish": (
                    first_in.strftime("%H:%M") if pd.notna(first_in) else ""
                ),
                "Oxirgi chiqish": (
                    last_out.strftime("%H:%M") if pd.notna(last_out) else ""
                ),
                "Kechikish (min)": late_min,
                "Oldin chiqish (min)": early_min,
            }
        )

    result = pd.DataFrame(per_day).sort_values("Дата").reset_index(drop=True)
    total_absent = round(result["Ishxonada bo‘lmagan (soat)"].sum(), 2)

    st.subheader("📅 Kunlik hisob-kitob")
    st.metric("💡 Umumiy ishxonada bo‘lmagan vaqt", f"{total_absent} soat")
    st.dataframe(result, width="stretch")

    # === DASHBOARD CHARTS ===
    st.markdown("---")
    st.subheader("📈 Grafiklar")

    # parse date for charts
    try:
        result_chart = result.copy()
        result_chart["Дата_parsed"] = pd.to_datetime(
            result_chart["Дата"], errors="coerce"
        )
        result_chart = result_chart.sort_values("Дата_parsed")
        chart_indexed = result_chart.set_index("Дата_parsed")
    except Exception:
        chart_indexed = result.set_index("Дата")

    # sidebar controls for charts
    show_charts = st.sidebar.checkbox("Grafiklarni ko'rsatish", value=True)
    chart_type = st.sidebar.selectbox(
        "Grafik turi", ["Chiziqli (line)", "Maydon (area)"]
    )

    if show_charts:
        st.markdown("**Ish vaqti (kunlik)**")
        if chart_type == "Chiziqli (line)":
            st.line_chart(chart_indexed["Ish vaqti (soat)"])
        else:
            st.area_chart(chart_indexed["Ish vaqti (soat)"])

        st.markdown("**Ishxonada bo‘lmagan (soat) — kunlik**")
        st.bar_chart(chart_indexed["Ishxonada bo‘lmagan (soat)"])


else:
    st.info(
        "Excel faylni yuklang — dastur avtomatik 'Дата' sarlavhasidan boshlab o‘qiydi."
    )


# --- Footer ---
footer_html = """
<div style="width:100%; padding:12px 0; border-top:1px solid #e6e6e6; margin-top:24px; text-align:center; color:#6b7280; font-size:13px;">
    © 2025 MyWorkTime by Davronbek. Barcha huquqlar himoyalangan.
</div>
"""
st.markdown(footer_html, unsafe_allow_html=True)
