"""
Former Executive Tracker
========================
Enter any US equity ticker → surface former C-suite, EVPs, and SVPs,
sorted by division with LinkedIn search links.

Data sources
------------
  SEC EDGAR  — 10-K executive officer lists + DEF 14A proxy compensation
               tables (free, always available, covers Named Exec Officers)
  Revelio Labs — workforce intelligence API (optional API key; broadens
               SVP coverage beyond SEC's ~5 named officers per year)

Run:  streamlit run app.py
"""

import os
import streamlit as st
import pandas as pd

import sec_parser as sec
import revelio_client as revelio
from classify import classify, linkedin_search_url

# ── Page config ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Former Executive Tracker",
    page_icon="🔍",
    layout="wide",
)

# ── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("🔍 Former Exec Tracker")
    st.caption("Surface senior executives who have left a company.")

    ticker = st.text_input("Ticker symbol", placeholder="e.g. ABT, JNJ, PFE").strip().upper()

    st.divider()
    st.subheader("Filters")
    seniority_filter = st.multiselect(
        "Seniority levels",
        ["CEO", "President", "COO", "CFO", "CTO", "CIO", "CMO", "CLO / General Counsel",
         "CHRO", "CRO", "CSO", "CAO / Controller", "Chief Commercial Officer",
         "Chief Medical Officer", "Chief Scientific Officer",
         "Chief Compliance Officer", "Chief Supply Chain Officer", "CPO",
         "Vice Chairman", "Chairman", "EVP", "SVP"],
        default=["CEO", "President", "COO", "CFO", "CTO", "CIO", "CMO",
                 "CLO / General Counsel", "CHRO", "CRO", "CSO", "EVP", "SVP",
                 "CAO / Controller", "Chief Commercial Officer",
                 "Chief Medical Officer", "Chief Scientific Officer"],
    )
    lookback = st.slider("Look back (years)", 5, 15, 10)

    st.divider()
    st.subheader("Revelio Labs (optional)")
    st.caption(
        "Adds SVP-level coverage beyond SEC proxy filings. "
        "[Get API access](https://www.reveliolabs.com)"
    )
    revelio_key = st.text_input("Revelio API key", type="password",
                                value=os.environ.get("REVELIO_API_KEY", ""))

    search_clicked = st.button("Search", type="primary", use_container_width=True)

# ── Main ─────────────────────────────────────────────────────────────────────

st.title("Former Executive Tracker")

if not search_clicked or not ticker:
    st.info("Enter a ticker in the sidebar and click **Search** to get started.")
    st.markdown("""
**How it works**
1. Fetches the last 10 years of SEC 10-K and DEF 14A filings via EDGAR
2. Parses executive officer tables and compensation disclosures
3. Cross-references year-over-year to identify who has departed
4. Optionally enriches with Revelio Labs workforce data for broader SVP coverage
5. Generates LinkedIn search links for each person

**Coverage**
| Source | What you get |
|--------|-------------|
| SEC 10-K | All disclosed executive officers (typically 10–20 per year) |
| SEC DEF 14A | Named Executive Officers in comp tables (~5/year) |
| Revelio Labs | SVP+ from LinkedIn-sourced workforce data (requires API key) |
""")
    st.stop()

# ── Lookup ───────────────────────────────────────────────────────────────────

with st.spinner(f"Looking up {ticker} on SEC EDGAR…"):
    cik_result = sec.get_cik_and_name(ticker)

if not cik_result:
    st.error(f"Ticker **{ticker}** not found on SEC EDGAR. Try the full company name, or check the ticker.")
    st.stop()

cik, company_name = cik_result
st.header(f"{company_name} ({ticker})")
st.caption(f"SEC CIK: {int(cik)}")

# ── Fetch filings ─────────────────────────────────────────────────────────

progress = st.progress(0, text="Fetching SEC filings…")

with st.spinner("Fetching 10-K and DEF 14A filing index…"):
    filings_10k = sec.get_filings(cik, "10-K", max_results=lookback)
    filings_14a = sec.get_filings(cik, "DEF 14A", max_results=lookback)

all_filings = filings_10k + filings_14a
total = len(all_filings)

if not all_filings:
    st.warning("No 10-K or DEF 14A filings found. The company may file under a different form type.")
    st.stop()

# ── Parse each filing ─────────────────────────────────────────────────────

yearly: dict[int, list[dict]] = {}
errors = []

for idx, filing in enumerate(all_filings):
    year = filing["year"]
    form = "10-K" if filing in filings_10k else "DEF 14A"
    progress.progress((idx + 1) / total, text=f"Parsing {form} {year}…")

    html = sec.fetch_filing_html(cik, filing)
    if not html:
        errors.append(f"Could not download {form} {year}")
        continue

    execs = sec.parse_executives(html)
    if year not in yearly:
        yearly[year] = []
    # Merge (same person may appear in both 10-K and DEF 14A for same year)
    existing_keys = {e["name"].lower() for e in yearly[year]}
    for e in execs:
        if e["name"].lower() not in existing_keys:
            yearly[year].append(e)
            existing_keys.add(e["name"].lower())

progress.empty()

if not yearly:
    st.error("Could not parse executive data from any filings. Try a larger company with full HTML filings.")
    st.stop()

current_year = max(yearly.keys())

# ── Revelio enrichment ───────────────────────────────────────────────────

revelio_rows: list[dict] = []
if revelio_key:
    with st.spinner("Fetching Revelio Labs workforce data…"):
        try:
            revelio_rows = revelio.fetch_former_executives(company_name, revelio_key)
            st.success(f"Revelio: {len(revelio_rows)} records fetched.")
        except RuntimeError as e:
            st.warning(f"Revelio error: {e}")

# ── Build former exec table ───────────────────────────────────────────────

sec_formers = sec.find_formers(yearly, current_year, lookback_years=lookback)

# Merge SEC + Revelio; deduplicate by name (Revelio wins for richer data)
merged: dict[str, dict] = {}
for row in sec_formers:
    merged[row["name"].lower()] = {**row, "source": "SEC EDGAR"}

for row in revelio_rows:
    key = row["name"].lower()
    if key not in merged:
        merged[key] = row
    else:
        # Revelio adds current employer and precise dates
        merged[key].update({k: v for k, v in row.items() if v})

records = list(merged.values())

# ── Classify and filter ───────────────────────────────────────────────────

enriched = []
for r in records:
    seniority, division, rank = classify(r.get("title", ""))
    if seniority_filter and seniority not in seniority_filter:
        continue
    enriched.append({
        "Name": r["name"],
        "Former Title": r.get("title", ""),
        "Seniority": seniority,
        "Division": r.get("division") or division,
        "_rank": rank,
        "First Seen": r.get("first_seen"),
        "Last Seen": r.get("last_seen"),
        "Departed After": r.get("departed_after"),
        "Current Employer": r.get("current_employer", ""),
        "Source": r.get("source", "SEC EDGAR"),
        "LinkedIn": linkedin_search_url(r["name"], company_name),
    })

if not enriched:
    st.warning("No former executives found matching your filters. Try widening seniority or lookback range.")
    st.stop()

df = (
    pd.DataFrame(enriched)
    .sort_values(["_rank", "Division", "Last Seen"], ascending=[True, True, False])
    .drop(columns=["_rank"])
    .reset_index(drop=True)
)

# ── Stats bar ─────────────────────────────────────────────────────────────

col1, col2, col3, col4 = st.columns(4)
col1.metric("Former Executives Found", len(df))
col2.metric("Divisions Represented", df["Division"].nunique())
col3.metric("Earliest Departure", df["Departed After"].min())
col4.metric("Most Recent Departure", df["Departed After"].max())

if errors:
    with st.expander(f"⚠️ {len(errors)} filing(s) could not be downloaded"):
        for e in errors:
            st.write(f"- {e}")

# ── Tabs ──────────────────────────────────────────────────────────────────

tab_div, tab_timeline, tab_raw = st.tabs(["By Division", "Timeline", "Full Table"])

with tab_div:
    st.subheader("Former Executives by Division")
    divisions = sorted(df["Division"].unique())
    for div in divisions:
        div_df = df[df["Division"] == div].copy()
        st.markdown(f"### {div} &nbsp; <small style='color:grey'>({len(div_df)})</small>", unsafe_allow_html=True)

        display = div_df[["Name", "Former Title", "Seniority", "Last Seen", "Departed After", "Current Employer", "LinkedIn"]].copy()

        # Render LinkedIn as clickable link
        def _link(url):
            return f'<a href="{url}" target="_blank">Search ↗</a>'

        display["LinkedIn"] = display["LinkedIn"].apply(_link)
        st.write(
            display.to_html(escape=False, index=False),
            unsafe_allow_html=True,
        )
        st.divider()

with tab_timeline:
    st.subheader("Departure Timeline")
    timeline_df = (
        df.dropna(subset=["Departed After"])
        .groupby("Departed After")
        .agg(Count=("Name", "count"), Names=("Name", lambda x: ", ".join(x)))
        .reset_index()
        .sort_values("Departed After")
    )
    st.bar_chart(timeline_df.set_index("Departed After")["Count"])

    st.dataframe(
        timeline_df.rename(columns={"Departed After": "Year"}),
        use_container_width=True,
        hide_index=True,
    )

with tab_raw:
    st.subheader("All Results")
    display_df = df.copy()
    display_df["LinkedIn"] = display_df["LinkedIn"].apply(
        lambda u: f'<a href="{u}" target="_blank">Search ↗</a>'
    )
    st.write(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

    csv = df.drop(columns=["LinkedIn"]).to_csv(index=False)
    st.download_button(
        "Download CSV",
        data=csv,
        file_name=f"{ticker}_former_executives.csv",
        mime="text/csv",
    )

# ── Footer ────────────────────────────────────────────────────────────────

st.divider()
st.caption(
    f"Data from SEC EDGAR (10-K + DEF 14A filings, {current_year - lookback}–{current_year})"
    + (" + Revelio Labs" if revelio_key else "")
    + " · For internal research use only · Not investment advice"
)
