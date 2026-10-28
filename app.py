# ── DEPLOYMENT ──────────────────────────────────────────────────────────
# 1. Push this repo to GitHub (app.py, requirements.txt, and the data/ folder
#    with the four CSVs — they're small, static GSC exports used for Tab 2).
# 2. Go to share.streamlit.io -> "New app"
# 3. Select repo, main file = app.py, Python 3.11 -> Deploy
#    -> Instant public shareable URL, no login required to view
#
# PREREQUISITE: The Google Sheets below must be "Anyone with the link ->
# Viewer". As of this writing the two Page&Query sheets (used by the
# Per-Page Keyword Explorer) are NOT shared that way yet — fix that in
# Google Sheets (Share -> General access) or the Explorer tab will show an
# error instead of data.
#
# NOTE ON HOW SHEETS ARE READ: these workbooks are not single flat CSVs.
# Each is a multi-tab workbook with one tab per month (e.g. "Jun'26"), plus
# a "Consolidated Data Final" tab that already has the monthly
# position-bucket breakdown pre-computed. There's no stable public "gid"
# worth hardcoding (gids don't survive a sheet reorg, and Google's own xlsx
# export reassigns them anyway), but Google Sheets' gviz endpoint lets you
# fetch any tab BY NAME instead — see gviz_csv_url() below. That's what
# this app uses everywhere instead of gid-based export links.
# ────────────────────────────────────────────────────────────────────────

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from urllib.parse import quote
import datetime

st.set_page_config(page_title="Canvas & Weaves SEO Dashboard", layout="wide")

# ── Constants ─────────────────────────────────────────────────────────────

BUCKETS = ['1-5', '6-10', '11-20', '21-30', '31-60', '61-90', '90+']
BUCKET_COLORS = ['#1d4ed8', '#0ea5e9', '#d946ef', '#f97316', '#eab308', '#22c55e', '#94a3b8']
BUCKET_COLOR_MAP = dict(zip(BUCKETS, BUCKET_COLORS))

PAGE_TYPE_COLORS = {
    'fashion-collection': '#1d4ed8', 'art-collection': '#f97316',
    'designer': '#d946ef', 'artist': '#eab308',
    'product': '#16a34a', 'blog': '#0ea5e9', 'other': '#94a3b8'
}
PAGE_TYPE_FILTERS = {
    'All': None, 'Fashion': 'fashion-collection', 'Art Collections': 'art-collection',
    'Designers': 'designer', 'Artists': 'artist', 'Blog': 'blog', 'Product': 'product'
}

ARTIST_SLUGS = {'ajay-de', 'alexandra-gallagher', 'anand-bekwad', 'anjali-mehta', 'anu-kulkarni',
                'arty-guava', 'asit-patnaik', 'bhuwal-prasad', 'bijoy-basak', 'buwa-shete', 'debasish-dutta',
                'deepa-vedpathak', 'em-co-art', 'emi-avora', 'fawad-tamkanat', 'gurmeet-marwah', 'hemavathy-guha',
                'hiral-chaudhari', 'irene-hoff', 'jaya-charan', 'kandi-narsimulu', 'ketki-fadnis', 'kosha-shah',
                'kr-santhana-krishnan', 'lekha-shastry', 'lou-lauwerys', 'louise-hill', 'manisha-agrawal',
                'medha-easwar', 'n-harinath', 'nagesh-ghodke', 'namrata-kumar', 'nandini-choudhary',
                'nayanaa-kanodia', 'nitin-nangare', 'pappu-bardhan', 'phan-thu-trang', 'pooja-acharya',
                'pr-narvekar', 'prapti', 'raghu-rai', 'richa-kashelkar', 'sachin-jaltare', 'satish-kale',
                'shiv-kumar-soni', 'sushil-soni', 'talking-textiles-by-deborah-mckellar', 'tanya-timble',
                'the-east-story', 'tilla', 'vraj-bhoomi', 'yesha-sheth', 'zoya-chaudhary', 'amol-pawar'}

DESIGNER_SLUGS = {'145-east', 'aarjavee', 'aavidi', 'akaaro', 'akaaro-men', 'aomidori-shimai',
                   'ashepa', 'capisvirleo', 'celestia-home', 'chola', 'cord', 'doodlage', 'dot', 'dyelogue',
                   'facets-by-garima', 'gado-living', 'homegrown-by-alterego', 'indigo-dreams', 'ituvana',
                   'jodilife', 'jodi-men', 'ka-sha', 'khara-kapas', 'khara-kapas-men', 'khara-kapas-jewellery',
                   'krithaa-jewellery', 'kritika-murarka', 'label-mogra', 'leh-studios', 'mac-duggal',
                   'maison-shefali', 'mallika-mathur', 'manan', 'marche', 'mason-and-mill', 'mati', 'mati-men',
                   'minimal-indian', 'moborr', 'myoho', 'nadiya-paar', 'nirjara', 'nirjara-men', 'nirjara-jewellery',
                   'niti-bothra', 'noib', 'ode-to-odd', 'ofrida', 'paulmi-harsh', 'prahnaaya', 'priti-prashant',
                   'rajiramniq', 'rashika-mittal', 'rias', 'sabo-by-sabina', 'saphed', 'saphed-men', 'shades-of-india',
                   'sheetal-batra', 'sobariko', 'stoique', 'sui', 'suparna-som', 'surmeyi', 'swatti-kapoor',
                   'the-burlap-people', 'the-summer-house', 'the-terra-tribe', 'three', 'three-men', 'tokree',
                   'translate-handwoven-ikat', 'translate-men', 'turn-black', 'world-of-crow', 'world-of-sisa',
                   'yam', 'zanzibar', 'zanzibar-living-jewellery'}

ART_GENRE_SLUGS = {'art', 'prints', 'fine-art', 'abstract', 'botanical', 'digital-art', 'figurative',
                    'landscape', 'limited-edition-prints', 'pichwai', 'portrait', 'tribal', 'curated-by-prabha-saha'}

SHEET_IDS = {
    'com_pages': '1S4QuLgs341vqXXEVYgodqqQoEIeP_PC6f9P3qVZgwoc',
    'com_keywords': '17l99A1I1BAibLzkVb13FeBeoqaZcSAMSwLz5BmP8ZOY',
    'sg_pages': '1jGyaVQNd3cpwTdHX01f1FpiOwjV938Y7XuTTU3Z-bVc',
    'sg_keywords': '1wcHFWP9ec1OijVnkBDxwcu8fbQlH87faJhlIYSIH_sY',
    'com_pq': '1mq0JzaRl_st9Kx_zoklDzUYxp_ErQKyvuE5P1HK6CO0',
    'sg_pq': '1FiUZQX8u4d738FsvOzRLzXsNM4R61K19AEvPx-UcMuU',
}

# Scoped window for the per-page/per-page+query monthly drill-down
# (Chart 3 and the Per-Page Keyword Explorer). These sheets go back to
# 2023 with huge per-month tabs (10k-30k rows); pulling full history on
# every page load isn't practical, so this mirrors the same "last 6
# months" window Fazeel pulled for the Tab 2 CSVs. Extend this list (and,
# if needed, MONTH_TAB_OVERRIDES below) as new months are needed.
MONTH_WINDOW = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']
MONTH_YEAR_SUFFIX = "'26"

# One-off tab-naming exceptions (the .com property's April'26 tab is named
# this way because that's the month the GSC property migration landed).
MONTH_TAB_OVERRIDES = {
    ('com', 'Apr'): "Apr'26 (migrated tracking)",
}


def month_tab_name(store, month_abbr):
    return MONTH_TAB_OVERRIDES.get((store, month_abbr), f"{month_abbr}{MONTH_YEAR_SUFFIX}")


def get_page_type(url):
    if not isinstance(url, str) or not url:
        return 'other'
    path = url.split('/', 3)[-1] if '://' in url else url
    if path.startswith('products/'):
        return 'product'
    if path.startswith('blogs/'):
        return 'blog'
    if path.startswith('collections/'):
        slug = path.replace('collections/', '').split('/')[0].split('?')[0]
        if slug in ARTIST_SLUGS:
            return 'artist'
        if slug in DESIGNER_SLUGS:
            return 'designer'
        if slug in ART_GENRE_SLUGS:
            return 'art-collection'
        return 'fashion-collection'
    return 'other'


def bucket_for_position(pos):
    try:
        pos = float(pos)
    except (TypeError, ValueError):
        return None
    if pos <= 5:
        return '1-5'
    if pos <= 10:
        return '6-10'
    if pos <= 20:
        return '11-20'
    if pos <= 30:
        return '21-30'
    if pos <= 60:
        return '31-60'
    if pos <= 90:
        return '61-90'
    return '90+'


def clean_number(series):
    """Numbers in these sheets come as display strings: '23,762', '1.9%'."""
    return pd.to_numeric(
        series.astype(str).str.replace(',', '', regex=False).str.replace('%', '', regex=False).str.strip(),
        errors='coerce'
    )


# ── Data loading ────────────────────────────────────────────────────────

def gviz_csv_url(sheet_id, sheet_name):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}"


@st.cache_data(ttl=300)
def load_sheet_tab(sheet_id, sheet_name):
    url = gviz_csv_url(sheet_id, sheet_name)
    df = pd.read_csv(url)
    # gviz sometimes prefixes helper columns with blank headers (pandas
    # names them "Unnamed: N") ahead of the real columns - harmless, just
    # ignored by callers that select columns by name.
    return df


@st.cache_data(ttl=300)
def load_consolidated_final(sheet_id):
    """The ready-made monthly bucket-count table each Pages workbook
    already has. Returns columns: month, <buckets...>, total. 'total' is
    computed as the sum of the bucket columns rather than read from a
    named "Total ..." column - that header's exact text isn't consistent
    across these sheets (confirmed: the .sg Keywords workbook's own total
    column is literally labelled "Total Pages"), while bucket-sum always
    matches it exactly since the sheet's own total is itself a sum."""
    df = load_sheet_tab(sheet_id, 'Consolidated Data Final')
    out = pd.DataFrame()
    out['month'] = df['Month (TEXT)']
    for b in BUCKETS:
        out[b] = clean_number(df[b])
    out['total'] = out[BUCKETS].sum(axis=1)
    # Keep chronological order as given by the sheet, oldest first, and
    # drop rows with no data at all (future months not yet reached).
    out = out.iloc[::-1].reset_index(drop=True)
    out = out[out['total'] > 0].reset_index(drop=True)
    return out


@st.cache_data(ttl=300)
def load_keyword_buckets(sheet_id, store):
    """Keyword-by-position bucket counts, computed from the scoped monthly
    tabs (one row per ranking keyword) rather than the workbook's own
    Consolidated Data Final tab. That roll-up is unreliable for Keywords
    specifically (the .com Keywords workbook's version is zeroed out for
    every month), so this counts keywords per position bucket directly -
    the same approach already used for the page-type-filtered Pages
    chart. Returns columns: month, <buckets...>, total."""
    raw = load_monthly_raw(sheet_id, store, ['Query', 'Position Bucket'])
    if raw.empty:
        return pd.DataFrame(columns=['month'] + BUCKETS + ['total'])
    counts = raw.groupby(['month', 'Position Bucket']).size().unstack(fill_value=0)
    for b in BUCKETS:
        if b not in counts.columns:
            counts[b] = 0
    counts = counts[BUCKETS]
    counts['total'] = counts.sum(axis=1)
    counts = counts.reset_index()
    month_order = [f"{m}{MONTH_YEAR_SUFFIX}" for m in MONTH_WINDOW]
    counts['month'] = pd.Categorical(counts['month'], categories=month_order, ordered=True)
    counts = counts.sort_values('month').reset_index(drop=True)
    counts = counts[counts['total'] > 0].reset_index(drop=True)
    return counts


@st.cache_data(ttl=300)
def load_monthly_raw(sheet_id, store, columns):
    """Loads and concatenates the scoped window of monthly tabs for a Pages
    or Page&Query workbook. `columns` is the expected column list to
    normalize to (handles the odd month or two missing a Position Bucket
    column by recomputing it)."""
    frames = []
    for month_abbr in MONTH_WINDOW:
        tab = month_tab_name(store, month_abbr)
        try:
            df = load_sheet_tab(sheet_id, tab)
        except Exception:
            continue
        if df.empty:
            continue
        df = df.rename(columns={'Bucket': 'Position Bucket'})
        if 'Position' in df.columns:
            df['Position'] = clean_number(df['Position'])
        if 'Position Bucket' not in df.columns and 'Position' in df.columns:
            df['Position Bucket'] = df['Position'].apply(bucket_for_position)
        if 'Clicks' in df.columns:
            df['Clicks'] = clean_number(df['Clicks'])
        if 'Impressions' in df.columns:
            df['Impressions'] = clean_number(df['Impressions'])
        df['month'] = f"{month_abbr}{MONTH_YEAR_SUFFIX}"
        keep = [c for c in columns if c in df.columns] + ['month']
        frames.append(df[keep])
    if not frames:
        return pd.DataFrame(columns=columns + ['month'])
    return pd.concat(frames, ignore_index=True)


@st.cache_data(ttl=300)
def load_daily_performance(store):
    all_df = pd.read_csv(f"data/{store}_l6m_all.csv", parse_dates=['Date'])
    all_df['CTR'] = all_df['CTR'].astype(str).str.replace('%', '', regex=False).astype(float)
    ai_df = pd.read_csv(f"data/{store}_l6m_ai.csv", parse_dates=['Date'])
    ai_df = ai_df.rename(columns={'Impressions': 'AI_Impressions'})
    merged = all_df.merge(ai_df, on='Date', how='left')
    return merged


# ── Sidebar ─────────────────────────────────────────────────────────────

st.markdown("""
<style>
.block-container {padding-top: 1.5rem;}
div[data-testid="stMetricValue"] {color: #1d4ed8; font-size: 22px; font-weight: 700;}
div[data-testid="stMetricLabel"] {color: #94a3b8; font-size: 11px;}
div.stRadio > div {flex-direction: row; gap: 6px;}
div.stRadio label {
    background: white; border: 1px solid #e2e8f0; border-radius: 999px;
    padding: 4px 14px; margin-right: 4px; cursor: pointer;
}
</style>
""", unsafe_allow_html=True)

st.sidebar.title("Canvas & Weaves SEO")
store_label = st.sidebar.radio("Store", ["canvasandweaves.com", "canvasandweaves.sg"])
store = 'com' if store_label == "canvasandweaves.com" else 'sg'

if st.sidebar.button("↻ Refresh data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(f"Last fetched: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")

if store == 'com':
    st.warning(
        "**Migration note:** Nov '25-Dec '25 data is from "
        "[www.canvasandweaves.com](https://www.canvasandweaves.com) (old property). "
        "Jan '26 wasn't pulled and Feb '26 is a migration gap. "
        "Mar '26 onward is canvasandweaves.com (new property, migrated 29 Mar)."
    )

pages_sheet_id = SHEET_IDS[f'{store}_pages']
keywords_sheet_id = SHEET_IDS[f'{store}_keywords']
pq_sheet_id = SHEET_IDS[f'{store}_pq']

tab1, tab2 = st.tabs(["Rank Distribution", "Site Performance"])

# ── Tab 1: Rank Distribution ─────────────────────────────────────────────

with tab1:
    pages_final = load_consolidated_final(pages_sheet_id)
    keywords_final = load_keyword_buckets(keywords_sheet_id, store)

    latest_pages = pages_final.iloc[-1] if not pages_final.empty else None
    latest_keywords = keywords_final.iloc[-1] if not keywords_final.empty else None

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pages in Google", f"{int(latest_pages['total']):,}" if latest_pages is not None else "—")
    c2.metric("Unique Keywords", f"{int(latest_keywords['total']):,}" if latest_keywords is not None else "—")
    top10_pages = int(latest_pages['1-5'] + latest_pages['6-10']) if latest_pages is not None else 0
    top10_kw = int(latest_keywords['1-5'] + latest_keywords['6-10']) if latest_keywords is not None else 0
    c3.metric("Pages in Top 10", f"{top10_pages:,}")
    c4.metric("Keywords in Top 10", f"{top10_kw:,}")

    page_type_filter_label = st.radio(
        "Page type", list(PAGE_TYPE_FILTERS.keys()), horizontal=True, label_visibility="collapsed"
    )
    page_type_filter = PAGE_TYPE_FILTERS[page_type_filter_label]

    def stacked_bucket_chart(df, title, total_label):
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        pct = df[BUCKETS].div(df['total'], axis=0) * 100
        for b in BUCKETS:
            fig.add_trace(go.Bar(
                x=df['month'], y=pct[b], name=b,
                marker_color=BUCKET_COLOR_MAP[b],
                hovertemplate=f"{b}: %{{y:.1f}}%<extra></extra>"
            ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=df['month'], y=df['total'], name=total_label, mode='lines+markers',
            line=dict(color='#0f172a', width=2), marker=dict(size=6),
            hovertemplate=f"{total_label}: %{{y:,.0f}}<extra></extra>"
        ), secondary_y=True)
        fig.update_layout(
            barmode='stack', title=title, template='plotly_white', height=360,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0),
            margin=dict(t=70, b=10, l=10, r=10)
        )
        fig.update_yaxes(title_text='% of total', range=[0, 100], secondary_y=False)
        fig.update_yaxes(title_text=total_label, secondary_y=True, showgrid=False)
        return fig

    # Chart 1: Pages in Google by Position. The page-type filter needs
    # per-page rows (only available in the scoped recent-month window),
    # so when a filter other than "All" is active we rebuild the monthly
    # bucket counts from that raw data instead of the full-history
    # Consolidated Data Final table.
    if page_type_filter is None:
        chart1_df = pages_final
    else:
        raw_pages = load_monthly_raw(
            pages_sheet_id, store, ['Page', 'Impressions', 'Position Bucket']
        )
        raw_pages['page_type'] = raw_pages['Page'].apply(get_page_type)
        filtered = raw_pages[raw_pages['page_type'] == page_type_filter]
        counts = filtered.groupby(['month', 'Position Bucket']).size().unstack(fill_value=0)
        for b in BUCKETS:
            if b not in counts.columns:
                counts[b] = 0
        counts = counts[BUCKETS]
        counts['total'] = counts.sum(axis=1)
        counts = counts.reset_index().rename(columns={'month': 'month'})
        month_order = [f"{m}{MONTH_YEAR_SUFFIX}" for m in MONTH_WINDOW]
        counts['month'] = pd.Categorical(counts['month'], categories=month_order, ordered=True)
        chart1_df = counts.sort_values('month').reset_index(drop=True)

    left, right = st.columns([3, 1])
    with left:
        st.plotly_chart(stacked_bucket_chart(chart1_df, "Pages in Google by Position", "Total Pages"),
                         use_container_width=True, config={"displayModeBar": False})
    with right:
        st.caption(f"Latest month ({chart1_df.iloc[-1]['month']})" if not chart1_df.empty else "")
        if not chart1_df.empty:
            last = chart1_df.iloc[-1]
            for b in BUCKETS:
                pct = (last[b] / last['total'] * 100) if last['total'] else 0
                st.markdown(
                    f"<span style='color:{BUCKET_COLOR_MAP[b]}'>&#9632;</span> "
                    f"**{b}**: {int(last[b]):,} ({pct:.1f}%)",
                    unsafe_allow_html=True
                )

    # Chart 2: Keywords by Position (no page-type filter applies here).
    # Built from the scoped monthly window (see load_keyword_buckets) -
    # the workbook's own Consolidated Data Final roll-up isn't reliable
    # for Keywords, so this chart's month range is Apr-Sep'26 regardless
    # of how far back the Pages chart above goes.
    st.plotly_chart(stacked_bucket_chart(keywords_final, "Keywords by Position", "Total Keywords"),
                     use_container_width=True, config={"displayModeBar": False})
    st.caption(f"Scoped to {MONTH_WINDOW[0]}–{MONTH_WINDOW[-1]}{MONTH_YEAR_SUFFIX} "
               "(the window with per-keyword data available).")

    # Chart 3: Pages Indexed by Type (only meaningful with no filter active)
    if page_type_filter is None:
        raw_pages_all = load_monthly_raw(pages_sheet_id, store, ['Page', 'Impressions'])
        raw_pages_all['page_type'] = raw_pages_all['Page'].apply(get_page_type)
        type_counts = raw_pages_all.groupby(['month', 'page_type']).size().unstack(fill_value=0)
        month_order = [f"{m}{MONTH_YEAR_SUFFIX}" for m in MONTH_WINDOW]
        type_counts = type_counts.reindex(month_order).fillna(0)
        fig3 = go.Figure()
        for t, color in PAGE_TYPE_COLORS.items():
            if t in type_counts.columns:
                fig3.add_trace(go.Scatter(
                    x=type_counts.index, y=type_counts[t], name=t, mode='lines+markers',
                    line=dict(color=color)
                ))
        fig3.update_layout(
            title="Pages Indexed by Type", template='plotly_white', height=320,
            margin=dict(t=40, b=10, l=10, r=10)
        )
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})
        st.caption(f"Scoped to {MONTH_WINDOW[0]}–{MONTH_WINDOW[-1]}{MONTH_YEAR_SUFFIX} "
                   "(the window with per-page data available).")

    st.divider()
    st.subheader("Per-Page Keyword Explorer")
    st.caption("Select a page or paste a URL to see position distribution and keyword movement across months.")

    try:
        pq_raw = load_monthly_raw(pq_sheet_id, store, ['Page', 'Query', 'Clicks', 'Impressions', 'Position'])
        pq_load_error = None
    except Exception as e:
        pq_raw = pd.DataFrame()
        pq_load_error = str(e)

    if pq_load_error:
        st.error(
            "Couldn't load the Page+Query sheet for this store — it likely isn't shared "
            "\"Anyone with the link\" yet. Share it from Google Sheets, then refresh. "
            f"(Details: {pq_load_error})"
        )
    elif pq_raw.empty:
        st.info("No Page+Query data available for this store's scoped window yet.")
    else:
        top_pages = (pq_raw.groupby('Page')['Impressions'].sum()
                     .sort_values(ascending=False).head(20).index.tolist())

        if 'explorer_url' not in st.session_state:
            st.session_state.explorer_url = top_pages[0] if top_pages else ""

        chip_cols = st.columns(4)
        for i, p in enumerate(top_pages):
            label = p.split('/')[-1][:30] or p
            if chip_cols[i % 4].button(label, key=f"chip_{i}"):
                st.session_state.explorer_url = p

        url_input = st.text_input("Page URL", value=st.session_state.explorer_url, key="explorer_input")
        if st.button("Analyse Page"):
            st.session_state.explorer_url = url_input

        selected_url = st.session_state.explorer_url
        page_rows = pq_raw[pq_raw['Page'] == selected_url]

        if page_rows.empty:
            st.info("No data for that URL in the scoped window.")
        else:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Clicks", f"{int(page_rows['Clicks'].sum()):,}")
            m2.metric("Total Impressions", f"{int(page_rows['Impressions'].sum()):,}")
            m3.metric("Avg CTR", f"{(page_rows['Clicks'].sum() / page_rows['Impressions'].sum() * 100):.2f}%"
                      if page_rows['Impressions'].sum() else "—")
            m4.metric("Avg Position", f"{page_rows['Position'].mean():.1f}")

            page_rows = page_rows.copy()
            page_rows['bucket'] = page_rows['Position'].apply(bucket_for_position)
            dist = page_rows.groupby(['month', 'bucket']).size().unstack(fill_value=0)
            for b in BUCKETS:
                if b not in dist.columns:
                    dist[b] = 0
            dist = dist[BUCKETS]
            month_order = [f"{m}{MONTH_YEAR_SUFFIX}" for m in MONTH_WINDOW]
            dist = dist.reindex(month_order).fillna(0)
            dist['total'] = dist.sum(axis=1)
            dist = dist.reset_index().rename(columns={'index': 'month'})
            st.plotly_chart(stacked_bucket_chart(dist, "Position Distribution — by Month", "Keywords Tracked"),
                             use_container_width=True, config={"displayModeBar": False})

            top_queries = (page_rows.groupby('Query')['Impressions'].sum()
                           .sort_values(ascending=False).head(30).index.tolist())
            pivot = page_rows[page_rows['Query'].isin(top_queries)].pivot_table(
                index='Query', columns='month', values='Position', aggfunc='mean'
            )
            pivot = pivot.reindex(columns=month_order)
            pivot = pivot.reindex(top_queries)

            def cell_style(v):
                if pd.isna(v):
                    return 'background-color: white'
                if v <= 5:
                    return 'background-color: #dbeafe'
                if v <= 10:
                    return 'background-color: #dcfce7'
                if v <= 20:
                    return 'background-color: #fef9c3'
                return 'background-color: white'

            st.markdown("**Top Keywords · Position by Month**")
            styled = pivot.style.map(cell_style).format("{:.1f}", na_rep="—")
            st.dataframe(styled, use_container_width=True)

# ── Tab 2: Site Performance ───────────────────────────────────────────────

with tab2:
    daily = load_daily_performance(store)
    min_d, max_d = daily['Date'].min().date(), daily['Date'].max().date()
    date_range = st.slider("Date range", min_value=min_d, max_value=max_d, value=(min_d, max_d))
    ranged = daily[(daily['Date'].dt.date >= date_range[0]) & (daily['Date'].dt.date <= date_range[1])]

    total_clicks = ranged['Clicks'].sum()
    total_impr = ranged['Impressions'].sum()
    avg_ctr = (total_clicks / total_impr * 100) if total_impr else 0
    avg_pos = ranged['Position'].mean()
    total_ai_impr = ranged['AI_Impressions'].sum()
    ai_share = (total_ai_impr / total_impr * 100) if total_impr else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Clicks", f"{total_clicks:,.0f}")
    k2.metric("Total Impressions", f"{total_impr:,.0f}")
    k3.metric("Avg CTR", f"{avg_ctr:.2f}%")
    k4.metric("Avg Position", f"{avg_pos:.1f}")
    k5.metric("AI Overview Impressions", f"{total_ai_impr:,.0f}", f"{ai_share:.1f}% of impressions")

    monthly = ranged.set_index('Date').resample('MS').agg(
        Clicks=('Clicks', 'sum'), Impressions=('Impressions', 'sum'),
        AI_Impressions=('AI_Impressions', 'sum')
    )
    monthly['CTR'] = monthly['Clicks'] / monthly['Impressions'] * 100
    pos_weighted = ranged.set_index('Date').resample('MS').apply(
        lambda d: (d['Position'] * d['Impressions']).sum() / d['Impressions'].sum() if d['Impressions'].sum() else None
    )
    monthly['Position'] = pos_weighted
    monthly = monthly.reset_index()

    row1 = st.columns(2)
    row2 = st.columns(2)

    with row1[0]:
        st.caption("Total clicks from search results")
        fig = go.Figure(go.Bar(x=monthly['Date'], y=monthly['Clicks'], marker_color='#1d4ed8'))
        fig.update_layout(title="Monthly Clicks", template='plotly_white', height=220,
                           margin=dict(t=30, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with row1[1]:
        st.caption("Total search result appearances")
        fig = go.Figure(go.Bar(x=monthly['Date'], y=monthly['Impressions'], marker_color='#2F5496'))
        fig.update_layout(title="Monthly Impressions", template='plotly_white', height=220,
                           margin=dict(t=30, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with row2[0]:
        st.caption("Clicks ÷ impressions")
        fig = go.Figure(go.Scatter(x=monthly['Date'], y=monthly['CTR'], mode='lines+markers',
                                    line=dict(color='#0ea5e9')))
        fig.update_layout(title="Click-Through Rate", template='plotly_white', height=220,
                           yaxis=dict(ticksuffix='%'), margin=dict(t=30, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with row2[1]:
        st.caption("Impression-weighted · axis reversed")
        fig = go.Figure(go.Scatter(x=monthly['Date'], y=monthly['Position'], mode='lines+markers',
                                    line=dict(color='#d946ef'), connectgaps=False))
        fig.update_layout(title="Avg. Position", template='plotly_white', height=220,
                           yaxis=dict(autorange='reversed'), margin=dict(t=30, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.divider()
    st.caption("AI Overview impressions — daily search appearances where Google showed an AI Overview "
               "(tracked from 18 May 2026 onward; no data before that date).")
    fig_ai = go.Figure(go.Scatter(x=ranged['Date'], y=ranged['AI_Impressions'], mode='lines',
                                   line=dict(color='#f97316'), connectgaps=False))
    fig_ai.update_layout(title="AI Overview Impressions (daily)", template='plotly_white', height=240,
                          margin=dict(t=30, b=10, l=10, r=10))
    st.plotly_chart(fig_ai, use_container_width=True, config={"displayModeBar": False})
