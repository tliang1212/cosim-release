import os
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st


APP_TITLE = "Cosim Release Dashboard"
DEFAULT_EXCEL_PATH = os.getenv("COSIM_EXCEL_PATH", "data/cosim_release_plan.xlsx").strip()

FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    "cadence": ("cadence", "release cadence", "cycle", "build cadence"),
    "program": ("program", "vehicle program", "project", "platform"),
    "domain": ("domain", "area", "workstream", "system domain"),
    "release": ("release", "release name", "milestone", "drop", "version"),
    "release_date": (
        "release date",
        "target date",
        "planned date",
        "timing",
        "eta",
        "availability date",
    ),
    "content": ("content", "release content", "scope", "deliverables", "features"),
    "issues": ("issues", "known issues", "risks", "blockers", "open issues"),
    "owner": ("owner", "lead", "dri", "responsible"),
    "status": ("status", "state", "health"),
    "notes": ("notes", "comments", "remarks"),
}

FILTER_FIELDS = ("cadence", "program", "domain")
DETAIL_FIELDS = (
    "cadence",
    "program",
    "domain",
    "release",
    "release_date",
    "status",
    "owner",
    "content",
    "issues",
    "notes",
)


def normalize_label(value: object) -> str:
    return " ".join(str(value).replace("_", " ").replace("-", " ").lower().split())


def is_blank(value: object) -> bool:
    return pd.isna(value) or str(value).strip() == ""


def display_value(value: object, default: str = "Not specified") -> str:
    if is_blank(value):
        return default
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def resolve_path(path_value: str) -> Path:
    if path_value.startswith("dbfs:/"):
        return Path("/dbfs") / path_value.removeprefix("dbfs:/")
    return Path(path_value).expanduser()


def infer_column(columns: Iterable[str], field_name: str) -> Optional[str]:
    normalized_columns = {normalize_label(column): column for column in columns}
    for alias in FIELD_ALIASES[field_name]:
        match = normalized_columns.get(normalize_label(alias))
        if match:
            return match

    for column in columns:
        normalized_column = normalize_label(column)
        if any(normalize_label(alias) in normalized_column for alias in FIELD_ALIASES[field_name]):
            return column
    return None


@st.cache_data(show_spinner=False)
def workbook_sheets_from_bytes(file_bytes: bytes) -> List[str]:
    return pd.ExcelFile(BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def read_sheet_from_bytes(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)


@st.cache_data(show_spinner=False)
def workbook_sheets_from_path(path: str, modified_time: float) -> List[str]:
    del modified_time
    return pd.ExcelFile(path).sheet_names


@st.cache_data(show_spinner=False)
def read_sheet_from_path(path: str, modified_time: float, sheet_name: str) -> pd.DataFrame:
    del modified_time
    return pd.read_excel(path, sheet_name=sheet_name)


def load_workbook() -> Tuple[Optional[str], Optional[bytes], List[str]]:
    uploaded_file = st.sidebar.file_uploader(
        "Upload release plan Excel",
        type=("xlsx", "xlsm", "xls"),
        help="Use this for local development or quick validation before publishing the file to Databricks.",
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        return uploaded_file.name, file_bytes, workbook_sheets_from_bytes(file_bytes)

    if not DEFAULT_EXCEL_PATH:
        return None, None, []

    workbook_path = resolve_path(DEFAULT_EXCEL_PATH)
    if not workbook_path.exists():
        return str(workbook_path), None, []

    modified_time = workbook_path.stat().st_mtime
    return (
        str(workbook_path),
        None,
        workbook_sheets_from_path(str(workbook_path), modified_time),
    )


def read_selected_sheet(source_name: str, file_bytes: Optional[bytes], sheet_name: str) -> pd.DataFrame:
    if file_bytes is not None:
        return read_sheet_from_bytes(file_bytes, sheet_name)

    workbook_path = resolve_path(source_name)
    return read_sheet_from_path(str(workbook_path), workbook_path.stat().st_mtime, sheet_name)


def build_column_mapping(columns: List[str]) -> Dict[str, Optional[str]]:
    options = ["<missing>"] + columns
    mapping: Dict[str, Optional[str]] = {}

    with st.sidebar.expander("Column mapping", expanded=False):
        st.caption("Adjust these if your Excel headers use different names.")
        for field_name in DETAIL_FIELDS:
            inferred = infer_column(columns, field_name)
            default_index = options.index(inferred) if inferred in options else 0
            selected = st.selectbox(
                field_name.replace("_", " ").title(),
                options,
                index=default_index,
                key=f"mapping_{field_name}",
            )
            mapping[field_name] = None if selected == "<missing>" else selected

    return mapping


def normalize_release_plan(raw_df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    normalized = pd.DataFrame(index=raw_df.index)

    for field_name in DETAIL_FIELDS:
        source_column = mapping.get(field_name)
        if source_column:
            normalized[field_name] = raw_df[source_column]
        else:
            normalized[field_name] = pd.NA

    for field_name in FILTER_FIELDS:
        normalized[field_name] = normalized[field_name].apply(display_value)

    normalized["release_date"] = pd.to_datetime(normalized["release_date"], errors="coerce")
    normalized["search_text"] = normalized.apply(
        lambda row: " ".join(display_value(row[field], "") for field in DETAIL_FIELDS).lower(),
        axis=1,
    )
    return normalized


def sorted_filter_values(series: pd.Series) -> List[str]:
    values = [display_value(value) for value in series.dropna().unique()]
    return sorted(values, key=str.lower)


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    filtered = df.copy()

    with st.sidebar:
        st.subheader("Filters")
        for field_name in FILTER_FIELDS:
            values = sorted_filter_values(filtered[field_name])
            selected_values = st.multiselect(
                field_name.title(),
                values,
                default=values,
                key=f"filter_{field_name}",
            )
            filtered = filtered[filtered[field_name].isin(selected_values)]

        show_upcoming_only = st.checkbox("Upcoming releases only", value=False)
        if show_upcoming_only:
            today = pd.Timestamp(date.today())
            filtered = filtered[filtered["release_date"].isna() | (filtered["release_date"] >= today)]

        search_term = st.text_input("Search content or issues", placeholder="e.g. blocker, ADAS, next drop")
        if search_term:
            filtered = filtered[
                filtered["search_text"].str.contains(search_term.lower(), na=False, regex=False)
            ]

    return filtered


def render_metrics(df: pd.DataFrame) -> None:
    issue_count = df["issues"].apply(lambda value: not is_blank(value)).sum()
    upcoming_dates = df.loc[df["release_date"].notna() & (df["release_date"] >= pd.Timestamp(date.today()))]
    next_release = (
        upcoming_dates["release_date"].min().strftime("%Y-%m-%d")
        if not upcoming_dates.empty
        else "Not scheduled"
    )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Release rows", f"{len(df):,}")
    col2.metric("Domains", f"{df['domain'].nunique():,}")
    col3.metric("Rows with issues", f"{issue_count:,}")
    col4.metric("Next release", next_release)


def render_timing_chart(df: pd.DataFrame) -> None:
    st.subheader("Release timing")
    chart_df = df[df["release_date"].notna()].copy()

    if chart_df.empty:
        count_df = df.groupby(["domain", "cadence"], dropna=False).size().reset_index(name="rows")
        fig = px.bar(
            count_df,
            x="domain",
            y="rows",
            color="cadence",
            title="Release rows by domain and cadence",
        )
    else:
        chart_df["release_label"] = chart_df["release"].apply(lambda value: display_value(value, "Release"))
        fig = px.scatter(
            chart_df,
            x="release_date",
            y="domain",
            color="cadence",
            symbol="program",
            hover_name="release_label",
            hover_data={
                "program": True,
                "status": True,
                "owner": True,
                "release_date": "|%Y-%m-%d",
                "content": True,
                "issues": True,
            },
            title="Planned release dates by domain",
        )
        fig.add_vline(
            x=pd.Timestamp(date.today()),
            line_dash="dash",
            line_color="gray",
            annotation_text="Today",
        )

    fig.update_layout(height=430, margin=dict(l=10, r=10, t=60, b=10))
    st.plotly_chart(fig, use_container_width=True)


def render_release_table(df: pd.DataFrame) -> None:
    display_df = df[[field for field in DETAIL_FIELDS if field in df.columns]].copy()
    display_df["release_date"] = display_df["release_date"].dt.strftime("%Y-%m-%d").fillna("")
    display_df = display_df.applymap(lambda value: "" if is_blank(value) else value)

    st.subheader("Release plan")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv_bytes = display_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered plan as CSV",
        data=csv_bytes,
        file_name="filtered_cosim_release_plan.csv",
        mime="text/csv",
    )


def render_domain_details(df: pd.DataFrame) -> None:
    st.subheader("Domain details")

    for domain_name in sorted_filter_values(df["domain"]):
        domain_df = df[df["domain"] == domain_name].sort_values(
            by=["release_date", "program", "cadence"],
            na_position="last",
        )
        with st.expander(f"{domain_name} ({len(domain_df)} rows)", expanded=False):
            for _, row in domain_df.iterrows():
                release_date = display_value(row["release_date"], "Date TBD")
                release_name = display_value(row["release"], "Release")
                st.markdown(f"**{release_name}** - {release_date}")
                st.write(
                    f"Program: {display_value(row['program'])} | "
                    f"Cadence: {display_value(row['cadence'])} | "
                    f"Status: {display_value(row['status'])}"
                )
                if not is_blank(row["content"]):
                    st.markdown(f"Content: {display_value(row['content'], '')}")
                if not is_blank(row["issues"]):
                    st.warning(f"Issues: {display_value(row['issues'], '')}")
                if not is_blank(row["notes"]):
                    st.caption(f"Notes: {display_value(row['notes'], '')}")
                st.divider()


def render_issue_view(df: pd.DataFrame) -> None:
    issue_df = df[df["issues"].apply(lambda value: not is_blank(value))].copy()
    st.subheader("Known issues")

    if issue_df.empty:
        st.success("No issue text is present for the current filters.")
        return

    issue_df["release_date"] = issue_df["release_date"].dt.strftime("%Y-%m-%d").fillna("")
    st.dataframe(
        issue_df[["domain", "program", "cadence", "release", "release_date", "owner", "status", "issues"]],
        use_container_width=True,
        hide_index=True,
    )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Filter cosim release timing, content, and known issues by cadence, program, and domain.")

    with st.sidebar:
        st.header("Data source")
        st.caption(
            "Set COSIM_EXCEL_PATH for Databricks deployment, or upload an Excel file for local testing."
        )

    try:
        source_name, file_bytes, sheets = load_workbook()
    except Exception as error:
        st.error(f"Could not open the Excel workbook: {error}")
        st.stop()

    if not sheets:
        st.info(
            "Upload the current release-plan Excel file, or set COSIM_EXCEL_PATH to a file that the "
            "Databricks app can read."
        )
        st.code("streamlit run app.py", language="bash")
        st.stop()

    st.sidebar.caption(f"Workbook: {source_name}")
    selected_sheet = st.sidebar.selectbox("Worksheet", sheets)

    try:
        raw_df = read_selected_sheet(str(source_name), file_bytes, selected_sheet)
    except Exception as error:
        st.error(f"Could not read worksheet '{selected_sheet}': {error}")
        st.stop()

    if raw_df.empty:
        st.warning("The selected worksheet has no rows.")
        st.stop()

    mapping = build_column_mapping(list(raw_df.columns))
    df = normalize_release_plan(raw_df, mapping)
    filtered_df = apply_filters(df)

    if filtered_df.empty:
        st.warning("No release rows match the current filters.")
        st.stop()

    render_metrics(filtered_df)
    render_timing_chart(filtered_df)

    table_tab, domain_tab, issue_tab = st.tabs(("Plan table", "Domain details", "Known issues"))
    with table_tab:
        render_release_table(filtered_df)
    with domain_tab:
        render_domain_details(filtered_df)
    with issue_tab:
        render_issue_view(filtered_df)


if __name__ == "__main__":
    main()
