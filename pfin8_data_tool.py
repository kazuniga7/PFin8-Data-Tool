"""
P-Fin 8 Data Exploration Tool
Built with Streamlit

To run: streamlit run pfin8_data_tool.py
Required: streamlit, pandas, numpy, plotly, openpyxl
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ==============================================================================
# CONFIGURATION
# ==============================================================================
DEBUG_MODE = True
DATA_DIR = Path(".")  # Change to the folder containing the xlsx files

# ==============================================================================
# CONSTANTS
# ==============================================================================
TOPIC_NAMES = {
    "Earnings": "pfin8_earningsQCorrect",
    "Consuming": "pfin8_consumingQCorrect",
    "Savings": "pfin8_savingsQCorrect",
    "Investing": "pfin8_investingQCorrect",
    "Borrowing/Managing Debt": "pfin8_borrowingQCorrect",
    "Insuring": "pfin8_insuringQCorrect",
    "Comprehending Uncertainty": "pfin8_compreUncertQCorrect",
    "Go-to Information Sources": "pfin8_infoSourcesQCorrect",
}

TOPIC_CAT3_NAMES = {
    "Earnings": "pfin8_earnings_cat3",
    "Consuming": "pfin8_consuming_cat3",
    "Savings": "pfin8_savings_cat3",
    "Investing": "pfin8_investing_cat3",
    "Borrowing/Managing Debt": "pfin8_borrowing_cat3",
    "Insuring": "pfin8_insuring_cat3",
    "Comprehending Uncertainty": "pfin8_compreUncert_cat3",
    "Go-to Information Sources": "pfin8_infoSources_cat3",
}

DEMOGRAPHIC_VARIABLES = {
    "Age (Buckets)": "age_category",
    "Age (Custom Range)": "reported_age",
    "Generation": "generation_category",
    "Gender": "gender",
    "Race/Ethnicity": "race_ethnicity_category",
    "Marital Status": "marital_status",
    "Dependent Children Under 18": "has_dependent_children",
    "Education": "education_category",
    "Financial Education Course": "took_Financial_Education",
    "Employment Status": "employment_category",
    "Income": "income_category",
}

DEMOGRAPHIC_DISPLAY_LABELS = {
    "age_category": "Age",
    "reported_age": "Age",
    "generation_category": "Generation",
    "gender": "Gender",
    "race_ethnicity_category": "Race/Ethnicity",
    "marital_status": "Marital Status",
    "has_dependent_children": "Do you have dependent children under the age of 18?",
    "education_category": "Education",
    "took_Financial_Education": "Have you ever participated in a financial education class or program?",
    "employment_category": "Employment Status",
    "income_category": "Income",
}

FINANCIAL_WELLBEING_VARIABLES = {
    "Debt Constraint": "debt_constrained_responses",
    "Financial Fragility": "is_Financially_Fragile",
    "Non-Retirement Savings (One Month)": "suffretirement_savings_responses",
    "Time Thinking About Finances (Overall)": "time_thinking_finances",
    "Time Thinking About Finances (At Work)": "worktime_thinking_finances",
}

FINANCIAL_WELLBEING_LABELS = {
    "debt_constrained_responses": "Do debt and debt payments prevent you from adequately addressing other financial priorities?",
    "is_Financially_Fragile": "How confident are you that you could come up with $2,000 if an unexpected need arose within the next month?",
    "suffretirement_savings_responses": "Do you have non-retirement savings sufficient to cover one month of living expenses if needed?",
    "time_thinking_finances": "How much time (hours per week) do you typically spend thinking about and dealing with issues related to your personal finances?",
    "worktime_thinking_finances": "How many of these hours occur at work?",
}

TOTAL_CORRECT_LABELS = {
    0: "None Correct", 1: "1", 2: "2", 3: "3", 4: "4",
    5: "5", 6: "6", 7: "7", 8: "All Correct",
}

BINARY_DISPLAY = {0: "No", 1: "Yes"}

THINKING_TIME_ORDER = ["No hours", "1h", "2h", "3-4h", "5-9h", "10-19h", "20h+"]

AGE_BUCKET_ORDER = ["18-32", "33-47", "48-62", "63-79", "80+"]
GENERATION_ORDER = ["genZ", "genY", "genX", "boomer", "silent"]
EDUCATION_ORDER = ["Less than HS", "High School", "Some College", "Bachelor's degree or higher"]
INCOME_ORDER = ["<$25K", "$25-50K", "$50-100K", ">$100K"]

MIN_SAMPLE_SIZE = 30

# ==============================================================================
# DATA LOADING
# ==============================================================================
@st.cache_data
def load_all_years():
    df = pd.read_excel(DATA_DIR / "allYearsPFin8.xlsx")
    cat3_cols = [c for c in df.columns if "cat3" in c]
    for col in cat3_cols:
        df[col] = df[col].replace("Corrcet", "Correct")
    return df


@st.cache_data
def load_genpop():
    df = pd.read_excel(DATA_DIR / "PFin2026_GenPop.xlsx")
    cat3_cols = [c for c in df.columns if "cat3" in c]
    for col in cat3_cols:
        df[col] = df[col].replace("Corrcet", "Correct")
    # Fix "None" in thinking columns using openpyxl for raw read
    from openpyxl import load_workbook
    wb = load_workbook(DATA_DIR / "PFin2026_GenPop.xlsx")
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    for col_name in ["time_thinking_finances", "worktime_thinking_finances"]:
        col_idx = headers.index(col_name) + 1
        vals = []
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col_idx).value
            if val == "None":
                vals.append("No hours")
            elif val == "" or val is None:
                vals.append(np.nan)
            else:
                vals.append(val)
        df[col_name] = vals
    # Map binary display for has_dependent_children and took_Financial_Education
    df["has_dependent_children_display"] = df["has_dependent_children"].map(BINARY_DISPLAY)
    df["took_Financial_Education_display"] = df["took_Financial_Education"].map(BINARY_DISPLAY)
    return df


# ==============================================================================
# SANITY CHECK FUNCTIONS
# ==============================================================================
def check_percentages_valid(percentages, label=""):
    issues = []
    for p in percentages:
        if pd.notna(p) and (p < 0 or p > 100):
            issues.append(f"Percentage out of range [0,100]: {p:.2f} in {label}")
    return issues


def check_group_counts(df, group_col, total_n):
    group_n = df.groupby(group_col).size().sum()
    if group_n != total_n:
        return [f"Group counts ({group_n}) don't sum to total ({total_n}) for {group_col}"]
    return []


def check_sample_size(n, label=""):
    if n < MIN_SAMPLE_SIZE:
        return f"⚠️ Small sample size (n={n}) for {label}. Results may be unreliable."
    return None


def run_sanity_checks(df, weight_col, result_data, environment, analysis_variable=None):
    checks = {"passed": [], "warnings": [], "errors": []}

    # Check weights are positive
    if (df[weight_col] <= 0).any():
        checks["errors"].append("Some survey weights are zero or negative")
    else:
        checks["passed"].append("All survey weights are positive")

    # Check total rows
    n = len(df)
    checks["passed"].append(f"Total observations in filtered data: {n}")

    # Check percentages in results
    if "percentage" in result_data.columns:
        pct_issues = check_percentages_valid(result_data["percentage"].values, "results")
        if pct_issues:
            checks["errors"].extend(pct_issues)
        else:
            checks["passed"].append("All percentages are within [0, 100]")

    # Check sample sizes per group
    if analysis_variable and analysis_variable in df.columns:
        for group, group_df in df.groupby(analysis_variable):
            warning = check_sample_size(len(group_df), str(group))
            if warning:
                checks["warnings"].append(warning)

    # Check that weighted percentages sum reasonably
    if "percentage" in result_data.columns:
        group_check_col = None
        cat_check_col = None
        for col in ["group_value", "group"]:
            if col in result_data.columns:
                group_check_col = col
                break
        for col in ["response_category", "category"]:
            if col in result_data.columns:
                cat_check_col = col
                break
        if group_check_col and cat_check_col:
            for grp_name, grp_data in result_data.groupby(group_check_col):
                total = grp_data.groupby(cat_check_col)["percentage"].sum().sum()
                # Only check if there's a facet (topic) or not
                if "topic" in result_data.columns:
                    for topic, topic_data in grp_data.groupby("topic"):
                        t = topic_data["percentage"].sum()
                        if abs(t - 100) > 1:
                            checks["warnings"].append(
                                f"Percentages for {grp_name}/{topic} sum to {t:.1f}%, expected ~100%"
                            )

    return checks


# ==============================================================================
# WEIGHTED CALCULATION FUNCTIONS
# ==============================================================================
def weighted_mean(df, col, weight_col="survey_weight"):
    valid = df.dropna(subset=[col, weight_col])
    if len(valid) == 0:
        return np.nan
    return np.average(valid[col], weights=valid[weight_col])


def weighted_percentage_binary(df, col, weight_col="survey_weight"):
    return weighted_mean(df, col, weight_col) * 100


def weighted_percentage_cat3(df, col, weight_col="survey_weight"):
    valid = df.dropna(subset=[col, weight_col])
    if len(valid) == 0:
        return pd.DataFrame()
    total_weight = valid[weight_col].sum()
    result = valid.groupby(col)[weight_col].sum() / total_weight * 100
    return result.reset_index().rename(columns={col: "category", weight_col: "percentage"})


def weighted_total_correct_distribution(df, weight_col="survey_weight"):
    valid = df.dropna(subset=["pfin8_totalCorrect", weight_col])
    if len(valid) == 0:
        return pd.DataFrame()
    total_weight = valid[weight_col].sum()
    result = valid.groupby("pfin8_totalCorrect")[weight_col].sum() / total_weight * 100
    result = result.reset_index().rename(
        columns={"pfin8_totalCorrect": "score", weight_col: "percentage"}
    )
    result["score_label"] = result["score"].map(TOTAL_CORRECT_LABELS)
    return result


# ==============================================================================
# CHART TYPE VALIDATION
# ==============================================================================
def get_valid_chart_types(analysis_type, view_mode, environment, axis_legend=None):
    valid = []
    if analysis_type == "Topic Bucket":
        if view_mode == "3-Category (Correct / Incorrect / Don't Know)":
            valid = ["Grouped Bar Chart"]
            # Stacked is valid only when legend represents parts of a whole
            if axis_legend == "Response Category":
                valid.append("Stacked Bar Chart")
        else:
            valid = ["Grouped Bar Chart", "Bar Chart", "Line Chart"]
    else:  # Total Correct
        valid = ["Bar Chart", "Grouped Bar Chart", "Line Chart"]
        # Stacked is valid only when Total Correct scores are in the legend
        if axis_legend == "Total Correct":
            valid.append("Stacked Bar Chart")
    return valid


# ==============================================================================
# CHART GENERATION
# ==============================================================================
def create_chart(chart_data, chart_type, title, x_label, y_label, color_col=None,
                 category_orders=None, group_label="group", hover_mode="binary",
                 legend_label=None, facet_col=None):
    fig = None
    try:
        label_map = {
            "x": x_label,
            "percentage": y_label,
            "group_value": group_label,
            "topic": "Topic",
            "response_category": "Response Category",
            "score_label": "Total Correct",
        }
        facet_args = {"facet_col": facet_col, "facet_col_wrap": 4} if facet_col else {}

        if chart_type == "Bar Chart":
            fig = px.bar(
                chart_data, x="x", y="percentage",
                color=color_col, barmode="group",
                title=title, labels=label_map,
                category_orders=category_orders,
                **facet_args,
            )
        elif chart_type == "Grouped Bar Chart":
            fig = px.bar(
                chart_data, x="x", y="percentage",
                color=color_col, barmode="group",
                title=title, labels=label_map,
                category_orders=category_orders,
                **facet_args,
            )
        elif chart_type == "Stacked Bar Chart":
            fig = px.bar(
                chart_data, x="x", y="percentage",
                color=color_col, barmode="stack",
                title=title, labels=label_map,
                category_orders=category_orders,
                **facet_args,
            )
        elif chart_type == "Line Chart":
            fig = px.line(
                chart_data, x="x", y="percentage",
                color=color_col, markers=True,
                title=title, labels=label_map,
                category_orders=category_orders,
                **facet_args,
            )

        if fig:
            # Custom hover templates based on mode
            if hover_mode == "cat3":
                cat3_hover_labels = {
                    "Correct": "% Correct",
                    "Incorrect": "% Incorrect",
                    "Don't Know": "% Don't Know",
                }
                for trace in fig.data:
                    cat_name = trace.name
                    pct_label = cat3_hover_labels.get(cat_name, "% of Respondents")
                    trace.hovertemplate = (
                        f"{x_label}: %{{x}}<br>"
                        f"{pct_label}: %{{y:.1f}}%<br>"
                        f"<extra></extra>"
                    )
            elif hover_mode == "total_correct":
                for trace in fig.data:
                    trace.hovertemplate = (
                        f"{x_label}: %{{x}}<br>"
                        f"% of Respondents: %{{y:.1f}}%<br>"
                        f"{group_label}: {trace.name}<br>"
                        f"<extra></extra>"
                    )
            else:  # binary
                for trace in fig.data:
                    trace.hovertemplate = (
                        f"{x_label}: %{{x}}<br>"
                        f"% Correct: %{{y:.1f}}%<br>"
                        f"{group_label}: {trace.name}<br>"
                        f"<extra></extra>"
                    )

            fig.update_layout(
                yaxis_title=y_label,
                xaxis_title=x_label,
                legend_title_text=legend_label if legend_label else (group_label if color_col == "group" else (color_col if color_col else "")),
                template="plotly_white",
                font=dict(size=12),
                title_font=dict(size=16),
                height=800 if facet_col else 500,
            )
            fig.update_yaxes(range=[0, 105])

            # Clean up facet subplot titles (remove "topic=" prefix)
            if facet_col:
                fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
    except Exception as e:
        st.error(f"Could not create chart: {str(e)}")
        return None
    return fig


# ==============================================================================
# DESCRIPTIVE NOTE GENERATION
# ==============================================================================
def generate_note(environment, analysis_type, view_mode, selected_topics, selected_range,
                  analysis_variable, subgroups, n_obs, dataset_name, year_range=None):
    parts = []
    parts.append(f"**Data Source:** {dataset_name}")

    if environment == "Over the Years":
        yr_text = f"{year_range[0]}–{year_range[1]}" if year_range else "2017–2026"
        parts.append(f"**Years:** {yr_text}")
    else:
        parts.append("**Survey Year:** 2026")

    if analysis_type == "Topic Bucket":
        topics_text = ", ".join(selected_topics) if selected_topics else "All Topics"
        mode_text = "binary (correct vs. not correct)" if "Binary" in view_mode else "3-category (correct, incorrect, don't know)"
        parts.append(f"**Topics:** {topics_text}")
        parts.append(f"**View:** {mode_text}")
    else:
        range_text = f"{selected_range[0]}–{selected_range[1]}" if selected_range else "0–8"
        parts.append(f"**Total Correct Range:** {range_text}")

    if environment != "Over the Years" and analysis_variable:
        parts.append(f"**Analysis Variable:** {analysis_variable}")
        if subgroups:
            parts.append(f"**Subgroups:** {', '.join(str(s) for s in subgroups)}")

    parts.append(f"**Sample Size (filtered):** n={n_obs:,}")
    parts.append("All results are weighted to represent the population using survey weights. "
                  "Missing, refused, and inapplicable responses have been excluded.")
    return " | ".join(parts[:4]) + "\n\n" + " | ".join(parts[4:])


# ==============================================================================
# DATA PREPARATION FUNCTIONS
# ==============================================================================
def prepare_topic_binary_data(df, topics, group_col, group_label, weight_col="survey_weight"):
    rows = []
    for group_val, group_df in df.groupby(group_col):
        for topic_name, topic_col in topics.items():
            pct = weighted_percentage_binary(group_df, topic_col, weight_col)
            rows.append({
                "topic": topic_name,
                "group_value": str(group_val),
                "percentage": pct,
                "n": len(group_df.dropna(subset=[topic_col])),
            })
    return pd.DataFrame(rows)


def prepare_topic_cat3_data(df, topics, group_col, group_label, weight_col="survey_weight"):
    rows = []
    for group_val, group_df in df.groupby(group_col):
        for topic_name, cat3_col in topics.items():
            cat3_data = weighted_percentage_cat3(group_df, cat3_col, weight_col)
            for _, row in cat3_data.iterrows():
                rows.append({
                    "topic": topic_name,
                    "group_value": str(group_val),
                    "percentage": row["percentage"],
                    "response_category": row["category"],
                    "n": len(group_df.dropna(subset=[cat3_col])),
                })
    return pd.DataFrame(rows)


def prepare_total_correct_data(df, group_col, weight_col="survey_weight", score_range=None):
    rows = []
    for group_val, group_df in df.groupby(group_col):
        dist = weighted_total_correct_distribution(group_df, weight_col)
        if score_range:
            dist = dist[dist["score"].between(score_range[0], score_range[1])]
        for _, row in dist.iterrows():
            rows.append({
                "score_label": row["score_label"],
                "score": row["score"],
                "percentage": row["percentage"],
                "group_value": str(group_val),
                "n": len(group_df.dropna(subset=["pfin8_totalCorrect"])),
            })
    return pd.DataFrame(rows)


def get_category_order(col_name):
    orders = {
        "age_category": AGE_BUCKET_ORDER,
        "generation_category": GENERATION_ORDER,
        "education_category": EDUCATION_ORDER,
        "income_category": INCOME_ORDER,
        "time_thinking_finances": THINKING_TIME_ORDER,
        "worktime_thinking_finances": THINKING_TIME_ORDER,
    }
    return orders.get(col_name)


# ==============================================================================
# SIDEBAR FILTERS
# ==============================================================================
def render_sidebar(df_years, df_genpop):
    with st.sidebar:
        st.title("P-Fin 8 Data Tool")
        st.markdown("---")

        # Environment selection
        environment = st.radio(
            "Exploration Type",
            ["Over the Years", "Demographics", "Financial Well-Being"],
            help="Choose how you want to explore the P-Fin 8 data",
        )

        st.markdown("---")

        # Analysis type
        analysis_type = st.radio(
            "Analysis Type",
            ["Topic Bucket", "Total Correct"],
            help="Analyze by individual topic questions or total correct score",
        )

        st.markdown("---")

        # View mode (only for Topic Bucket)
        view_mode = None
        selected_topics = None
        selected_range = None

        if analysis_type == "Topic Bucket":
            view_mode = st.radio(
                "View Mode",
                ["Binary (Correct / Not Correct)", "3-Category (Correct / Incorrect / Don't Know)"],
            )
            all_topics = list(TOPIC_NAMES.keys())
            selected_topics = st.multiselect(
                "Select Topics",
                all_topics,
                default=all_topics,
                help="Choose which P-Fin 8 topics to include",
            )
            if not selected_topics:
                st.warning("Please select at least one topic.")
        else:
            selected_range = st.slider(
                "Total Correct Range",
                min_value=0, max_value=8, value=(0, 8),
                help="Filter the range of total correct scores",
            )

        st.markdown("---")

        # Environment-specific filters
        analysis_variable = None
        analysis_col = None
        subgroups = None
        year_range = None
        custom_age_range = None

        if environment == "Over the Years":
            years = sorted(df_years["survey_year"].unique())
            year_range = st.slider(
                "Year Range",
                min_value=int(min(years)), max_value=int(max(years)),
                value=(int(min(years)), int(max(years))),
            )

        elif environment == "Demographics":
            analysis_variable = st.selectbox(
                "Demographic Variable",
                list(DEMOGRAPHIC_VARIABLES.keys()),
            )
            analysis_col = DEMOGRAPHIC_VARIABLES[analysis_variable]

            if analysis_variable == "Age (Custom Range)":
                min_age = int(df_genpop["reported_age"].min())
                max_age = int(df_genpop["reported_age"].max())

                num_groups = st.selectbox(
                    "Number of age groups",
                    [1, 2, 3, 4, 5, 6, 7, 8],
                    index=2,
                )

                st.markdown("**Define your age groups**")
                st.caption(f"Min age: {min_age} · Max age: {max_age}")
                custom_age_groups = []
                age_errors = []

                for i in range(num_groups):
                    st.markdown(f"**Group {i+1}:**")
                    col1, col2 = st.columns(2)
                    with col1:
                        start = st.number_input(
                            "Start",
                            min_value=min_age, max_value=max_age,
                            value=min(min_age + i * ((max_age - min_age) // num_groups), max_age),
                            key=f"age_start_{i}",
                        )
                    with col2:
                        default_end = min(min_age + (i + 1) * ((max_age - min_age) // num_groups) - 1, max_age)
                        if i == num_groups - 1:
                            default_end = max_age
                        end = st.number_input(
                            "End",
                            min_value=min_age, max_value=max_age,
                            value=default_end,
                            key=f"age_end_{i}",
                        )

                    # Validate: end must be >= start
                    if end < start:
                        st.markdown(f'<p style="color: red; font-size: 0.85rem; margin: -10px 0 5px 0;">⚠️ Group {i+1}: end age must be ≥ start age</p>', unsafe_allow_html=True)
                        age_errors.append(f"Group {i+1}: end < start")

                    custom_age_groups.append((start, end))

                # Validate: check for overlaps
                for i in range(len(custom_age_groups)):
                    for j in range(i + 1, len(custom_age_groups)):
                        g1_start, g1_end = custom_age_groups[i]
                        g2_start, g2_end = custom_age_groups[j]
                        if g1_start <= g2_end and g2_start <= g1_end:
                            overlap_start = max(g1_start, g2_start)
                            overlap_end = min(g1_end, g2_end)
                            st.markdown(f'<p style="color: red; font-size: 0.85rem; margin: 0 0 5px 0;">⚠️ Groups {i+1} and {j+1} overlap (ages {overlap_start}–{overlap_end})</p>', unsafe_allow_html=True)
                            age_errors.append(f"Groups {i+1} and {j+1} overlap")

                if age_errors:
                    st.error("Invalid groups — please adjust ranges")

                # Build labels
                custom_age_labels = []
                for i, (s, e) in enumerate(custom_age_groups):
                    if i == len(custom_age_groups) - 1 and e == max_age:
                        custom_age_labels.append(f"{s}+")
                    else:
                        custom_age_labels.append(f"{s}-{e}")

                custom_age_range = {
                    "groups": custom_age_groups,
                    "labels": custom_age_labels,
                    "errors": age_errors,
                }
            elif analysis_variable == "Dependent Children Under 18":
                subgroups = st.multiselect(
                    "Select Groups",
                    ["Yes", "No"],
                    default=["Yes", "No"],
                )
            elif analysis_variable == "Financial Education Course":
                subgroups = st.multiselect(
                    "Select Groups",
                    ["Yes", "No"],
                    default=["Yes", "No"],
                )
            else:
                available_values = sorted(df_genpop[analysis_col].dropna().unique().tolist())
                order = get_category_order(analysis_col)
                if order:
                    available_values = [v for v in order if v in available_values]
                subgroups = st.multiselect(
                    f"Select {analysis_variable} Groups",
                    available_values,
                    default=available_values,
                )

        elif environment == "Financial Well-Being":
            analysis_variable = st.selectbox(
                "Financial Well-Being Variable",
                list(FINANCIAL_WELLBEING_VARIABLES.keys()),
            )
            analysis_col = FINANCIAL_WELLBEING_VARIABLES[analysis_variable]
            available_values = df_genpop[analysis_col].dropna().unique().tolist()
            order = get_category_order(analysis_col)
            if order:
                available_values = [v for v in order if v in available_values]
            else:
                available_values = sorted(available_values, key=str)
            subgroups = st.multiselect(
                f"Select Groups",
                available_values,
                default=available_values,
            )

        st.markdown("---")

        # Axis assignment
        # Determine the group dimension label
        if environment == "Over the Years":
            group_dim_label = "Year"
        elif environment == "Demographics":
            group_dim_label = analysis_variable if analysis_variable else "Demographic"
        else:
            group_dim_label = analysis_variable if analysis_variable else "Financial Well-Being"

        axis_x = None
        axis_legend = None
        axis_facet = None

        if analysis_type == "Topic Bucket" and view_mode and "3-Category" in view_mode:
            # 3 dimensions: Topic, Group, Response Category
            dimensions = ["Topic", group_dim_label, "Response Category"]
            st.markdown("**Axis Assignment**")
            axis_x = st.selectbox("X-Axis", dimensions, index=1)
            remaining_for_legend = [d for d in dimensions if d != axis_x]
            axis_legend = st.selectbox("Legend", remaining_for_legend, index=0)
            axis_facet = [d for d in dimensions if d != axis_x and d != axis_legend][0]
            st.caption(f"Facet (panels): **{axis_facet}**")
        elif analysis_type == "Topic Bucket":
            # 2 dimensions: Topic and Group
            dimensions = ["Topic", group_dim_label]
            st.markdown("**Axis Assignment**")
            axis_x = st.selectbox("X-Axis", dimensions, index=0)
            axis_legend = [d for d in dimensions if d != axis_x][0]
            st.caption(f"Legend: **{axis_legend}**")
        else:
            # 2 dimensions: Total Correct and Group
            dimensions = ["Total Correct", group_dim_label]
            st.markdown("**Axis Assignment**")
            axis_x = st.selectbox("X-Axis", dimensions, index=0)
            axis_legend = [d for d in dimensions if d != axis_x][0]
            st.caption(f"Legend: **{axis_legend}**")

        st.markdown("---")

        # Chart type selection (after axis assignment so stacked bar validity can be checked)
        valid_charts = get_valid_chart_types(analysis_type, view_mode, environment, axis_legend)
        chart_type = st.selectbox("Chart Type", valid_charts)

        return {
            "environment": environment,
            "analysis_type": analysis_type,
            "view_mode": view_mode,
            "selected_topics": selected_topics,
            "selected_range": selected_range,
            "analysis_variable": analysis_variable,
            "analysis_col": analysis_col,
            "subgroups": subgroups,
            "year_range": year_range,
            "custom_age_range": custom_age_range,
            "chart_type": chart_type,
            "axis_x": axis_x,
            "axis_legend": axis_legend,
            "axis_facet": axis_facet,
            "group_dim_label": group_dim_label,
        }


# ==============================================================================
# MAIN ANALYSIS AND VISUALIZATION
# ==============================================================================
def run_analysis(config, df_years, df_genpop):
    environment = config["environment"]
    analysis_type = config["analysis_type"]
    view_mode = config["view_mode"]
    chart_type = config["chart_type"]

    # Select dataset and apply filters
    if environment == "Over the Years":
        df = df_years.copy()
        year_range = config["year_range"]
        if year_range:
            df = df[(df["survey_year"] >= year_range[0]) & (df["survey_year"] <= year_range[1])]
        group_col = "survey_year"
        group_label = "Year"
        dataset_name = "allYearsPFin8 (2017–2026)"
        analysis_col = "survey_year"
    else:
        df = df_genpop.copy()
        analysis_col = config["analysis_col"]
        dataset_name = "PFin2026_GenPop (2026)"

        if config["analysis_variable"] == "Age (Custom Range)":
            age_config = config["custom_age_range"]
            if age_config and age_config.get("errors"):
                st.error("Invalid groups — please adjust ranges")
                return None, None, None

            groups = age_config["groups"]
            labels = age_config["labels"]

            # Assign each respondent to a group based on their age
            def assign_age_group(age):
                for i, (start, end) in enumerate(groups):
                    if start <= age <= end:
                        return labels[i]
                return None  # Age not in any group

            df["custom_age_group"] = df["reported_age"].apply(assign_age_group)
            df = df.dropna(subset=["custom_age_group"])

            group_col = "custom_age_group"
            group_label = "Age Range"
            analysis_col = "custom_age_group"

            # Set category order to match the label order (used later via config)
        elif config["analysis_variable"] == "Dependent Children Under 18":
            df["has_dependent_children_display"] = df["has_dependent_children"].map(BINARY_DISPLAY)
            group_col = "has_dependent_children_display"
            group_label = DEMOGRAPHIC_DISPLAY_LABELS.get("has_dependent_children", config["analysis_variable"])
            analysis_col = "has_dependent_children_display"
            if config["subgroups"]:
                df = df[df[group_col].isin(config["subgroups"])]
        elif config["analysis_variable"] == "Financial Education Course":
            df["took_Financial_Education_display"] = df["took_Financial_Education"].map(BINARY_DISPLAY)
            group_col = "took_Financial_Education_display"
            group_label = DEMOGRAPHIC_DISPLAY_LABELS.get("took_Financial_Education", config["analysis_variable"])
            analysis_col = "took_Financial_Education_display"
            if config["subgroups"]:
                df = df[df[group_col].isin(config["subgroups"])]
        else:
            group_col = analysis_col
            if environment == "Demographics":
                group_label = DEMOGRAPHIC_DISPLAY_LABELS.get(analysis_col, config["analysis_variable"])
            else:
                group_label = FINANCIAL_WELLBEING_LABELS.get(analysis_col, config["analysis_variable"])

            # Filter by subgroups
            if config["subgroups"]:
                df = df[df[group_col].isin(config["subgroups"])]

        # Drop missing values in the group column
        df = df.dropna(subset=[group_col])

    # Check if data remains after filtering
    if len(df) == 0:
        st.warning("No data available for the selected filters. Please adjust your selections.")
        return None, None, None

    # Determine category orders for the group column
    cat_order = get_category_order(group_col)
    category_orders = {}
    if group_col == "custom_age_group" and config.get("custom_age_range"):
        # Use the user-defined label order for custom age groups
        category_orders["group_value"] = config["custom_age_range"]["labels"]
    elif cat_order:
        available = [v for v in cat_order if v in df[group_col].unique()]
        category_orders["group_value"] = [str(v) for v in available]
    if environment == "Over the Years":
        category_orders["group_value"] = [str(v) for v in sorted(df[group_col].unique())]

    # Build chart data
    chart_data = None
    color_col = "group_value"
    x_label = ""
    hover_mode = "binary"
    use_facet = None
    group_dim_label = config["group_dim_label"]
    axis_x = config["axis_x"]
    axis_legend = config["axis_legend"]
    axis_facet = config.get("axis_facet")

    # Map dimension names to data columns
    def dim_to_col(dim_name, mode="binary"):
        if dim_name == "Topic":
            return "topic"
        elif dim_name == "Total Correct":
            return "score_label"
        elif dim_name == "Response Category":
            return "response_category"
        else:  # group dimension
            return "group_value"

    if analysis_type == "Topic Bucket":
        selected_topics = config["selected_topics"]
        if not selected_topics:
            return None, None, None

        if view_mode and "Binary" in view_mode:
            topics_map = {k: v for k, v in TOPIC_NAMES.items() if k in selected_topics}
            chart_data = prepare_topic_binary_data(df, topics_map, group_col, group_label)
            hover_mode = "binary"
            y_label = "% Correct"

            # Assign axes
            x_col = dim_to_col(axis_x)
            legend_col = dim_to_col(axis_legend)
            x_dim_label = axis_x if axis_x == "Topic" else group_label
            legend_dim_label = axis_legend if axis_legend == "Topic" else group_label

            chart_data["x"] = chart_data[x_col]
            color_col = legend_col
            x_label = x_dim_label
            title = f"P-Fin 8: % Correct — {x_dim_label} × {legend_dim_label}"

        else:
            topics_map = {k: v for k, v in TOPIC_CAT3_NAMES.items() if k in selected_topics}
            chart_data = prepare_topic_cat3_data(df, topics_map, group_col, group_label)
            hover_mode = "cat3"
            y_label = "% of Respondents"

            # Assign axes for 3 dimensions
            x_col = dim_to_col(axis_x, "cat3")
            legend_col = dim_to_col(axis_legend, "cat3")
            facet_dim = axis_facet
            facet_col = dim_to_col(axis_facet, "cat3") if axis_facet else None

            x_dim_label = "Topic" if axis_x == "Topic" else ("Response Category" if axis_x == "Response Category" else group_label)
            legend_dim_label = "Topic" if axis_legend == "Topic" else ("Response Category" if axis_legend == "Response Category" else group_label)

            chart_data["x"] = chart_data[x_col]
            color_col = legend_col
            use_facet = facet_col
            x_label = x_dim_label
            title = f"P-Fin 8: Response Distribution by {group_label}"

            # Set category orders for response_category if used
            if "response_category" in [x_col, legend_col, facet_col]:
                category_orders["response_category"] = ["Correct", "Incorrect", "Don't Know"]

    else:
        selected_range = config["selected_range"]
        chart_data = prepare_total_correct_data(df, group_col, score_range=selected_range)
        hover_mode = "total_correct"
        y_label = "% of Respondents"

        # Assign axes
        x_col = dim_to_col(axis_x)
        legend_col = dim_to_col(axis_legend)
        x_dim_label = "Total Correct" if axis_x == "Total Correct" else group_label
        legend_dim_label = "Total Correct" if axis_legend == "Total Correct" else group_label

        chart_data["x"] = chart_data[x_col]
        color_col = legend_col
        x_label = x_dim_label
        title = f"P-Fin 8: Distribution of Total Correct — {x_dim_label} × {legend_dim_label}"

        # Ensure score order
        if not chart_data.empty:
            score_labels = [TOTAL_CORRECT_LABELS[i] for i in range(
                selected_range[0] if selected_range else 0,
                (selected_range[1] if selected_range else 8) + 1
            )]
            category_orders["score_label"] = score_labels
            if x_col == "score_label":
                category_orders["x"] = score_labels

    if chart_data is None or chart_data.empty:
        st.warning("No data available for the selected combination. Please adjust your filters.")
        return None, None, None

    # Set legend label
    legend_label_text = None
    if environment == "Financial Well-Being" and legend_col == "group_value":
        legend_label_text = "Response"
    elif legend_col == "response_category":
        legend_label_text = "Response Category"
    elif legend_col == "topic":
        legend_label_text = "Topic"
    elif legend_col == "score_label":
        legend_label_text = "Total Correct"
    else:
        legend_label_text = group_label

    # Create chart
    fig = create_chart(chart_data, chart_type, title, x_label, y_label, color_col,
                       category_orders, group_label=legend_label_text, hover_mode=hover_mode,
                       legend_label=legend_label_text, facet_col=use_facet)

    # Generate note
    note = generate_note(
        environment=environment,
        analysis_type=analysis_type,
        view_mode=view_mode,
        selected_topics=config["selected_topics"],
        selected_range=config["selected_range"],
        analysis_variable=config.get("analysis_variable"),
        subgroups=config.get("subgroups"),
        n_obs=len(df),
        dataset_name=dataset_name,
        year_range=config.get("year_range"),
    )

    # Run sanity checks
    checks = run_sanity_checks(df, "survey_weight", chart_data, environment, analysis_col)

    return fig, note, checks


# ==============================================================================
# DEBUG PANEL
# ==============================================================================
def render_debug_panel(checks):
    if not DEBUG_MODE or checks is None:
        return

    with st.expander("🔧 Debug: Validation Panel", expanded=False):
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("**✅ Passed**")
            for item in checks["passed"]:
                st.write(f"• {item}")

        with col2:
            st.markdown("**⚠️ Warnings**")
            if checks["warnings"]:
                for item in checks["warnings"]:
                    st.warning(item)
            else:
                st.write("No warnings")

        with col3:
            st.markdown("**❌ Errors**")
            if checks["errors"]:
                for item in checks["errors"]:
                    st.error(item)
            else:
                st.write("No errors")


# ==============================================================================
# MAIN APP
# ==============================================================================
def main():
    st.set_page_config(
        page_title="P-Fin 8 Data Exploration Tool",
        page_icon=None,
        layout="wide",
    )

    # Load data
    try:
        df_years = load_all_years()
        df_genpop = load_genpop()
    except FileNotFoundError as e:
        st.error(
            f"Data file not found: {e}. Please ensure 'allYearsPFin8.xlsx' and "
            f"'PFin2026_GenPop.xlsx' are in the directory: {DATA_DIR.resolve()}"
        )
        return

    # Render sidebar and get config
    config = render_sidebar(df_years, df_genpop)

    # Main content
    st.title("P-Fin 8 Data Exploration Tool")
    st.markdown("**What is the P-Fin?**")
    st.markdown(
        "The TIAA Institute–GFLEC Personal Finance Index (P-Fin) is an annual survey of U.S. adults "
        "designed to measure the knowledge and understanding needed for sound financial decision-making "
        "and effective management of personal finances. It was first fielded in 2016 for the inaugural "
        "2017 report, and it has since been used each year as a broad measure of financial literacy in "
        "the United States."
    )

    # Custom CSS to make the toggle buttons look like inline text links
    # Safe to target all stButton because Show more/less are the only st.button calls
    # (download uses st.download_button which has a different data-testid)
    st.markdown("""
        <style>
        [data-testid="stButton"] button {
            background: none !important;
            border: none !important;
            padding: 0 !important;
            margin: -10px 0 0 0 !important;
            color: inherit !important;
            font-weight: 800 !important;
            font-size: 1rem !important;
            cursor: pointer !important;
            box-shadow: none !important;
            min-height: 0 !important;
            line-height: 1.5 !important;
        }
        [data-testid="stButton"] button:hover {
            text-decoration: underline !important;
            color: #1f4e79 !important;
        }
        [data-testid="stButton"] button p {
            font-weight: 800 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    if "show_full_summary" not in st.session_state:
        st.session_state.show_full_summary = False

    if not st.session_state.show_full_summary:
        if st.button("Show more", key="show_more"):
            st.session_state.show_full_summary = True
            st.rerun()
    else:
        st.markdown(
            "The P-Fin measures financial literacy using 28 multiple-choice questions grouped across "
            "eight functional areas of personal finance: earning, consuming, saving, investing, "
            "borrowing/managing debt, insuring, comprehending uncertainty, and go-to information sources. "
            "In this way, the survey does not treat financial literacy as a single narrow concept, but "
            "rather as a broad set of skills tied to the major areas in which people routinely make "
            "financial decisions.\n\n"
            "Researchers recently created a smaller reduced-form financial literacy measure by selecting "
            "one question from each of the eight topic buckets. This produces an 8-question index, which "
            "can be referred to as the P-Fin 8. The P-Fin 8 preserves the broad topic coverage of the "
            "original survey while offering a shorter summary measure of financial literacy.\n\n"
            "Using this tool, you will be able to explore the P-Fin 8 with visuals in three different "
            "ways: responses throughout the years, by demographics, and by financial well-being. For each "
            "environment, you will be able to create visuals of the data for further analysis.\n\n"
            "For further information, visit "
            "[The TIAA Institute-GFLEC](https://gflec.org/initiatives/personal-finance-index/#list)."
        )
        if st.button("Show less", key="show_less"):
            st.session_state.show_full_summary = False
            st.rerun()
    st.markdown("---")

    # Run analysis
    fig, note, checks = run_analysis(config, df_years, df_genpop)

    # Display chart
    if fig:
        st.plotly_chart(fig, use_container_width=True)

        # Display sample size warnings inline
        if checks and checks["warnings"]:
            for warning in checks["warnings"]:
                st.caption(warning)

        # Display note
        if note:
            st.markdown("---")
            st.caption(note)

        # Debug panel
        render_debug_panel(checks)

        # Export option
        st.markdown("---")
        st.download_button(
            label="📥 Download Chart as HTML",
            data=fig.to_html(include_plotlyjs="cdn"),
            file_name="pfin8_chart.html",
            mime="text/html",
        )


if __name__ == "__main__":
    main()
