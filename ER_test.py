import streamlit as st
import pandas as pd
import altair as alt

# ===== ADD THESE CONSTANTS =====
CHART_WIDTH = 550
CHART_HEIGHT = 300
# ==============================


st.set_page_config(page_title="ER Automation", layout="wide")
st.title("🎯 ER Automated Report")

uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx with Sheet1, Sheet2 & Sheet3)", type=["xlsx"])

if uploaded_file:
    try:
        sheet_dict = pd.read_excel(uploaded_file, sheet_name=None)
        df1 = sheet_dict.get("Sheet1")
        df2 = sheet_dict.get("Sheet2")
        df3 = sheet_dict.get("Sheet3")
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
        st.stop()

    if df1 is None or df2 is None or df3 is None:
        st.error("❌ 'Sheet1', 'Sheet2', and 'Sheet3' are required in the Excel file.")
        st.stop()

    # Combine unique site names from all sheets
    sites_sheet1 = df1["Site Name"].dropna().unique() if "Site Name" in df1.columns else []
    sites_sheet2 = df2["Site Name"].dropna().unique() if "Site Name" in df2.columns else []
    sites_sheet3 = df3["Site Name"].dropna().unique() if "Site Name" in df3.columns else []
    site_list = sorted(set(sites_sheet1).union(sites_sheet2).union(sites_sheet3))
    site_list.insert(0, "Enterprise")

    selected_site = st.selectbox("🏢 Select Site", site_list, index=0)

    ### ---------- SHEET1 ANALYSIS ----------
    st.header("📊 Sheet1: ER Scores")
    # ... (rest of your Sheet1 analysis code remains the same) ...
    col1 = "Was the correct process article for the primary player issue captured?"
    col2 = "Was the case linked to the correct process in the Knowledge Base?"
    if col1 in df1.columns and col2 in df1.columns:
        def merge_bi(row):
            val1 = row.get(col1)
            val2 = row.get(col2)
            if pd.isna(val1) and pd.isna(val2):
                return None
            elif val1 == 5 or val2 == 5:
                return 5
            elif val1 == 0 or val2 == 0:
                return 0
            else:
                return None
        df1["Business Intelligence"] = df1.apply(merge_bi, axis=1)
        df1.drop([col1, col2], axis=1, inplace=True)
    elif col1 in df1.columns:
        df1.rename(columns={col1: "Business Intelligence"}, inplace=True)
    elif col2 in df1.columns:
        df1.rename(columns={col2: "Business Intelligence"}, inplace=True)
    quality_params = [
        "Professional Greeting", "Understand", "Assurance", "Intent",
        "Correct Process", "Complete Process", "Time Management", "Additional Needs",
        "Professional Close", "Engaged Tone", "Mirror Persona", "Player Privacy",
        "Legal and Safety", "Positive Positioning", "Business Intelligence"
    ]
    if selected_site == "Enterprise":
        df1_filtered = df1.copy()
    else:
        df1_filtered = df1[df1["Site Name"] == selected_site]
    st.subheader("📌 Evaluation Summary")
    er_reviews = len(df1_filtered)
    er_score = round(df1_filtered['Avg. SCORE_AVG'].mean(), 2) if 'Avg. SCORE_AVG' in df1_filtered.columns else "N/A"
    colA, colB = st.columns(2)
    colA.metric("🧾 ER Reviews", er_reviews)
    colB.metric("⭐ ER Score (Avg. SCORE_AVG)", er_score)
    avg_scores = (df1_filtered[quality_params] == 5).sum() / df1_filtered[quality_params].notna().sum() * 100
    avg_scores = avg_scores.round(2)
    df1_result = pd.DataFrame([avg_scores], index=[selected_site])
    st.subheader(f"📊 Standardwise Score - {selected_site}")
    st.dataframe(df1_result)
    # Filter out standards with no data for the selected site
    valid_standards = [std for std in quality_params if
                       std in df1_result.columns and not pd.isna(df1_result[std].iloc[0])]
    df1_viz = df1_result[valid_standards].reset_index().melt(
        id_vars='index',
        value_vars=valid_standards,
        var_name='Standard',
        value_name='Score (%)'
    )

    # Create and display the bar chart (only if data exists)
    if not df1_viz.empty:
        # 1. Preserve original order
        original_order = df1_viz['Standard'].unique().tolist()

        # 2. Create base chart with PROPERLY CLOSED STRINGS/PARENTHESES
        base = alt.Chart(df1_viz).encode(
            x=alt.X(
                'Standard:N',
                sort=original_order,
                axis=alt.Axis(labelAngle=-45, labelLimit=200)
            ),  # Comma added here
            y=alt.Y(
                'Score (%):Q',
                scale=alt.Scale(domain=[0, 105])
            )
        )

        # 3. Create bars with FIXED QUOTES
        bars = base.mark_bar(
            color='steelblue',
            size=20
        ).properties(
            width=750,
            height=500,
            title={
                "text": "Standard scores",
                "anchor": "middle",
                "fontSize": 16
            }
                # Increase font size if needed
                  # Adjust dy to move the title up or down

        )

        # 4. Labels with guaranteed visibility
        text = base.mark_text(
            align='center',
            baseline='bottom',
            dy=-12,
            fontSize=12,
            fontWeight='bold',
            color='red'  # Temporary for verification
        ).encode(
            text=alt.Text('Score (%):Q', format='.1f')
        )

        # 5. Combine and render
        chart = (bars + text).properties(
            # Set overall chart dimensions with buffer space
            width=750,
            height=450  # Increased from 450 to 500
        ).configure_view(
            strokeOpacity=0
        ).configure_title(
            # These settings will pull the title up
            anchor='middle',
            dy=10

        )

        st.altair_chart(chart, use_container_width=False)

    st.download_button(
        label=f"⬇️ Download Sheet1 QA Report ({selected_site})",
        data=df1_result.to_csv().encode("utf-8"),
        file_name=f"{selected_site.lower()}_qa_report.csv",
        mime="text/csv"
    )
    st.subheader(f"🕵️ Quality Evaluator Performance - {selected_site}")
    if 'Quality Evaluator Name' in df1_filtered.columns and 'Avg. SCORE_AVG' in df1_filtered.columns:
        evaluator_stats = df1_filtered.groupby('Quality Evaluator Name').agg(
            Evaluation_Count=('Quality Evaluator Name', 'size'),
            Average_ER_Score=('Avg. SCORE_AVG', 'mean')
        ).sort_values(by='Average_ER_Score', ascending=False).reset_index()
        evaluator_stats['Average_ER_Score'] = evaluator_stats['Average_ER_Score'].round(2)
        st.dataframe(evaluator_stats)
        st.subheader(f"📈 ER Leaderboard - {selected_site}")
        base = alt.Chart(evaluator_stats).encode(
            x=alt.X('Quality Evaluator Name', sort=alt.EncodingSortField(field='Average_ER_Score', op='mean', order='descending'), title='Quality Evaluator')
        )
        bars = base.mark_bar(width=15).encode(
            y=alt.Y('Average_ER_Score', title='Average ER Score', axis=alt.Axis(titleColor='steelblue')),
            color=alt.value('steelblue'),
            tooltip=['Quality Evaluator Name', 'Average_ER_Score']
        )
        text_bars = base.mark_text(dy=-10).encode(
            y=alt.Y('Average_ER_Score', axis=None),
            text=alt.Text('Average_ER_Score', format='.1f'),
            color=alt.value('Red')
        )
        line = base.mark_line(color='orange').encode(
            y=alt.Y('Evaluation_Count', title='No. of Evaluations', axis=alt.Axis(orient='right', titleColor='orange')),
            tooltip=['Quality Evaluator Name', 'Evaluation_Count']
        )
        combo_chart = alt.layer(bars + text_bars, line).resolve_scale(
            y='independent'
        ).properties(
            width=600,
            height=350,
            title='ER Leaderboard'
        ).configure_title(
            anchor='middle'
        )
        st.altair_chart(combo_chart, use_container_width=False)
    else:
        st.warning("⚠️ 'Quality Evaluator Name' or 'Avg. SCORE_AVG' column not found in Sheet1 to display evaluator performance.")

    ### ---------- SHEET2 ANALYSIS ----------
    st.markdown("---")
    st.header("⚖️ Standard Disputes")

    if selected_site == "Enterprise":
        df3_filtered = df3.copy()
    else:
        df3_filtered = df3[df3["Site Name"] == selected_site]

    if not df3_filtered.empty and all(
            col in df3_filtered.columns for col in ["Standard", "Overturned", "Upheld", "Partial Overturn"]):
        # Calculate total disputes and filter out standards with no disputes
        df3_filtered['Total Disputes'] = df3_filtered[['Overturned', 'Upheld', 'Partial Overturn']].sum(axis=1)
        df3_filtered = df3_filtered[df3_filtered['Total Disputes'] > 0].drop(columns=['Total Disputes'])

        # Create pivot table for display
        standard_disputes_pivot = df3_filtered.pivot_table(
            index='Standard',
            values=['Overturned', 'Upheld', 'Partial Overturn'],
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        st.subheader("Standard Disputes Table")
        st.dataframe(standard_disputes_pivot)

        # Download button
        st.download_button(
            label=f"⬇️ Download Standard Disputes ({selected_site})",
            data=standard_disputes_pivot.to_csv(index=False).encode("utf-8"),
            file_name=f"{selected_site.lower()}_standard_disputes.csv",
            mime="text/csv"
        )

        st.subheader("📊 Standards Dispute Status")

        # NEW AGGREGATION CODE (REPLACE EXISTING MELT CODE)
        dispute_totals = df3_filtered.groupby('Standard')[
            ['Overturned', 'Upheld', 'Partial Overturn']].sum().reset_index()

        melted = dispute_totals.melt(
            id_vars='Standard',
            value_vars=['Overturned', 'Upheld', 'Partial Overturn'],
            var_name='Dispute Status',
            value_name='Count'
        ).query('Count > 0')  # Only show positive counts

        # Calculate positions for labels (NEW)
        melted['cumulative'] = melted.groupby('Standard')['Count'].cumsum()
        melted['previous'] = melted.groupby('Standard')['Count'].shift().fillna(0)
        melted['middle'] = (melted['previous'] + melted['cumulative']) / 2

        # 3. Create the bars with centered title
        bars = alt.Chart(melted).mark_bar().encode(
            x=alt.X('Standard:N', title='Standard', axis=alt.Axis(labelAngle=-45)),
            y=alt.Y('Count:Q', stack='zero', title='Dispute Count'),
            color=alt.Color('Dispute Status:N',
                            scale=alt.Scale(domain=['Overturned', 'Upheld', 'Partial Overturn'],
                                            range=['#FF6F61', '#6BA292', '#F4A261'])),
            order=alt.Order('Dispute Status:N', sort='ascending'),
            tooltip=['Standard', 'Dispute Status', 'Count']
        ).properties(
            width=700,
            height=400,
            title={
                "text": "Standards Dispute",
                "anchor": "middle",
                "fontSize": 16,
                "fontWeight": "bold",
                "offset": 20
            }
        )

        # 4. Add the improved labels (THIS IS WHERE YOUR LABEL CODE GOES)
        labels = alt.Chart(melted).mark_text(
            align='center',
            baseline='middle',
            dy=0,
            fontSize=12,
            fontWeight='bold',
            color='black'
        ).encode(
            x=alt.X('Standard:N'),
            y=alt.Y('middle:Q'),
            text=alt.Text('Count:Q', format='.0f')
        )

        # 5. Final composition and configuration (REPLACES st.altair_chart)
        final_chart = (bars + labels).configure_view(
            strokeOpacity=0
        ).configure_title(
            anchor='middle'
        ).configure_axisX(
            titlePadding=25
        )

        # 6. Display the chart
        st.altair_chart(final_chart, use_container_width=False)

    else:
        st.info(f"No standard dispute data available for {selected_site}.")

    ### ---------- SHEET2 ANALYSIS ----------
    st.markdown("---")
    st.header("📝 Slack Queries Analysis")

    # Define constants if not already defined at top
    CHART_WIDTH = 500
    CHART_HEIGHT = 450
    TITLE_FONT_SIZE = 16
    BAR_COLOR = "#4B9AC7"  # Slack-like blue

    if selected_site == "Enterprise":
        df2_filtered = df2.copy()
    else:
        df2_filtered = df2[df2["Site Name"] == selected_site]

    # Get all standard columns (excluding Site Name)
    standards = [col for col in df2.columns if col != "Site Name"]

    # Calculate counts for each standard
    standard_counts = df2_filtered[standards].sum().reset_index()
    standard_counts.columns = ['Standard', 'Count']
    standard_counts = standard_counts[standard_counts['Count'] > 0]  # Filter standards with at least 1 count

    if not standard_counts.empty:
        # Display the counts table
        st.subheader(f"📌 Standard Query Counts - {selected_site}")
        st.dataframe(standard_counts.sort_values('Count', ascending=False))

        # Create the bar chart
        st.subheader("📊 Slack Queries")

        chart = alt.Chart(standard_counts).mark_bar(
            color=BAR_COLOR,
            size=20
        ).encode(
            x=alt.X('Standard:N',
                    sort='-y',  # Sort by count descending
                    axis=alt.Axis(labelAngle=-45, title='Standard')),
            y=alt.Y('Count:Q', title='Number of Queries')
        ).properties(
            width=CHART_WIDTH,
            height=CHART_HEIGHT,
            title={
                "text": "Slack Queries",
                "anchor": "middle",
                "fontSize": TITLE_FONT_SIZE
            }
        )

        # Add data labels
        text = chart.mark_text(
            align='center',
            baseline='bottom',
            dy=-10,  # Adjust label position
            fontSize=12,
            fontWeight='bold',
            color='black'
        ).encode(
            text='Count:Q'
        )

        final_chart = (chart + text).configure_view(
            strokeOpacity=0
        )

        st.altair_chart(final_chart, use_container_width=False)

        # Download button
        st.download_button(
            label=f"⬇️ Download Slack Query Counts ({selected_site})",
            data=standard_counts.to_csv(index=False).encode("utf-8"),
            file_name=f"{selected_site.lower()}_slack_queries.csv",
            mime="text/csv"
        )
    else:
        st.info(f"No slack query data available for {selected_site}.")
