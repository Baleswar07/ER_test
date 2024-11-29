import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Title of the app
st.title("Standards Score Analysis")

# Upload CSV file
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file is not None:
    # Read the CSV file with the correct delimiter and encoding
    df = pd.read_csv(uploaded_file, encoding='utf-8', delimiter=',')

    # Strip spaces from column names
    df.columns = df.columns.str.strip()

    # Ensure the correct column name is used
    site_name_column = 'Site Name'  # Ensure this matches exactly

    # Check if the column exists
    if site_name_column in df.columns:
        # Select the columns for standards scores
        standards_score = df.loc[:, 'Professional Greeting':'Was the case linked to the correct process in the Knowledge Base?']

        # Group by 'Site Name' and calculate the mean for each site
        grouped = df.groupby(site_name_column)[standards_score.columns].mean()

        # Divide each average by 5 and multiply by 100 to get the percentage
        Standardwise_score = ((grouped / 5) * 100).round(2)

        # Ensure the index is not empty
        if not Standardwise_score.empty:
            # Dropdown to select a site name
            site_names = Standardwise_score.index.tolist()
            selected_site = st.selectbox("Select a Site Name", site_names)

            # Count the number of evaluations for the selected site
            evaluation_count = df[df[site_name_column] == selected_site].shape[0]

            # Display evaluation count in a small box
            st.markdown(
                f"""
                <div style="border: 1px solid #ccc; padding: 10px; width: 150px; text-align: center; border-radius: 5px; background-color: #f0f0f0;">
                    <strong>Evaluation Number:</strong><br>{evaluation_count}
                </div>
                """,
                unsafe_allow_html=True
            )

            # Display the result for the selected site
            st.write(f"Standardwise_score (%) for {selected_site}:")
            site_scores = Standardwise_score.loc[selected_site]
            st.dataframe(site_scores)

            # Identify top 3 and bottom 3 standards
            top_3 = site_scores.nlargest(3)
            bottom_3 = site_scores.nsmallest(3)

            st.write("Top 3 Standards:")
            st.dataframe(top_3)

            st.write("Bottom 3 Standards:")
            st.dataframe(bottom_3)

            # Plot the bar graph with data labels
            st.write(f"Bar Graph for {selected_site}:")
            plt.figure(figsize=(12, 6))
            ax = site_scores.plot(kind='bar', width=0.6)
            plt.title(f"Standards Score for {selected_site}")
            plt.ylabel('Score (%)')
            plt.xticks(rotation=45, ha='right')

            # Rename the column for display in the graph
            ax.set_xticklabels(
                [label.replace('Was the case linked to the correct process in the Knowledge Base?',
                               'Business Intelligence') for label in site_scores.index]
            )

            # Add data labels
            for i, v in enumerate(site_scores):
                ax.text(i, v + 3, f"{v:.1f}%", ha='center', va='bottom')  # Adjust label position

            # Set y-axis limit to provide space for labels
            ax.set_ylim(0, site_scores.max() + 10)

            plt.tight_layout()
            st.pyplot(plt)
        else:
            st.write("No data available for the selected criteria.")