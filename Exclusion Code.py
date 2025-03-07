import streamlit as st
import pandas as pd
import io
import openpyxl

# Streamlit App Title
st.title("Company Filtering & Exclusion App")

# File Uploader
uploaded_file = st.file_uploader("📂 Upload an S&P file", type=["xlsx"])

if uploaded_file:
    # Sidebar for exclusion selection
    st.sidebar.header("🔧 Exclusion Criteria")
    sector_exclusion = st.sidebar.checkbox("Exclude companies involved in sector")

    exclusion_categories = {
        "Nuclear Weapons": 0,
        "Depleted Uranium": 0,
        "Incendiary Weapons": 0,
        "Blinding Laser Weapons": 0,
        "Cluster Munitions": 0,
        "Anti-Personnel Mines": 0,
        "Biological and Chemical Weapons": 0,
        "Tobacco": 0,
        "Production (Tobacco)": 0,
        "Alcohol": 10,
        "Gambling": 5,
        "Adult Entertainment": 5,
        "Palm Oil": 5,
        "Retail (Cannabis - Recreational)": 0,
        "Wholesale (Cannabis - Recreational)": 0,
        "Pesticides": 20
    }

    # User-defined thresholds
    user_thresholds = {}
    for category, default_value in exclusion_categories.items():
        if st.sidebar.checkbox(f"Exclude {category}", value=True):
            user_thresholds[category] = st.sidebar.number_input(
                f"{category} Threshold (%)", min_value=0, max_value=100, value=default_value
            )

    # Sidebar button to trigger processing
    run_processing = st.sidebar.button("Run Processing")

    if run_processing:
        # Load Excel file with formatting preserved
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet_name = workbook.sheetnames[0]  # Assuming data is in the first sheet
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=5, engine='openpyxl')

        # Preserve original column names
        original_columns = df.columns.tolist()

        # Rename unnamed columns based on expected format
        rename_dict = {
            "Unnamed: 1": "SP_ENTITY_ID",
            "Unnamed: 2": "SP_COMPANY_ID",
            "Unnamed: 3": "SP_ISIN",
            "Unnamed: 4": "SP_LEI"
        }
        df.rename(columns=rename_dict, inplace=True)

        # Identify exclusion columns dynamically
        exclusion_columns = [col for col in df.columns if "SP_ESG_BUS_INVOLVE_REV_PCT" in col]

        # Convert exclusion columns to numeric
        df[exclusion_columns] = df[exclusion_columns].replace({',': '.', ' ': ''}, regex=True)
        df[exclusion_columns] = df[exclusion_columns].apply(pd.to_numeric, errors='coerce')
        
        # Apply exclusion criteria
        df["Exclusion Reason"] = ""
        exclusion_counts = {category: 0 for category in user_thresholds.keys()}
        
        for category, threshold in user_thresholds.items():
            if category in df.columns:
                mask = df[category] > threshold if not sector_exclusion else df[category].notna()
                df.loc[mask, "Exclusion Reason"] += f"{category} > {threshold}%; "
                exclusion_counts[category] += mask.sum()

        # Separate included and excluded companies
        excluded_df = df[df["Exclusion Reason"] != ""].copy()
        retained_df = df[df["Exclusion Reason"] == ""].copy()
        retained_df = retained_df.drop(columns=["Exclusion Reason"], errors='ignore')

        # Statistics
        st.subheader("📈 Exclusion Statistics")
        total_companies = len(df)
        excluded_companies = len(excluded_df)
        retained_companies = len(retained_df)
        
        st.write(f"Total Companies: {total_companies}")
        st.write(f"Excluded Companies: {excluded_companies}")
        st.write(f"Retained Companies: {retained_companies}")
        
        # Display exclusion counts per category
        st.subheader("Companies Excluded by Sector")
        for category, count in exclusion_counts.items():
            st.write(f"{category}: {count} companies excluded")

        # Save results to an in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
        output.seek(0)
        
        # Download button for the exclusion file
        st.download_button(
            label="📥 Download Exclusion File",
            data=output,
            file_name="Filtered_SPGlobal_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ Exclusion process complete! You can now download the results.")
