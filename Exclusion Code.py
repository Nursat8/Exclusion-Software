import streamlit as st
import pandas as pd
import io
import openpyxl

# Streamlit App Title
st.title("📊 Company Filtering & Exclusion App")

# File Uploader
uploaded_file = st.file_uploader("📂 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Load Excel file with formatting preserved
    workbook = openpyxl.load_workbook(uploaded_file)
    sheet_name = workbook.sheetnames[0]  # Assuming data is in the first sheet
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=5, engine='openpyxl')

    # Preserve original column names (spaces and formatting)
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

    # Apply filtering: Remove companies where any exclusion column has revenue > 0%
    filtered_df = df[~(df[exclusion_columns] > 0).any(axis=1)].copy()

    st.success("✅ Filtering completed!")

    # Display preview of filtered companies
    st.subheader("Filtered Companies Preview:")
    st.dataframe(filtered_df.head())

    # Sidebar for threshold adjustments
    st.sidebar.header("🔧 Adjust Exclusion Thresholds")
    exclusion_thresholds = {
        "Alcohol": st.sidebar.number_input("Alcohol Threshold (%)", min_value=0, max_value=100, value=10),
        "Gambling": st.sidebar.number_input("Gambling Threshold (%)", min_value=0, max_value=100, value=5),
        "Adult Entertainment": st.sidebar.number_input("Adult Entertainment Threshold (%)", min_value=0, max_value=100, value=5),
        "Palm Oil": st.sidebar.number_input("Palm Oil Threshold (%)", min_value=0, max_value=100, value=5),
        "Pesticides": st.sidebar.number_input("Pesticides Threshold (%)", min_value=0, max_value=100, value=20)
    }

    # Convert relevant columns to numeric if they exist in the dataset
    for category, threshold in exclusion_thresholds.items():
        if category in df.columns:
            df[category] = pd.to_numeric(df[category], errors="coerce")

    # Initialize exclusion tracking
    df["Exclusion Reason"] = ""
    exclusion_counts = {category: 0 for category in exclusion_thresholds.keys()}

    # Apply exclusion criteria
    for category, threshold in exclusion_thresholds.items():
        if category in df.columns:
            mask = df[category] > threshold
            df.loc[mask, "Exclusion Reason"] += f"{category} revenue > {threshold}%; "
            exclusion_counts[category] += mask.sum()

    # Separate included and excluded companies
    excluded_df = df[df["Exclusion Reason"] != ""].copy()
    retained_df = df[df["Exclusion Reason"] == ""].copy()

    # Remove "Exclusion Reason" from retained companies
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
    st.subheader("📊 Companies Excluded by Reason")
    for category, count in exclusion_counts.items():
        st.write(f"{category}: {count} companies excluded")

    # Save results to an in-memory Excel file while preserving format
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
