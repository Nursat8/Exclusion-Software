import streamlit as st
import pandas as pd
import io

# Streamlit App Title
st.title("📊 Company Filtering & Exclusion App")

# File Uploader
uploaded_file = st.file_uploader("📂 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Load Excel file
    df = pd.read_excel(uploaded_file, skiprows=5)  # Skip extra rows

    # Rename the first column to "Company Name"
    df.rename(columns={"Unnamed: 0": "Company Name"}, inplace=True)

    # Strip spaces from column names
    df.columns = df.columns.str.strip()

    # Identify exclusion columns dynamically
    exclusion_columns = [col for col in df.columns if "SP_ESG_BUS_INVOLVE_REV_PCT" in col]

    # Convert exclusion columns to numeric
    df[exclusion_columns] = df[exclusion_columns].replace({',': '.', ' ': ''}, regex=True)
    df[exclusion_columns] = df[exclusion_columns].apply(pd.to_numeric, errors='coerce')

    # Apply filtering: Remove companies where any exclusion column has revenue > 0%
    filtered_df = df[~(df[exclusion_columns] > 0).any(axis=1)]

    st.success("✅ Filtering completed!")

    # Display preview of filtered companies
    st.subheader("Filtered Companies Preview:")
    st.dataframe(filtered_df.head())

    # **💡 Adjustable Exclusion Thresholds (Using Sliders)**
    st.sidebar.header("🔧 Adjust Exclusion Thresholds")

    alcohol_threshold = st.sidebar.slider("Alcohol Threshold (%)", 0, 100, 10)
    gambling_threshold = st.sidebar.slider("Gambling Threshold (%)", 0, 100, 5)
    adult_entertainment_threshold = st.sidebar.slider("Adult Entertainment Threshold (%)", 0, 100, 5)
    palm_oil_threshold = st.sidebar.slider("Palm Oil Threshold (%)", 0, 100, 5)
    pesticides_threshold = st.sidebar.slider("Pesticides Threshold (%)", 0, 100, 20)

    # Define Exclusion Rules (Using User-Selected Thresholds)
    exclusion_rules = {
        "Alcohol": ("Alcohol", alcohol_threshold),  
        "Gambling": ("Gambling", gambling_threshold),
        "Adult Entertainment": ("Adult Entertainment", adult_entertainment_threshold),
        "Palm Oil": ("Palm Oil", palm_oil_threshold),
        "Pesticides": ("Pesticides", pesticides_threshold)
    }

    # Convert relevant columns to numeric
    for category, (column, threshold) in exclusion_rules.items():
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")

    # Initialize exclusion tracking
    df["Exclusion Reason"] = ""

    # Apply exclusion criteria
    for category, (column, threshold) in exclusion_rules.items():
        if column in df.columns:
            df.loc[df[column] > threshold, "Exclusion Reason"] += f"{category} revenue > {threshold}%; "

    # Separate included and excluded companies
    excluded_df = df[df["Exclusion Reason"] != ""].copy()
    retained_df = df[df["Exclusion Reason"] == ""].copy()

    # Remove "Exclusion Reason" from retained companies
    retained_df = retained_df.drop(columns=["Exclusion Reason"])

    # Display preview of excluded companies
    st.subheader("Excluded Companies Preview:")
    st.dataframe(excluded_df[["Company Name", "Exclusion Reason"]].head())

    # Save results to an in-memory Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
    
    # Download button for the exclusion file
    st.download_button(
        label="📥 Download Exclusion File",
        data=output.getvalue(),
        file_name="Excluded_Companies.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ Exclusion process complete! You can now download the results.")
