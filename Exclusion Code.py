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
    pesticides_threshold
