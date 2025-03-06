import streamlit as st
import pandas as pd
import io

# Streamlit app title
st.title("📊 Company Exclusion & Filtering App")

# File uploader
uploaded_file = st.file_uploader("📂 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Load the uploaded Excel file
    df = pd.read_excel(uploaded_file, skiprows=5)  # Skip extra rows

    # Strip spaces from column names
    df.columns = df.columns.str.strip()

    # Identify exclusion columns dynamically
    exclusion_columns = [col for col in df.columns if "SP_ESG_BUS_INVOLVE_REV_PCT" in col]

    # Convert exclusion columns to numeric (fix commas/spaces)
    df[exclusion_columns] = df[exclusion_columns].replace({',': '.', ' ': ''}, regex=True)
    df[exclusion_columns] = df[exclusion_columns].apply(pd.to_numeric, errors='coerce')

    # Apply exclusion: Remove companies where any exclusion column has revenue > 0%
    filtered_df = df[~(df[exclusion_columns] > 0).any(axis=1)]

    st.success("✅ Filtering completed!")

    # Display the first few rows of the filtered dataset
    st.subheader("Filtered Companies Preview:")
    st.dataframe(filtered_df.head())

    # Step 2: Apply Exclusion Rules
    exclusion_rules = {
        "Alcohol": ("Alcohol", 10),  
        "Gambling": ("Gambling", 5),
        "Adult Entertainment": ("Adult Entertainment", 5),
        "Palm Oil": ("Palm Oil", 5),
        "Pesticides": ("Pesticides", 20)
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

    # Remove the exclusion column for retained companies
    retained_df = retained_df.drop(columns=["Exclusion Reason"])

    st.subheader("Excluded Companies Preview:")
    st.dataframe(excluded_df[["SP_ENTITY_NAME", "Exclusion Reason"]].head())

    # Save results to an in-memory Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
    
    # Provide a download button for the final Excel file
    st.download_button(
        label="📥 Download Processed Excel",
        data=output.getvalue(),
        file_name="Excluded_Companies.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ Exclusion process complete! You can now download the results.")

