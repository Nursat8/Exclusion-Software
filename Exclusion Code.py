import streamlit as st
import pandas as pd
import io
import openpyxl

# Streamlit App Title
st.title("Company Filtering & Exclusion App")

# File Uploader
uploaded_file = st.file_uploader("📂 Upload an S&P file", type=["xlsx"])

if uploaded_file:
    # --- 1. Sidebar for exclusion thresholds ---
    st.sidebar.header("🔧 Exclusion Criteria")

    # Define the categories and their default threshold
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

    # --- 2. Individual category thresholds ---
    st.sidebar.subheader("Exclude by Individual Category")
    user_thresholds = {}
    for category, default_value in exclusion_categories.items():
        if st.sidebar.checkbox(f"Exclude {category}", value=True):
            user_thresholds[category] = st.sidebar.number_input(
                f"{category} Threshold (%)", min_value=0, max_value=100, value=default_value
            )

    # --- 3. Custom sum thresholds ---
    st.sidebar.subheader("Exclude by Custom Sum of Categories")
    # Decide how many custom sums the user wants to define
    sum_count = st.sidebar.number_input(
        "Number of custom sum criteria", min_value=0, max_value=10, value=0, step=1
    )

    # We will store each custom sum definition as a tuple: (list_of_categories, threshold_value)
    custom_sum_definitions = []
    available_categories = list(exclusion_categories.keys())  # or filter from your final data

    for i in range(int(sum_count)):
        st.sidebar.write(f"**Custom Sum #{i+1}**")
        selected_categories = st.sidebar.multiselect(
            f"Select categories to sum for Custom Sum #{i+1}",
            available_categories
        )
        threshold = st.sidebar.number_input(
            f"Threshold (%) for Custom Sum #{i+1}",
            min_value=0, max_value=100, value=10
        )
        custom_sum_definitions.append((selected_categories, threshold))

    # --- 4. Sidebar button to trigger processing ---
    run_processing = st.sidebar.button("Run Processing")

    if run_processing:
        # --- 5. Load Excel file with formatting preserved ---
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

        # Identify exclusion columns dynamically (any column that has 'SP_ESG_BUS_INVOLVE_REV_PCT')
        exclusion_columns = [col for col in df.columns if "SP_ESG_BUS_INVOLVE_REV_PCT" in col]

        # Convert exclusion columns to numeric (if necessary)
        df[exclusion_columns] = df[exclusion_columns].replace({',': '.', ' ': ''}, regex=True)
        df[exclusion_columns] = df[exclusion_columns].apply(pd.to_numeric, errors='coerce')

        # --- 6. Apply exclusion criteria ---
        df["Exclusion Reason"] = ""

        # Track how many are excluded by each individual category
        exclusion_counts = {category: 0 for category in user_thresholds.keys()}

        # (a) Individual category thresholds
        for category, threshold in user_thresholds.items():
            if category in df.columns:
                mask = df[category] > threshold
                df.loc[mask, "Exclusion Reason"] += f"{category} > {threshold}%; "
                exclusion_counts[category] += mask.sum()

        # (b) Custom sum thresholds
        for i, (categories_list, threshold) in enumerate(custom_sum_definitions, start=1):
            if categories_list:
                # Ensure the columns exist and are numeric
                sum_series = df[categories_list].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                # Mask for rows exceeding the custom sum threshold
                mask = sum_series > threshold
                # Append a reason
                reason_str = ", ".join(categories_list)
                df.loc[mask, "Exclusion Reason"] += f"Sum of [{reason_str}] > {threshold}%; "

        # --- 7. Separate included and excluded companies ---
        excluded_df = df[df["Exclusion Reason"] != ""].copy()
        retained_df = df[df["Exclusion Reason"] == ""].copy()
        retained_df = retained_df.drop(columns=["Exclusion Reason"], errors='ignore')

        # --- 8. Statistics ---
        st.subheader("📈 Exclusion Statistics")
        total_companies = len(df)
        excluded_companies = len(excluded_df)
        retained_companies = len(retained_df)

        st.write(f"Total Companies: {total_companies}")
        st.write(f"Excluded Companies: {excluded_companies}")
        st.write(f"Retained Companies: {retained_companies}")

        # Display exclusion counts for individual categories
        st.subheader("Companies Excluded by Individual Category")
        for category, count in exclusion_counts.items():
            st.write(f"{category}: {count} companies excluded")

        st.subheader("Companies Excluded by Custom Sums")
        for i, (categories_list, threshold) in enumerate(custom_sum_definitions, start=1):
            if categories_list:
                sum_series = df[categories_list].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                mask = sum_series > threshold
                st.write(
                    f"Custom Sum #{i} (Categories: {', '.join(categories_list)}; "
                    f"Threshold: {threshold}%): {mask.sum()} companies excluded"
                )

        # --- 9. Save results to an in-memory Excel file ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
        output.seek(0)

        # --- 10. Download button for the exclusion file ---
        st.download_button(
            label="📥 Download Filtered File",
            data=output,
            file_name="Filtered_SPGlobal_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ Exclusion process complete! You can now download the results.")
