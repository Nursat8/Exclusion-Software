import streamlit as st
import pandas as pd
import io
import openpyxl

# ---------- 0. Title ----------
st.title("Company Filtering & Exclusion App")

# ---------- 1. File upload ----------
uploaded_file = st.file_uploader("📂 Upload an S&P file", type=["xlsx"])

# --------------------------------------------------------------------------- #
# Everything that depends on a file goes INSIDE this block
# --------------------------------------------------------------------------- #
if uploaded_file:

    # 2. 🔹 Exclusion settings. A dictionary with sectors and default percentage values set 🔹
    st.sidebar.header("🔧 Exclusion Criteria")

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
        "Retail (Cannabis - Recreational)": 10,
        "Wholesale (Cannabis - Recreational)": 5,
        "Pesticides": 20,
    }
    # 🔹 Sectors that have more and equal to in the beginning 🔹
    default_inclusive = {
        "Gambling",
        "Retail (Cannabis - Recreational)",
        "Adult Entertainment",
    }

    # 3. 🔹 Individual thresholds.🔹
    st.sidebar.subheader("Exclude by Individual Category")
    
    # 🔹 Creating open dictionaries for threshold and possible "more and equal to" condition (inclusive_flags). Will be filled dynamically 🔹
    user_thresholds  = {}
    inclusive_flags  = {}

    # 🔹 Sidebar UI for Streamlit. It creates two columns in the sidebar. The first column is for name of category and second "≥" checkbox. Sets "category" as first value of dictionary in exclusion_category and "default_val" as second value in dictionary 🔹    
    for category, default_val in exclusion_categories.items():
        # Row layout:  [Exclude ☐ Category name.............]  [≥ ☐]
        col_lbl, col_geq = st.sidebar.columns([7, 1])
        
        # 🔹 Sidebar UI for Streamlit. It checks all categories in the first column, which is name of category. Thus, thresholds for all sectors are activated in the beginning. 🔹 
        # 🔹 key=f"chk_{category}" creates widgets by following this logic: chk_Alcohol (if categort is Alcohol). It is needed to create a special ID for each sector and loop it 🔹 
        apply_flag = col_lbl.checkbox(
            category,
            value=True,
            key=f"chk_{category}",
        )
        # 🔹 Sidebar UI for Streamlit. It checks only sectors in default_inclusive. If category is in default_inclusive, it returns True; thus, value = True 🔹 
        inclusive_flags[category] = col_geq.checkbox(
            "≥",
            value=category in default_inclusive,
            key=f"inc_{category}",
        )
      
        # 🔹 Sidebar UI for Streamlit. Input of threshold. After all these, user_threshold is filled with data which have apply_flag == True and inclusive_flags == True. in the beginning, value is set to default_val. 🔹 
        if apply_flag:
            user_thresholds[category] = st.sidebar.number_input(
                f"{category} Threshold (%)",
                min_value=0,
                max_value=100,
                value=default_val,
                key=f"num_thr_{category}",    
            )

    # 🔹 4. Custom sum rules 🔹
    st.sidebar.subheader("Exclude by Custom Sum of Categories")
    
    # 🔹 Sidebar UI for Streamlit 🔹
    sum_count = st.sidebar.number_input(
        "Number of custom sum criteria",
        min_value=0,
        max_value=10,
        value=0,
        step=1,
    )
    
    # 🔹 Empty list custom_sum_definitions which is tuple that would consists of category (cats), threshold value (thr) and bolean (inc) for "equal and more than" 🔹
    # 🔹 Available_category extracts only category names from exclusion_categories; These values are used to populate the multiselect widgets for the user to pick categories to sum 🔹
    custom_sum_definitions = []
    available_categories   = list(exclusion_categories.keys())

    # 🔹 Sidebar UI for Streamlit. Shows categories for users and allows to set threshold and put checker for tuple custom_sum_definitions 🔹
    for i in range(int(sum_count)):
        st.sidebar.write(f"**Custom Sum #{i+1}**")
        cats = st.sidebar.multiselect(
            f"Select categories for Sum #{i+1}",
            available_categories,
            key=f"cats_{i}",
        )
        thr = st.sidebar.number_input(
            f"Threshold (%) for Sum #{i+1}",
            min_value=0,
            max_value=100,
            value=10,
            key=f"sum_thr_{i}",
        )
        inc = st.sidebar.checkbox(
            "≥",
            value=False,
            key=f"sum_inc_{i}",
        )
        custom_sum_definitions.append((cats, thr, inc))

    # 🔹 5. Run button 🔹
    run_processing = st.sidebar.button("Run Processing")

    # 🔹 6. Processing 🔹
    if run_processing:
        # 🔹 6-a. Load file 🔹
        wb = openpyxl.load_workbook(uploaded_file)
        sheet_name = wb.sheetnames[0]
        df = pd.read_excel(
            uploaded_file, sheet_name=sheet_name, skiprows=5, engine="openpyxl"
        )

        # 🔹 6-b. Rename unnamed ID columns (optional, keep original list). The output file had unnamed columns; thus, they are renamed 🔹
        rename_dict = {
            "Unnamed: 1": "SP_ENTITY_ID",
            "Unnamed: 2": "SP_COMPANY_ID",
            "Unnamed: 3": "SP_ISIN",
            "Unnamed: 4": "SP_LEI",
        }
        df.rename(columns=rename_dict, inplace=True)

        # 🔹 6-c. Find numeric business-involvement columns and data cleaning 🔹
        excl_cols = [c for c in df.columns if "SP_ESG_BUS_INVOLVE_REV_PCT" in c]
        df[excl_cols] = (
            df[excl_cols].replace({",": ".", " ": ""}, regex=True)
            .apply(pd.to_numeric, errors="coerce")
        )

        # 🔹🔹🔹 6-d. Apply exclusions. Pre-create an empty string column; later we concatenate reason. Initialise per-category counters 🔹🔹🔹
        df["Exclusion Reason"] = ""
        exclusion_counts = {cat: 0 for cat in user_thresholds}

        #  🔹 Individual categories. Check whether it is "equal or more than" or not by checking if inc = True (in the inclusive_flags list). And applies logic later through if else. 🔹
        #  🔹 df.loc[mask, "Exclusion Reason"] += f"{cat} {op} {thr}%; " is needed for output file to see Exclusion Reason column in the output file 🔹
        for cat, thr in user_thresholds.items():
            if cat in df.columns:
                inc = inclusive_flags[cat]         
                op = ">=" if inc else ">"
                mask = df[cat] >= thr if inc else df[cat] > thr
                df.loc[mask, "Exclusion Reason"] += f"{cat} {op} {thr}%; "
                exclusion_counts[cat] += mask.sum()

        # 🔹 Custom sums. Row-wise sum across the selected category columns. df[cats] → selects just the columns listed in cats and ensures that every value is number 🔹
        for idx, (cats, thr, inc) in enumerate(custom_sum_definitions, 1):
            if cats:
                sums = df[cats].apply(pd.to_numeric, errors="coerce").sum(axis=1)
                op = ">=" if inc else ">"
                mask = sums >= thr if inc else sums > thr
                reason = ", ".join(cats)
                df.loc[mask, "Exclusion Reason"] += (
                    f"Sum of [{reason}] {op} {thr}%; "
                )

        # 🔹 6-e. Split retained / excluded in the output file in the seperate sheets🔹
        excluded_df = df[df["Exclusion Reason"] != ""].copy()
        retained_df = df[df["Exclusion Reason"] == ""].drop(
            columns=["Exclusion Reason"], errors="ignore"
        )

        # 🔹 7. Stats. Shows exclusion_counts value for each category 🔹
        st.subheader("📈 Exclusion Statistics")
        st.write(f"Total Companies: {len(df)}")
        st.write(f"Excluded Companies: {len(excluded_df)}")
        st.write(f"Retained Companies: {len(retained_df)}")

        st.subheader("Companies Excluded by Individual Category")
        for cat, cnt in exclusion_counts.items():
            st.write(f"{cat}: {cnt}")

        st.subheader("Companies Excluded by Custom Sums")
        for i, (cats, thr, inc) in enumerate(custom_sum_definitions, 1):
            if cats:
                sums = df[cats].apply(pd.to_numeric, errors="coerce").sum(axis=1)
                mask = sums >= thr if inc else sums > thr
                op = ">=" if inc else ">"
                st.write(
                    f"Custom Sum #{i} (Categories: {', '.join(cats)}; "
                    f"Threshold: {thr}% {op}): {mask.sum()} companies excluded"
                )

        # 🔹 8. Download 🔹
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            retained_df.to_excel(writer, "Retained Companies", index=False)
            excluded_df.to_excel(writer, "Excluded Companies", index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download Filtered File",
            data=output,
            file_name="Filtered_SPGlobal_Output.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        st.success("✅ Exclusion process complete! You can now download the results.")
