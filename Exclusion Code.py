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

    # ---------- 2. Exclusion settings ----------
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

    default_inclusive = {
        "Gambling",
        "Retail (Cannabis - Recreational)",
        "Adult Entertainment",
    }

    # ---------- 3. Individual thresholds ----------
    st.sidebar.subheader("Exclude by Individual Category")

    user_thresholds  = {}
    inclusive_flags  = {}

    for category, default_val in exclusion_categories.items():
        # Row layout:  [Exclude ☐ Category name.............]  [≥ ☐]
        col_lbl, col_geq = st.sidebar.columns([7, 1])

        apply_flag = col_lbl.checkbox(
            category,
            value=True,
            key=f"chk_{category}",
        )

        inclusive_flags[category] = col_geq.checkbox(
            "≥",
            value=category in default_inclusive,
            key=f"inc_{category}",
        )

        if apply_flag:
            user_thresholds[category] = st.sidebar.number_input(
                f"{category} Threshold (%)",
                min_value=0,
                max_value=100,
                value=default_val,
                key=f"num_thr_{category}",     # <-- unique key
            )

    # ---------- 4. Custom sum rules ----------
    st.sidebar.subheader("Exclude by Custom Sum of Categories")

    sum_count = st.sidebar.number_input(
        "Number of custom sum criteria",
        min_value=0,
        max_value=10,
        value=0,
        step=1,
    )

    custom_sum_definitions = []
    available_categories   = list(exclusion_categories.keys())

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

    # ---------- 5. Run button ----------
    run_processing = st.sidebar.button("Run Processing")

    # ---------- 6. Processing ----------
    if run_processing:
        # (everything from your step 6 onward remains unchanged)
        # 6-a. Load file
        wb = openpyxl.load_workbook(uploaded_file)
        sheet_name = wb.sheetnames[0]
        df = pd.read_excel(
            uploaded_file, sheet_name=sheet_name, skiprows=5, engine="openpyxl"
        )

        # 6-b. Rename unnamed ID columns (optional, keep original list)
        rename_dict = {
            "Unnamed: 1": "SP_ENTITY_ID",
            "Unnamed: 2": "SP_COMPANY_ID",
            "Unnamed: 3": "SP_ISIN",
            "Unnamed: 4": "SP_LEI",
        }
        df.rename(columns=rename_dict, inplace=True)

        # 6-c. Find numeric business-involvement columns
        excl_cols = [c for c in df.columns if "SP_ESG_BUS_INVOLVE_REV_PCT" in c]
        df[excl_cols] = (
            df[excl_cols].replace({",": ".", " ": ""}, regex=True)
            .apply(pd.to_numeric, errors="coerce")
        )

        # 6-d. Apply exclusions
        df["Exclusion Reason"] = ""
        exclusion_counts = {cat: 0 for cat in user_thresholds}

        # (i) Individual categories
        for cat, thr in user_thresholds.items():
            if cat in df.columns:
                inc = inclusive_flags[cat]          # 🔹
                op = ">=" if inc else ">"
                mask = df[cat] >= thr if inc else df[cat] > thr
                df.loc[mask, "Exclusion Reason"] += f"{cat} {op} {thr}%; "
                exclusion_counts[cat] += mask.sum()

        # (ii) Custom sums
        for idx, (cats, thr, inc) in enumerate(custom_sum_definitions, 1):
            if cats:
                sums = df[cats].apply(pd.to_numeric, errors="coerce").sum(axis=1)
                op = ">=" if inc else ">"
                mask = sums >= thr if inc else sums > thr
                reason = ", ".join(cats)
                df.loc[mask, "Exclusion Reason"] += (
                    f"Sum of [{reason}] {op} {thr}%; "
                )

        # 6-e. Split retained / excluded
        excluded_df = df[df["Exclusion Reason"] != ""].copy()
        retained_df = df[df["Exclusion Reason"] == ""].drop(
            columns=["Exclusion Reason"], errors="ignore"
        )

        # ---------- 7. Stats ----------
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

        # ---------- 8. Download ----------
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
