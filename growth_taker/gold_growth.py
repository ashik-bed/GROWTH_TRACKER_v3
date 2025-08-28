import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io
import streamlit as st

# ---------------- GOOGLE SHEETS SETTINGS ----------------
SERVICE_ACCOUNT_FILE = "sheetconnector-468508-1e0052475ae2.json"
SPREADSHEET_ID = "1gJUFsC0WTohZvo1gVF925dpQMVNXSR3GmjrOUXf9cFU"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

ADMIN_PASSWORD = "ASHph7#"  # Admin password for upload

# ---------------- FUNCTIONS (UNCHANGED) ----------------
def connect_to_gsheet():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def upload_dataframe_to_specific_tab(df, sheet_name):
    try:
        gc = connect_to_gsheet()
        try:
            ws = gc.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            ws = gc.add_worksheet(title=sheet_name, rows="1000", cols="20")

        ws.clear()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.update("A1", [["Last Updated:", timestamp]])
        ws.update("A3", [df.columns.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        st.error(f"❌ Failed to upload to Google Sheets: {e}")
        return False

def read_file(uploaded_file):
    if uploaded_file.name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    elif uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    elif uploaded_file.name.endswith(".tsv"):
        return pd.read_csv(uploaded_file, sep="\t")
    else:
        raise ValueError("Unsupported file format.")

# ---------------- STREAMLIT APP ----------------
st.set_page_config(page_title="Growth & Pending Analyzer", layout="centered")
st.title("📊 Growth & Pending Analyzer")
st.markdown("---")

# Step 1: Report selection
report_type = st.selectbox("📁 Select Report Type", ["Gold", "Subdebt", "SS Pending Report", "NPA"])
mode = None
if report_type in ["Gold", "Subdebt"]:
    mode = st.radio("📌 Select Report View", ["Branch-wise", "Staff-wise"], horizontal=True)

# Step 2: File uploads
file_types = ["xlsx", "xls", "csv", "tsv"]

if report_type == "Gold":
    old_file = st.file_uploader("📤 Upload OLD Gold Outstanding File", type=file_types, key="gold_old")
    new_file = st.file_uploader("📤 Upload NEW Gold Outstanding File", type=file_types, key="gold_new")

elif report_type == "Subdebt":
    old_file = st.file_uploader("📤 Upload OLD Subdebt Outstanding File", type=file_types, key="subdebt_old")
    new_file = st.file_uploader("📤 Upload NEW Subdebt Outstanding File", type=file_types, key="subdebt_new")

elif report_type == "SS Pending Report":
    pending_file = st.file_uploader("📤 Upload Gold Outstanding File", type=file_types, key="ss_pending")

elif report_type == "NPA":
    uploaded = st.file_uploader("📤 Upload Gold Outstanding File", type=file_types, key="npa_file")

# Step 3: Column mapping defaults
if report_type == "Gold":
    value_column = "PRINCIPAL OS"
    staff_column = "CANVASSER ID"
    branch_column = "BRANCH NAME"
elif report_type == "Subdebt":
    value_column = "Deposit Amount"
    staff_column = "Canvassed By"
    branch_column = "Branch Name"

# ---------------- GOLD / SUBDEBT REPORT ----------------
if report_type in ["Gold", "Subdebt"] and old_file and new_file:
    include_branches = False
    if report_type == "Subdebt" and mode == "Staff-wise":
        include_branches = st.checkbox("✅ Include Branches")

    if st.button("▶️ Run Report"):
        try:
            old_df = read_file(old_file)
            new_df = read_file(new_file)

            # 🔹 Remove Predator scheme
            if "SCHEME NAME" in old_df.columns:
                old_df = old_df[old_df["SCHEME NAME"].str.strip().str.upper() != "RCIL PREDATOR 18%"]
            if "SCHEME NAME" in new_df.columns:
                new_df = new_df[new_df["SCHEME NAME"].str.strip().str.upper() != "RCIL PREDATOR 18%"]

            required_cols = [value_column, staff_column, branch_column]
            missing_cols_old = [col for col in required_cols if col not in old_df.columns]
            missing_cols_new = [col for col in required_cols if col not in new_df.columns]

            if missing_cols_old or missing_cols_new:
                st.error(f"❌ Missing columns: {missing_cols_old + missing_cols_new}")
            else:
                # Grouping
                if report_type == "Subdebt" and mode == "Staff-wise":
                    group_column = [staff_column, branch_column] if include_branches else [staff_column]
                else:
                    group_column = branch_column if mode == "Branch-wise" else [staff_column, branch_column]

                old_group = old_df.groupby(group_column)[value_column].sum().reset_index()
                new_group = new_df.groupby(group_column)[value_column].sum().reset_index()

                merged = pd.merge(new_group, old_group, on=group_column, suffixes=('_New', '_Old'))
                merged["Growth"] = merged[f"{value_column}_New"] - merged[f"{value_column}_Old"]

                if report_type == "Subdebt" and mode == "Staff-wise" and "Canvasser Name" in new_df.columns:
                    merged = pd.merge(
                        merged,
                        new_df[[staff_column, "Canvasser Name"]].drop_duplicates(),
                        on=staff_column, how="left"
                    )

                # Column order
                col_order = []
                if staff_column in merged.columns: col_order.append(staff_column)
                if "Canvasser Name" in merged.columns: col_order.append("Canvasser Name")
                if branch_column in merged.columns and (include_branches or mode == "Branch-wise"):
                    col_order.append(branch_column)
                for col in merged.columns:
                    if col not in col_order: col_order.append(col)
                merged = merged[col_order].sort_values("Growth", ascending=False)

                st.session_state["merged_df"] = merged
                st.success("✅ Report generated successfully!")
                st.dataframe(merged)

                # Excel download
                output = io.BytesIO()
                merged.to_excel(output, index=False, sheet_name=f"{report_type} Report")
                st.download_button(
                    "⬇️ Download Excel Report",
                    data=output.getvalue(),
                    file_name=f"{report_type.lower()}_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Error processing files: {e}")

# ---------------- SS PENDING REPORT ----------------
if report_type == "SS Pending Report" and pending_file:
    show_customer_profile = st.checkbox("👤 Show Customer Profile (Pending Customers Only)")

    if st.button("▶️ Run Report"):
        try:
            df = read_file(pending_file)
            df.columns = df.columns.str.strip().str.upper()

            # 🔹 Remove Predator
            if "SCHEME NAME" in df.columns:
                df = df[df["SCHEME NAME"].str.strip().str.upper() != "RCIL PREDATOR 18%"]

            required_cols = ["BRANCH NAME", "DUE DAYS", "SCHEME NAME",
                             "PRINCIPAL OS", "INTEREST OS", "CUSTOMER NAME", "CUSTOMER ID"]
            missing_cols = [c for c in required_cols if c not in df.columns]
            if missing_cols:
                st.error(f"❌ Missing columns: {missing_cols}")
            else:
                # Allowed schemes filter
                allowed_schemes = [
                    "BIG SPL @20% KAR", "BIG SPL 20%", "BIG SPL 22%", "BUSINESS GOLD 12 MNTH SPL",
                    "RCIL SPL $24", "RCIL SPL $24 HYD", "RCIL SPL @24 KL-T", "RCIL SPL 2024(24%)",
                    "RCIL SPL 2025@24", "RCIL SPL- 22%", "RCIL SPL HT@24", "RCIL SPL HYD@24",
                    "RCIL SPL KAR @24", "RCIL SPL KAR@24", "RCIL SPL KL @24", "RCIL SPL KL@24",
                    "RCIL SPL KL@24-T", "RCIL SPL TAKEOVER @24", "RCIL SPL TAKEOVER 24%",
                    "RCIL SPL TAKEOVER@24", "RCIL SPL@ 20", "RCIL SPL@24", "RCIL SPL@24 KAR",
                    "RCIL SPL@24 KL", "RCIL SPL@24 OCT", "RCIL SPL@24 TAKEOVER",
                    "RCIL SPL@24 TAKEOVER KAR", "RCIL SPL24 KL", "RCIL SPL24 KL-T",
                    "RCIL SPL-5 KAR - 22%", "RCIL TAKEOVER SPL@24"
                ]
                allowed_schemes = [s.upper() for s in allowed_schemes]
                df = df[df["SCHEME NAME"].str.upper().isin(allowed_schemes)]

                if show_customer_profile:
                    # 🔹 Only pending customers, with selected columns
                    pending_customers = df[df["DUE DAYS"] > 30][
                        ["BRANCH NAME", "CUSTOMER NAME", "CUSTOMER ID",
                         "PRINCIPAL OS", "INTEREST OS", "DUE DAYS"]
                    ].copy()

                    st.success("✅ SS Pending Customer Profile generated successfully!")
                    st.dataframe(pending_customers, use_container_width=True)

                    output = io.BytesIO()
                    pending_customers.to_excel(output, index=False, sheet_name="SS Pending Customers")
                    st.download_button(
                        "⬇️ Download SS Pending Customers",
                        data=output.getvalue(),
                        file_name="ss_pending_customers.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                else:
                    # 🔹 Branch summary (default view)
                    grouped = df.groupby("BRANCH NAME")
                    report = []
                    for branch, data in grouped:
                        total_count = len(data)
                        total_amount = data["PRINCIPAL OS"].sum()
                        pending = data[data["DUE DAYS"] > 30]
                        pending_count = len(pending)
                        pending_amount = pending["PRINCIPAL OS"].sum()
                        pending_interest = pending["INTEREST OS"].sum()
                        pending_pct = (pending_count / total_count * 100) if total_count > 0 else 0
                        report.append({
                            "BRANCH NAME": branch,
                            "Total_Count": total_count,
                            "Total_Amount": round(total_amount, 2),
                            "Pending_Count": pending_count,
                            "Pending_Amount": round(pending_amount, 2),
                            "Pending_Interest": round(pending_interest, 2),
                            "Pending %": f"{int(round(pending_pct, 0))}%"
                        })

                    final = pd.DataFrame(report)
                    st.session_state["merged_df"] = final
                    st.success("✅ SS Pending Report generated successfully!")
                    st.dataframe(final, use_container_width=True)

                    output = io.BytesIO()
                    final.to_excel(output, index=False, sheet_name="SS Pending Report")
                    st.download_button(
                        "⬇️ Download SS Pending Report",
                        data=output.getvalue(),
                        file_name="ss_pending_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"❌ Error processing SS Pending Report: {e}")

# ---------------- NPA + MATURITY REPORT ----------------
if report_type == "NPA" and uploaded:
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)

    required_cols = [
        "BRANCH NAME", "STATE", "NEW ACCOUNT NO", "CUSTOMER NAME", "CUSTOMER ID",
        "SCHEME NAME", "LOAN PURPOSE", "SANCTIONED DATE",
        "PRINCIPAL OS", "INTEREST OS", "MATURITY DATE", "TENURE OF THE LOAN"
    ]
    df = df[[col for col in required_cols if col in df.columns]]

    if "SCHEME NAME" in df.columns:
        df = df[df["SCHEME NAME"].str.strip().str.upper() != "RCIL PREDATOR 18%"]

    # Calculate CR Maturity
    special_schemes = ["BUSINESS GOLD 12 MNTH SPL","INTEREST SAVER -6%","OUTSIDE SWEEPER - 20",
                       "RELIANT GRABBER 11.8%","BUSINESS GOLD NEW-12"]
    if "SANCTIONED DATE" in df.columns and "TENURE OF THE LOAN" in df.columns:
        def calc_cr(x):
            scheme = str(x["SCHEME NAME"]).strip().upper()
            if scheme in [s.upper() for s in special_schemes]:
                return pd.to_datetime(x["MATURITY DATE"], dayfirst=True, errors="coerce").strftime("%d-%m-%Y") \
                    if pd.notnull(x["MATURITY DATE"]) else None
            else:
                if pd.notnull(pd.to_datetime(x["SANCTIONED DATE"], dayfirst=True, errors="coerce")) and pd.notnull(x["TENURE OF THE LOAN"]):
                    return (pd.to_datetime(x["SANCTIONED DATE"], dayfirst=True, errors="coerce") + pd.Timedelta(days=int(x["TENURE OF THE LOAN"]))).strftime("%d-%m-%Y")
                return None
        df["CR_MATURITY"] = df.apply(calc_cr, axis=1)

    current_date = st.date_input("📅 Select Current Date", datetime.today().date())
    df["CURRENT_DATE"] = pd.to_datetime(current_date, dayfirst=True).strftime("%d-%m-%Y")
    as_on_maturity = st.date_input("📅 Select As On Maturity Date", datetime.today().date())

    if "CR_MATURITY" in df.columns:
        df["METURITY"] = (
            pd.to_datetime(df["CURRENT_DATE"], format="%d-%m-%Y", dayfirst=True, errors="coerce")
            - pd.to_datetime(df["CR_MATURITY"], format="%d-%m-%Y", dayfirst=True, errors="coerce")
        ).dt.days

    st.session_state["processed_df"] = df.copy()

    # --- Maturity ---
    if st.button("▶️ Run Maturity Report"):
        maturity_df = df[pd.to_datetime(df["CR_MATURITY"], dayfirst=True, errors="coerce") <= pd.to_datetime(as_on_maturity, dayfirst=True)].copy()
        maturity_df.rename(columns={"METURITY": "Maturity"}, inplace=True)
        st.session_state["maturity_df"] = maturity_df

        st.subheader("📄 Maturity Report")
        st.dataframe(maturity_df, use_container_width=True)

        output = io.BytesIO()
        maturity_df.to_excel(output, index=False, sheet_name="Maturity Report")
        st.download_button(
            "⬇️ Download Maturity Report",
            data=output.getvalue(),
            file_name="maturity_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        consolidated = maturity_df.groupby("BRANCH NAME").size().reset_index(name="Maturity Count")
        st.subheader("📊 Maturity Consolidated Report")
        st.dataframe(consolidated, use_container_width=True)

        output_cons = io.BytesIO()
        consolidated.to_excel(output_cons, index=False, sheet_name="Maturity Consolidated")
        st.download_button(
            "⬇️ Download Maturity Consolidated",
            data=output_cons.getvalue(),
            file_name="maturity_consolidated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- NPA ---
    if st.button("⚠️ Run NPA Report"):
        if "maturity_df" not in st.session_state:
            st.warning("⚠️ Please run Maturity Report first!")
        else:
            npa_df = st.session_state["maturity_df"].copy()
            npa_df = npa_df[npa_df["Maturity"] > 90].rename(columns={"Maturity": "NPA"})
            st.session_state["npa_df"] = npa_df

            st.subheader("⚠️ NPA Report (Overdue > 90 Days)")
            st.dataframe(npa_df, use_container_width=True)

            output = io.BytesIO()
            npa_df.to_excel(output, index=False, sheet_name="NPA Report")
            st.download_button(
                "⬇️ Download NPA Report",
                data=output.getvalue(),
                file_name="npa_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ---------------- GOOGLE SHEET UPLOAD PANEL ----------------
if "merged_df" in st.session_state:
    merged_df = st.session_state["merged_df"]

    # CSV backup download
    csv_data = merged_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Download CSV Backup",
        data=csv_data,
        file_name=f"{report_type}_{mode if mode else ''}_Report.csv",
        mime="text/csv"
    )

    # Map correct sheet names
    if report_type == "Gold" and mode == "Branch-wise":
        sheet_name = "BRANCH_GL"
    elif report_type == "Gold" and mode == "Staff-wise":
        sheet_name = "STAFF_GL"
    elif report_type == "Subdebt":
        if mode == "Branch-wise":
            sheet_name = "BRANCH_SD"
        elif mode == "Staff-wise":
            # ✅ Fix: if staff-wise includes branches → send to BRANCH_SD
            if "Branch Name" in merged_df.columns:
                sheet_name = "BRANCH_SD"
            else:
                sheet_name = "STAFF_SD"
        else:
            sheet_name = "STAFF_SD"
    elif report_type == "SS Pending Report":
        sheet_name = "SS_PENDING"
    elif report_type == "NPA":
        sheet_name = "NPA_REPORT"
    else:
        sheet_name = "UNKNOWN"

    with st.expander("🔐 Admin Upload to Google Sheet"):
        password_input = st.text_input("Enter Admin Password", type="password")
        if st.button("🔗 Connect to Google Sheet"):
            if password_input == ADMIN_PASSWORD:
                with st.spinner(f"🔄 Uploading report to {sheet_name}... Please wait"):
                    if upload_dataframe_to_specific_tab(merged_df, sheet_name):
                        st.success(f"✅ Report uploaded to Google Sheet tab: {sheet_name}")
            else:
                st.error("❌ Incorrect password. Access denied.")
else:
    st.info("📎 Please upload and run a report before connecting to Google Sheets.")
