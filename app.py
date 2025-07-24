import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import os
import json

BACKUP_FILE = "attendance_backup.json"

# Ensure backup file exists
if not os.path.exists(BACKUP_FILE):
    with open(BACKUP_FILE, "w") as f:
        json.dump({"final_data_dict": {}, "current_index": 0}, f)

# Load backup
with open(BACKUP_FILE, "r") as f:
    try:
        backup_data = json.load(f)
        if 'final_data_dict' not in st.session_state:
            st.session_state.final_data_dict = backup_data.get("final_data_dict", {})
        if 'current_index' not in st.session_state:
            st.session_state.current_index = backup_data.get("current_index", 0)
    except Exception as e:
        st.warning(f"⚠️ Failed to load backup: {e}")

# Auto-save function
def save_backup():
    with open(BACKUP_FILE, "w") as f:
        json.dump({
            "final_data_dict": st.session_state.final_data_dict,
            "current_index": st.session_state.current_index
        }, f)

st.set_page_config(page_title="Attendance Entry", layout="wide")
st.title(u"\U0001F4CB Employee Attendance Sheet Generator")

uploaded_file = st.file_uploader("📄 Upload Excel file (with 'Employee Code' & 'Employee Name')", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]

        if 'Employee Code' not in df.columns or 'Employee Name' not in df.columns:
            st.error("❌ Excel file must contain 'Employee Code' and 'Employee Name' columns.")
        else:
            month = st.selectbox("🗓️ Select Month", list(range(1, 13)), index=datetime.now().month - 1)
            year = st.selectbox("📆 Select Year", list(range(2020, 2031)), index=5)

            employee_list = df[['Employee Code', 'Employee Name']].drop_duplicates().reset_index(drop=True)
            total_employees = len(employee_list)

            current_index = st.session_state.current_index
            if current_index < total_employees:
                current = employee_list.iloc[current_index]
                st.subheader(f"🧑‍💼 {current['Employee Name']} (Employee Code: {current['Employee Code']})")

                days_in_month = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day
                saved_data = st.session_state.final_data_dict.get(str(current_index), {})

                row_data = {
                    'Employee Code': current['Employee Code'],
                    'Employee Name': current['Employee Name']
                }

                total_ot_hours = 0
                count_P = count_A = count_L = count_WO = count_HL = count_PH = 0
                st.markdown("### Select Attendance for Each Day")

                for day in range(1, days_in_month + 1):
                    date_str = f"{day:02d}-{month:02d}"
                    st.markdown(f"#### 📅 {date_str}")
                    att_col, = st.columns(1)

                    saved_status = saved_data.get(f'{day:02d}_Status', "P")
                    status = att_col.selectbox(
                        f"Attendance Type ({date_str})",
                        ["P", "A", "L", "WO", "HL", "PH"],
                        key=f"status_{day}_{current_index}",
                        index=["P", "A", "L", "WO", "HL", "PH"].index(saved_status)
                    )

                    if status in ["P", "PH"]:
                        if status == "P":
                            count_P += 1
                        else:
                            count_PH += 1

                        col1, col2 = st.columns(2)
                        default_ci = saved_data.get(f'{day:02d}_Check-in', "09:00")
                        default_co = saved_data.get(f'{day:02d}_Check-out', "18:00")

                        with col1:
                            ci_str = st.text_input(f"Check-in ({date_str})", value=default_ci, key=f"ci_txt_{day}_{current_index}")
                            try:
                                ci = datetime.strptime(ci_str, "%H:%M").time()
                            except:
                                ci = datetime.strptime("09:00", "%H:%M").time()
                                st.warning("⏰ Invalid check-in")

                        with col2:
                            co_str = st.text_input(f"Check-out ({date_str})", value=default_co, key=f"co_txt_{day}_{current_index}")
                            try:
                                co = datetime.strptime(co_str, "%H:%M").time()
                            except:
                                co = datetime.strptime("18:00", "%H:%M").time()
                                st.warning("⏰ Invalid check-out")

                        check_in_dt = datetime.combine(datetime(year, month, day), ci)
                        check_out_dt = datetime.combine(datetime(year, month, day), co)
                        if check_out_dt <= check_in_dt:
                            check_out_dt += timedelta(days=1)

                        hours = round((check_out_dt - check_in_dt).total_seconds() / 3600, 2)
                        ot = 0 if status == "PH" else round(max(0, hours - 8), 2)

                    elif status == "A":
                        count_A += 1
                        ci = co = datetime.strptime("00:00", "%H:%M").time()
                        ot = 0

                    elif status == "L":
                        count_L += 1
                        ci = co = datetime.strptime("00:00", "%H:%M").time()
                        ot = 0

                    elif status == "WO":
                        count_WO += 1
                        ci = datetime.strptime("09:00", "%H:%M").time()
                        co = datetime.strptime("17:00", "%H:%M").time()
                        ot = 0

                    elif status == "HL":
                        count_HL += 1
                        ci = datetime.strptime("09:00", "%H:%M").time()
                        co = datetime.strptime("13:00", "%H:%M").time()
                        ot = 0

                    total_ot_hours += ot
                    row_data[f'{day:02d}_Check-in'] = ci.strftime("%H:%M")
                    row_data[f'{day:02d}_Check-out'] = co.strftime("%H:%M")
                    row_data[f'{day:02d}_Status'] = status
                    row_data[f'{day:02d}_OT'] = ot

                row_data["Total P"] = count_P
                row_data["Total A"] = count_A
                row_data["Total L"] = count_L
                row_data["Total WO"] = count_WO
                row_data["Total HL"] = count_HL
                row_data["Total PH"] = count_PH
                row_data["Total Attendance"] = count_P + count_HL + count_L + count_PH
                row_data["OT Hours"] = round(total_ot_hours, 2)

                # Save user edits and backup
                st.session_state.final_data_dict[str(current_index)] = row_data
                save_backup()

                st.markdown("### 🧾 Preview:")
                st.dataframe(pd.DataFrame([row_data]), use_container_width=True)

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("⏮ Previous"):
                        if current_index > 0:
                            st.session_state.current_index -= 1
                            save_backup()
                            st.rerun()

                with col2:
                    if st.button("💾 Save & Next"):
                        if current_index < total_employees - 1:
                            st.session_state.current_index += 1
                            save_backup()
                            st.rerun()
                        else:
                            # Last employee saved — show download!
                            st.success("✅ All employee data entered!")
                            final_df = pd.DataFrame([
                                v for k, v in sorted(st.session_state.final_data_dict.items(), key=lambda x: int(x[0]))
                            ])
                            st.dataframe(final_df, use_container_width=True)
                            towrite = io.BytesIO()
                            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                                final_df.to_excel(writer, index=False, sheet_name='Attendance')
                            st.download_button("📥 Download Final Excel", data=towrite.getvalue(), file_name="attendance_sheet.xlsx")

            elif current_index >= total_employees:
                # Show download if rerun happened and all employees are already done
                st.success("✅ All employee data entered!")
                final_df = pd.DataFrame([
                    v for k, v in sorted(st.session_state.final_data_dict.items(), key=lambda x: int(x[0]))
                ])
                st.dataframe(final_df, use_container_width=True)
                towrite = io.BytesIO()
                with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Attendance')
                st.download_button("📥 Download Final Excel", data=towrite.getvalue(), file_name="attendance_sheet.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")

