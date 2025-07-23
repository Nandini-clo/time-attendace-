import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="Attendance Entry", layout="wide")
st.title("📋 Employee Attendance Sheet Generator")

uploaded_file = st.file_uploader("📤 Upload Excel file (with 'Post Code' & 'Employee Name')", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]

        if 'Post Code' not in df.columns or 'Employee Name' not in df.columns:
            st.error("❌ Excel file must contain 'Post Code' and 'Employee Name' columns.")
        else:
            month = st.selectbox("📅 Select Month", list(range(1, 13)), index=datetime.now().month - 1)
            year = st.selectbox("📆 Select Year", list(range(2020, 2031)), index=5)

            employee_list = df[['Pos Code', 'Employee Name']].drop_duplicates().reset_index(drop=True)

            if 'current_index' not in st.session_state:
                st.session_state.current_index = 0
            if 'final_data_dict' not in st.session_state:
                st.session_state.final_data_dict = {}

            total_employees = len(employee_list)

            if st.session_state.current_index < total_employees:
                current = employee_list.iloc[st.session_state.current_index]
                st.subheader(f"🧑‍💼 {current['Employee Name']} (Post Code: {current['Post Code']})")

                days_in_month = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day

                row_data = {
                    'Post Code': current['Post Code'],
                    'Employee Name': current['Employee Name']
                }
                total_ot_hours = 0
                count_P = count_A = count_L = count_WO = count_HL = count_PH = 0

                st.markdown("### Select Attendance for Each Day")

                for day in range(1, days_in_month + 1):
                    date_str = f"{day:02d}-{month:02d}"
                    st.markdown(f"#### 📅 {date_str}")
                    att_col, = st.columns(1)
                    status = att_col.selectbox(
                        f"Attendance Type ({date_str})", 
                        ["P", "A", "L", "WO", "HL", "PH"], 
                        key=f"status_{day}"
                    )

                    if status in ["P", "PH"]:
                        if status == "P":
                            count_P += 1
                        else:
                            count_PH += 1

                        col1, col2 = st.columns(2)
                        with col1:
                            ci_str = st.text_input(f"Check-in ({date_str})", value="09:00", key=f"ci_txt_{day}")
                            try:
                                ci = datetime.strptime(ci_str, "%H:%M").time()
                            except:
                                st.warning("⏰ Invalid check-in format. Using default 09:00")
                                ci = datetime.strptime("09:00", "%H:%M").time()
                        with col2:
                            co_str = st.text_input(f"Check-out ({date_str})", value="18:00", key=f"co_txt_{day}")
                            try:
                                co = datetime.strptime(co_str, "%H:%M").time()
                            except:
                                st.warning("⏰ Invalid check-out format. Using default 18:00")
                                co = datetime.strptime("18:00", "%H:%M").time()

                        date = datetime(year, month, day)
                        check_in_dt = datetime.combine(date, ci)
                        check_out_dt = datetime.combine(date, co)
                        if check_out_dt <= check_in_dt:
                            check_out_dt += timedelta(days=1)

                        hours = round((check_out_dt - check_in_dt).total_seconds() / 3600, 2)
                        ot = 0 if status == "PH" else round(max(0, hours - 8), 2)

                    elif status == "A":
                        count_A += 1
                        ci = co = datetime.strptime("00:00", "%H:%M").time()
                        hours = ot = 0

                    elif status == "L":
                        count_L += 1
                        ci = co = datetime.strptime("00:00", "%H:%M").time()
                        hours = ot = 0

                    elif status == "WO":
                        count_WO += 1
                        ci = datetime.strptime("09:00", "%H:%M").time()
                        co = datetime.strptime("17:00", "%H:%M").time()
                        hours = 8
                        ot = 0

                    elif status == "HL":
                        count_HL += 1
                        ci = datetime.strptime("09:00", "%H:%M").time()
                        co = datetime.strptime("13:00", "%H:%M").time()
                        hours = 4
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

                preview_df = pd.DataFrame([row_data])
                st.markdown("### 🧾 Preview:")
                st.dataframe(preview_df, use_container_width=True)

                # Navigation Buttons
                col_prev, col_next = st.columns([1, 1])
                with col_prev:
                    if st.button("⏮ Previous"):
                        if st.session_state.current_index > 0:
                            st.session_state.final_data_dict[st.session_state.current_index] = row_data
                            st.session_state.current_index -= 1

                with col_next:
                    if st.button("💾 Save & Next"):
                        st.session_state.final_data_dict[st.session_state.current_index] = row_data
                        st.session_state.current_index += 1

            else:
                st.success("✅ All employee data entered!")

                final_df = pd.DataFrame([data for idx, data in sorted(st.session_state.final_data_dict.items())])
                st.dataframe(final_df, use_container_width=True)

                towrite = io.BytesIO()
                with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Attendance')
                st.download_button("📥 Download Final Excel", data=towrite.getvalue(), file_name="attendance_sheet.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
