import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
import re
import os
import matplotlib.pyplot as plt

st.set_page_config(page_title="Teacher Resource Summary", layout="centered")
st.title("📊 Teacher Resource Summary Tool")

uploaded_files = st.file_uploader(
    "Upload resource files (Excel, CSV, XML)",
    type=["xls", "xlsx", "csv", "xml"],
    accept_multiple_files=True,
)

raw_data = []

if uploaded_files:
    for file in uploaded_files:
        try:
            file_head = file.read(2048).decode("utf-8", errors="ignore").lower()
            file.seek(0)

            is_xml = "<?xml" in file_head and (
                "<workbook" in file_head or "urn:schemas-microsoft-com:office:spreadsheet" in file_head
            )

            if is_xml:
                tree = ET.parse(file)
                root = tree.getroot()
                ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
                table = root.find(".//ss:Table", ns)

                if table is not None:
                    rows = table.findall("ss:Row", ns)
                    data = []
                    for row in rows:
                        values = []
                        for cell in row.findall("ss:Cell", ns):
                            data_elem = cell.find("ss:Data", ns)
                            values.append(data_elem.text.strip() if data_elem is not None and data_elem.text else "")
                        data.append(values)

                    if len(data) < 2:
                        st.warning(f"⚠️ Skipped {file.name} due to insufficient rows.")
                        continue

                    df = pd.DataFrame(data[1:], columns=data[0])
                else:
                    st.warning(f"⚠️ Skipped {file.name} due to missing XML table.")
                    continue
            elif file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            df.columns = df.columns.str.strip()

            if "Title" not in df.columns:
                st.warning(f"⚠️ Skipped {file.name} due to missing Title column.")
                continue

            if "Subject" not in df.columns:
                st.warning(f"⚠️ Skipped {file.name} due to missing Subject column.")
                continue

            if "Teacher Name" not in df.columns:
                if "Created By" in df.columns:
                    df.rename(columns={"Created By": "Teacher Name"}, inplace=True)
                else:
                    st.warning(f"⚠️ Skipped {file.name} due to unknown teacher column.")
                    continue

            df["Teacher Name"] = df["Teacher Name"].fillna("").astype(str).str.strip()
            df = df[df["Teacher Name"] != ""]

            if "Created Date" in df.columns:
                df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")

            raw_data.append(df)

        except Exception as e:
            st.error(f"❌ Error processing {file.name}: {e}")

if raw_data:
    combined_df = pd.concat(raw_data, ignore_index=True)

    st.subheader("🔎 Filter Options")

    # Filter by Subject
    subjects = combined_df["Subject"].dropna().unique().tolist()
    subjects.sort()
    all_option = "All (Select All)"
    subjects_with_all = [all_option] + subjects
    selected_subjects = st.multiselect("📚 Filter by Subject", subjects_with_all, default=all_option)

    if all_option in selected_subjects:
        selected_subjects = subjects

    filtered_df = combined_df[combined_df["Subject"].isin(selected_subjects)]

    # Filter by Teacher
    teacher_names = filtered_df["Teacher Name"].dropna().unique().tolist()
    teacher_names.sort()
    teacher_names_with_all = [all_option] + teacher_names
    selected_teachers = st.multiselect("👤 Filter by Teacher Name", teacher_names_with_all, default=all_option)

    if all_option in selected_teachers:
        selected_teachers = teacher_names

    filtered_df = filtered_df[filtered_df["Teacher Name"].isin(selected_teachers)]

    # Filter by Date
    if "Created Date" in filtered_df.columns and not filtered_df["Created Date"].isna().all():
        min_date = filtered_df["Created Date"].min().date()
        max_date = filtered_df["Created Date"].max().date()
        date_range = st.date_input("📅 Filter by Created Date Range", [min_date, max_date])
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

        filtered_df = filtered_df[
            (filtered_df["Created Date"] >= start_date) & (filtered_df["Created Date"] <= end_date)
        ]

    if filtered_df.empty:
        st.warning("⚠️ No data after filtering.")
    else:
        # Display filtered data table
        st.subheader(f"📋 All Resources ({len(filtered_df)} total)")
        display_cols = ["Title", "Subject", "Teacher Name", "Created Date"]
        if "Total Access" in filtered_df.columns:
            display_cols.append("Total Access")
        
        st.dataframe(filtered_df[display_cols], use_container_width=True, hide_index=True)

        # Teacher Summary
        st.subheader("👩‍🏫 Teacher Resource Summary")
        teacher_summary = filtered_df.groupby("Teacher Name").size().reset_index(name="Total Resources")
        teacher_summary = teacher_summary.sort_values("Total Resources", ascending=False)
        st.dataframe(teacher_summary, use_container_width=True, hide_index=True)

        # Subject Summary
        st.subheader("📚 Subject Resource Summary")
        subject_summary = filtered_df.groupby("Subject").size().reset_index(name="Total Resources")
        subject_summary = subject_summary.sort_values("Total Resources", ascending=False)
        st.dataframe(subject_summary, use_container_width=True, hide_index=True)

        # Combined Summary by Teacher and Subject
        st.subheader("📊 Detailed Summary")
        summary = filtered_df.groupby(["Teacher Name", "Subject"]).size().unstack(fill_value=0).reset_index()
        summary = summary.sort_values("Teacher Name").reset_index(drop=True)
        summary["Total"] = summary.iloc[:, 1:].sum(axis=1)
        total_row = ["Total"] + summary.iloc[:, 1:].sum(numeric_only=True).tolist()
        summary.loc[len(summary)] = total_row

        st.dataframe(summary, use_container_width=True, hide_index=True)

        # Bar Chart
        st.subheader("📊 Bar Chart - Resources by Teacher")
        chart_data = summary.iloc[:-1].set_index("Teacher Name").drop(columns=["Total"], errors="ignore")
        if not chart_data.empty:
            st.bar_chart(chart_data)

        # Pie Chart
        st.subheader("🥧 Pie Chart - Total Resources by Teacher")
        pie_data = summary.iloc[:-1].set_index("Teacher Name")["Total"]
        if not pie_data.empty:
            fig, ax = plt.subplots()
            ax.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=90)
            ax.axis("equal")
            st.pyplot(fig)

            chart_img = io.BytesIO()
            fig.savefig(chart_img, format='png')
            chart_img.seek(0)
        else:
            chart_img = None

        # Download Excel
        st.subheader("⬇️ Download Results")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # All resources
            filtered_df[display_cols].to_excel(writer, index=False, sheet_name="All Resources")
            
            # Teacher summary
            teacher_summary.to_excel(writer, index=False, sheet_name="Teacher Summary")
            
            # Subject summary
            subject_summary.to_excel(writer, index=False, sheet_name="Subject Summary")
            
            # Combined summary
            summary.to_excel(writer, index=False, sheet_name="Detailed Summary")

        st.download_button(
            "⬇️ Download Filtered Excel",
            data=output.getvalue(),
            file_name="teacher_resource_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload one or more resource files.")
