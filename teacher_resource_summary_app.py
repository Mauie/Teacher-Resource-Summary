import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
import re
import os
import hashlib
import matplotlib.pyplot as plt

# --- MAIN APP ---
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
                        continue

                    df = pd.DataFrame(data[1:], columns=data[0])
                else:
                    continue
            elif file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            df.columns = df.columns.str.strip()

            if "Teacher Name" not in df.columns:
                if "Created By" in df.columns:
                    df.rename(columns={"Created By": "Teacher Name"}, inplace=True)
                elif df.shape[1] == 1:
                    df.columns = ["Teacher Name"]
                else:
                    continue

            if "Created Date" in df.columns:
                df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")

            df["Resource Type"] = os.path.splitext(file.name)[0].strip().capitalize()

            raw_data.append(df)

        except Exception as e:
            st.error(f"❌ Error processing {file.name}: {e}")

# Filtering and Summary
if raw_data:
    combined_df = pd.concat(raw_data, ignore_index=True)

    # Filters
    st.subheader("🔎 Filter Options")

    teacher_names = combined_df["Teacher Name"].dropna().unique().tolist()
    teacher_names.sort()
    all_option = "All (Select All)"
    teacher_names_with_all = [all_option] + teacher_names
    selected_teachers = st.multiselect("👤 Filter by Teacher Name", teacher_names_with_all, default=all_option)

    if all_option in selected_teachers:
        selected_teachers = teacher_names

    filtered_df = combined_df[combined_df["Teacher Name"].isin(selected_teachers)]

    if "Created Date" in filtered_df.columns:
        min_date = filtered_df["Created Date"].min()
        max_date = filtered_df["Created Date"].max()
        date_range = st.date_input("📅 Filter by Created Date Range", [min_date, max_date])
        start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        filtered_df = filtered_df[
            (filtered_df["Created Date"] >= start_date) & (filtered_df["Created Date"] <= end_date)
        ]

    if filtered_df.empty:
        st.warning("⚠️ No data after filtering.")
    else:
        # Summary by Teacher and Resource Type
        summary = filtered_df.groupby(["Teacher Name", "Resource Type"]).size().unstack(fill_value=0).reset_index()
        summary = summary.sort_values("Teacher Name").reset_index(drop=True)

        # Add total per teacher
        summary["Total"] = summary.iloc[:, 1:].sum(axis=1)

        # Add total row
        total_row = ["Total"] + summary.iloc[:, 1:].sum(numeric_only=True).tolist()
        summary.loc[len(summary)] = total_row

        # Insert row numbers
        summary.insert(0, "No.", list(range(1, len(summary))) + [""])

        st.subheader("📋 Filtered Resource Summary")
        st.dataframe(summary, use_container_width=True)

        # 📊 Bar Chart
        st.subheader("📊 Bar Chart")
        chart_data = summary.iloc[:-1].set_index("Teacher Name").drop(columns=["No.", "Total"], errors="ignore")
        st.bar_chart(chart_data)

        # 🥧 Pie Chart
        st.subheader("🥧 Pie Chart")
        pie_data = summary.iloc[:-1].set_index("Teacher Name")["Total"]
        fig, ax = plt.subplots()
        ax.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%', startangle=90)
        ax.axis("equal")
        st.pyplot(fig)

        # Downloadable Excel with chart
        chart_img = io.BytesIO()
        fig.savefig(chart_img, format='png')
        chart_img.seek(0)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            chart_data.to_excel(writer, sheet_name="Chart Data")
            workbook = writer.book
            worksheet = writer.sheets["Summary"]
            worksheet.insert_image("H2", "chart.png", {"image_data": chart_img})

        st.download_button(
            "⬇️ Download Filtered Excel with Chart",
            data=output.getvalue(),
            file_name="teacher_resource_summary_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload one or more resource files.")
