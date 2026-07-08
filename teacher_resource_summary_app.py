import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
import re
import os
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="Teacher Resource Summary",
    layout="centered"
)

st.title("📊 Teacher Resource Summary Tool")

uploaded_files = st.file_uploader(
    "Upload resource files (Excel, CSV, XML)",
    type=["xls", "xlsx", "csv", "xml"],
    accept_multiple_files=True
)

# Resource columns
resource_columns = [
    "Lesson",
    "Quiz",
    "Assignment",
    "Shared Blog",
    "Survey",
    "Forum",
    "Announcement",
    "Message",
    "Flashcard"
]


# Function to clean names
def clean_name(name):
    if pd.isna(name):
        return ""

    name = str(name).strip()
    name = re.sub(r"\s+", " ", name)

    return name


# Function to read files
def read_file(uploaded_file):

    file_name = uploaded_file.name.lower()

    try:

        if file_name.endswith(".csv"):
            return pd.read_csv(uploaded_file)


        elif file_name.endswith(".xlsx"):
            return pd.read_excel(
                uploaded_file,
                engine="openpyxl"
            )


        elif file_name.endswith(".xls"):

            try:
                return pd.read_excel(
                    uploaded_file,
                    engine="xlrd"
                )

            except Exception:

                uploaded_file.seek(0)

                # Handles Excel XML saved as .xls
                return pd.read_xml(
                    uploaded_file
                )


        elif file_name.endswith(".xml"):

            return pd.read_xml(
                uploaded_file
            )


    except Exception as e:

        st.error(
            f"Cannot read {uploaded_file.name}: {e}"
        )

        return pd.DataFrame()


    return pd.DataFrame()


if uploaded_files:

    all_data = []

    for file in uploaded_files:

        df = read_file(file)

        if not df.empty:

            df.columns = [
                str(col).strip()
                for col in df.columns
            ]

            all_data.append(df)


    if all_data:

        data = pd.concat(
            all_data,
            ignore_index=True
        )


        st.success(
            f"{len(uploaded_files)} file(s) uploaded successfully"
        )


        st.subheader("Preview Data")

        st.dataframe(data.head())


        # Continue to Part 2...
                # Detect important columns

        teacher_column = None
        subject_column = None
        date_column = None


        for col in data.columns:

            col_lower = str(col).lower()

           if any(keyword in col_lower for keyword in [
                "teacher",
                "created by",
                "createdby"
       ]):
              teacher_column = col
            if "subject" in col_lower:
                subject_column = col

            if "date" in col_lower or "created" in col_lower:
                date_column = col



        if teacher_column:

            data[teacher_column] = (
                data[teacher_column]
                .apply(clean_name)
            )


            # Subject filter and sorting

            if subject_column:

                st.sidebar.header("🔎 Subject Filter")

                subjects = sorted(
                    data[subject_column]
                    .dropna()
                    .unique()
                    .tolist()
                )


                selected_subject = st.sidebar.multiselect(
                    "Select Subject",
                    subjects
                )


                if selected_subject:

                    data = data[
                        data[subject_column]
                        .isin(selected_subject)
                    ]


            # Date filter

            if date_column:

                data[date_column] = pd.to_datetime(
                    data[date_column],
                    errors="coerce"
                )


                st.sidebar.header("📅 Date Filter")


                min_date = (
                    data[date_column]
                    .min()
                )

                max_date = (
                    data[date_column]
                    .max()
                )


                if pd.notna(min_date) and pd.notna(max_date):

                    date_range = st.sidebar.date_input(
                        "Select Date Range",
                        value=(
                            min_date,
                            max_date
                        )
                    )


                    if len(date_range) == 2:

                        start_date, end_date = date_range


                        data = data[
                            (
                                data[date_column]
                                >= pd.Timestamp(start_date)
                            )
                            &
                            (
                                data[date_column]
                                <= pd.Timestamp(end_date)
                            )
                        ]



            # Teacher summary

            summary = (
                data
                .groupby(teacher_column)
                .size()
                .reset_index(name="Total Resources")
            )


            summary = summary.sort_values(
                by="Total Resources",
                ascending=False
            )


            st.subheader("👩‍🏫 Teacher Resource Summary")

            st.dataframe(
                summary,
                use_container_width=True
            )


            # Subject Summary

            if subject_column:

                st.subheader("📚 Subject Resource Summary")


                subject_summary = (
                    data
                    .groupby(subject_column)
                    .size()
                    .reset_index(
                        name="Total Resources"
                    )
                    .sort_values(
                        by="Total Resources",
                        ascending=False
                    )
                )


                st.dataframe(
                    subject_summary,
                    use_container_width=True
                )



            # Download Excel

            output = io.BytesIO()


            with pd.ExcelWriter(
                output,
                engine="xlsxwriter"
            ) as writer:

                summary.to_excel(
                    writer,
                    index=False,
                    sheet_name="Teacher Summary"
                )


                if subject_column:

                    subject_summary.to_excel(
                        writer,
                        index=False,
                        sheet_name="Subject Summary"
                    )


            st.download_button(
                label="⬇️ Download Summary Excel",
                data=output.getvalue(),
                file_name="Teacher_Resource_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )



        else:

            st.warning(
                "Teacher Name column was not detected."
            )
