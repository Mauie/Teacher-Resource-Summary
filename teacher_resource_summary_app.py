import streamlit as st
import pandas as pd
import io
from io import BytesIO

st.set_page_config(
    page_title="Teacher Resource Summary",
    layout="wide"
)

st.title("📊 Teacher Resource Summary Tool")

uploaded_files = st.file_uploader(
    "Upload resource files (Excel, CSV, XML)",
    type=["xls", "xlsx", "csv", "xml"],
    accept_multiple_files=True
)

def read_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    
    try:
        if file_name.endswith(".csv"):
            return pd.read_csv(uploaded_file)
        elif file_name.endswith(".xlsx"):
            return pd.read_excel(uploaded_file, engine="openpyxl")
        elif file_name.endswith(".xls"):
            try:
                return pd.read_excel(uploaded_file, engine="xlrd")
            except Exception:
                try:
                    uploaded_file.seek(0)
                    return pd.read_excel(uploaded_file, engine="openpyxl")
                except Exception:
                    uploaded_file.seek(0)
                    return pd.read_xml(uploaded_file)
        elif file_name.endswith(".xml"):
            return pd.read_xml(uploaded_file)
    except Exception as e:
        st.error(f"Cannot read {uploaded_file.name}: {e}")
        return pd.DataFrame()
    
    return pd.DataFrame()

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        df = read_file(file)
        if not df.empty:
            df.columns = [str(col).strip() for col in df.columns]
            all_data.append(df)
    
    if all_data:
        data = pd.concat(all_data, ignore_index=True)
        
        st.success(f"{len(uploaded_files)} file(s) uploaded successfully")
        
        # Detect key columns
        title_col = None
        subject_col = None
        created_by_col = None
        created_date_col = None
        access_col = None
        
        for col in data.columns:
            col_lower = col.lower().strip()
            if col_lower == "title":
                title_col = col
            elif col_lower == "subject":
                subject_col = col
            elif col_lower == "created by":
                created_by_col = col
            elif col_lower == "created date":
                created_date_col = col
            elif "total access" in col_lower:
                access_col = col
        
        # Sidebar filters
        st.sidebar.header("🔍 Filters")
        
        filtered_data = data.copy()
        
        # Subject filter
        if subject_col:
            subjects = sorted(filtered_data[subject_col].dropna().unique().tolist())
            selected_subjects = st.sidebar.multiselect(
                "Select Subject(s)",
                subjects,
                default=subjects
            )
            if selected_subjects:
                filtered_data = filtered_data[filtered_data[subject_col].isin(selected_subjects)]
        
        # Teacher filter
        if created_by_col:
            teachers = sorted(filtered_data[created_by_col].dropna().unique().tolist())
            selected_teachers = st.sidebar.multiselect(
                "Select Teacher(s)",
                teachers,
                default=teachers
            )
            if selected_teachers:
                filtered_data = filtered_data[filtered_data[created_by_col].isin(selected_teachers)]
        
        # Date filter
        if created_date_col:
            try:
                filtered_data[created_date_col] = pd.to_datetime(filtered_data[created_date_col], errors="coerce")
                min_date = filtered_data[created_date_col].min()
                max_date = filtered_data[created_date_col].max()
                
                if pd.notna(min_date) and pd.notna(max_date):
                    date_range = st.sidebar.date_input(
                        "Select Date Range",
                        value=(min_date.date(), max_date.date()),
                        min_value=min_date.date(),
                        max_value=max_date.date()
                    )
                    
                    if len(date_range) == 2:
                        start_date, end_date = date_range
                        filtered_data = filtered_data[
                            (filtered_data[created_date_col] >= pd.Timestamp(start_date)) &
                            (filtered_data[created_date_col] <= pd.Timestamp(end_date))
                        ]
            except Exception as e:
                st.sidebar.warning(f"Could not parse dates: {e}")
        
        # Display filtered resources
        st.subheader(f"📚 Resources ({len(filtered_data)} total)")
        
        # Build display columns
        display_cols = []
        if title_col:
            display_cols.append(title_col)
        if subject_col:
            display_cols.append(subject_col)
        if created_by_col:
            display_cols.append(created_by_col)
        if created_date_col:
            display_cols.append(created_date_col)
        if access_col:
            display_cols.append(access_col)
        
        if display_cols:
            display_df = filtered_data[display_cols].copy()
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        # Create summaries
        col1, col2 = st.columns(2)
        
        with col1:
            if created_by_col:
                st.subheader("👩‍🏫 Teacher Summary")
                teacher_summary = filtered_data.groupby(created_by_col).size().reset_index(name="Total Resources")
                teacher_summary = teacher_summary.sort_values("Total Resources", ascending=False)
                st.dataframe(teacher_summary, use_container_width=True, hide_index=True)
        
        with col2:
            if subject_col:
                st.subheader("📖 Subject Summary")
                subject_summary = filtered_data.groupby(subject_col).size().reset_index(name="Total Resources")
                subject_summary = subject_summary.sort_values("Total Resources", ascending=False)
                st.dataframe(subject_summary, use_container_width=True, hide_index=True)
        
        # Download button
        st.subheader("⬇️ Export Results")
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            if display_cols:
                filtered_data[display_cols].to_excel(
                    writer,
                    index=False,
                    sheet_name="Resources"
                )
            
            if created_by_col:
                teacher_summary = filtered_data.groupby(created_by_col).size().reset_index(name="Total Resources")
                teacher_summary = teacher_summary.sort_values("Total Resources", ascending=False)
                teacher_summary.to_excel(
                    writer,
                    index=False,
                    sheet_name="Teacher Summary"
                )
            
            if subject_col:
                subject_summary = filtered_data.groupby(subject_col).size().reset_index(name="Total Resources")
                subject_summary = subject_summary.sort_values("Total Resources", ascending=False)
                subject_summary.to_excel(
                    writer,
                    index=False,
                    sheet_name="Subject Summary"
                )
        
        output.seek(0)
        
        st.download_button(
            label="📊 Download as Excel",
            data=output.getvalue(),
            file_name="Teacher_Resource_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data found in uploaded files")
else:
    st.info("👆 Upload a file to get started")
