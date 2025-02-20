import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Excel to CSV Converter", layout="wide")

# Custom CSS
st.markdown(
    """
    <style>
    .stApp{
        background-color: #f0f0f5;
        color: #000000;
    }      
    </style>
    """,
    unsafe_allow_html=True
)

# Title and Description
st.title("Excel to CSV Converter")
st.write("Upload an Excel file and convert it to CSV format.")

# File Uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["csv", "xlsx"], accept_multiple_files=True)

if uploaded_file:
    for file in uploaded_file:
        file_ext = os.path.splitext(file.name)[-1].lower().strip()

        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file, engine='openpyxl')  # Specify engine for xlsx
        else:
            st.error("Invalid file format.")
            continue

        # File Details
        st.write("Preview of the file:")
        st.dataframe(df.head())  # Fixed typo

        # Data Cleaning Options
        st.subheader("Data Cleaning Options")
        if st.checkbox(f"Clean data for {file.name}"):
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button(f"Remove Duplicates for {file.name}"):
                    df.drop_duplicates(inplace=True)
                    st.write("Duplicates removed.")

            with col2:
                if st.button(f"Fill missing values for {file.name}"):
                    numeric_cols = df.select_dtypes(include=['number']).columns
                    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                    st.write("Missing values filled.")

        st.subheader("Select Columns to Keep")
        columns = st.multiselect(f"Choose columns for {file.name}", df.columns, default=df.columns)
        df = df[columns]

        # Data Visualization
        st.subheader("Data Visualization")
        if st.checkbox(f"Show data visualization for {file.name}"):
            st.bar_chart(df.select_dtypes(include=['number']).iloc[:, :2])

        # Conversion Options
        st.subheader("Conversion Options")
        conversion_type = st.radio(f"Choose conversion type for {file.name}", ["CSV", "Excel"], key=file.name)
        if st.button(f"Convert {file.name}"):
            buffer = BytesIO()
            if conversion_type == "CSV":
                df.to_csv(buffer, index=False)
                file_name = file.name.replace(file_ext, ".csv")
                mime_type = "text/csv"

            elif conversion_type == "Excel":
                df.to_excel(buffer, index=False)  # Fixed typo
                file_name = file.name.replace(file_ext, ".xlsx")
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            buffer.seek(0)

            st.download_button(
                label=f"Download {file.name} as {conversion_type}",
                data=buffer,
                file_name=file_name,
                mime=mime_type
            )

            st.success(f"{file.name} converted successfully.")
