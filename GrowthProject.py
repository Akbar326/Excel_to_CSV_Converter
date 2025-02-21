import streamlit as st
import pandas as pd
import os
from io import BytesIO
from streamlit_lottie import st_lottie
import requests

st.set_page_config(page_title="Excel Wizard: CSV Converter", layout="wide", initial_sidebar_state="expanded")

# Custom CSS
st.markdown(
    """
    <style>
    .stApp {
        background-color: #f0f8ff;
        color: #333333;
    }
    .main-header {
        font-size: 2.5rem;
        color: #2c3e50;
        text-align: center;
        padding: 1rem;
        border-bottom: 2px solid #3498db;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.8rem;
        color: #34495e;
        border-left: 4px solid #3498db;
        padding-left: 1rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .info-text {
        font-size: 1.1rem;
        color: #7f8c8d;
        font-style: italic;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Lottie Animation
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

lottie_excel = load_lottieurl("https://assets5.lottiefiles.com/packages/lf20_qp1q7mct.json")

# Title and Description
st.markdown("<h1 class='main-header'>🧙‍♂️ Excel Wizard: CSV Converter</h1>", unsafe_allow_html=True)
col1, col2 = st.columns([2, 1])
with col1:
    st.markdown("<p class='info-text'>Transform your Excel files into CSV format with ease. Upload, clean, and convert your data in just a few clicks!</p>", unsafe_allow_html=True)
with col2:
    st_lottie(lottie_excel, height=150, key="excel_animation")

# File Uploader
st.markdown("<h2 class='sub-header'>📁 Upload Your Files</h2>", unsafe_allow_html=True)
uploaded_files = st.file_uploader("Choose Excel or CSV files", type=["csv", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower().strip()

        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file, engine='openpyxl')
        else:
            st.error("Invalid file format.")
            continue

        # File Details
        st.markdown(f"<h3 class='sub-header'>📊 Preview: {file.name}</h3>", unsafe_allow_html=True)
        st.dataframe(df.head())

        # Data Cleaning Options
        st.markdown("<h3 class='sub-header'>🧹 Data Cleaning Options</h3>", unsafe_allow_html=True)
        if st.checkbox(f"Clean data for {file.name}"):
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button(f"🔄 Remove Duplicates for {file.name}"):
                    df.drop_duplicates(inplace=True)
                    st.success("Duplicates removed successfully!")

            with col2:
                if st.button(f"🔢 Fill missing values for {file.name}"):
                    numeric_cols = df.select_dtypes(include=['number']).columns
                    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                    st.success("Missing values filled successfully!")

        st.markdown("<h3 class='sub-header'>🎛️ Column Selection</h3>", unsafe_allow_html=True)
        columns = st.multiselect(f"Choose columns for {file.name}", df.columns, default=df.columns)
        df = df[columns]

        # Data Visualization
        st.markdown("<h3 class='sub-header'>📈 Data Visualization</h3>", unsafe_allow_html=True)
        if st.checkbox(f"Show data visualization for {file.name}"):
            numerical_data = df.select_dtypes(include=['number']).iloc[:, :2]
            
            numerical_data = numerical_data.apply(pd.to_numeric, errors='coerce')
            numerical_data = numerical_data.dropna()

            if not numerical_data.empty:
                st.bar_chart(numerical_data)
            else:
                st.info("No numerical data available for visualization.")

        # Conversion Options
        st.markdown("<h3 class='sub-header'>🔄 Conversion Options</h3>", unsafe_allow_html=True)
        conversion_type = st.radio(f"Choose conversion type for {file.name}", ["CSV", "Excel"], key=file.name)
        if st.button(f"🚀 Convert {file.name}"):
            buffer = BytesIO()
            if conversion_type == "CSV":
                df.to_csv(buffer, index=False)
                file_name = file.name.replace(file_ext, ".csv")
                mime_type = "text/csv"

            elif conversion_type == "Excel":
                df.to_excel(buffer, index=False) 
                file_name = file.name.replace(file_ext, ".xlsx")
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            buffer.seek(0)

            st.download_button(
                label=f"📥 Download {file.name} as {conversion_type}",
                data=buffer,
                file_name=file_name,
                mime=mime_type
            )

            st.success(f"🎉 {file.name} converted successfully!")

else:
    st.info("👆 Upload your Excel or CSV files to get started!")