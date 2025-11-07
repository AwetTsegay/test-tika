import streamlit as st
from PIL import Image
import pathlib
import io

import extraction_process.extract_docx_info as ed
import extraction_process.convert_files as cf


def clear_submit():
    st.session_state["submit"] = False

st.set_page_config(layout="wide")
lm, main, rm = st.columns([1,8,1], gap="small")

with st.sidebar:
    st.title("Docx Embedded file Extractor")
    st.header("Instructions")
    st.markdown("""
    1. Upload a file (pdf, docx, pptx).
    2. Click on the 'Submit' button.
    3. View the extracted embedded files and their content.
    4. Convert the extracted data to various formats (JSON, Parquet, Pickle, CSV, Excel).
    5. Download the converted files.
    """)

    uploaded_file = st.file_uploader(
        "Upload a file (docx, doc, pdf, xlsx, etc.)", 
        type=["pdf", "docx", "pptx", "xlsx", "docx"],
        help="Any file are supported",
        on_change=clear_submit,
        key="file_uploader"
    )

with main:
    st.header("Output Results")
    tab1, tab2, tab3, tab4 = st.tabs(["Extracted Data", "Attached Files", "Attached Images", "Image"])

if uploaded_file:
    with st.sidebar:
        progress_bar = st.progress(100)
        status_text = st.empty

    with open(uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    processed_data = ed.extract_all_info(uploaded_file.name)
    combined_data = ed.combine_extracted_data(uploaded_file.name)
    json_data = cf.convert_to_json(combined_data)
    polars_data = cf.create_polars_df(combined_data)
    parquet_data = cf.create_parquet_file(polars_data)
    pickle_data = cf.create_pickle_file(combined_data)

    with tab1:
        with st.container(height=800):
            st.text(processed_data["content"])

    with tab2:
        with st.container(height=800):
            for key, value in processed_data["attachments"].items():
                file_extension = pathlib.Path(key).suffix.lower()

                if file_extension != '.png' and file_extension != '.emf' and file_extension != '.jpeg':
                    st.download_button(
                        label=f"**Download the fiel**: {key}",
                        data=value,
                        file_name=key
                    )


    with tab3:
        with st.container(height=800):
            for key, value in processed_data["attachments"].items():
                file_extension = pathlib.Path(key).suffix.lower()

                if file_extension == '.png' or file_extension == '.emf' or file_extension == '.jpeg':
                    st.download_button(
                        label=f"**Download the fiel**: {key}",
                        data=value,
                        file_name=key
                    )

    with tab4:
        with st.container(height=800):
            for key, value in processed_data["attachments"].items():
                file_extension = pathlib.Path(key).suffix.lower()

                if file_extension == '.png' or file_extension == '.jpeg':
                    with st.expander(f"**Image for** | {key}"):
                        image = Image.open(io.BytesIO(value))
                        st.image(image)


    with st.sidebar:
        st.header("Convert & Download")
        table_format = st.selectbox(
            "Select conversion format",
            options=["Json", "Parquet", "Pickle"],
            help="Choose the format"
        )

        if table_format == "Json":
            st.download_button(
                label="**Download the Json File.**",
                data=json_data,
                file_name="Json_data.json"
            )

        if table_format == "Parquet":
            st.download_button(
                label="**Download the Parquet File.**",
                data=parquet_data,
                file_name="Parquet_data.json"
            )

        if table_format == "Pickle":
            st.download_button(
                label="**Download the Pickle File.**",
                data=pickle_data,
                file_name="Pickle_data.json"
            )

