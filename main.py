import os
import certifi
os.environ['SSL_CERT_FILE'] = certifi.where()
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()

# Suppress SSL certificate verification for urllib3/requests (not recommended for production)
import ssl
import urllib3
ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

os.environ['CURL_CA_BUNDLE'] = ''
os.environ['REQUESTS_CA_BUNDLE'] = ''

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
import requests
import re

# --- UI ---
st.title("📄 Form Metadata Q&A (Remote Vector Store API)")

tab1, tab2 = st.tabs(["Upload File to API", "Ask a Question"])

API_UPLOAD_URL = "http://iatsvctest01.ofcwic.com/IATAIAPI/v1/Upload"
API_QA_URL = "http://iatsvctest01.ofcwic.com/IATAIAPI/v1/AzureAI"
API_KEY = "Use Coforge Key"
VECTOR_STORE_ID = "vs_eFNZIoLgz8GP7ydIV2WTfmNi"

def clean_answer(text):
    # Remove citations like 【6:0†source】
    text = re.sub(r'【\d+:\d+†source】', '', text)
    # Remove HTML tags
    text = re.sub(r'<.*?>', '', text)
    # Remove [thread:...] if present
    text = re.sub(r'\[thread:[^\]]+\]', '', text)
    # Remove [source:...] converted.txt and similar
    text = re.sub(r'\[source:[^\]]+\] converted\.txt', '', text)
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

with tab1:
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    st.info(f"File will be converted to text and uploaded to API, associated with Vector Store ID: {VECTOR_STORE_ID}")

    text_data = None  # Track converted text for download

    if uploaded_file is not None:
        if st.button("Convert to Text & Upload to API"):
            # Convert Excel to text (one row per line, columns as key: value)
            try:
                df = pd.read_excel(uploaded_file)
                text_rows = []
                for _, row in df.iterrows():
                    text = "\n".join([f"{col}: {val}" for col, val in row.items() if pd.notna(val)])
                    text_rows.append(text)
                text_data = "\n\n".join(text_rows)
            except Exception as e:
                st.error(f"Error reading Excel file: {e}")
                st.stop()

            # Prepare text as a file-like object for upload
            from io import BytesIO
            text_file = BytesIO(text_data.encode("utf-8"))

            files = {
                "uploaded_file": ("converted.txt", text_file, "text/plain")
            }
            headers = {
                "X-API-Key": API_KEY
            }
            params = {
                "storeId": VECTOR_STORE_ID
            }
            with st.spinner("Uploading converted text file to API..."):
                response = requests.post(API_UPLOAD_URL, headers=headers, params=params, files=files)
            if response.status_code == 200:
                st.success("✅ File (converted to text) uploaded to API successfully!")
                try:
                    st.json(response.json())
                except Exception:
                    st.write(response.text)
                # Download button for the converted text file
                st.download_button(
                    label="Download Converted Text File",
                    data=text_data,
                    file_name="converted.txt",
                    mime="text/plain"
                )
            else:
                st.error(f"❌ Upload failed! Status code: {response.status_code}")
                try:
                    st.json(response.json())
                except Exception:
                    st.write(response.text)

with tab2:
    store_id = st.text_input("Enter your Vector Store ID to use:", value=VECTOR_STORE_ID, key="query_store_id")
    question = st.text_area("❓ Ask a question about the forms:")
    submit = st.button("Submit Question")

    if submit and question and store_id:
        headers = {"X-API-Key": API_KEY}
        body = {
            "threadId": "",
            "prompt": question,
            "temperature": 0.15,
            "storeId": store_id,
            "uploadedFiles": []
        }
        with st.spinner("Getting answer from API..."):
            response = requests.post(API_QA_URL, headers=headers, json=body, stream=True)
            answer = ""
            for chunk in response.iter_lines():
                if chunk:
                    answer += chunk.decode("utf-8")
            st.markdown("### ✅ Answer")
            st.write(clean_answer(answer))
