import os
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from langchain_community.vectorstores import FAISS
from langchain_openai import AzureChatOpenAI, AzureOpenAIEmbeddings
from langchain.schema import HumanMessage
from langchain.docstore.document import Document

import regex
import json

# Load env
load_dotenv()
from openpyxl import Workbook
import io

def generate_excel_from_json(data: dict, sheet_name: str = "Form Metadata") -> io.BytesIO:
    """
    Converts a dictionary into an Excel file and returns it as a BytesIO object.

    Parameters:
    - data (dict): The JSON-like dictionary to write into Excel.
    - sheet_name (str): The name of the Excel sheet.

    Returns:
    - io.BytesIO: An in-memory Excel file stream.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Write headers
    headers = list(data.keys())
    ws.append(headers)

    # Write values
    values = list(data.values())
    ws.append(values)

    # Save to memory buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer


# Azure OpenAI Configuration
deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")
embedding_deployment_name = os.getenv("AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME")
api_version = os.getenv("AZURE_OPENAI_API_VERSION")
api_key = os.getenv("AZURE_OPENAI_API_KEY")
azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")

# Init Chat Model
chat = AzureChatOpenAI(
    azure_deployment=deployment_name,
    azure_endpoint=azure_endpoint,
    openai_api_version=api_version,
    openai_api_key=api_key,
    temperature=0.2
)

# Init Embeddings
embedding_model = AzureOpenAIEmbeddings(
    azure_deployment=embedding_deployment_name,
    azure_endpoint=azure_endpoint,
    openai_api_version=api_version,
    openai_api_key=api_key
)

# Streamlit UI
st.title("📄 Form Metadata Q&A (Vector-Based)")
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])



def extract_json_from_text(text):  
    # Regular expression pattern to identify JSON objects  
    json_pattern = regex.compile(r'\{(?:[^{}]|(?R))*\}')  
      
    # Find all matches in the text  
    json_matches = json_pattern.findall(text)  
      
    json_objects = []  
      
    for match in json_matches:  
        try:  
            # Parse JSON string to a Python dictionary  
            json_object = json.loads(match)  
            json_objects.append(json_object)  
        except json.JSONDecodeError:  
            # Skip invalid JSON strings  
            continue  
      
    return json_objects

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ File uploaded successfully")
    # st.dataframe(df.head())

    # Convert each row to a string representation
    docs = []
    for _, row in df.iterrows():
        text = "\n".join([f"{col}: {val}" for col, val in row.items() if pd.notna(val)])
        docs.append(Document(page_content=text))

    # Create FAISS Vector Store
    with st.spinner("🔍 Creating vector index..."):
        vector_store = FAISS.from_documents(docs, embedding_model)

    # # User Question Input
    query = st.text_area("❓ Ask a question about the forms:")
    sumit_button = st.button("Submit Question")
    if sumit_button:
        if not query:
            st.warning("Please enter a question.")
            st.stop()

        if query:
            with st.spinner("🔎 Searching relevant rows..."):
                similar_docs = vector_store.similarity_search(query, k=5)

            context_text = "\n\n".join(doc.page_content for doc in similar_docs)

        #     prompt = f"""
        #     You are a helpful assistant that reads insurance form metadata and answers questions based on the data.

        #     Here are some rows from the data relevant to the user's question:

        #     {context_text}

        #     Now answer the following question in a clear, human-readable paragraph:
        #     {query}
        # """


            # prompt = f"""
            #     You are a data assistant that reads structured insurance form metadata from Excel rows.

            #     Below are the rows related to the user's query:

            #     {context_text}

            #     The user query is:
            #     {query}

            #     Instructions:
            #     1. If the user is asking to generate structured data (Excel/JSON) for a specific **Form Number**, then:
            #         - Identify the row(s) where the 'Form Number' matches the one mentioned in the query.
            #         - Extract and map the following fields directly from the row:
            #             - 'Form Number' → Formcode
            #             - 'Form Title' → FormDesc
            #             - 'Eff Date' → StartEffectiveDate
            #             - 'Exp Date' → EndEffectiveDate
            #             - 'Premium Bearing?' → IsMandatory
            #             - 'Line of Business' → LineOfBusiness
            #             - 'IAT Product' → IATProduct

            #     2. Now derive the following columns based on the text in the column **"Policy Forms Attach. Rules"**.
            #         For each derived field, do **not just output a keyword** like "Terrorism", "Risk State", etc.
            #         Instead, do the following:
            #         - If you find a relevant clause, sentence, or explanation in **Policy Forms Attach. Rules**, return that full matching line or phrase as the value.
            #         - If nothing is found, return `"NA"`.

            #         Derived Fields:
            #         - **TerrorismCheck**: Any clause or note related to terrorism endorsements, TRIA coverage, or terrorism conditions.
            #         - **PrimaryRatingStateCheck**: Any rule related to the Primary Rating State, its applicability, or its conditions.
            #         - **RiskStateCheck**: Any language referring to "Risk State" or its specific requirements or exclusions.
            #         - **ExposureClassCodeCheck**: Any checks or rules that include Exposure Class Codes or mention class-based eligibility.
            #         - **CoverageExtraDataCheck**: Any requirements for extra data, supplemental coverage fields, or additional inputs.
            #         - **ExposureClassCodeExclude**: Any exclusions or limitations based on Exposure Class Codes.
            #         - **LImitDeductibleCheck**: Any mention of deductible thresholds, coverage limits, or related clauses.

            #     3. Return the final result as **JSON only**, using the following keys exactly:
            #     Formcode, FormDesc, StartEffectiveDate, EndEffectiveDate, IsMandatory,
            #     LineOfBusiness, IATProduct, TerrorismCheck, PrimaryRatingStateCheck,
            #     RiskStateCheck, ExposureClassCodeCheck, CoverageExtraDataCheck,
            #     ExposureClassCodeExclude, LImitDeductibleCheck

            #     4. If the user query is not asking for such structured output, just respond naturally in paragraph format.
            #     """



            prompt = f"""
                You are a data assistant that reads structured insurance form metadata from Excel rows.

                Below are the rows related to the user's query:

                {context_text}

                The user query is:
                {query}

                Instructions:
                1. If the user is asking to generate structured data (Excel/JSON) for a specific **Form Number**, then:
                    - Identify the row(s) where the 'Form Number' matches the one mentioned in the query.
                    - Extract and map the following fields directly from the row:
                        - 'Form Number' → Formcode
                        - 'Form Title' → FormDesc
                        - 'Eff Date' → StartEffectiveDate
                        - 'Exp Date' → EndEffectiveDate  
                        → If no Exp Date is provided or is blank, return "31-12-9999" as the default end date.
                        - 'Premium Bearing?' → IsMandatory
                        - 'Line of Business' → LineOfBusiness
                        - 'IAT Product' → IATProduct

                2. Now derive the following columns based on the text in the column **"Policy Forms Attach. Rules"**.
                    For each derived field:
                    - Do **not** return just a keyword like "Terrorism" or "Risk State".
                    - Instead, return the **actual full clause, sentence, or phrase** from "Policy Forms Attach. Rules" that mentions the concept.
                    - Rephrase or summarize it if needed, but keep the original context.
                    - If no relevant info is found, return `"NA"`.

                    Derived Fields:
                    - **TerrorismCheck**: Any clause related to terrorism endorsements, TRIA coverage, or related conditions.
                    - **PrimaryRatingStateCheck**: Any rule about Primary Rating State applicability.
                    - **RiskStateCheck**: Any language referring to Risk State eligibility, exceptions, or handling.
                    - **ExposureClassCodeCheck**: Any inclusion rule based on exposure class codes.
                    - **CoverageExtraDataCheck**: Any requirement for additional data fields or supplemental coverage info.
                    - **ExposureClassCodeExclude**: Any exclusion conditions based on exposure class codes.
                    - **LImitDeductibleCheck**: Any rule about minimum/maximum limits or deductible amounts.

                3. Return the result in **valid JSON format only** with the following keys:
                Formcode, FormDesc, StartEffectiveDate, EndEffectiveDate, IsMandatory,
                LineOfBusiness, IATProduct, TerrorismCheck, PrimaryRatingStateCheck,
                RiskStateCheck, ExposureClassCodeCheck, CoverageExtraDataCheck,
                ExposureClassCodeExclude, LImitDeductibleCheck

                4. If the user's query is not asking for JSON or structured Excel generation, then ignore all formatting and respond normally in paragraph form.
                5. If user asks for multiple Form Numbers, then generate a JSON for each Form Number mentioned in the query.
                """



            with st.spinner("🧠 Thinking..."):
                answer = chat.invoke([HumanMessage(content=prompt)]).content
                if extract_json_from_text(answer):
                    try:
                        parsed_json = extract_json_from_text(answer)[0]
                        excel_buffer = generate_excel_from_json(parsed_json)

                        st.markdown("### 📥 Download Generated Excel")
                        st.download_button(
                            label="Download Excel File",
                            data=excel_buffer,
                            file_name=f"Form_Metadata_{parsed_json.get('Formcode', 'output')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
                    except json.JSONDecodeError as e:
                        st.error(f"Error parsing JSON: {e}")

            
            
                

            st.markdown("### ✅ Answer")
            st.write(answer)
