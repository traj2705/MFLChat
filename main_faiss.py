import os
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from langchain_community.vectorstores import FAISS
from langchain_openai import AzureChatOpenAI, AzureOpenAIEmbeddings
from langchain.schema import HumanMessage
from langchain.docstore.document import Document

# Load env
load_dotenv()

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

    # User Question Input
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

            prompt = f"""
    You are a helpful assistant that reads insurance form metadata and answers questions based on the data.

    Here are some rows from the data relevant to the user's question:

    {context_text}

    Now answer the following question in a clear, human-readable paragraph:
    {query}
    """

            with st.spinner("🧠 Thinking..."):
                answer = chat.invoke([HumanMessage(content=prompt)]).content

            st.markdown("### ✅ Answer")
            st.write(answer)
