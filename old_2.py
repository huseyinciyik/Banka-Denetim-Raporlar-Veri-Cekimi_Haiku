import streamlit as st
import os, tempfile, re
from io import BytesIO
from openpyxl import Workbook
from pydantic import BaseModel, Field
from typing import Optional

# --- CORE KÜTÜPHANELER ---
from docling.document_converter import DocumentConverter
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_ollama import ChatOllama
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.retrievers import BM25Retriever
from langchain_classic.retrievers import EnsembleRetriever

# --- 1. VERİ MODELİ ---
class DenetimVerileri(BaseModel):
    banka_adi: str = Field(description="Banka tam adı.")
    donem: str = Field(description="Rapor dönemi (Örn: 30 Haziran 2025).")
    nakit_benzerleri_tp: Optional[int] = Field(description="Bilanço 1.1 Nakit TP.")
    nakit_benzerleri_yp: Optional[int] = Field(description="Bilanço 1.1 Nakit YP.")
    varliklar_toplami_tp: Optional[int] = Field(description="Varlıklar Toplamı TP.")
    varliklar_toplami_yp: Optional[int] = Field(description="Varlıklar Toplamı YP.")

# --- 2. AKILLI TABLO ÇIKARIMI (SMART EXTRACTION) ---
def extract_table_blocks(markdown_text: str) -> str:
    lines = markdown_text.split('\n')
    table_blocks, current_block = [], []
    context_header, in_table = "", False

    for line in lines:
        if line.startswith('#'): context_header = line
        if '|' in line:
            if not in_table:
                in_table = True
                if context_header: current_block.append(context_header)
            current_block.append(line)
        else:
            if in_table and line.strip() == "":
                table_blocks.append('\n'.join(current_block))
                current_block, in_table = [], False
    
    if current_block: table_blocks.append('\n'.join(current_block))
    return "\n\n---TABLO_AYRACI---\n\n".join(table_blocks)

# --- 3. HİBRİT RETRIEVAL (BM25 + VEKTÖR) ---
def run_hybrid_retrieval(markdown_text: str):
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=3000, 
        chunk_overlap=250,
        separators=["\n# ", "\n## ", "\n### ", "\n\n"] 
    )
    chunks = text_splitter.split_text(markdown_text)
    
    # 🔥 OPTİMİZASYON 1: Akıllı Tablo Filtreleme
    raw_tables = extract_table_blocks(markdown_text)
    relevant_tables = []
    
    for table in raw_tables.split("\n\n---TABLO_AYRACI---\n\n"):
        # Sadece finansal anahtar kelimeleri içeren tabloları tut (Context Window'u koru)
        if any(keyword in table.upper() for keyword in ["VARLIK", "BİLANÇO", "AKTİF", "NAKİT"]):
            relevant_tables.append(table)
            
    table_context = "\n\n---FİLTRELENMİŞ_TABLO---\n\n".join(relevant_tables)
    
    # M3 GPU Destekli Embedding
    embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2",
        model_kwargs={'device': 'mps'}
    )
    
    # Retriever'lar
    vectorstore = Chroma.from_texts(texts=chunks, embedding=embeddings)
    v_retriever = vectorstore.as_retriever(search_kwargs={"k": 4})
    bm25_retriever = BM25Retriever.from_texts(chunks)
    bm25_retriever.k = 4
    
    ensemble_retriever = EnsembleRetriever(
        retrievers=[bm25_retriever, v_retriever], weights=[0.5, 0.5]
    )
    
    queries = [
        "KONSOLİDE OLMAYAN BİLANÇO FİNANSAL DURUM TABLOSU AKTİFLER",
        "nakit ve nakit benzerleri TP YP cari dönem",
        "VARLIKLAR TOPLAMI bilanço aktif"
    ]
    
    retrieved_chunks = set()
    for q in queries:
        docs = ensemble_retriever.invoke(q)
        for d in docs: retrieved_chunks.add(d.page_content)
    
    final_context = f"## DOĞRUDAN VE FİLTRELENMİŞ TABLOLAR:\n{table_context}\n\n"
    final_context += "## EK BAĞLAM:\n" + "\n---\n".join(retrieved_chunks)
    return final_context

# --- 4. DOĞRULAMA VE YENİDEN DENEME ---
def validate_and_retry(result: DenetimVerileri, structured_llm, prompt: str) -> DenetimVerileri:
    issues = []
    if result.varliklar_toplami_tp and result.nakit_benzerleri_tp:
        if result.nakit_benzerleri_tp > result.varliklar_toplami_tp:
            issues.append("MANTIK HATASI: Nakit, toplam varlıktan büyük olamaz.")
    
    if not result.banka_adi or result.banka_adi.strip() == "":
        issues.append("Hata: Banka adı boş.")
    
    if issues:
        st.warning(f"⚠️ Hata tespit edildi, tekrar analiz ediliyor: {issues}")
        retry_prompt = prompt + f"\n\n# ÖNCEKİ HATALAR:\n{issues}\nLütfen düzelt."
        return structured_llm.invoke(retry_prompt)
    return result

# --- STREAMLIT ARAYÜZÜ ---
st.set_page_config(page_title="Banka Analiz Pipeline V3", layout="wide")
st.title("🏦 Profesyonel Banka Analiz Pipeline (Memory Optimized)")

with st.sidebar:
    model_name = st.selectbox("LLM Model", ["gemma2", "llama3.2"], index=0)

uploaded_file = st.file_uploader("PDF Rapor Yükle", type="pdf")

if uploaded_file is not None:
    if st.button("Pipeline'ı Başlat"):
        with st.spinner('Docling + Memory Optimized Hybrid Extraction çalışıyor...'):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                converter = DocumentConverter()
                md_text = converter.convert(tmp_path).document.export_to_markdown()
                os.remove(tmp_path)
                
                context = run_hybrid_retrieval(md_text)
                
                with st.expander("🔍 AI'nın Gördüğü Filtrelenmiş Veri (Debug)"):
                    st.text(context)

                # 🔥 OPTİMİZASYON 2: Hafıza (Context Window) Genişletme
                llm = ChatOllama(
                    model=model_name, 
                    temperature=0,
                    num_ctx=8192 # Hafızayı 8K token'a çıkardık
                )
                structured_llm = llm.with_structured_output(DenetimVerileri)
                
                # 🔥 OPTİMİZASYON 3: Recency Bias (Sondan Hatırlama) Mantıklı Prompt
                prompt = f"""
                # BAĞLAM (CONTEXT):
                {context}

                ---
                # ROLE: Senior Bank Auditor
                Yukarıdaki verileri inceledin. Şimdi şu adımları eksiksiz uygula:
                1. SADECE 'Bilanço / Varlıklar' tablosuna odaklan.
                2. 'Cari Dönem' sütununu bul.
                3. '1.1 Nakit' ve 'VARLIKLAR TOPLAMI' değerlerini çek.
                4. Rakamları tam sayı yap. Null ise uydurma.
                """
                
                raw_result = structured_llm.invoke(prompt)
                final_result = validate_and_retry(raw_result, structured_llm, prompt)
                
                st.subheader(f"Banka: {final_result.banka_adi}")
                st.json(final_result.model_dump())
                
            except Exception as e:
                st.error(f"Sistem Hatası: {str(e)}")