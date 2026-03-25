import streamlit as st
import os
import tempfile
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from pydantic import BaseModel, Field
from typing import Optional

# Pipeline ve RAG Kütüphaneleri
from docling.document_converter import DocumentConverter
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_ollama import ChatOllama
from langchain_huggingface import HuggingFaceEmbeddings


# --- 1. VERİ MODELİ (PYDANTIC) ---
class DenetimVerileri(BaseModel):
    banka_adi: str = Field(description="Raporun en üstünde yer alan bankanın tam ticari unvanı.")
    donem: str = Field(description="Raporun ait olduğu tarih ve dönem bilgisi.")
    
    
    nakit_benzerleri_tp: Optional[int] = Field(description="Bilanço tablosu 1.1 satırında, dipnottan sonraki İLK sayısal değer.")
    nakit_benzerleri_yp: Optional[int] = Field(description="Bilanço tablosu 1.1 satırında, dipnottan sonraki İKİNCİ sayısal değer.")
    
    varliklar_toplami_tp: Optional[int] = Field(description="'VARLIKLAR TOPLAMI' satırındaki İLK sayısal değer (TP).")
    varliklar_toplami_yp: Optional[int] = Field(description="'VARLIKLAR TOPLAMI' satırındaki İKİNCİ sayısal değer (YP).")   


# --- 2. EXCEL OLUŞTURMA FONKSİYONU ---
def create_excel_report(data: DenetimVerileri):
    wb = Workbook()
    ws = wb.active
    ws.title = "Denetim Verileri"
    
    bold_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    
    ws['A1'] = "BANKA ADI:"
    ws['B1'] = data.banka_adi
    ws['A2'] = "DÖNEM:"
    ws['B2'] = data.donem
    
    headers = ["KALEMLER", "TP (Türk Parası)", "YP (Yabancı Para)"]
    ws.append([]) 
    ws.append(headers)
    
    rows = [
        ("Nakit ve Nakit Benzerleri", data.nakit_benzerleri_tp, data.nakit_benzerleri_yp),
        ("VARLIKLAR TOPLAMI", data.varliklar_toplami_tp, data.varliklar_toplami_yp)
    ]
    for row in rows:
        ws.append(row)
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- 3. RETRIEVAL PIPELINE (DÜZELTİLMİŞ) ---
def run_advanced_retrieval(markdown_text):
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=2000, 
        chunk_overlap=500, 
        separators=["\n# ", "\n## ", "\n\n", "|"] 
    )
    chunks = text_splitter.split_text(markdown_text)
    
    model_name = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"    
    embeddings = HuggingFaceEmbeddings(
        model_name=model_name,
        model_kwargs={'device': 'mps'} # M3 işlemcinin GPU gücünü kullanır
    ) # <--- EKSİK OLAN KAPANIŞ PARANTEZİ BURASIYDI!

    vectorstore = Chroma.from_texts(texts=chunks, embedding=embeddings)
    
    queries = [
        "KONSOLİDE OLMAYAN BİLANÇO FİNANSAL DURUM TABLOSU AKTİFLER",
        "VARLIKLAR TOPLAMI nakit ve nakit benzerleri tablosu",
        "Bilanço tablosu Cari Dönem TP ve YP sütunları"
    ]
    
    final_context = ""
    for q in queries:
        docs = vectorstore.as_retriever(search_kwargs={"k": 3}).invoke(q)
        final_context += "\n".join([d.page_content for d in docs]) + "\n---\n"
        
    return final_context

# --- 4. ANA UYGULAMA (STREAMLIT) ---
st.set_page_config(page_title="Banka Analiz Pipeline", layout="wide")
st.title("🏦 Profesyonel Banka Analiz Pipeline (RAG)")

with st.sidebar:
    st.header("⚙️ Ayarlar")
    model_name = st.selectbox("LLM Model", ["gemma2", "llama3.2", "llama3.1"], index=0)

uploaded_file = st.file_uploader("Banka Denetim Raporu PDF", type="pdf")

if uploaded_file is not None:
    if st.button("Pipeline'ı Başlat ve Excel Üret"):
        with st.spinner('Docling + Retrieval Pipeline çalışıyor...'):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                converter = DocumentConverter()
                markdown_text = converter.convert(tmp_path).document.export_to_markdown()
                os.remove(tmp_path)
                
                context = run_advanced_retrieval(markdown_text)
                
                # --- BUGÜN EKLEDİĞİMİZ DEBUG KISMI BURASI ---
                with st.expander("🔍 AI'nın Gördüğü Ham Tablo Verisi (Debug)"):
                    st.markdown("### LLM'e Gönderilen Ham Metin Parçaları:")
                    st.text(context) # Bu satır sayesinde kutunun içinde ham metni göreceksin.
                # --------------------------------------------

                llm = ChatOllama(model=model_name, temperature=0)
                
                structured_llm = llm.with_structured_output(DenetimVerileri)
                
                prompt = f"""
                # ROLE: Senior Bank Auditor
                Görevin, aşağıdaki metinden Bilanço (Finansal Durum) tablosundaki verileri çekmektir.

                # KESİN KURALLAR:
                1. **TABLO AYRIMI:** Sadece 'VARLIKLAR' (Assets) ile başlayan tabloyu kullan. 'ÖZ KAYNAKLAR' tablosuna ASLA bakma.
                2. **SÜTUN MANTIĞI:** - Tablo satırında açıklama metninden sonra gelen ilk sayısal sütun 'TP' (Türk Parası) değeridir.
                - İkinci sayısal sütun 'YP' (Yabancı Para) değeridir.
                3. **DİKKAT:** Sadece 'Cari Dönem' (en güncel tarihli) sütunlarındaki verileri al.
                4. **TEMİZLİK:** Rakamlardaki binlik ayraçlarını (nokta veya virgül) silerek sadece tam sayı (integer) döndür.
                5. **DOĞRULUK:** Veriyi bulamazsan 'null' döndür. Tahmin yürütme.

                # CONTEXT:
                {context}
                """
                # Invoke satırı aşağı alındı ve temizlendi
                result = structured_llm.invoke(prompt)
                
                st.subheader(f"Analiz: {result.banka_adi}")
                st.json(result.model_dump())
                
                excel_file = create_excel_report(result)
                st.download_button(
                    label="Excel Raporunu İndir",
                    data=excel_file,
                    file_name=f"analiz_{result.donem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Pipeline Hatası: {str(e)}")