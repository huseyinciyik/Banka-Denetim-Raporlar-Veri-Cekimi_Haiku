import streamlit as st
import os, tempfile, re
from io import BytesIO
from openpyxl import Workbook
from pydantic import BaseModel, Field
from typing import Optional

from docling.document_converter import DocumentConverter
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_ollama import ChatOllama
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.retrievers import BM25Retriever
from langchain_classic.retrievers import EnsembleRetriever

# =============================================================================
# --- 1. VERİ MODELİ (GÜNCELLENDİ: Rakamlar birebir format için 'str' yapıldı) ---
# =============================================================================
class DenetimVerileri(BaseModel):
    banka_adi: str = Field(
        description="Raporun başlığındaki bankanın tam resmi adı."
    )
    donem: str = Field(
        description="Tablonun 'Cari dönem' tarihi. Örnek: '30 Haziran 2025'."
    )
    nakit_benzerleri_tp: Optional[str] = Field(
        default=None,
        description=(
            "'1.1 Nakit ve Nakit Benzerleri' satırının 'Cari Dönem TP' (TL) değeri. "
            "Tabloda nasılsa (noktalı/virgüllü) BİREBİR aynı metin olarak al. "
            "SADECE Cari Dönem sütunu — Önceki Dönem sütununu kesinlikle alma. "
            "Tabloda net göremiyorsan None döndür."
        )
    )
    nakit_benzerleri_yp: Optional[str] = Field(
        default=None,
        description=(
            "'1.1 Nakit ve Nakit Benzerleri' satırının 'Cari Dönem YP' (döviz) değeri. "
            "TP sütununun hemen sağındaki sütundur. Tabloda nasılsa BİREBİR aynı metin olarak al. "
            "SADECE Cari Dönem sütunu — Önceki Dönem sütununu kesinlikle alma. "
            "Tabloda net göremiyorsan None döndür."
        )
    )
    varliklar_toplami_tp: Optional[str] = Field(
        default=None,
        description=(
            "Tablonun en altındaki 'VARLIKLAR TOPLAMI' satırının 'Cari Dönem TP' değeri. "
            "Tabloda nasılsa (noktalı/virgüllü) BİREBİR aynı metin olarak al. "
            "SADECE Cari Dönem sütunu — Önceki Dönem sütununu kesinlikle alma. "
            "Tabloda net göremiyorsan None döndür."
        )
    )
    varliklar_toplami_yp: Optional[str] = Field(
        default=None,
        description=(
            "Tablonun en altındaki 'VARLIKLAR TOPLAMI' satırının 'Cari Dönem YP' değeri. "
            "TP sütununun hemen sağındaki sütundur. Tabloda nasılsa BİREBİR aynı metin olarak al. "
            "SADECE Cari Dönem sütunu — Önceki Dönem sütununu kesinlikle alma. "
            "Tabloda net göremiyorsan None döndür."
        )
    )


# --- 2. TABLO BLOKLARINI ÇIKAR ---
def extract_table_blocks(markdown_text: str) -> str:
    lines = markdown_text.split('\n')
    table_blocks, current_block = [], []
    context_header, in_table = "", False

    for line in lines:
        if line.startswith('#'):
            context_header = line
        if '|' in line:
            if not in_table:
                in_table = True
                if context_header:
                    current_block.append(context_header)
            current_block.append(line)
        else:
            if in_table and line.strip() == "":
                table_blocks.append('\n'.join(current_block))
                current_block, in_table = [], False

    if current_block:
        table_blocks.append('\n'.join(current_block))

    return "\n\n---TABLO---\n\n".join(table_blocks)


# --- 3. HİBRİT RETRIEVAL ---
def run_hybrid_retrieval(markdown_text: str) -> str:
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=3000, chunk_overlap=250,
        separators=["\n# ", "\n## ", "\n### ", "\n\n"]
    )
    chunks = text_splitter.split_text(markdown_text)

    raw_tables = extract_table_blocks(markdown_text)
    relevant_tables = [
        t for t in raw_tables.split("\n\n---TABLO---\n\n")
        if any(k in t.upper() for k in ["VARLIK", "BİLANÇO", "AKTİF", "NAKİT", "FİNANSAL DURUM"])
    ]
    table_context = "\n\n---\n\n".join(relevant_tables)

    embeddings = HuggingFaceEmbeddings(
        model_name="intfloat/multilingual-e5-base",
        model_kwargs={'device': 'mps'},
        encode_kwargs={'normalize_embeddings': True}
    )
    vectorstore = Chroma.from_texts(texts=chunks, embedding=embeddings)
    v_retriever = vectorstore.as_retriever(search_kwargs={"k": 4})
    bm25_retriever = BM25Retriever.from_texts(chunks)
    bm25_retriever.k = 4
    ensemble_retriever = EnsembleRetriever(
        retrievers=[bm25_retriever, v_retriever],
        weights=[0.7, 0.3]
    )

    queries = [
        "KONSOLİDE OLMAYAN BİLANÇO FİNANSAL DURUM TABLOSU VARLIKLAR",
        "1.1 Nakit ve Nakit Benzerleri Cari Dönem TP YP",
        "VARLIKLAR TOPLAMI Cari Dönem TP YP",
        "I. FİNANSAL VARLIKLAR nakit cari dönem",
        "AKTİF TOPLAMI bilanço cari dönem",
    ]
    retrieved = set()
    for q in queries:
        for d in ensemble_retriever.invoke(q):
            retrieved.add(d.page_content)

    return f"## FİLTRELENMİŞ TABLOLAR:\n{table_context}\n\n## EK BAĞLAM:\n" + "\n---\n".join(retrieved)


# --- 4. ANA EXTRACTION PROMPT (GÜNCELLENDİ: Orijinal formatı koru) ---
def build_extraction_prompt(context: str, raw_text_head: str) -> str:
    return f"""
Sen bir kıdemli banka denetçisisin. Aşağıdaki Türk bankası bilançosunu analiz et.

## 1. RAPOR KİMLİĞİ (Banka Adı ve Dönem):
Aşağıdaki ham rapor başlangıcından bankanın TAM RESMİ ADINI ve rapor dönemini bul.
ASLA "XYZ Bankası" gibi uydurma isimler kullanma. Bulamazsan "Bilinmeyen Banka" yaz.
[KAPAK METNİ]:
{raw_text_head}

## 2. BİLANÇO VERİLERİ (Rakamlar):
Aşağıdaki tablolardan istenen verileri çek.
[TABLO METNİ]:
{context}

## HEDEF DEĞERLER VE KURALLAR:
1. "1.1 Nakit ve Nakit Benzerleri" satırını bul.
2. "VARLIKLAR TOPLAMI" (veya AKTİF TOPLAMI) satırını bul.
3. Sütunları soldan sağa saymaya çalışma! Sütun başlıklarında "Cari Dönem TP" ve "Cari Dönem YP" yazan hizadaki rakamları al. "Önceki Dönem" rakamlarını KESİNLİKLE alma.
4. Rakamları tabloda nasıl görüyorsan (noktalı, virgüllü, boşluklu) BİREBİR AYNI METİN olarak kopyala. Kesinlikle tam sayıya çevirme veya temizleme yapma. (Örn: "5.754.499" veya "5,754,499" ise aynen öyle yaz).
5. Değeri tabloda net göremiyorsan None döndür.
"""


# --- 5. DOĞRULAMA (GÜNCELLENDİ: Metinsel sayılar için matematiksel dönüşüm) ---
def validate_and_retry(
    result: DenetimVerileri,
    structured_llm,
    context: str,
    prompt: str
) -> DenetimVerileri:
    issues = []

    # Metin halindeki sayıları gizlice integer'a çeviren yardımcı fonksiyon (sadece doğrulama için)
    def to_num(val_str):
        if not val_str: return 0
        clean = re.sub(r'[^\d]', '', str(val_str)) # Sadece rakamları tut
        return int(clean) if clean else 0

    # Kontrol 1: Nakit > Varlık Toplamı olamaz
    if result.nakit_benzerleri_tp and result.varliklar_toplami_tp:
        if to_num(result.nakit_benzerleri_tp) > to_num(result.varliklar_toplami_tp):
            issues.append("HATA: Nakit TP, Varlık Toplamı TP'den büyük olamaz.")
            
    if result.nakit_benzerleri_yp and result.varliklar_toplami_yp:
        if to_num(result.nakit_benzerleri_yp) > to_num(result.varliklar_toplami_yp):
            issues.append("HATA: Nakit YP, Varlık Toplamı YP'den büyük olamaz.")

    # Kontrol 2: Varlık Toplamı TP ve YP aynı olmamalı
    if result.varliklar_toplami_tp and result.varliklar_toplami_yp:
        if result.varliklar_toplami_tp == result.varliklar_toplami_yp:
            issues.append("UYARI: Varlık Toplamı TP ve YP değerleri aynı — TP/YP karışmış olabilir.")

    # Kontrol 3: Uydurma tespiti (Doğrudan tam metin eşleşmesi)
    for label, val in [
        ("nakit_benzerleri_tp", result.nakit_benzerleri_tp),
        ("varliklar_toplami_tp", result.varliklar_toplami_tp),
    ]:
        if val is not None:
            if str(val).strip() not in context:
                issues.append(f"UYDURMA TESPİTİ: '{val}' belgede bu tam formatıyla bulunamadı.")

    if issues:
        st.warning(f"⚠️ {len(issues)} sorun bulundu, tekrar analiz ediliyor...")
        for i in issues:
            st.caption(f"  → {i}")

        retry_prompt = prompt + f"""

## ÖNCEKİ ÇALIŞMADA TESPİT EDİLEN HATALAR:
{chr(10).join(f'- {i}' for i in issues)}

Bu hataları düzelt. Doğrulayamazsan ilgili alanı None yap.
HATIRLATMA: Sayı formatını ASLA bozma, tabloda nasılsa öyle al.
"""
        return structured_llm.invoke(retry_prompt)

    st.success("✅ Doğrulama geçti.")
    return result


# =============================================================================
# --- STREAMLIT ARAYÜZÜ ---
# =============================================================================
st.set_page_config(page_title="Banka Analiz Pipeline V5.2", layout="wide")
st.title("🏦 Banka Analiz Pipeline V5.2 (Orijinal Format)")
st.caption("Hedef: 1.1 Nakit ve Nakit Benzerleri (TP/YP) + VARLIKLAR TOPLAMI (TP/YP) — Cari Dönem")

with st.sidebar:
    st.subheader("⚙️ Model Ayarları")
    model_name = st.selectbox(
        "LLM Model",
        ["qwen2.5:7b", "gemma2", "mistral-nemo", "llama3.1:8b"],
        index=0,
        help="qwen2.5:7b tablo parsing'de en başarılı modeldir."
    )
    num_ctx = st.slider("Context Window", 4096, 16384, 8192, 1024)

    st.divider()
    st.subheader("🎯 Beklenen Değerler (Debug)")
    st.caption("Görselden doğrulanan referans değerler:")
    st.code(
        "Nakit TP:       5,754,499\n"
        "Nakit YP:      21,043,165\n"
        "Varlık Top TP: 53,910,706\n"
        "Varlık Top YP: 66,547,291",
        language="text"
    )

uploaded_file = st.file_uploader("📄 PDF Rapor Yükle", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 Analiz Et"):
        with st.spinner("Analiz başlatılıyor..."):
            try:
                # PDF → Markdown
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                st.info("📄 PDF → Markdown dönüşümü...")
                converter = DocumentConverter()
                md_text = converter.convert(tmp_path).document.export_to_markdown()
                os.remove(tmp_path)
                st.success(f"✅ Markdown hazır ({len(md_text):,} karakter)")

                with st.expander("📝 Ham Markdown (ilk 3000 karakter)"):
                    st.text(md_text[:3000])

                # Context hazırla
                st.info("🔍 İlgili tablo bölümleri çekiliyor...")
                context = run_hybrid_retrieval(md_text)

                with st.expander("🔍 LLM'e Gönderilen Context"):
                    st.text(context[:6000])

                # LLM
                llm = ChatOllama(
                    model=model_name,
                    temperature=0,
                    num_ctx=num_ctx,
                    format="json",
                    seed=42,
                    top_p=0.1
                )
                structured_llm = llm.with_structured_output(DenetimVerileri)

                # Extraction
                st.info("🤖 LLM extraction başlatılıyor...")
                
                # Rapor kapağı ile prompt oluştur
                prompt = build_extraction_prompt(context, md_text[:1500])

                with st.expander("📋 LLM'e Gönderilen Prompt"):
                    st.text(prompt[:4000])

                raw_result = structured_llm.invoke(prompt)
                st.info("🔎 Doğrulama yapılıyor...")
                final = validate_and_retry(raw_result, structured_llm, context, prompt)

                # --- SONUÇLAR ---
                st.divider()
                st.subheader(f"🏦 {final.banka_adi}")
                st.caption(f"📅 Cari Dönem: {final.donem}")

                col1, col2 = st.columns(2)

                # DEĞİŞİKLİK: Metin olduğu için formatlayıcı (:,) kaldırıldı, direkt str basılıyor
                def fmt(v): return str(v) if v else "⚠️ Bulunamadı"

                with col1:
                    st.metric(
                        "1.1 Nakit ve Nakit Benzerleri — TP",
                        fmt(final.nakit_benzerleri_tp)
                    )
                    st.metric(
                        "VARLIKLAR TOPLAMI — TP",
                        fmt(final.varliklar_toplami_tp)
                    )
                with col2:
                    st.metric(
                        "1.1 Nakit ve Nakit Benzerleri — YP",
                        fmt(final.nakit_benzerleri_yp)
                    )
                    st.metric(
                        "VARLIKLAR TOPLAMI — YP",
                        fmt(final.varliklar_toplami_yp)
                    )

                with st.expander("📋 Ham JSON"):
                    st.json(final.model_dump())

                # Excel export
                wb = Workbook()
                ws = wb.active
                ws.title = "Denetim Verileri"
                ws.append(["Banka Adı", "Dönem",
                            "Nakit TP", "Nakit YP",
                            "Varlık Top TP", "Varlık Top YP"])
                ws.append([
                    final.banka_adi, final.donem,
                    final.nakit_benzerleri_tp, final.nakit_benzerleri_yp,
                    final.varliklar_toplami_tp, final.varliklar_toplami_yp,
                ])
                buf = BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button(
                    "📥 Excel İndir", buf,
                    file_name=f"{final.banka_adi}_{final.donem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Hata: {e}")
                st.exception(e)