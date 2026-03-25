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
# VERİ MODELİ
# =============================================================================
class DenetimVerileri(BaseModel):
    banka_adi: str = Field(description="Bankanın tam resmi adı.")
    donem: str = Field(description="Cari dönem tarihi. Örnek: '30 Haziran 2025'.")
    nakit_benzerleri_tp: Optional[str] = Field(default=None)
    nakit_benzerleri_yp: Optional[str] = Field(default=None)
    varliklar_toplami_tp: Optional[str] = Field(default=None)
    varliklar_toplami_yp: Optional[str] = Field(default=None)


# =============================================================================
# YARDIMCI: Sayı metnini temizle (doğrulama için)
# =============================================================================
def to_num(val_str) -> int:
    if not val_str:
        return 0
    clean = re.sub(r'[^\d]', '', str(val_str))
    return int(clean) if clean else 0


# =============================================================================
# AŞAMA 1: DETERMINISTIK TABLO PARSER
#
# Docling'in ürettiği markdown tablosunu satır satır okur.
# Sütun pozisyonlarını başlık satırından çıkarır → TP/YP karışması imkansız.
# LLM kullanmaz, regex ile doğrudan değeri alır.
# =============================================================================

def _find_tp_yp_columns(header_row: list[str]) -> tuple[int, int]:
    """
    Başlık satırından Cari Dönem TP ve YP sütun indekslerini bulur.
    Docling'in ürettiği format: [satır_kodu, açıklama, dipnot, TP, YP, Toplam, TP, YP, Toplam]
    Cari dönem her zaman ilk TP/YP çiftidir.
    """
    tp_idx, yp_idx = -1, -1
    for i, cell in enumerate(header_row):
        cell_clean = cell.strip().upper()
        if cell_clean == "TP" and tp_idx == -1:
            tp_idx = i
        elif cell_clean == "YP" and yp_idx == -1 and tp_idx != -1:
            yp_idx = i
            break  # İlk TP/YP çiftini bulduk, dur
    return tp_idx, yp_idx


def _parse_table_row(row: str) -> list[str]:
    """Markdown tablo satırını hücrelere böler."""
    cells = row.strip().strip('|').split('|')
    return [c.strip() for c in cells]


def deterministic_parse(markdown_text: str) -> dict:
    """
    Markdown tablosundan hedef değerleri deterministik olarak çeker.
    Döndürür: {nakit_tp, nakit_yp, varlik_tp, varlik_yp} veya None değerler.
    """
    result = {
        "nakit_tp": None,
        "nakit_yp": None,
        "varlik_tp": None,
        "varlik_yp": None,
        "tp_col": -1,
        "yp_col": -1,
    }

    lines = markdown_text.split('\n')
    in_bilanco = False
    tp_col, yp_col = -1, -1
    header_found = False

    for i, line in enumerate(lines):
        # Bilanço tablosunu bul
        if 'BİLANÇO' in line.upper() or 'FİNANSAL DURUM TABLOSU' in line.upper():
            in_bilanco = True

        if not in_bilanco:
            continue

        if '|' not in line:
            continue

        cells = _parse_table_row(line)

        # Ayırıcı satırları atla (--- içerenler)
        if all(re.match(r'^[-:]+$', c) or c == '' for c in cells):
            continue

        # TP/YP başlık satırını bul
        upper_cells = [c.upper() for c in cells]
        if 'TP' in upper_cells and 'YP' in upper_cells and not header_found:
            tp_col, yp_col = _find_tp_yp_columns(cells)
            result["tp_col"] = tp_col
            result["yp_col"] = yp_col
            header_found = True
            st.caption(f"  📍 Sütun tespiti: TP={tp_col}, YP={yp_col} | Başlık: {cells}")
            continue

        if not header_found or tp_col == -1:
            continue

        # Satır içeriğini kontrol et
        row_text = ' '.join(cells).upper()

        # 1.1 Nakit ve Nakit Benzerleri satırı
        # Dikkat: "1.1.1", "1.1.2" gibi alt kalemleri ALMA
        if result["nakit_tp"] is None:
            # Satırın ilk anlamlı hücresi "1.1" ile başlamalı, "1.1." ile başlamamalı
            first_meaningful = next((c for c in cells if c.strip()), "")
            is_11_row = (
                first_meaningful.strip() == "1.1" or
                (
                    'NAKİT VE NAKİT BENZERLERİ' in row_text and
                    '1.1.1' not in row_text and
                    '1.1.2' not in row_text and
                    '1.1.3' not in row_text and
                    '1.1.4' not in row_text
                )
            )
            if is_11_row and len(cells) > max(tp_col, yp_col):
                nakit_tp = cells[tp_col].strip()
                nakit_yp = cells[yp_col].strip()
                if nakit_tp and nakit_tp != '-' and nakit_tp != '':
                    result["nakit_tp"] = nakit_tp
                    result["nakit_yp"] = nakit_yp if nakit_yp and nakit_yp != '-' else None
                    st.caption(f"  ✅ 1.1 Nakit satırı bulundu: TP={nakit_tp}, YP={nakit_yp}")

        # VARLIKLAR TOPLAMI satırı (en alt, kalın — ara toplamları alma)
        is_varlik_toplami = (
            'VARLIKLAR TOPLAMI' in row_text or
            'AKTİF TOPLAMI' in row_text
        ) and 'FİNANSAL VARLIKLAR' not in row_text  # Ara toplam değil

        if is_varlik_toplami and len(cells) > max(tp_col, yp_col):
            varlik_tp = cells[tp_col].strip()
            varlik_yp = cells[yp_col].strip()
            if varlik_tp and varlik_tp != '-':
                result["varlik_tp"] = varlik_tp
                result["varlik_yp"] = varlik_yp if varlik_yp and varlik_yp != '-' else None
                st.caption(f"  ✅ VARLIKLAR TOPLAMI satırı bulundu: TP={varlik_tp}, YP={varlik_yp}")

    return result


# =============================================================================
# AŞAMA 2: BANKA ADI VE DÖNEMİ LLM İLE ÇEK (sadece metadata)
# =============================================================================
def extract_metadata_llm(llm, raw_text_head: str) -> tuple[str, str]:
    prompt = f"""Aşağıdaki banka raporunun başlangıç metninden:
1. Bankanın TAM RESMİ ADINI
2. Raporun Cari Dönem TARİHİNİ (örn: "30 Haziran 2025")
bul ve JSON olarak döndür.

Format: {{"banka_adi": "...", "donem": "..."}}

Bulamazsan: {{"banka_adi": "Bilinmeyen Banka", "donem": "Bilinmiyor"}}

METİN:
{raw_text_head[:2000]}
"""
    try:
        response = llm.invoke(prompt)
        content = response.content if hasattr(response, 'content') else str(response)
        import json
        # JSON bloğunu çıkar
        match = re.search(r'\{[^{}]+\}', content, re.DOTALL)
        if match:
            data = json.loads(match.group())
            return data.get("banka_adi", "Bilinmeyen Banka"), data.get("donem", "Bilinmiyor")
    except Exception:
        pass
    return "Bilinmeyen Banka", "Bilinmiyor"


# =============================================================================
# AŞAMA 3: DETERMINISTIK BAŞARISIZ OLURSA LLM FALLBACK
# =============================================================================
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
        "1.1 Nakit ve Nakit Benzerleri Cari Dönem TP YP",
        "VARLIKLAR TOPLAMI Cari Dönem TP YP",
        "KONSOLİDE OLMAYAN BİLANÇO FİNANSAL DURUM TABLOSU VARLIKLAR",
    ]
    retrieved = set()
    for q in queries:
        for d in ensemble_retriever.invoke(q):
            retrieved.add(d.page_content)
    return f"## FİLTRELENMİŞ TABLOLAR:\n{table_context}\n\n## EK BAĞLAM:\n" + "\n---\n".join(retrieved)


def llm_fallback_extract(structured_llm, context: str) -> DenetimVerileri:
    """Deterministik parse başarısız olursa LLM ile dene."""
    prompt = f"""
Sen kıdemli bir banka denetçisisin.

## TABLO YAPISI:
Bu Türk bankası bilançosu tablosunda sütun sırası SOLDAN SAĞA:
  [Satır No] | [Açıklama] | [Dipnot] | Cari TP | Cari YP | Cari Toplam | Önceki TP | Önceki YP | Önceki Toplam

## GÖREV:
1. "1.1  Nakit ve Nakit Benzerleri" satırını bul (1.1.1, 1.1.2, 1.1.3 DEĞİL — tam "1.1")
   → nakit_benzerleri_tp = Cari TP değeri (tabloda nasılsa birebir al)
   → nakit_benzerleri_yp = Cari YP değeri (tabloda nasılsa birebir al)

2. "VARLIKLAR TOPLAMI" satırını bul (tablonun en son satırı)
   → varliklar_toplami_tp = Cari TP değeri
   → varliklar_toplami_yp = Cari YP değeri

## KESİN KURALLAR:
- SADECE "Cari Dönem" sütunları. "Önceki Dönem" = YASAK.
- Rakamları tabloda nasılsa birebir al ("1,234,567") ise öyle yaz).
- Net göremiyorsan None döndür. Asla tahmin etme.

## TABLO:
{context[:8000]}
"""
    return structured_llm.invoke(prompt)


# =============================================================================
# DOĞRULAMA
# =============================================================================
def validate(result: DenetimVerileri, context: str) -> list[str]:
    issues = []
    if result.nakit_benzerleri_tp and result.varliklar_toplami_tp:
        if to_num(result.nakit_benzerleri_tp) > to_num(result.varliklar_toplami_tp):
            issues.append("HATA: Nakit TP > Varlık Toplamı TP — mantık hatası.")
    if result.nakit_benzerleri_yp and result.varliklar_toplami_yp:
        if to_num(result.nakit_benzerleri_yp) > to_num(result.varliklar_toplami_yp):
            issues.append("HATA: Nakit YP > Varlık Toplamı YP — mantık hatası.")
    if result.varliklar_toplami_tp and result.varliklar_toplami_yp:
        if result.varliklar_toplami_tp == result.varliklar_toplami_yp:
            issues.append("UYARI: TP ve YP değerleri aynı — sütun karışmış olabilir.")
    return issues


# =============================================================================
# STREAMLIT ARAYÜZÜ
# =============================================================================
st.set_page_config(page_title="Banka Analiz Pipeline V6", layout="wide")
st.title("🏦 Banka Analiz Hüseyin Çiyik V6")
st.caption("Deterministik Parser → LLM Metadata → LLM Fallback (gerekirse)")

with st.sidebar:
    st.subheader("⚙️ Model Ayarları")
    model_name = st.selectbox(
        "LLM Model",
        ["qwen2.5:7b", "gemma2", "mistral-nemo", "llama3.1:8b"],
        index=0,
    )
    num_ctx = st.slider("Context Window", 4096, 16384, 8192, 1024)
    

uploaded_file = st.file_uploader("📄 PDF Rapor Yükle", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 Analiz Et"):
        with st.spinner("Analiz başlatılıyor..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                st.info("📄 PDF → Markdown...")
                converter = DocumentConverter()
                md_text = converter.convert(tmp_path).document.export_to_markdown()
                os.remove(tmp_path)
                st.success(f"✅ Markdown hazır ({len(md_text):,} karakter)")

                with st.expander("📝 Ham Markdown (ilk 4000 karakter)"):
                    st.text(md_text[:4000])

                # --- AŞAMA 1: Deterministik Parse ---
                st.info("🔢 Aşama 1: Deterministik tablo parsing...")
                with st.expander("🔍 Parser Detayları"):
                    parsed = deterministic_parse(md_text)

                nakit_tp  = parsed["nakit_tp"]
                nakit_yp  = parsed["nakit_yp"]
                varlik_tp = parsed["varlik_tp"]
                varlik_yp = parsed["varlik_yp"]

                det_success = all([nakit_tp, nakit_yp, varlik_tp, varlik_yp])

                # --- LLM (metadata + fallback) ---
                llm = ChatOllama(
                    model=model_name,
                    temperature=0,
                    num_ctx=num_ctx,
                    format="json",
                    seed=42,
                )
                structured_llm = llm.with_structured_output(DenetimVerileri)

                # Banka adı ve dönem her zaman LLM ile
                st.info("🤖 Aşama 2: Banka adı ve dönem tespiti...")
                banka_adi, donem = extract_metadata_llm(llm, md_text)
                st.success(f"✅ Banka: {banka_adi} | Dönem: {donem}")

                if det_success:
                    st.success("✅ Deterministik parse başarılı — LLM'e gerek yok.")
                    final = DenetimVerileri(
                        banka_adi=banka_adi,
                        donem=donem,
                        nakit_benzerleri_tp=nakit_tp,
                        nakit_benzerleri_yp=nakit_yp,
                        varliklar_toplami_tp=varlik_tp,
                        varliklar_toplami_yp=varlik_yp,
                    )
                else:
                    st.warning(
                        f"⚠️ Deterministik parse eksik "
                        f"(nakit_tp={nakit_tp}, nakit_yp={nakit_yp}, "
                        f"varlik_tp={varlik_tp}, varlik_yp={varlik_yp}). "
                        f"LLM fallback başlatılıyor..."
                    )
                    context = run_hybrid_retrieval(md_text)
                    with st.expander("🔍 LLM Fallback Context"):
                        st.text(context[:5000])
                    llm_result = llm_fallback_extract(structured_llm, context)
                    # Deterministik başarılı olanları koru, başarısızları LLM ile doldur
                    final = DenetimVerileri(
                        banka_adi=banka_adi,
                        donem=donem,
                        nakit_benzerleri_tp=nakit_tp or llm_result.nakit_benzerleri_tp,
                        nakit_benzerleri_yp=nakit_yp or llm_result.nakit_benzerleri_yp,
                        varliklar_toplami_tp=varlik_tp or llm_result.varliklar_toplami_tp,
                        varliklar_toplami_yp=varlik_yp or llm_result.varliklar_toplami_yp,
                    )

                # Doğrulama
                issues = validate(final, md_text)
                if issues:
                    st.warning("⚠️ Doğrulama sorunları:")
                    for i in issues:
                        st.caption(f"  → {i}")
                else:
                    st.success("✅ Doğrulama geçti.")

                # --- SONUÇLAR ---
                st.divider()
                st.subheader(f"🏦 {final.banka_adi}")
                st.caption(f"📅 {final.donem}")

                
                col1, col2 = st.columns(2)

                def fmt(v): return str(v) if v else "⚠️ Bulunamadı"

                with col1:
                    st.metric("1.1 Nakit Benzerleri — TP", fmt(final.nakit_benzerleri_tp))
                    st.metric("VARLIKLAR TOPLAMI — TP", fmt(final.varliklar_toplami_tp))
                with col2:
                    st.metric("1.1 Nakit Benzerleri — YP", fmt(final.nakit_benzerleri_yp))
                    st.metric("VARLIKLAR TOPLAMI — YP", fmt(final.varliklar_toplami_yp))

                with st.expander("📋 Ham JSON"):
                    st.json(final.model_dump())

                # Excel
                wb = Workbook()
                ws = wb.active
                ws.title = "Denetim Verileri"
                ws.append(["Banka Adı", "Dönem", "Nakit ve Nakit Benzerleri TP", "Nakit ve Nakit Benzerleri YP", "Varlık Toplam TP", "Varlık Toplam YP"])
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