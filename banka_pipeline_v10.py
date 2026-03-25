import streamlit as st
import os, tempfile, re, json, time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from rapidfuzz import fuzz
from docling.document_converter import DocumentConverter
from langchain_groq import ChatGroq

# =============================================================================
# YAPILANDIRMA
# =============================================================================
FIELDS_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "fields_config.json")
FUZZY_THRESHOLD = 72

GROQ_MODELS = {
    "qwen-qwq-32b":              "⭐ Qwen QwQ 32b — Finansal tablo için önerilen",
    "llama-3.3-70b-versatile":   "Llama 3.3 70b — Güçlü genel amaçlı",
    "llama-3.1-70b-versatile":   "Llama 3.1 70b — Stabil alternatif",
    "mixtral-8x7b-32768":        "Mixtral 8x7b — 32k token context",
    "gemma2-9b-it":              "Gemma2 9b — Hafif ve hızlı",
}


# =============================================================================
# FIELDS CONFIG
# =============================================================================
@st.cache_data
def load_fields_config() -> list[dict]:
    with open(FIELDS_CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# =============================================================================
# YARDIMCI
# =============================================================================
def to_num(val_str) -> int:
    if not val_str:
        return 0
    clean = re.sub(r'[^\d]', '', str(val_str))
    return int(clean) if clean else 0


# =============================================================================
# DETERMINISTIK PARSER
# =============================================================================
def parse_markdown_table_rows(markdown_text: str) -> list[dict]:
    """
    Docling markdown tablo formatı:
    | satır_no | açıklama | dipnot | Cari_TP | Cari_YP | Cari_Toplam | Önceki_TP | Önceki_YP | Önceki_Toplam |

    Başlık satırı tam olarak 'TP' ve 'YP' hücrelerini içerir.
    Veri satırlarında bazı hücreler boş olabilir — bunu pipe sayısıyla değil
    header'dan tespit edilen sabit indeksle çözüyoruz.
    """
    lines = markdown_text.split('\n')
    rows = []
    tp_col, yp_col, donem_col = -1, -1, -1
    expected_cell_count = -1  # Header satırındaki toplam hücre sayısı
    current_section = "VARLIKLAR"
    header_found = False
    in_bilanco = False

    for line in lines:
        if not line.strip():
            continue
        upper = line.upper()

        # Bilanço tablosuna giriş
        if 'BİLANÇO' in upper or 'FINANSAL DURUM TABLOSU' in upper:
            in_bilanco = True

        # Bölüm takibi
        if 'VARLIKLAR' in upper and '|' not in line:
            current_section = "VARLIKLAR"
        elif 'YÜKÜMLÜLÜK' in upper and '|' not in line:
            current_section = "YUKUMLULUKLER"
        elif 'ÖZKAYNAK' in upper and '|' not in line:
            current_section = "OZKAYNAK"
        elif ('KÂR' in upper or 'KAR' in upper) and 'ZARAR' in upper and '|' not in line:
            current_section = "GELIR_TABLOSU"

        if '|' not in line:
            continue

        cells = [c.strip() for c in line.strip().strip('|').split('|')]

        # Ayırıcı satırları atla
        if all(re.match(r'^[-:]+$', c) or c == '' for c in cells):
            continue

        upper_cells = [c.upper() for c in cells]

        # TP/YP başlık satırını bul — tam eşleşme (sadece 'TP', 'YP' olan hücreler)
        if not header_found and 'TP' in upper_cells and 'YP' in upper_cells:
            for i, c in enumerate(upper_cells):
                if c.strip() == 'TP' and tp_col == -1:
                    tp_col = i
                elif c.strip() == 'YP' and yp_col == -1 and tp_col != -1:
                    yp_col = i
                    break
            if tp_col != -1 and yp_col != -1:
                header_found = True
                expected_cell_count = len(cells)
            continue

        if not header_found:
            continue

        # Veri satırı — hücre sayısı farklıysa padding uygula
        if len(cells) < expected_cell_count:
            cells = cells + [''] * (expected_cell_count - len(cells))

        row_text = ' '.join(cells).strip()
        if not row_text:
            continue

        # Sadece sayısal veri içeren satırları al (en az bir rakam olmalı)
        has_number = any(re.search(r'\d{3,}', c) for c in cells)
        if not has_number and not any(c.strip() for c in cells[:2]):
            continue

        rows.append({
            "cells": cells,
            "row_text": row_text,
            "section": current_section,
            "tp_col": tp_col,
            "yp_col": yp_col,
            "donem_col": donem_col,
        })

    return rows


def fuzzy_match_row(target_variants: list[str], row_text: str,
                    exclude_variants: list[str] = None) -> int:
    row_clean = re.sub(r'[^\w\s]', '', row_text.lower())

    # Exclude kontrolü: eğer satır bir exclude varyantını içeriyorsa → 0 döndür
    if exclude_variants:
        for excl in exclude_variants:
            excl_clean = re.sub(r'[^\w\s]', '', excl.lower())
            if excl_clean and excl_clean in row_clean:
                return 0

    best = 0
    for variant in target_variants:
        score = fuzz.partial_ratio(
            re.sub(r'[^\w\s]', '', variant.lower()), row_clean
        )
        if score > best:
            best = score
    return best


def extract_value_from_row(row: dict, column_type: str) -> str | None:
    cells = row["cells"]
    tp_col, yp_col, donem_col = row["tp_col"], row["yp_col"], row["donem_col"]

    idx = -1
    if column_type == "cari_tp" and tp_col != -1:
        idx = tp_col
    elif column_type == "cari_yp" and yp_col != -1:
        idx = yp_col
    elif column_type == "cari_donem":
        idx = donem_col if donem_col != -1 else tp_col

    if idx == -1 or idx >= len(cells):
        return None
    val = cells[idx].strip()
    return None if not val or val in ['-', '', 'None', 'null'] else val


def extract_all_fields(markdown_text: str, fields_config: list[dict],
                       threshold: int) -> tuple[dict, list[dict]]:
    rows = parse_markdown_table_rows(markdown_text)
    results, match_log = {}, []

    for field in fields_config:
        fid = field["id"]
        section = field["section"]
        candidate_rows = [r for r in rows if r["section"] == section] or rows

        best_score, best_value, best_row_text = 0, None, None
        exclude = field.get("exclude_variants", [])
        for row in candidate_rows:
            score = fuzzy_match_row(field["variants"], row["row_text"], exclude)
            if score > best_score:
                val = extract_value_from_row(row, field["column"])
                if val:
                    best_score, best_value = score, val
                    best_row_text = row["row_text"][:80]

        if best_score >= threshold and best_value:
            results[fid] = best_value
            match_log.append({
                "field": fid, "label": field["label"], "value": best_value,
                "score": best_score, "matched_row": best_row_text, "status": "✅"
            })
        else:
            results[fid] = None
            match_log.append({
                "field": fid, "label": field["label"], "value": None,
                "score": best_score, "matched_row": best_row_text or "—",
                "status": "⚠️" if best_score > 50 else "❌"
            })

    return results, match_log


# =============================================================================
# GROQ LLM
# =============================================================================
def get_groq_llm(api_key: str, model: str) -> ChatGroq:
    return ChatGroq(api_key=api_key, model=model, temperature=0, max_tokens=4096)


def extract_metadata_llm(llm: ChatGroq, raw_text_head: str) -> tuple[str, str]:
    prompt = f"""Aşağıdaki banka raporundan bankanın TAM RESMİ ADINI ve Cari dönem tarihini bul.
SADECE JSON döndür: {{"banka_adi": "...", "donem": "..."}}
METİN: {raw_text_head[:2000]}"""
    try:
        response = llm.invoke(prompt)
        content = response.content if hasattr(response, 'content') else str(response)
        match = re.search(r'\{[^{}]+\}', content, re.DOTALL)
        if match:
            data = json.loads(match.group())
            return data.get("banka_adi", "Bilinmeyen"), data.get("donem", "Bilinmiyor")
    except Exception:
        pass
    return "Bilinmeyen Banka", "Bilinmiyor"


def llm_fill_missing(llm: ChatGroq, missing_fields: list[dict],
                     context_snippet: str) -> tuple[dict, float]:
    if not missing_fields:
        return {}, 0.0

    fields_desc = "\n".join([f"- {f['id']}: {f['label']}" for f in missing_fields[:30]])

    prompt = f"""Türk bankası bilançosundan (KONSOLİDE OLMAYAN BİLANÇO / FİNANSAL DURUM TABLOSU) eksik değerleri bul.

## TABLO SÜTUN YAPISI (soldan sağa):
[Satır No] | [Açıklama] | [Dipnot] | Cari TP | Cari YP | Cari Toplam | Önceki TP | Önceki YP | Önceki Toplam

## KRİTİK KURALLAR:
1. SADECE "Cari Dönem" sütunları (1. ve 2. sayısal sütun). Önceki Dönem = KESİNLİKLE ALMA.
2. "2.4 İtfa Edilmiş Maliyeti ile Ölçülen" için SADECE 2.4 satırını al — "II. İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar (Net)" ana başlığını ALMA.
3. "1.1 Nakit ve Nakit Benzerleri" için SADECE 1.1 satırını al — 1.1.1, 1.1.2, 1.1.3, 1.1.4 alt satırlarını ALMA.
4. "2.5 Beklenen Zarar Karşılıkları" için SADECE 2.5 satırını al — 1.1.4 satırını ALMA.
5. "IV. Ortaklık Yatırımları" için SADECE IV. ana satırını al — 4.1, 4.2, 4.3 alt satırlarını ALMA.
6. Değeri net göremiyorsan null döndür. Tahmin etme.
7. Sayıları tabloda nasılsa birebir al (örn: "5,754,499").
8. SADECE JSON döndür: {{"alan_id": "deger_veya_null"}}

## BULUNAMAYAN ALANLAR:
{fields_desc}

## TABLO:
{context_snippet[:6000]}"""

    t0 = time.time()
    try:
        response = llm.invoke(prompt)
        elapsed = time.time() - t0
        content = response.content if hasattr(response, 'content') else str(response)
        match = re.search(r'\{.*\}', content, re.DOTALL)
        if match:
            return json.loads(match.group()), elapsed
    except Exception as e:
        st.caption(f"LLM fallback hatası: {e}")
    return {}, time.time() - t0


# =============================================================================
# TEK RAPOR İŞLEME (hem tekli hem çoklu için ortak)
# =============================================================================
def process_single_report(pdf_bytes: bytes, filename: str,
                           fields_config: list[dict], llm: ChatGroq,
                           fuzzy_threshold: int,
                           use_llm_fallback: bool) -> dict:
    """
    Tek bir PDF'i işler. Döndürür:
    {banka_adi, donem, filename, results, match_log,
     det_found, llm_filled, total, elapsed_sec, errors}
    """
    t_start = time.time()
    errors = []

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        converter = DocumentConverter()
        md_text = converter.convert(tmp_path).document.export_to_markdown()
        os.remove(tmp_path)
    except Exception as e:
        return {"filename": filename, "errors": [f"PDF parse hatası: {e}"],
                "banka_adi": "HATA", "donem": "—",
                "results": {}, "match_log": [],
                "det_found": 0, "llm_filled": 0,
                "total": len(fields_config), "elapsed_sec": 0}

    # Deterministik
    results, match_log = extract_all_fields(md_text, fields_config, fuzzy_threshold)
    det_found = sum(1 for v in results.values() if v)

    # Metadata
    try:
        banka_adi, donem = extract_metadata_llm(llm, md_text)
    except Exception as e:
        banka_adi, donem = filename, "Bilinmiyor"
        errors.append(f"Metadata LLM hatası: {e}")

    # LLM Fallback
    llm_filled = 0
    if use_llm_fallback:
        missing_fields = [f for f in fields_config if results.get(f["id"]) is None]
        if missing_fields:
            bilanco_start = md_text.find('BİLANÇO')
            ctx = (md_text[max(0, bilanco_start):bilanco_start + 12000]
                   if bilanco_start != -1 else md_text[:12000])

            for i in range(0, len(missing_fields), 30):
                batch = missing_fields[i:i+30]
                try:
                    llm_results, _ = llm_fill_missing(llm, batch, ctx)
                    for fid, val in llm_results.items():
                        if val and val not in ["null", "None", ""] \
                                and results.get(fid) is None:
                            results[fid] = val
                            for log in match_log:
                                if log["field"] == fid:
                                    log["value"] = val
                                    log["status"] = "🤖"
                                    break
                            llm_filled += 1
                except Exception as e:
                    errors.append(f"LLM batch hatası: {e}")
                time.sleep(0.5)  # Groq rate limit

    return {
        "filename": filename,
        "banka_adi": banka_adi,
        "donem": donem,
        "results": results,
        "match_log": match_log,
        "det_found": det_found,
        "llm_filled": llm_filled,
        "total": len(fields_config),
        "elapsed_sec": round(time.time() - t_start, 1),
        "errors": errors,
        "md_text": md_text,
    }


# =============================================================================
# EXCEL EXPORT — ÇOK SAYFALALI (her rapor ayrı sayfa)
# =============================================================================
def create_multi_excel(reports: list[dict], fields_config: list[dict]) -> BytesIO:
    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    section_fill = PatternFill("solid", fgColor="D6E4F0")
    missing_fill = PatternFill("solid", fgColor="FFE0E0")
    llm_fill_style = PatternFill("solid", fgColor="FFF3CD")

    section_labels = {
        "VARLIKLAR": "VARLIKLAR",
        "YUKUMLULUKLER": "YÜKÜMLÜLÜKLER",
        "OZKAYNAK": "ÖZKAYNAKLAR",
        "GELIR_TABLOSU": "GELİR TABLOSU",
    }

    # --- Özet sayfası ---
    ws_ozet = wb.active
    ws_ozet.title = "Özet"
    ozet_headers = ["Dosya", "Banka Adı", "Dönem", "Toplam Alan",
                    "Deterministik", "LLM ile", "Bulunamayan", "Süre (sn)"]
    ws_ozet.append(ozet_headers)
    for cell in ws_ozet[1]:
        cell.fill = header_fill
        cell.font = header_font

    for r in reports:
        found_total = r["det_found"] + r["llm_filled"]
        ws_ozet.append([
            r["filename"], r["banka_adi"], r["donem"], r["total"],
            r["det_found"], r["llm_filled"],
            r["total"] - found_total, r["elapsed_sec"]
        ])

    for col, width in zip("ABCDEFGH", [30, 35, 18, 12, 14, 10, 12, 10]):
        ws_ozet.column_dimensions[col].width = width

    # --- Her rapor için ayrı sayfa ---
    for r in reports:
        safe_name = re.sub(r'[\\/*?:\[\]]', '_', r["banka_adi"])[:25]
        donem_short = r["donem"].replace(" ", "_")[:12]
        sheet_name = f"{safe_name}_{donem_short}"[:31]

        ws = wb.create_sheet(title=sheet_name)
        ws.append(["Banka Adı", r["banka_adi"]])
        ws.append(["Dönem", r["donem"]])
        ws.append(["Dosya", r["filename"]])
        ws.append([])
        ws.append(["Alan ID", "Açıklama", "Bölüm", "Değer", "Kaynak"])

        for cell in ws[5]:
            cell.fill = header_fill
            cell.font = header_font

        current_section = None
        for field in fields_config:
            fid = field["id"]
            section = field["section"]
            value = r["results"].get(fid)
            log_entry = next((l for l in r["match_log"] if l["field"] == fid), {})
            status = log_entry.get("status", "❓")

            if section != current_section:
                ws.append([])
                ws.append([section_labels.get(section, section)])
                for cell in ws[ws.max_row]:
                    cell.fill = section_fill
                    cell.font = Font(bold=True)
                current_section = section

            kaynak = "Deterministik" if status == "✅" else ("LLM" if status == "🤖" else "—")
            ws.append([fid, field["label"], section, value or "—", kaynak])

            row_idx = ws.max_row
            if not value:
                for cell in ws[row_idx]:
                    cell.fill = missing_fill
            elif kaynak == "LLM":
                for cell in ws[row_idx]:
                    cell.fill = llm_fill_style

        for col, width in [('A', 35), ('B', 60), ('C', 18), ('D', 20), ('E', 15)]:
            ws.column_dimensions[col].width = width

    # --- Karşılaştırma sayfası (tüm raporlar yan yana) ---
    if len(reports) > 1:
        ws_kars = wb.create_sheet("Karşılaştırma")
        # Başlık satırı: Alan | Rapor1 | Rapor2 | ...
        header = ["Alan ID", "Açıklama"] + [
            f"{r['banka_adi']}\n{r['donem']}" for r in reports
        ]
        ws_kars.append(header)
        for cell in ws_kars[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(wrap_text=True)
        ws_kars.row_dimensions[1].height = 35

        for field in fields_config:
            fid = field["id"]
            row_data = [fid, field["label"]]
            for r in reports:
                row_data.append(r["results"].get(fid) or "—")
            ws_kars.append(row_data)

        ws_kars.column_dimensions['A'].width = 35
        ws_kars.column_dimensions['B'].width = 55
        for i in range(len(reports)):
            col_letter = chr(ord('C') + i)
            ws_kars.column_dimensions[col_letter].width = 20

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# =============================================================================
# STREAMLIT ARAYÜZÜ
# =============================================================================
st.set_page_config(page_title="Banka Analiz Pipeline V10", layout="wide")
st.title("🏦 Banka Analiz Pipeline V10")
st.caption("14 hedef satır | 28 alan (TP+YP) | Exclude logic | Groq API | Çoklu rapor")

# --- Sidebar ---
with st.sidebar:
    st.subheader("🔑 Groq API")
    groq_api_key = st.text_input(
        "Groq API Key", type="password",
        value=os.getenv("GROQ_API_KEY", ""),
        help="console.groq.com → ücretsiz"
    )
    if groq_api_key:
        st.success("✅ API key girildi")
    else:
        st.warning("⚠️ API key gerekli")
        st.markdown("[→ Groq API Key Al](https://console.groq.com)")

    st.divider()
    st.subheader("⚙️ Model Seçimi")
    selected_model = st.selectbox(
        "Groq Model",
        options=list(GROQ_MODELS.keys()),
        format_func=lambda x: GROQ_MODELS[x],
        index=0,
    )
    st.caption("Qwen QwQ 32b, Türkçe finansal tablolarda en başarılı.")

    st.divider()
    st.subheader("🔧 Parser Ayarları")
    fuzzy_threshold = st.slider("Fuzzy Eşleşme Eşiği", 50, 95, FUZZY_THRESHOLD, 1)
    use_llm_fallback = st.toggle("LLM Fallback", value=True)

    st.divider()
    st.subheader("ℹ️ Groq Limitleri")
    st.caption("Ücretsiz:\n• 30 req/dk\n• Qwen/Llama 70b dahil\n• Çoklu raporda otomatik bekleme")

# Fields config yükle
try:
    fields_config = load_fields_config()
    st.sidebar.success(f"✅ {len(fields_config)} alan yüklendi")
except FileNotFoundError:
    st.sidebar.error("❌ fields_config.json bulunamadı!")
    st.stop()

# --- Ana alan: çoklu dosya yükleme ---
uploaded_files = st.file_uploader(
    "📄 PDF Rapor(lar) Yükle",
    type="pdf",
    accept_multiple_files=True,
    help="Birden fazla rapor seçebilirsiniz — her biri ayrı sayfa olarak Excel'e aktarılır."
)

if uploaded_files:
    st.info(f"**{len(uploaded_files)} rapor** yüklendi: "
            f"{', '.join(f.name for f in uploaded_files)}")

    if not groq_api_key:
        st.error("❌ Groq API key gerekli.")
        st.stop()

    if st.button("🚀 Tümünü Analiz Et"):
        llm = get_groq_llm(groq_api_key, selected_model)
        all_reports = []

        overall_progress = st.progress(0, f"0/{len(uploaded_files)} rapor işlendi")

        for idx, uploaded_file in enumerate(uploaded_files):
            st.divider()
            st.subheader(f"📄 {uploaded_file.name}")

            with st.status(f"İşleniyor: {uploaded_file.name}", expanded=True) as status:
                st.write("📄 PDF → Markdown...")
                st.write("🔢 Deterministik parser çalışıyor...")
                st.write("🤖 Groq LLM (gerekirse)...")

                report = process_single_report(
                    pdf_bytes=uploaded_file.getvalue(),
                    filename=uploaded_file.name,
                    fields_config=fields_config,
                    llm=llm,
                    fuzzy_threshold=fuzzy_threshold,
                    use_llm_fallback=use_llm_fallback,
                )

                if report["errors"]:
                    status.update(label=f"⚠️ {uploaded_file.name} — hatalar var",
                                  state="error")
                    for err in report["errors"]:
                        st.error(err)
                else:
                    found_total = report["det_found"] + report["llm_filled"]
                    status.update(
                        label=f"✅ {report['banka_adi']} | {report['donem']} "
                              f"— {found_total}/{report['total']} alan | "
                              f"{report['elapsed_sec']}sn",
                        state="complete"
                    )

            all_reports.append(report)

            # Sonuç özeti
            found_total = report["det_found"] + report["llm_filled"]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Banka", report["banka_adi"])
            c2.metric("Dönem", report["donem"])
            c3.metric(
                "Bulunan Alan",
                f"{found_total}/{report['total']}",
                delta=f"Det:{report['det_found']} LLM:{report['llm_filled']}"
            )
            c4.metric("Süre", f"{report['elapsed_sec']}sn")

            # Alan detayları
            sections = {}
            for field in fields_config:
                sections.setdefault(field["section"], []).append(field)

            section_labels_ui = {
                "VARLIKLAR": "📊 Varlıklar",
                "YUKUMLULUKLER": "📋 Yükümlülükler",
                "OZKAYNAK": "💰 Özkaynaklar",
                "GELIR_TABLOSU": "📈 Gelir Tablosu",
            }

            for sec, fields in sections.items():
                sec_found = sum(1 for f in fields if report["results"].get(f["id"]))
                with st.expander(
                    f"{section_labels_ui.get(sec, sec)} — {sec_found}/{len(fields)}"
                ):
                    cols = st.columns(3)
                    for i, field in enumerate(fields):
                        val = report["results"].get(field["id"])
                        log = next(
                            (l for l in report["match_log"] if l["field"] == field["id"]), {}
                        )
                        src = "🤖" if log.get("status") == "🤖" \
                            else f"✅{log.get('score','')}" if val else "❌"
                        with cols[i % 3]:
                            st.metric(
                                field["label"][:40],
                                val if val else "—",
                                delta=src, delta_color="off"
                            )

            # Raporlar arası bekleme (Groq rate limit)
            if idx < len(uploaded_files) - 1:
                with st.spinner("⏳ Groq rate limit bekleniyor (2sn)..."):
                    time.sleep(2)

            overall_progress.progress(
                (idx + 1) / len(uploaded_files),
                f"{idx + 1}/{len(uploaded_files)} rapor işlendi"
            )

        # =============================================================
        # KARŞILAŞTIRMA PANELİ (2+ rapor varsa)
        # =============================================================
        if len(all_reports) > 1:
            st.divider()
            st.subheader("📊 Raporlar Arası Karşılaştırma")

            # Kritik alanları karşılaştır
            key_fields = [
                "nakit_benzerleri_tp", "nakit_benzerleri_yp",
                "varliklar_toplami_tp", "varliklar_toplami_yp",
                "krediler_tp", "krediler_yp",
                "toplam_yukumlulukler_tp", "toplam_ozkaynak_tp",
            ]

            kf_labels = {f["id"]: f["label"] for f in fields_config}

            cols_header = st.columns([2] + [1] * len(all_reports))
            cols_header[0].markdown("**Alan**")
            for i, r in enumerate(all_reports):
                cols_header[i+1].markdown(f"**{r['donem']}**")

            for fid in key_fields:
                if fid not in kf_labels:
                    continue
                row_cols = st.columns([2] + [1] * len(all_reports))
                row_cols[0].caption(kf_labels[fid][:45])
                for i, r in enumerate(all_reports):
                    val = r["results"].get(fid) or "—"
                    row_cols[i+1].metric("", val)

        # =============================================================
        # EXCEL İNDİR
        # =============================================================
        st.divider()
        excel_buf = create_multi_excel(all_reports, fields_config)
        total_found = sum(r["det_found"] + r["llm_filled"] for r in all_reports)
        total_fields = sum(r["total"] for r in all_reports)

        st.success(
            f"✅ **{len(all_reports)} rapor** tamamlandı — "
            f"toplam {total_found}/{total_fields} alan bulundu"
        )
        st.download_button(
            f"📥 Excel İndir ({len(all_reports)} Rapor + Karşılaştırma)",
            excel_buf,
            file_name=f"banka_analiz_{len(all_reports)}_rapor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )