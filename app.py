import streamlit as st
import os
import tempfile
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from langchain_community.document_loaders import PyPDFLoader
from langchain_ollama import ChatOllama
from pydantic import BaseModel, Field
from typing import Optional

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Banka Denetim Analizörü", layout="wide")

st.title("🏦 Banka Denetim Raporu Analizörü (AI)")
st.markdown("""
Bu uygulama, yüklediğiniz banka denetim raporu (PDF) içerisindeki finansal verileri 
**Llama 3.2** modeli ile analiz eder ve formatlı bir Excel dosyasına dönüştürür.
""")

# --- 1. VERİ MODELİ (PYDANTIC) ---
class DenetimVerileri(BaseModel):
    banka_adi: str = Field(description="Raporun en üstünde yer alan bankanın tam ticari unvanı.")
    donem: str = Field(description="Raporun ait olduğu tarih ve dönem bilgisi (Örn: 30 Haziran 2025).")
    
    nakit_benzerleri_tp: Optional[int] = Field(description="Cari Dönem 'Nakit ve Nakit Benzerleri' TP değeri.")
    nakit_benzerleri_yp: Optional[int] = Field(description="Cari Dönem 'Nakit ve Nakit Benzerleri' YP değeri.")
    
    fv_kar_zarar_tp: Optional[int] = Field(description="Cari Dönem 'Gerçeğe Uygun Değer Farkı Kâr Zarara Yansıtılan' TP değeri.")
    fv_kar_zarar_yp: Optional[int] = Field(description="Cari Dönem 'Gerçeğe Uygun Değer Farkı Kâr Zarara Yansıtılan' YP değeri.")
    
    fv_diger_kapsamli_tp: Optional[int] = Field(description="Cari Dönem 'Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan' TP değeri.")
    fv_diger_kapsamli_yp: Optional[int] = Field(description="Cari Dönem 'Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan' YP değeri.")
    
    itfa_edilmis_tp: Optional[int] = Field(description="Cari Dönem 'İtfa Edilmiş Maliyeti ile Ölçülen' TP değeri.")
    itfa_edilmis_yp: Optional[int] = Field(description="Cari Dönem 'İtfa Edilmiş Maliyeti ile Ölçülen' YP değeri.")
    
    beklenen_zarar_tp: Optional[int] = Field(description="Cari Dönem 'Beklenen Zarar Karşılıkları (-)' TP değeri.")
    beklenen_zarar_yp: Optional[int] = Field(description="Cari Dönem 'Beklenen Zarar Karşılıkları (-)' YP değeri.")
    
    ortaklik_yatirimlari_tp: Optional[int] = Field(description="Cari Dönem 'ORTAKLIK YATIRIMLARI' veya 'İştirakler' TP değeri.")
    ortaklik_yatirimlari_yp: Optional[int] = Field(description="Cari Dönem 'ORTAKLIK YATIRIMLARI' veya 'İştirakler' YP değeri.")
    
    maddi_duran_tp: Optional[int] = Field(description="Cari Dönem 'MADDİ DURAN VARLIKLAR (Net)' TP değeri.")
    maddi_duran_yp: Optional[int] = Field(description="Cari Dönem 'MADDİ DURAN VARLIKLAR (Net)' YP değeri.")
    
    maddi_olmayan_tp: Optional[int] = Field(description="Cari Dönem 'MADDİ OLMAYAN DURAN VARLIKLAR (Net)' TP değeri.")
    maddi_olmayan_yp: Optional[int] = Field(description="Cari Dönem 'MADDİ OLMAYAN DURAN VARLIKLAR (Net)' YP değeri.")
    
    diger_aktifler_tp: Optional[int] = Field(description="Cari Dönem 'DİĞER AKTİFLER' TP değeri.")
    diger_aktifler_yp: Optional[int] = Field(description="Cari Dönem 'DİĞER AKTİFLER' YP değeri.")
    
    varliklar_toplami_tp: Optional[int] = Field(description="Cari Dönem 'VARLIKLAR TOPLAMI' TP değeri.")
    varliklar_toplami_yp: Optional[int] = Field(description="Cari Dönem 'VARLIKLAR TOPLAMI' YP değeri.")

# --- 2. EXCEL OLUŞTURMA FONKSİYONU ---
def create_excel_report(data: DenetimVerileri):
    wb = Workbook()
    ws = wb.active
    ws.title = "Denetim Verileri"

    # Tasarım
    bold_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20

    # Genel Bilgiler
    ws['A1'] = "BANKA ADI:"
    ws['B1'] = data.banka_adi
    ws['A2'] = "DÖNEM:"
    ws['B2'] = data.donem
    
    # Başlıklar
    headers = ["KALEMLER", "TP (Türk Parası)", "YP (Yabancı Para)"]
    ws.append([]) 
    ws.append(headers) # 4. Satır

    for cell in ws[4]:
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = center_align

    # Satır Etiketleri ve Değerler
    rows_map = [
        ("Nakit ve Nakit Benzerleri", data.nakit_benzerleri_tp, data.nakit_benzerleri_yp),
        ("G.U.D Farkı K/Z Yansıtılan Finansal Varlıklar", data.fv_kar_zarar_tp, data.fv_kar_zarar_yp),
        ("G.U.D Farkı Diğer Kapsamlı Gelire Yansıtılan", data.fv_diger_kapsamli_tp, data.fv_diger_kapsamli_yp),
        ("İtfa Edilmiş Maliyeti ile Ölçülen", data.itfa_edilmis_tp, data.itfa_edilmis_yp),
        ("Beklenen Zarar Karşılıkları (-)", data.beklenen_zarar_tp, data.beklenen_zarar_yp),
        ("Ortaklık Yatırımları", data.ortaklik_yatirimlari_tp, data.ortaklik_yatirimlari_yp),
        ("Maddi Duran Varlıklar (Net)", data.maddi_duran_tp, data.maddi_duran_yp),
        ("Maddi Olmayan Duran Varlıklar (Net)", data.maddi_olmayan_tp, data.maddi_olmayan_yp),
        ("Diğer Aktifler", data.diger_aktifler_tp, data.diger_aktifler_yp),
        ("VARLIKLAR TOPLAMI", data.varliklar_toplami_tp, data.varliklar_toplami_yp)
    ]

    for i, (label, val_tp, val_yp) in enumerate(rows_map, start=5):
        ws[f'A{i}'] = label
        ws[f'A{i}'].font = Font(bold=True)
        ws[f'B{i}'] = val_tp
        ws[f'C{i}'] = val_yp

    # Excel'i belleğe kaydet
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 3. ANA UYGULAMA AKIŞI ---

# Kenar Çubuğu Ayarları
with st.sidebar:
    st.header("Ayarlar")
    model_name = st.selectbox("Model Seçimi", ["llama3.2", "llama3", "mistral"], index=0)
    page_start = st.number_input("Başlangıç Sayfası (0 indexli)", value=9, min_value=0)
    page_end = st.number_input("Bitiş Sayfası", value=11, min_value=1)
    st.info("Not: Varsayılan olarak kodundaki gibi 9-11 arası sayfaları okur. Tüm PDF için genişletebilirsin.")

# EKSİK OLAN SATIR BUYDU 👇
uploaded_file = st.file_uploader("PDF Raporunu Yükleyin", type="pdf")

if uploaded_file is not None:
    if st.button("Analiz Et ve Excel Oluştur"):
        
        with st.spinner('PDF işleniyor ve AI analiz yapıyor...'):
            try:
                # 1. PDF'i geçici dosyaya kaydet (LangChain loader için gerekli)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name

                # 2. PDF Yükleme ve Metin Çıkarma
                loader = PyPDFLoader(tmp_file_path)
                pages = loader.load()
                
                # Sayfa aralığını kontrol et
                if page_end > len(pages):
                    st.warning(f"PDF sadece {len(pages)} sayfa. Bitiş sayfası ayarlandı.")
                    page_end = len(pages)

                selected_text = "".join([p.page_content for p in pages[page_start:page_end]])
                
                # Geçici dosyayı temizle
                os.remove(tmp_file_path)

                if not selected_text.strip():
                    st.error("Seçilen sayfalardan metin okunamadı!")
                    st.stop()

                # 3. LLM Kurulumu ve Çağrı
                llm = ChatOllama(model=model_name, temperature=0)
                structured_llm = llm.with_structured_output(DenetimVerileri)
                
                prompt = f"""
                Aşağıdaki banka denetim raporu metnini analiz et ve istenen finansal verileri çıkar.
                Eğer veriyi bulamazsan ilgili alan için null/None döndür. Asla tahmin yürütme.
                
                RAPOR METNİ:
                {selected_text}
                """
                
                # AI Analizi
                result = structured_llm.invoke(prompt)
                
                # 4. Sonuçları Göster (Opsiyonel Preview)
                with st.expander("Çıkarılan Ham Verileri Gör"):
                    st.json(result.model_dump())

                # 5. Excel Oluşturma
                excel_file = create_excel_report(result)
                
                st.success("İşlem Başarılı! Excel dosyanız hazır.")
                
                # İndirme Butonu
                st.download_button(
                    label="📥 Excel Raporunu İndir",
                    data=excel_file,
                    file_name=f"analiz_sonucu_{result.donem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Bir hata oluştu: {str(e)}")