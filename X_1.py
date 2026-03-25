
import os
from typing import Optional
from langchain_openai import ChatOpenAI
from langchain_community.document_loaders import PyPDFLoader
from pydantic import BaseModel, Field
from openpyxl import load_workbook

# 1. ADIM: Veri Şemasını Tanımlama (Pydantic)
# LLM'in çıktıyı kafasına göre değil, tam bu formatta vermesini zorunluyoruz.
class DenetimVerileri(BaseModel):
    banka_adi: str = Field(description="Raporun ait olduğu bankanın tam ticari unvanı")
    donem: str = Field(description="Raporun ait olduğu cari yıl ve dönem (Örn: 2024 3. Çeyrek)")
    net_kar: Optional[int] = Field(description="Dönem net kârı veya zararı. Sadece sayısal değer.")
    toplam_aktifler: Optional[int] = Field(description="Bilançodaki toplam aktifler tutarı.")

# 2. ADIM: Ayarlar ve Model Kurulumu
# API Key'inizi buraya veya environment variable olarak ekleyin
os.environ["OPENAI_API_KEY"] = "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" 

# Finansal verilerde hata payı olmaması için temperature=0 (yaratıcılık kapalı) kullanıyoruz.
llm = ChatOpenAI(model="gpt-4o", temperature=0)

# Modeli, yukarıdaki şemayı (JSON) çıktı verecek şekilde ayarlıyoruz.
structured_llm = llm.with_structured_output(DenetimVerileri)

def veriyi_cek_ve_yaz(pdf_yolu, excel_sablon_yolu, cikti_yolu):
    print(f"{pdf_yolu} işleniyor...")

    # --- A. PDF OKUMA ---
    try:
        loader = PyPDFLoader(pdf_yolu)
        pages = loader.load()
        
        # Maliyet ve hız için tüm rapor yerine sadece ilgili sayfaları birleştirebilirsiniz.
        # Şimdilik ilk 10 sayfayı alalım örnek olarak:
        full_text = "".join([p.page_content for p in pages[:11]])
        
    except Exception as e:
        print(f"PDF okuma hatası: {e}")
        return

    # --- B. AI İLE VERİ ÇIKARMA ---
    print("AI verileri analiz ediyor...")
    prompt = f"""
    Aşağıdaki banka denetim raporu metnini analiz et ve istenen finansal verileri çıkar.
    Eğer veriyi bulamazsan ilgili alan için null/None döndür. Asla tahmin yürütme.
    
    RAPOR METNİ:
    {full_text}
    """
    
    # Yapılandırılmış veriyi alıyoruz
    veri_sonucu = structured_llm.invoke(prompt)
    print(f"Çıkarılan Veriler: {veri_sonucu}")

    # --- C. EXCEL'E YAZMA (OpenPyXL) ---
    print("Excel'e yazılıyor...")
    try:
        # Mevcut şablonu yükle (data_only=False formülleri korur)
        wb = load_workbook(excel_sablon_yolu)
        ws = wb.active  # Veya spesifik sayfa adı: wb["VeriSayfasi"]

        # Verileri spesifik hücrelere haritalıyoruz
        # Bu hücre adreslerini (B2, B3 vs.) kendi şablonunuza göre değiştirin.
        ws['B2'] = veri_sonucu.banka_adi
        ws['B3'] = veri_sonucu.donem
        ws['B4'] = veri_sonucu.net_kar
        ws['B5'] = veri_sonucu.toplam_aktifler

        # Yeni dosyayı kaydet
        wb.save(cikti_yolu)
        print(f"İşlem tamamlandı! Dosya kayıt edildi: {cikti_yolu}")

    except Exception as e:
        print(f"Excel yazma hatası: {e}")

# --- ÇALIŞTIRMA ---
if __name__ == "__main__":
    # Dosya yollarını kendi bilgisayarınıza göre düzenleyin
    pdf_dosyasi = "ornek_banka_raporu.pdf"
    sablon_dosyasi = "sablon.xlsx"
    sonuc_dosyasi = "sonuc_2024.xlsx"

    # Fonksiyonu çağır
    # (Dosyalar gerçekten varsa çalışır)
    # veriyi_cek_ve_yaz(pdf_dosyasi, sablon_dosyasi, sonuc_dosyasi)