import streamlit as st
import anthropic
import base64
import json
import re
import tempfile
import os
import time
from io import BytesIO
from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =============================================================================
# HEDEF VERİ YAPISI
# Tüm çekilecek alanlar — tablo tipine göre gruplandırılmış
# =============================================================================

EXTRACTION_GROUPS = [
    {
        "id": "aktif",
        "name": "AKTİF (VARLIKLAR)",
        "table_hint": "KONSOLİDE OLMAYAN BİLANÇO veya FİNANSAL DURUM TABLOSU - VARLIKLAR bölümü",
        "fields": [
            {"id": "nakit_tp", "label": "Nakit ve Nakit Benzerleri", "col": "TP"},
            {"id": "nakit_yp", "label": "Nakit ve Nakit Benzerleri", "col": "YP"},
            {"id": "menkul_tp", "label": "Menkul Değerler", "col": "TP"},
            {"id": "menkul_yp", "label": "Menkul Değerler", "col": "YP"},
            {"id": "gercege_kzara_tp", "label": "Gerçeğe Uygun Değer Farkı K/Z'a Yansıtılan Finansal Varlıklar", "col": "TP"},
            {"id": "gercege_kzara_yp", "label": "Gerçeğe Uygun Değer Farkı K/Z'a Yansıtılan Finansal Varlıklar", "col": "YP"},
            {"id": "gercege_diger_tp", "label": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar", "col": "TP"},
            {"id": "gercege_diger_yp", "label": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar", "col": "YP"},
            {"id": "itfa_tp", "label": "İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar", "col": "TP"},
            {"id": "itfa_yp", "label": "İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar", "col": "YP"},
            {"id": "krediler_tp", "label": "Krediler", "col": "TP"},
            {"id": "krediler_yp", "label": "Krediler", "col": "YP"},
            {"id": "beklenen_zarar_tp", "label": "Beklenen Zarar Karşılıkları (-)", "col": "TP"},
            {"id": "beklenen_zarar_yp", "label": "Beklenen Zarar Karşılıkları (-)", "col": "YP"},
            {"id": "satis_amacli_tp", "label": "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar (Net)", "col": "TP"},
            {"id": "satis_amacli_yp", "label": "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar (Net)", "col": "YP"},
            {"id": "ortaklik_tp", "label": "Ortaklık Yatırımları", "col": "TP"},
            {"id": "ortaklik_yp", "label": "Ortaklık Yatırımları", "col": "YP"},
            {"id": "maddi_duran_tp", "label": "Maddi Duran Varlıklar (Net)", "col": "TP"},
            {"id": "maddi_duran_yp", "label": "Maddi Duran Varlıklar (Net)", "col": "YP"},
            {"id": "maddi_olmayan_tp", "label": "Maddi Olmayan Duran Varlıklar (Net)", "col": "TP"},
            {"id": "maddi_olmayan_yp", "label": "Maddi Olmayan Duran Varlıklar (Net)", "col": "YP"},
            {"id": "yatirim_gay_tp", "label": "Yatırım Amaçlı Gayrimenkuller (Net)", "col": "TP"},
            {"id": "yatirim_gay_yp", "label": "Yatırım Amaçlı Gayrimenkuller (Net)", "col": "YP"},
            {"id": "cari_vergi_v_tp", "label": "Cari Vergi Varlığı", "col": "TP"},
            {"id": "cari_vergi_v_yp", "label": "Cari Vergi Varlığı", "col": "YP"},
            {"id": "ertelenmis_v_tp", "label": "Ertelenmiş Vergi Varlığı", "col": "TP"},
            {"id": "ertelenmis_v_yp", "label": "Ertelenmiş Vergi Varlığı", "col": "YP"},
            {"id": "diger_aktif_tp", "label": "Diğer Aktifler (Net)", "col": "TP"},
            {"id": "diger_aktif_yp", "label": "Diğer Aktifler (Net)", "col": "YP"},
            {"id": "varliklar_toplami_tp", "label": "VARLIKLAR TOPLAMI", "col": "TP"},
            {"id": "varliklar_toplami_yp", "label": "VARLIKLAR TOPLAMI", "col": "YP"},
        ]
    },
    {
        "id": "pasif",
        "name": "PASİF (YÜKÜMLÜLÜKLER)",
        "table_hint": "KONSOLİDE OLMAYAN BİLANÇO - YÜKÜMLÜLÜKLER bölümü",
        "fields": [
            {"id": "mevduat_tp", "label": "Mevduat", "col": "TP"},
            {"id": "mevduat_yp", "label": "Mevduat", "col": "YP"},
            {"id": "alinan_kred_tp", "label": "Alınan Krediler", "col": "TP"},
            {"id": "alinan_kred_yp", "label": "Alınan Krediler", "col": "YP"},
            {"id": "para_piy_bor_tp", "label": "Para Piyasalarına Borçlar", "col": "TP"},
            {"id": "para_piy_bor_yp", "label": "Para Piyasalarına Borçlar", "col": "YP"},
            {"id": "ihrac_menkul_tp", "label": "İhraç Edilen Menkul Kıymetler (Net)", "col": "TP"},
            {"id": "ihrac_menkul_yp", "label": "İhraç Edilen Menkul Kıymetler (Net)", "col": "YP"},
            {"id": "fonlar_tp", "label": "Fonlar", "col": "TP"},
            {"id": "fonlar_yp", "label": "Fonlar", "col": "YP"},
            {"id": "cari_vergi_b_tp", "label": "Cari Vergi Borcu", "col": "TP"},
            {"id": "cari_vergi_b_yp", "label": "Cari Vergi Borcu", "col": "YP"},
            {"id": "ertelenmis_b_tp", "label": "Ertelenmiş Vergi Borcu", "col": "TP"},
            {"id": "ertelenmis_b_yp", "label": "Ertelenmiş Vergi Borcu", "col": "YP"},
            {"id": "ozkaynak_tp", "label": "Özkaynaklar", "col": "TP"},
            {"id": "ozkaynak_yp", "label": "Özkaynaklar", "col": "YP"},
            {"id": "odenmis_ser_tp", "label": "Ödenmiş Sermaye", "col": "TP"},
            {"id": "odenmis_ser_yp", "label": "Ödenmiş Sermaye", "col": "YP"},
            {"id": "donem_net_kar_tp", "label": "Dönem Net Kar veya Zararı", "col": "TP"},
            {"id": "donem_net_kar_yp", "label": "Dönem Net Kar veya Zararı", "col": "YP"},
            {"id": "yukumluluk_top_tp", "label": "YÜKÜMLÜLÜKLER TOPLAMI", "col": "TP"},
            {"id": "yukumluluk_top_yp", "label": "YÜKÜMLÜLÜKLER TOPLAMI", "col": "YP"},
        ]
    },
    {
        "id": "gelir",
        "name": "GELİR TABLOSU",
        "table_hint": "KONSOLİDE OLMAYAN KAR VEYA ZARAR TABLOSU — Cari dönem sütunu",
        "fields": [
            {"id": "faiz_gel", "label": "Faiz Gelirleri", "col": "Cari Dönem"},
            {"id": "faiz_gid", "label": "Faiz Giderleri (-)", "col": "Cari Dönem"},
            {"id": "net_ucret_kom", "label": "Net Ücret ve Komisyon Gelirleri/Giderleri", "col": "Cari Dönem"},
            {"id": "alinan_ucret", "label": "Alınan Ücret ve Komisyonlar", "col": "Cari Dönem"},
            {"id": "verilen_ucret", "label": "Verilen Ücret ve Komisyonlar (-)", "col": "Cari Dönem"},
            {"id": "kar_yeniden_sinif_olmayan", "label": "Kar veya Zararda Yeniden Sınıflandırılmayacaklar", "col": "Cari Dönem"},
            {"id": "kar_yeniden_sinif_olan", "label": "Kar veya Zararda Yeniden Sınıflandırılacaklar", "col": "Cari Dönem"},
            {"id": "temettü", "label": "Temettü Gelirleri", "col": "Cari Dönem"},
            {"id": "ticari_kz", "label": "Ticari Kar/Zarar (Net)", "col": "Cari Dönem"},
            {"id": "diger_faal_gel", "label": "Diğer Faaliyet Gelirleri", "col": "Cari Dönem"},
            {"id": "beklenen_zarar_gid", "label": "Beklenen Zarar Karşılıkları (-)", "col": "Cari Dönem"},
            {"id": "diger_karsilik_gid", "label": "Diğer Karşılık Giderleri (-)", "col": "Cari Dönem"},
            {"id": "personel_gid", "label": "Personel Giderleri (-)", "col": "Cari Dönem"},
            {"id": "diger_faal_gid", "label": "Diğer Faaliyet Giderleri (-)", "col": "Cari Dönem"},
            {"id": "net_faal_kz", "label": "Net Faaliyet Karı/Zararı", "col": "Cari Dönem"},
            {"id": "surd_faal_vergi_oncesi", "label": "Sürdürülen Faaliyetler Vergi Öncesi K/Z", "col": "Cari Dönem"},
            {"id": "surd_faal_vergi", "label": "Sürdürülen Faaliyetler Vergi Karşılığı", "col": "Cari Dönem"},
            {"id": "donem_net_kz", "label": "Dönem Net Karı/Zararı", "col": "Cari Dönem"},
        ]
    },
    {
        "id": "vergi",
        "name": "VERGİ DETAYI",
        "table_hint": "Sürdürülen faaliyetler vergi karşılığına ilişkin dipnot tablosu",
        "fields": [
            {"id": "ozkaynak_ortaklik_kz", "label": "Özkaynak Yöntemi Uygulanan Ortaklıklardan Kar/Zarar", "col": "Cari Dönem"},
            {"id": "surd_vergi_oncesi", "label": "Sürdürülen Faaliyetler Vergi Öncesi Kar/Zarar", "col": "Cari Dönem"},
            {"id": "surd_vergi_karsiliği", "label": "Sürdürülen Faaliyetler Vergi Karşılığı", "col": "Cari Dönem"},
            {"id": "cari_vergi_karsiliği", "label": "Cari Vergi Karşılığı", "col": "Cari Dönem"},
            {"id": "ertelenmis_vergi_net", "label": "Ertelenmiş Vergi Net", "col": "Cari Dönem"},
            {"id": "ertelenmis_gider_etkisi", "label": "Ertelenmiş Vergi Gider Etkisi (+)", "col": "Cari Dönem"},
            {"id": "ertelenmis_gelir_etkisi", "label": "Ertelenmiş Vergi Gelir Etkisi (-)", "col": "Cari Dönem"},
        ]
    },
    {
        "id": "faiz_detay",
        "name": "FAİZ GELİR/GİDER DETAYI",
        "table_hint": "Faiz gelirlerine ilişkin açıklamalar dipnot tablosu — TL ve YP ayrımı",
        "fields": [
            {"id": "kred_faiz_gel_tl", "label": "Kredilerden Alınan Faiz Gelirleri", "col": "TL"},
            {"id": "kred_faiz_gel_yp", "label": "Kredilerden Alınan Faiz Gelirleri", "col": "YP"},
            {"id": "zorunlu_kars_faiz", "label": "Zorunlu Karşılıklardan Alınan Faizler", "col": "Cari Dönem"},
            {"id": "bankalar_faiz", "label": "Bankalardan Alınan Faizler", "col": "Cari Dönem"},
            {"id": "para_piy_faiz", "label": "Para Piyasası İşlemlerinden Alınan Faizler", "col": "Cari Dönem"},
            {"id": "menkul_faiz", "label": "Menkul Değerlerden Alınan Faizler", "col": "Cari Dönem"},
            {"id": "fin_kira_faiz", "label": "Finansal Kiralama Faiz Gelirleri", "col": "Cari Dönem"},
            {"id": "diger_faiz_gel", "label": "Diğer Faiz Gelirleri", "col": "Cari Dönem"},
            {"id": "mevduat_faiz_gid", "label": "Mevduata Verilen Faizler", "col": "Cari Dönem"},
            {"id": "kred_faiz_gid", "label": "Kullanılan Kredilere Verilen Faizler", "col": "Cari Dönem"},
            {"id": "para_piy_faiz_gid", "label": "Para Piyasası İşlemlerine Verilen Faizler", "col": "Cari Dönem"},
            {"id": "ihrac_menkul_faiz", "label": "İhraç Edilen Menkul Kıymetlere Verilen Faizler", "col": "Cari Dönem"},
            {"id": "kira_faiz_gid", "label": "Kiralama Faiz Giderleri", "col": "Cari Dönem"},
            {"id": "diger_faiz_gid", "label": "Diğer Faiz Giderleri", "col": "Cari Dönem"},
        ]
    },
    {
        "id": "ort_faiz",
        "name": "ORTALAMA FAİZ ORANLARI (%)",
        "table_hint": "Parasal finansal araçlara uygulanan ortalama faiz oranları tablosu — TL ve YP (USD) sütunları",
        "fields": [
            {"id": "vkred_tl", "label": "Verilen Krediler", "col": "TL"},
            {"id": "bank_mev_tl", "label": "Bankalar Mevduatı", "col": "TL"},
            {"id": "diger_mev_tl", "label": "Diğer Mevduat", "col": "TL"},
            {"id": "ihrac_menkul_tl", "label": "İhraç Edilen Menkul Değerler (Net)", "col": "TL"},
            {"id": "vkred_yp", "label": "Verilen Krediler", "col": "YP (USD)"},
            {"id": "bank_mev_yp", "label": "Bankalar Mevduatı", "col": "YP (USD)"},
            {"id": "diger_mev_yp", "label": "Diğer Mevduat", "col": "YP (USD)"},
            {"id": "ihrac_menkul_yp", "label": "İhraç Edilen Menkul Değerler (Net)", "col": "YP (USD)"},
        ]
    },
    {
        "id": "ticari",
        "name": "TİCARİ KAR/ZARAR DETAYI",
        "table_hint": "Ticari kar/zarara ilişkin açıklamalar dipnot tablosu",
        "fields": [
            {"id": "ticari_kar", "label": "Ticari Kar", "col": "Cari Dönem"},
            {"id": "ser_piy_kar", "label": "Sermaye Piyasası İşlemleri Karı", "col": "Cari Dönem"},
            {"id": "turev_kar", "label": "Türev Finansal İşlemlerden Kar", "col": "Cari Dönem"},
            {"id": "kambiyo_kar", "label": "Kambiyo İşlemlerinden Kar", "col": "Cari Dönem"},
            {"id": "ticari_zarar", "label": "Ticari Zarar", "col": "Cari Dönem"},
            {"id": "ser_piy_zarar", "label": "Sermaye Piyasası İşlemleri Zararı", "col": "Cari Dönem"},
            {"id": "turev_zarar", "label": "Türev Finansal İşlemlerden Zarar", "col": "Cari Dönem"},
            {"id": "kambiyo_zarar", "label": "Kambiyo İşlemlerinden Zarar", "col": "Cari Dönem"},
        ]
    },
    {
        "id": "bkzk",
        "name": "BEKLENEN KREDİ ZARAR KARŞILIKLARI",
        "table_hint": "Beklenen zarar karşılıkları ve diğer karşılık giderleri dipnot tablosu",
        "fields": [
            {"id": "birinci_asama", "label": "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)", "col": "Cari Dönem"},
            {"id": "ikinci_asama", "label": "Kredi Riskinde Önemli Artış (İkinci Aşama)", "col": "Cari Dönem"},
            {"id": "ucuncu_asama", "label": "Temerrüt (Üçüncü Aşama)", "col": "Cari Dönem"},
        ]
    },
    {
        "id": "faal_gider",
        "name": "DİĞER FAALİYET GİDERLERİ DETAYI",
        "table_hint": "Diğer faaliyet giderlerine ilişkin bilgiler dipnot tablosu",
        "fields": [
            {"id": "kidem_taz", "label": "Kıdem Tazminat Karşılığı", "col": "Cari Dönem"},
            {"id": "amortisman", "label": "Amortisman Giderleri", "col": "Cari Dönem"},
            {"id": "diger_isletme", "label": "Diğer İşletme Giderleri", "col": "Cari Dönem"},
            {"id": "tfrs16_kira", "label": "TFRS 16 İstisnalarına İlişkin Kiralama Giderleri", "col": "Cari Dönem"},
            {"id": "bakim_onarim", "label": "Bakım ve Onarım Giderleri", "col": "Cari Dönem"},
            {"id": "reklam_ilan", "label": "Reklam ve İlan Giderleri", "col": "Cari Dönem"},
            {"id": "diger_giderler_1", "label": "Diğer Giderler (1)", "col": "Cari Dönem"},
            {"id": "aktif_satis_zarar", "label": "Aktiflerin Satışından Doğan Zararlar", "col": "Cari Dönem"},
            {"id": "diger_2", "label": "Diğer (2)", "col": "Cari Dönem"},
        ]
    },
    {
        "id": "kredi_kalite",
        "name": "KREDİ KALİTE TABLOLARI",
        "table_hint": "Standart nitelikli ve yakın izlemedeki krediler ile sorunlu kredi gruplarına ilişkin dipnot tabloları",
        "fields": [
            {"id": "std_nitelikli", "label": "Standart Nitelikli Krediler", "col": "Cari Dönem"},
            {"id": "yakin_izleme_yeniden_yap", "label": "Yakın İzlemedeki Krediler - Yeniden Yapılandırma Kapsamında Yer Almayanlar", "col": "Cari Dönem"},
            {"id": "yakin_sozlesme_degisik", "label": "Yakın İzlemedeki Krediler - Sözleşme Koşullarında Değişiklik", "col": "Cari Dönem"},
            {"id": "yakin_yeniden_fin", "label": "Yakın İzlemedeki Krediler - Yeniden Finansman", "col": "Cari Dönem"},
            {"id": "konut_kred", "label": "Konut Kredileri (Tüketici+Personel)", "col": "Cari Dönem"},
            {"id": "tasit_kred", "label": "Taşıt Kredileri (Tüketici+Personel)", "col": "Cari Dönem"},
            {"id": "ihtiyac_kred", "label": "İhtiyaç Kredileri (Tüketici+Personel)", "col": "Cari Dönem"},
            {"id": "diger_kred", "label": "Diğer (Tüketici+Personel)", "col": "Cari Dönem"},
            {"id": "kredi_karti", "label": "Kredi Kartları", "col": "Cari Dönem"},
            {"id": "kmh", "label": "KMH", "col": "Cari Dönem"},
            {"id": "gr3_donem_bakiye", "label": "III. Grup - Dönem Sonu Bakiyesi", "col": "Cari Dönem"},
            {"id": "gr3_karsilik", "label": "III. Grup - Karşılık (-)", "col": "Cari Dönem"},
            {"id": "gr3_net_bakiye", "label": "III. Grup - Bilançodaki Net Bakiyesi", "col": "Cari Dönem"},
            {"id": "gr4_donem_bakiye", "label": "IV. Grup - Dönem Sonu Bakiyesi", "col": "Cari Dönem"},
            {"id": "gr4_karsilik", "label": "IV. Grup - Karşılık (-)", "col": "Cari Dönem"},
            {"id": "gr4_net_bakiye", "label": "IV. Grup - Bilançodaki Net Bakiyesi", "col": "Cari Dönem"},
            {"id": "gr5_donem_bakiye", "label": "V. Grup - Dönem Sonu Bakiyesi", "col": "Cari Dönem"},
            {"id": "gr5_karsilik", "label": "V. Grup - Karşılık (-)", "col": "Cari Dönem"},
            {"id": "gr5_net_bakiye", "label": "V. Grup - Bilançodaki Net Bakiyesi", "col": "Cari Dönem"},
        ]
    },
]

# =============================================================================
# SAYFA TESPİT HARİTASI
# Her grup için PDF'de aranacak anahtar kelimeler
# =============================================================================
PAGE_KEYWORDS = {
    "aktif":       ["VARLIKLAR TOPLAMI", "FİNANSAL VARLIKLAR", "Nakit ve Nakit Benzerleri"],
    "pasif":       ["YÜKÜMLÜLÜKLER TOPLAMI", "MEVDUAT", "ÖZKAYNAKLAR"],
    "gelir":       ["FAİZ GELİRLERİ", "FAİZ GİDERLERİ", "DÖNEM NET KARI"],
    "vergi":       ["Sürdürülen Faaliyetler Vergi Karşılığı", "Cari Vergi Karşılığı", "Ertelenmiş Vergi"],
    "faiz_detay":  ["Kredilerden Alınan Faiz", "Zorunlu Karşılıklardan", "Bankalardan Alınan Faizler"],
    "ort_faiz":    ["ortalama faiz", "Verilen Krediler", "Bankalar Mevduatı"],
    "ticari":      ["Ticari Kar", "Sermaye Piyasası İşlemleri", "Kambiyo İşlemlerinden"],
    "bkzk":        ["Beklenen Kredi Zarar", "Birinci Aşama", "İkinci Aşama", "Temerrüt"],
    "faal_gider":  ["Kıdem Tazminat", "Amortisman", "TFRS 16", "Reklam ve İlan"],
    "kredi_kalite":["Standart Nitelikli", "Yakın İzlemedeki", "Tahsil İmkânı Sınırlı", "Tahsili Şüpheli"],
}


# =============================================================================
# PDF → GÖRÜNTÜ DÖNÜŞÜMÜ
# =============================================================================
def pdf_to_images(pdf_bytes: bytes) -> list:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        images = convert_from_path(tmp_path, dpi=200)
    finally:
        os.remove(tmp_path)
    return images


def image_to_base64(image) -> str:
    buf = BytesIO()
    image.save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8")


def find_relevant_pages(images, keywords: list[str], max_pages: int = 4) -> list[int]:
    """
    pdf2image ile dönüştürülmüş sayfalarda metin araması yapamayız.
    Bunun yerine tüm sayfaları Claude'a göndermek yerine
    pdfplumber ile metin tabanlı sayfa tespiti yapıyoruz.
    """
    # Bu fonksiyon pdfplumber ile metin araması yapar
    # ve ilgili sayfa indekslerini döndürür
    return list(range(len(images)))  # Fallback: tümünü döndür


def find_pages_with_pdfplumber(pdf_bytes: bytes, keywords: list[str], max_pages: int = 5) -> list[int]:
    import pdfplumber
    relevant = []
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        with pdfplumber.open(tmp_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = (page.extract_text() or "").upper()
                score = sum(1 for kw in keywords if kw.upper() in text)
                if score > 0:
                    relevant.append((i, score))
        relevant.sort(key=lambda x: -x[1])
        return [idx for idx, _ in relevant[:max_pages]]
    finally:
        os.remove(tmp_path)


# =============================================================================
# CLAUDE API EXTRACTION
# =============================================================================
def extract_group_from_pages(
    client: anthropic.Anthropic,
    group: dict,
    page_images: list,
    banka_adi: str,
    donem: str
) -> dict:
    """
    Belirli bir grup için Claude API'ye ilgili sayfaları gönderir
    ve JSON çıktı alır.
    """
    fields_desc = "\n".join([
        f'- "{f["id"]}": {f["label"]} ({f["col"]} sütunu)'
        for f in group["fields"]
    ])

    prompt = f"""Sen kıdemli bir banka denetçisisin. 
Banka: {banka_adi}
Dönem: {donem}
Tablo: {group["table_hint"]}

Aşağıdaki sayfa görüntülerinden SADECE aşağıdaki değerleri çek:

{fields_desc}

KRİTİK KURALLAR:
1. SADECE Cari Dönem değerlerini al. Önceki dönem/31 Aralık 2024 gibi önceki dönemleri KESİNLİKLE ALMA.
2. TP ve YP sütunlarını karıştırma. Tabloda nasıl yazıyorsa birebir al.
3. Sayıları tabloda nasılsa birebir al (örn: "5,138,035").
4. Bir değeri tabloda net göremiyorsan null yaz.
5. Yeşil renkli veya TOPLAM olan satırları alma — sadece belirtilen kalemler.
6. SADECE JSON döndür, başka hiçbir şey yazma:
{{"field_id": "değer_veya_null", ...}}"""

    # Görüntüleri mesaj içeriğine ekle
    content = []
    for img in page_images[:6]:  # Max 6 sayfa per grup
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": image_to_base64(img)
            }
        })
    content.append({"type": "text", "text": prompt})

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": content}]
    )

    raw = response.content[0].text.strip()

    # JSON çıkar
    try:
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            return json.loads(match.group())
    except Exception:
        pass
    return {}


def extract_metadata(client: anthropic.Anthropic, first_page_img) -> tuple[str, str]:
    """İlk sayfadan banka adı ve dönem bilgisini çeker."""
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=200,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": image_to_base64(first_page_img)
                    }
                },
                {
                    "type": "text",
                    "text": 'Bu sayfadan bankanın tam resmi adını ve raporun dönemini bul. SADECE JSON: {"banka_adi": "...", "donem": "..."}'
                }
            ]
        }]
    )
    try:
        match = re.search(r'\{[^{}]+\}', response.content[0].text, re.DOTALL)
        if match:
            data = json.loads(match.group())
            return data.get("banka_adi", "Bilinmeyen"), data.get("donem", "Bilinmiyor")
    except Exception:
        pass
    return "Bilinmeyen Banka", "Bilinmiyor"


# =============================================================================
# EXCEL EXPORT
# =============================================================================
def create_excel(banka_adi: str, donem: str, all_results: dict) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Denetim Verileri"

    # Stiller
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    group_fill = PatternFill("solid", fgColor="D6E4F0")
    group_font = Font(bold=True, size=10)
    missing_fill = PatternFill("solid", fgColor="FFE0E0")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Başlık
    ws.merge_cells("A1:D1")
    ws["A1"] = f"{banka_adi} — {donem}"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    ws.append([])

    # Sütun başlıkları
    ws.append(["Alan ID", "Açıklama", "Sütun", "Değer"])
    for cell in ws[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    ws.row_dimensions[3].height = 20

    # Her grup
    for group in EXTRACTION_GROUPS:
        ws.append([])
        ws.append([group["name"], "", "", ""])
        row_idx = ws.max_row
        for cell in ws[row_idx]:
            cell.fill = group_fill
            cell.font = group_font
            cell.border = border

        for field in group["fields"]:
            val = all_results.get(field["id"])
            ws.append([field["id"], field["label"], field["col"], val or "—"])
            r = ws.max_row
            for cell in ws[r]:
                cell.border = border
                cell.alignment = Alignment(vertical="center")
            if not val:
                for cell in ws[r]:
                    cell.fill = missing_fill

    # Sütun genişlikleri
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 20

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# =============================================================================
# STREAMLIT ARAYÜZÜ
# =============================================================================
st.set_page_config(page_title="Banka Denetim Raporu Analizi", layout="wide")
st.title("🏦 Banka Denetim Raporu — Claude AI Veri Çıkarımı")
st.caption("PDF → Görüntü → Claude Vision API → Excel | ~104 veri noktası")

# Sidebar
with st.sidebar:
    st.subheader("🔑 Anthropic API Key")
    api_key = st.text_input(
        "API Key",
        type="password",
        value=os.getenv("ANTHROPIC_API_KEY", ""),
        help="console.anthropic.com → API Keys"
    )
    if api_key:
        st.success("✅ API key girildi")
    else:
        st.warning("⚠️ API key gerekli")
        st.markdown("[→ Anthropic API Key Al](https://console.anthropic.com)")

    st.divider()
    st.subheader("ℹ️ Sistem Bilgisi")
    total_fields = sum(len(g["fields"]) for g in EXTRACTION_GROUPS)
    st.metric("Toplam Veri Noktası", total_fields)
    st.metric("Grup Sayısı", len(EXTRACTION_GROUPS))
    st.caption("Model: Claude Sonnet\nMaliyet: ~$0.05/rapor")

    st.divider()
    st.subheader("📋 Gruplar")
    for g in EXTRACTION_GROUPS:
        st.caption(f"• {g['name']} ({len(g['fields'])} alan)")

# Ana alan
uploaded_file = st.file_uploader("📄 PDF Rapor Yükle", type="pdf")

if uploaded_file and api_key:
    if st.button("🚀 Analiz Et", type="primary"):
        client = anthropic.Anthropic(api_key=api_key)
        pdf_bytes = uploaded_file.getvalue()
        all_results = {}

        progress = st.progress(0, "Başlatılıyor...")

        # 1. PDF → Görüntüler
        progress.progress(5, "📄 PDF sayfalara dönüştürülüyor...")
        with st.spinner("PDF işleniyor..."):
            images = pdf_to_images(pdf_bytes)
        st.success(f"✅ {len(images)} sayfa tespit edildi")

        # 2. Metadata
        progress.progress(10, "🤖 Banka adı ve dönem tespiti...")
        banka_adi, donem = extract_metadata(client, images[0])
        st.info(f"🏦 **{banka_adi}** | 📅 **{donem}**")

        # 3. Her grup için ilgili sayfaları bul ve Claude'a gönder
        total_groups = len(EXTRACTION_GROUPS)
        for g_idx, group in enumerate(EXTRACTION_GROUPS):
            progress_pct = 10 + int((g_idx / total_groups) * 85)
            progress.progress(progress_pct, f"🔍 {group['name']} çekiliyor...")

            with st.status(f"📊 {group['name']}", expanded=False) as status:
                # İlgili sayfaları bul
                keywords = PAGE_KEYWORDS.get(group["id"], [])
                page_indices = find_pages_with_pdfplumber(pdf_bytes, keywords, max_pages=5)

                if not page_indices:
                    page_indices = list(range(min(5, len(images))))

                st.caption(f"İlgili sayfalar: {[i+1 for i in page_indices]}")
                relevant_images = [images[i] for i in page_indices if i < len(images)]

                try:
                    result = extract_group_from_pages(
                        client, group, relevant_images, banka_adi, donem
                    )
                    found = sum(1 for v in result.values() if v and v != "null")
                    all_results.update(result)
                    status.update(
                        label=f"✅ {group['name']} — {found}/{len(group['fields'])} bulundu",
                        state="complete"
                    )
                except Exception as e:
                    status.update(label=f"❌ {group['name']} — hata: {e}", state="error")

                # API rate limit
                time.sleep(0.5)

        progress.progress(100, "✅ Tamamlandı!")

        # 4. Sonuçlar
        st.divider()
        total_found = sum(1 for v in all_results.values() if v and v != "null")
        total_fields = sum(len(g["fields"]) for g in EXTRACTION_GROUPS)

        col1, col2, col3 = st.columns(3)
        col1.metric("Toplam Alan", total_fields)
        col2.metric("Bulunan", total_found, delta=f"%{total_found/total_fields*100:.0f}")
        col3.metric("Bulunamayan", total_fields - total_found)

        # Grup bazlı özet
        for group in EXTRACTION_GROUPS:
            found = sum(1 for f in group["fields"] if all_results.get(f["id"]) and all_results.get(f["id"]) != "null")
            with st.expander(f"{group['name']} — {found}/{len(group['fields'])}"):
                cols = st.columns(3)
                for i, field in enumerate(group["fields"]):
                    val = all_results.get(field["id"])
                    with cols[i % 3]:
                        st.metric(
                            f"{field['label'][:30]} ({field['col']})",
                            val if val and val != "null" else "⚠️ —"
                        )

        # Excel
        excel_buf = create_excel(banka_adi, donem, all_results)
        st.download_button(
            "📥 Excel İndir",
            excel_buf,
            file_name=f"{banka_adi}_{donem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif uploaded_file and not api_key:
    st.warning("⚠️ Sol sidebar'dan Anthropic API key girin.")
