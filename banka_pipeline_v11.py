import streamlit as st
import os, tempfile, re, json, time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docling.document_converter import DocumentConverter
from langchain_groq import ChatGroq

# =============================================================================
# YAPILANDIRMA
# =============================================================================
GROQ_MODELS = {
    "qwen-qwq-32b":            "⭐ Qwen QwQ 32b — Finansal tablolar için önerilen",
    "llama-3.3-70b-versatile": "Llama 3.3 70b — Güçlü genel amaçlı",
    "llama-3.1-70b-versatile": "Llama 3.1 70b — Stabil alternatif",
    "mixtral-8x7b-32768":      "Mixtral 8x7b — 32k token context",
    "gemma2-9b-it":            "Gemma2 9b — Hafif ve hızlı",
}

# =============================================================================
# HEDEF VERİ YAPISI (v7'den alındı — 9 grup, 100+ alan)
# =============================================================================
EXTRACTION_GROUPS = [
    {
        "id": "aktif",
        "name": "AKTİF (VARLIKLAR)",
        "hint": "KONSOLİDE OLMAYAN BİLANÇO - VARLIKLAR bölümü. Cari dönem TP ve YP sütunları.",
        "keywords": ["VARLIKLAR TOPLAMI", "FİNANSAL VARLIKLAR", "Nakit ve Nakit Benzerleri"],
        "fields": [
            {"id": "nakit_tp",           "label": "Nakit ve Nakit Benzerleri",                                                      "col": "TP"},
            {"id": "nakit_yp",           "label": "Nakit ve Nakit Benzerleri",                                                      "col": "YP"},
            {"id": "gercege_kzara_tp",   "label": "Gerçeğe Uygun Değer Farkı K/Z'a Yansıtılan Finansal Varlıklar",                "col": "TP"},
            {"id": "gercege_kzara_yp",   "label": "Gerçeğe Uygun Değer Farkı K/Z'a Yansıtılan Finansal Varlıklar",                "col": "YP"},
            {"id": "gercege_diger_tp",   "label": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar", "col": "TP"},
            {"id": "gercege_diger_yp",   "label": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar", "col": "YP"},
            {"id": "itfa_tp",            "label": "İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar (2.4)",                       "col": "TP"},
            {"id": "itfa_yp",            "label": "İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar (2.4)",                       "col": "YP"},
            {"id": "krediler_tp",        "label": "Krediler",                                                                       "col": "TP"},
            {"id": "krediler_yp",        "label": "Krediler",                                                                       "col": "YP"},
            {"id": "beklenen_zarar_tp",  "label": "Beklenen Zarar Karşılıkları (-)",                                                "col": "TP"},
            {"id": "beklenen_zarar_yp",  "label": "Beklenen Zarar Karşılıkları (-)",                                                "col": "YP"},
            {"id": "satis_amacli_tp",    "label": "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar",  "col": "TP"},
            {"id": "satis_amacli_yp",    "label": "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar",  "col": "YP"},
            {"id": "ortaklik_tp",        "label": "Ortaklık Yatırımları",                                                           "col": "TP"},
            {"id": "ortaklik_yp",        "label": "Ortaklık Yatırımları",                                                           "col": "YP"},
            {"id": "maddi_duran_tp",     "label": "Maddi Duran Varlıklar (Net)",                                                    "col": "TP"},
            {"id": "maddi_duran_yp",     "label": "Maddi Duran Varlıklar (Net)",                                                    "col": "YP"},
            {"id": "maddi_olmayan_tp",   "label": "Maddi Olmayan Duran Varlıklar (Net)",                                            "col": "TP"},
            {"id": "maddi_olmayan_yp",   "label": "Maddi Olmayan Duran Varlıklar (Net)",                                            "col": "YP"},
            {"id": "yatirim_gay_tp",     "label": "Yatırım Amaçlı Gayrimenkuller (Net)",                                           "col": "TP"},
            {"id": "yatirim_gay_yp",     "label": "Yatırım Amaçlı Gayrimenkuller (Net)",                                           "col": "YP"},
            {"id": "cari_vergi_v_tp",    "label": "Cari Vergi Varlığı",                                                             "col": "TP"},
            {"id": "cari_vergi_v_yp",    "label": "Cari Vergi Varlığı",                                                             "col": "YP"},
            {"id": "ertelenmis_v_tp",    "label": "Ertelenmiş Vergi Varlığı",                                                       "col": "TP"},
            {"id": "ertelenmis_v_yp",    "label": "Ertelenmiş Vergi Varlığı",                                                       "col": "YP"},
            {"id": "diger_aktif_tp",     "label": "Diğer Aktifler (Net)",                                                           "col": "TP"},
            {"id": "diger_aktif_yp",     "label": "Diğer Aktifler (Net)",                                                           "col": "YP"},
            {"id": "varliklar_top_tp",   "label": "VARLIKLAR TOPLAMI",                                                              "col": "TP"},
            {"id": "varliklar_top_yp",   "label": "VARLIKLAR TOPLAMI",                                                              "col": "YP"},
        ]
    },
    {
        "id": "pasif",
        "name": "PASİF (YÜKÜMLÜLÜKLER)",
        "hint": "KONSOLİDE OLMAYAN BİLANÇO - YÜKÜMLÜLÜKLER bölümü. Cari dönem TP ve YP sütunları.",
        "keywords": ["YÜKÜMLÜLÜKLER TOPLAMI", "MEVDUAT", "ÖZKAYNAKLAR", "Ödenmiş Sermaye"],
        "fields": [
            {"id": "mevduat_tp",         "label": "Mevduat",                                "col": "TP"},
            {"id": "mevduat_yp",         "label": "Mevduat",                                "col": "YP"},
            {"id": "alinan_kred_tp",     "label": "Alınan Krediler",                        "col": "TP"},
            {"id": "alinan_kred_yp",     "label": "Alınan Krediler",                        "col": "YP"},
            {"id": "para_piy_bor_tp",    "label": "Para Piyasalarına Borçlar",              "col": "TP"},
            {"id": "para_piy_bor_yp",    "label": "Para Piyasalarına Borçlar",              "col": "YP"},
            {"id": "ihrac_menkul_tp",    "label": "İhraç Edilen Menkul Kıymetler (Net)",   "col": "TP"},
            {"id": "ihrac_menkul_yp",    "label": "İhraç Edilen Menkul Kıymetler (Net)",   "col": "YP"},
            {"id": "fonlar_tp",          "label": "Fonlar",                                 "col": "TP"},
            {"id": "fonlar_yp",          "label": "Fonlar",                                 "col": "YP"},
            {"id": "cari_vergi_b_tp",    "label": "Cari Vergi Borcu",                       "col": "TP"},
            {"id": "cari_vergi_b_yp",    "label": "Cari Vergi Borcu",                       "col": "YP"},
            {"id": "ertelenmis_b_tp",    "label": "Ertelenmiş Vergi Borcu",                 "col": "TP"},
            {"id": "ertelenmis_b_yp",    "label": "Ertelenmiş Vergi Borcu",                 "col": "YP"},
            {"id": "ozkaynak_tp",        "label": "Özkaynaklar",                            "col": "TP"},
            {"id": "ozkaynak_yp",        "label": "Özkaynaklar",                            "col": "YP"},
            {"id": "odenmis_ser_tp",     "label": "Ödenmiş Sermaye",                        "col": "TP"},
            {"id": "odenmis_ser_yp",     "label": "Ödenmiş Sermaye",                        "col": "YP"},
            {"id": "donem_net_kar_tp",   "label": "Dönem Net Kar veya Zararı",              "col": "TP"},
            {"id": "donem_net_kar_yp",   "label": "Dönem Net Kar veya Zararı",              "col": "YP"},
            {"id": "yukumluluk_top_tp",  "label": "YÜKÜMLÜLÜKLER TOPLAMI",                 "col": "TP"},
            {"id": "yukumluluk_top_yp",  "label": "YÜKÜMLÜLÜKLER TOPLAMI",                 "col": "YP"},
        ]
    },
    {
        "id": "gelir",
        "name": "GELİR TABLOSU",
        "hint": "KONSOLİDE OLMAYAN KAR VEYA ZARAR TABLOSU. Sadece Cari Dönem sütunu.",
        "keywords": ["FAİZ GELİRLERİ", "DÖNEM NET KARI", "PERSONEL GİDERLERİ", "NET FAALIYET"],
        "fields": [
            {"id": "faiz_gel",           "label": "Faiz Gelirleri",                                              "col": "Cari Dönem"},
            {"id": "faiz_gid",           "label": "Faiz Giderleri (-)",                                          "col": "Cari Dönem"},
            {"id": "net_ucret_kom",      "label": "Net Ücret ve Komisyon Gelirleri/Giderleri",                   "col": "Cari Dönem"},
            {"id": "alinan_ucret",       "label": "Alınan Ücret ve Komisyonlar",                                 "col": "Cari Dönem"},
            {"id": "verilen_ucret",      "label": "Verilen Ücret ve Komisyonlar (-)",                            "col": "Cari Dönem"},
            {"id": "kar_sinif_olmayan",  "label": "Kar veya Zararda Yeniden Sınıflandırılmayacaklar",            "col": "Cari Dönem"},
            {"id": "kar_sinif_olan",     "label": "Kar veya Zararda Yeniden Sınıflandırılacaklar",               "col": "Cari Dönem"},
            {"id": "temettu",            "label": "Temettü Gelirleri",                                           "col": "Cari Dönem"},
            {"id": "ticari_kz",          "label": "Ticari Kar/Zarar (Net)",                                      "col": "Cari Dönem"},
            {"id": "diger_faal_gel",     "label": "Diğer Faaliyet Gelirleri",                                    "col": "Cari Dönem"},
            {"id": "beklenen_zarar_gid", "label": "Beklenen Zarar Karşılıkları (-)",                             "col": "Cari Dönem"},
            {"id": "diger_kars_gid",     "label": "Diğer Karşılık Giderleri (-)",                               "col": "Cari Dönem"},
            {"id": "personel_gid",       "label": "Personel Giderleri (-)",                                      "col": "Cari Dönem"},
            {"id": "diger_faal_gid",     "label": "Diğer Faaliyet Giderleri (-)",                               "col": "Cari Dönem"},
            {"id": "net_faal_kz",        "label": "Net Faaliyet Karı/Zararı",                                    "col": "Cari Dönem"},
            {"id": "surd_vergi_oncesi",  "label": "Sürdürülen Faaliyetler Vergi Öncesi K/Z",                    "col": "Cari Dönem"},
            {"id": "surd_vergi",         "label": "Sürdürülen Faaliyetler Vergi Karşılığı",                     "col": "Cari Dönem"},
            {"id": "donem_net_kz",       "label": "Dönem Net Karı/Zararı",                                       "col": "Cari Dönem"},
        ]
    },
    {
        "id": "vergi",
        "name": "VERGİ DETAYI",
        "hint": "Sürdürülen faaliyetler vergi karşılığına ilişkin dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Cari Vergi Karşılığı", "Ertelenmiş Vergi Gider", "Ertelenmiş Vergi Gelir"],
        "fields": [
            {"id": "ozkaynak_ortaklik",    "label": "Özkaynak Yöntemi Uygulanan Ortaklıklardan Kar/Zarar",  "col": "Cari Dönem"},
            {"id": "cari_vergi_karsiliği", "label": "Cari Vergi Karşılığı",                                 "col": "Cari Dönem"},
            {"id": "ertelenmis_net",       "label": "Ertelenmiş Vergi Net",                                  "col": "Cari Dönem"},
            {"id": "ertelenmis_gider",     "label": "Ertelenmiş Vergi Gider Etkisi (+)",                     "col": "Cari Dönem"},
            {"id": "ertelenmis_gelir",     "label": "Ertelenmiş Vergi Gelir Etkisi (-)",                     "col": "Cari Dönem"},
        ]
    },
    {
        "id": "faiz_detay",
        "name": "FAİZ GELİR/GİDER DETAYI",
        "hint": "Faiz gelirlerine ilişkin açıklamalar dipnot tablosu. TL ve YP sütunları ayrı.",
        "keywords": ["Kredilerden Alınan Faiz", "Zorunlu Karşılıklardan", "Mevduata Verilen Faizler"],
        "fields": [
            {"id": "kred_faiz_gel_tl",   "label": "Kredilerden Alınan Faiz Gelirleri",                 "col": "TL"},
            {"id": "kred_faiz_gel_yp",   "label": "Kredilerden Alınan Faiz Gelirleri",                 "col": "YP"},
            {"id": "zorunlu_kars_faiz",  "label": "Zorunlu Karşılıklardan Alınan Faizler",             "col": "Cari Dönem"},
            {"id": "bankalar_faiz",      "label": "Bankalardan Alınan Faizler",                         "col": "Cari Dönem"},
            {"id": "para_piy_faiz_gel",  "label": "Para Piyasası İşlemlerinden Alınan Faizler",        "col": "Cari Dönem"},
            {"id": "menkul_faiz",        "label": "Menkul Değerlerden Alınan Faizler",                  "col": "Cari Dönem"},
            {"id": "fin_kira_faiz",      "label": "Finansal Kiralama Faiz Gelirleri",                   "col": "Cari Dönem"},
            {"id": "diger_faiz_gel",     "label": "Diğer Faiz Gelirleri",                               "col": "Cari Dönem"},
            {"id": "mevduat_faiz_gid",   "label": "Mevduata Verilen Faizler",                           "col": "Cari Dönem"},
            {"id": "kred_faiz_gid",      "label": "Kullanılan Kredilere Verilen Faizler",               "col": "Cari Dönem"},
            {"id": "para_piy_faiz_gid",  "label": "Para Piyasası İşlemlerine Verilen Faizler",         "col": "Cari Dönem"},
            {"id": "ihrac_faiz_gid",     "label": "İhraç Edilen Menkul Kıymetlere Verilen Faizler",    "col": "Cari Dönem"},
            {"id": "kira_faiz_gid",      "label": "Kiralama Faiz Giderleri",                            "col": "Cari Dönem"},
            {"id": "diger_faiz_gid",     "label": "Diğer Faiz Giderleri",                               "col": "Cari Dönem"},
        ]
    },
    {
        "id": "ort_faiz",
        "name": "ORTALAMA FAİZ ORANLARI (%)",
        "hint": "Parasal finansal araçlara uygulanan ortalama faiz oranları tablosu. TL ve YP (USD) sütunları.",
        "keywords": ["ortalama faiz", "Verilen Krediler", "Bankalar Mevduatı", "İhraç Edilen Menkul"],
        "fields": [
            {"id": "vkred_tl",             "label": "Verilen Krediler",                      "col": "TL"},
            {"id": "bank_mev_tl",          "label": "Bankalar Mevduatı",                     "col": "TL"},
            {"id": "diger_mev_tl",         "label": "Diğer Mevduat",                         "col": "TL"},
            {"id": "ihrac_menkul_oran_tl", "label": "İhraç Edilen Menkul Değerler (Net)",   "col": "TL"},
            {"id": "vkred_yp",             "label": "Verilen Krediler",                      "col": "YP (USD)"},
            {"id": "bank_mev_yp",          "label": "Bankalar Mevduatı",                     "col": "YP (USD)"},
            {"id": "diger_mev_yp",         "label": "Diğer Mevduat",                         "col": "YP (USD)"},
            {"id": "ihrac_menkul_oran_yp", "label": "İhraç Edilen Menkul Değerler (Net)",   "col": "YP (USD)"},
        ]
    },
    {
        "id": "ticari",
        "name": "TİCARİ KAR/ZARAR DETAYI",
        "hint": "Ticari kar/zarara ilişkin açıklamalar dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Sermaye Piyasası İşlemleri Kârı", "Türev Finansal İşlemlerden", "Kambiyo İşlemlerinden"],
        "fields": [
            {"id": "ticari_kar",     "label": "Ticari Kar",                             "col": "Cari Dönem"},
            {"id": "ser_piy_kar",    "label": "Sermaye Piyasası İşlemleri Karı",        "col": "Cari Dönem"},
            {"id": "turev_kar",      "label": "Türev Finansal İşlemlerden Kar",          "col": "Cari Dönem"},
            {"id": "kambiyo_kar",    "label": "Kambiyo İşlemlerinden Kar",              "col": "Cari Dönem"},
            {"id": "ticari_zarar",   "label": "Ticari Zarar",                           "col": "Cari Dönem"},
            {"id": "ser_piy_zarar",  "label": "Sermaye Piyasası İşlemleri Zararı",      "col": "Cari Dönem"},
            {"id": "turev_zarar",    "label": "Türev Finansal İşlemlerden Zarar",        "col": "Cari Dönem"},
            {"id": "kambiyo_zarar",  "label": "Kambiyo İşlemlerinden Zarar",            "col": "Cari Dönem"},
        ]
    },
    {
        "id": "bkzk",
        "name": "BEKLENEN KREDİ ZARAR KARŞILIKLARI",
        "hint": "Beklenen zarar karşılıkları dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Birinci Aşama", "İkinci Aşama", "Temerrüt", "Beklenen Kredi Zarar"],
        "fields": [
            {"id": "birinci_asama", "label": "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)", "col": "Cari Dönem"},
            {"id": "ikinci_asama",  "label": "Kredi Riskinde Önemli Artış (İkinci Aşama)",        "col": "Cari Dönem"},
            {"id": "ucuncu_asama",  "label": "Temerrüt (Üçüncü Aşama)",                           "col": "Cari Dönem"},
        ]
    },
    {
        "id": "faal_gider",
        "name": "DİĞER FAALİYET GİDERLERİ DETAYI",
        "hint": "Diğer faaliyet giderlerine ilişkin bilgiler dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Kıdem Tazminat", "Amortisman", "TFRS 16", "Reklam ve İlan", "Bakım ve Onarım"],
        "fields": [
            {"id": "kidem_taz",         "label": "Kıdem Tazminat Karşılığı",                         "col": "Cari Dönem"},
            {"id": "amortisman",        "label": "Amortisman Giderleri",                             "col": "Cari Dönem"},
            {"id": "diger_isletme",     "label": "Diğer İşletme Giderleri",                          "col": "Cari Dönem"},
            {"id": "tfrs16_kira",       "label": "TFRS 16 İstisnalarına İlişkin Kiralama Giderleri", "col": "Cari Dönem"},
            {"id": "bakim_onarim",      "label": "Bakım ve Onarım Giderleri",                        "col": "Cari Dönem"},
            {"id": "reklam_ilan",       "label": "Reklam ve İlan Giderleri",                         "col": "Cari Dönem"},
            {"id": "diger_giderler_1",  "label": "Diğer Giderler (1)",                               "col": "Cari Dönem"},
            {"id": "aktif_satis_zarar", "label": "Aktiflerin Satışından Doğan Zararlar",             "col": "Cari Dönem"},
            {"id": "diger_2",           "label": "Diğer (2)",                                        "col": "Cari Dönem"},
        ]
    },
    {
        "id": "kredi_kalite",
        "name": "KREDİ KALİTE TABLOLARI",
        "hint": "Standart nitelikli/yakın izlemedeki krediler ve sorunlu kredi grupları dipnot tabloları.",
        "keywords": ["Standart Nitelikli", "Yakın İzlemedeki", "Tahsil İmkânı Sınırlı", "Tahsili Şüpheli"],
        "fields": [
            {"id": "std_nitelikli",     "label": "Standart Nitelikli Krediler",                                                   "col": "Cari Dönem"},
            {"id": "yakin_yap_disinda", "label": "Yakın İzlemedeki - Yeniden Yapılandırma Kapsamında Yer Almayanlar",             "col": "Cari Dönem"},
            {"id": "yakin_sozlesme",    "label": "Yakın İzlemedeki - Sözleşme Koşullarında Değişiklik",                          "col": "Cari Dönem"},
            {"id": "yakin_yeniden_fin", "label": "Yakın İzlemedeki - Yeniden Finansman",                                         "col": "Cari Dönem"},
            {"id": "konut_kred",        "label": "Konut Kredileri (Tüketici+Personel)",                                          "col": "Cari Dönem"},
            {"id": "tasit_kred",        "label": "Taşıt Kredileri (Tüketici+Personel)",                                          "col": "Cari Dönem"},
            {"id": "ihtiyac_kred",      "label": "İhtiyaç Kredileri (Tüketici+Personel)",                                        "col": "Cari Dönem"},
            {"id": "diger_kred_tipi",   "label": "Diğer (Tüketici+Personel)",                                                    "col": "Cari Dönem"},
            {"id": "kredi_karti",       "label": "Kredi Kartları",                                                               "col": "Cari Dönem"},
            {"id": "kmh",               "label": "KMH",                                                                          "col": "Cari Dönem"},
            {"id": "gr3_bakiye",        "label": "III. Grup Tahsil İmkânı Sınırlı - Dönem Sonu Bakiyesi",                       "col": "Cari Dönem"},
            {"id": "gr3_karsilik",      "label": "III. Grup Tahsil İmkânı Sınırlı - Karşılık (-)",                              "col": "Cari Dönem"},
            {"id": "gr3_net",           "label": "III. Grup Tahsil İmkânı Sınırlı - Bilançodaki Net Bakiyesi",                  "col": "Cari Dönem"},
            {"id": "gr4_bakiye",        "label": "IV. Grup Tahsili Şüpheli - Dönem Sonu Bakiyesi",                              "col": "Cari Dönem"},
            {"id": "gr4_karsilik",      "label": "IV. Grup Tahsili Şüpheli - Karşılık (-)",                                     "col": "Cari Dönem"},
            {"id": "gr4_net",           "label": "IV. Grup Tahsili Şüpheli - Bilançodaki Net Bakiyesi",                         "col": "Cari Dönem"},
            {"id": "gr5_bakiye",        "label": "V. Grup Zarar Niteliğindeki - Dönem Sonu Bakiyesi",                           "col": "Cari Dönem"},
            {"id": "gr5_karsilik",      "label": "V. Grup Zarar Niteliğindeki - Karşılık (-)",                                  "col": "Cari Dönem"},
            {"id": "gr5_net",           "label": "V. Grup Zarar Niteliğindeki - Bilançodaki Net Bakiyesi",                      "col": "Cari Dönem"},
        ]
    },
]


# =============================================================================
# YARDIMCI
# =============================================================================
def safe_parse_json(raw: str) -> dict:
    """Markdown kod bloklarını temizleyip güvenli JSON parse eder."""
    try:
        clean = re.sub(r'```(?:json)?', '', raw).strip().replace('```', '').strip()
        match = re.search(r'\{.*\}', clean, re.DOTALL)
        if match:
            return json.loads(match.group())
        return json.loads(clean)
    except Exception as e:
        print(f"JSON Parse Hatası: {e}\nHam: {raw[:200]}")
        return {}


def find_context_for_group(md_text: str, keywords: list[str],
                            window: int = 8000) -> str:
    """Markdown içinde keyword'lere göre en alakalı bölümü döndürür."""
    upper = md_text.upper()
    best_pos = -1
    for kw in keywords:
        pos = upper.find(kw.upper())
        if pos != -1 and (best_pos == -1 or pos < best_pos):
            best_pos = pos
    if best_pos == -1:
        return md_text[:window]
    start = max(0, best_pos - 500)
    return md_text[start: start + window]


# =============================================================================
# GROQ LLM
# =============================================================================
def get_groq_llm(api_key: str, model: str) -> ChatGroq:
    return ChatGroq(api_key=api_key, model=model, temperature=0, max_tokens=4096)


def extract_metadata_llm(llm: ChatGroq, md_head: str) -> tuple[str, str]:
    prompt = (
        "Aşağıdaki banka raporundan bankanın TAM RESMİ ADINI ve Cari dönem tarihini bul.\n"
        "SADECE JSON döndür: {\"banka_adi\": \"...\", \"donem\": \"...\"}\n\n"
        f"METİN:\n{md_head[:2000]}"
    )
    try:
        resp = llm.invoke(prompt)
        content = resp.content if hasattr(resp, "content") else str(resp)
        data = safe_parse_json(content)
        if data:
            return data.get("banka_adi", "Bilinmeyen"), data.get("donem", "Bilinmiyor")
    except Exception:
        pass
    return "Bilinmeyen Banka", "Bilinmiyor"


def extract_group_llm(llm: ChatGroq, group: dict,
                       context: str, banka_adi: str, donem: str) -> dict:
    """Bir grubu Groq LLM ile çıkarır."""
    fields_desc = "\n".join([
        f'- "{f["id"]}": {f["label"]} ({f["col"]} sütunu)'
        for f in group["fields"]
    ])

    prompt = f"""Sen kıdemli bir banka denetçisisin.
Banka: {banka_adi} | Dönem: {donem}
Tablo: {group["hint"]}

Aşağıdaki tablo metninden SADECE Cari Dönem değerlerini çek:

{fields_desc}

KRİTİK KURALLAR:
1. SADECE Cari Dönem. Önceki dönem = KESİNLİKLE ALMA.
2. TP ve YP sütunlarını karıştırma.
3. Değeri net göremiyorsan null yaz. Asla tahmin etme.
4. Sayıları tabloda nasılsa birebir al (örn: "5,138,035").
5. SADECE JSON döndür: {{"field_id": "deger_veya_null"}}

TABLO METNİ:
{context[:6000]}"""

    try:
        resp = llm.invoke(prompt)
        content = resp.content if hasattr(resp, "content") else str(resp)
        result = safe_parse_json(content)
        return result if result else {}
    except Exception as e:
        st.caption(f"⚠️ Groq hatası ({group['id']}): {e}")
        return {}


# =============================================================================
# PDF → MARKDOWN (Docling)
# =============================================================================
def pdf_to_markdown(pdf_bytes: bytes) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        converter = DocumentConverter()
        md = converter.convert(tmp_path).document.export_to_markdown()
    finally:
        os.remove(tmp_path)
    return md


# =============================================================================
# TEK RAPOR İŞLEME
# =============================================================================
def process_report(pdf_bytes: bytes, filename: str,
                   llm: ChatGroq) -> dict:
    t_start = time.time()
    errors = []

    # PDF → Markdown
    try:
        md_text = pdf_to_markdown(pdf_bytes)
    except Exception as e:
        return {
            "filename": filename, "banka_adi": "HATA", "donem": "—",
            "all_results": {}, "method_log": {}, "errors": [f"PDF hatası: {e}"],
            "elapsed_sec": 0
        }

    # Metadata
    try:
        banka_adi, donem = extract_metadata_llm(llm, md_text)
    except Exception as e:
        banka_adi, donem = filename, "Bilinmiyor"
        errors.append(f"Metadata hatası: {e}")

    all_results = {}
    method_log = {}

    for group in EXTRACTION_GROUPS:
        context = find_context_for_group(md_text, group["keywords"])
        result = extract_group_llm(llm, group, context, banka_adi, donem)
        all_results.update(result)
        found = sum(1 for v in result.values() if v and v not in ["null", "None", None, ""])
        method_log[group["id"]] = f"Groq ({found}/{len(group['fields'])})"
        time.sleep(0.8)  # Groq rate limit

    return {
        "filename": filename,
        "banka_adi": banka_adi,
        "donem": donem,
        "all_results": all_results,
        "method_log": method_log,
        "errors": errors,
        "elapsed_sec": round(time.time() - t_start, 1),
    }


# =============================================================================
# EXCEL EXPORT
# =============================================================================
def create_excel(report: dict) -> BytesIO:
    banka_adi = report["banka_adi"]
    donem = report["donem"]
    all_results = report["all_results"]
    method_log = report["method_log"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Denetim Verileri"

    h_fill  = PatternFill("solid", fgColor="1F4E79")
    h_font  = Font(color="FFFFFF", bold=True, size=11)
    g_fill  = PatternFill("solid", fgColor="D6E4F0")
    g_font  = Font(bold=True, size=10)
    miss_fill = PatternFill("solid", fgColor="FFE0E0")
    thin = Side(style="thin", color="CCCCCC")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:E1")
    ws["A1"] = f"{banka_adi} — {donem}"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28
    ws.append([])

    ws.append(["Alan ID", "Açıklama", "Sütun", "Değer", "Kaynak"])
    for cell in ws[3]:
        cell.fill = h_fill; cell.font = h_font
        cell.alignment = Alignment(horizontal="center"); cell.border = brd
    ws.row_dimensions[3].height = 20

    for group in EXTRACTION_GROUPS:
        ws.append([])
        method = method_log.get(group["id"], "—")
        ws.append([group["name"], "", "", "", method])
        r = ws.max_row
        for cell in ws[r]:
            cell.fill = g_fill; cell.font = g_font; cell.border = brd

        for field in group["fields"]:
            val = all_results.get(field["id"])
            display = val if val and val not in ["null", "None", ""] else None
            ws.append([field["id"], field["label"], field["col"],
                       display if display else "—", method])
            ri = ws.max_row
            for cell in ws[ri]:
                cell.border = brd
            if not display:
                for cell in ws[ri]:
                    cell.fill = miss_fill

    for col, w in [("A", 28), ("B", 65), ("C", 14), ("D", 22), ("E", 22)]:
        ws.column_dimensions[col].width = w

    # Özet sayfası
    ws2 = wb.create_sheet("Özet")
    ws2.append(["Grup", "Toplam Alan", "Bulunan", "Başarı %", "Kaynak"])
    for cell in ws2[1]:
        cell.fill = h_fill; cell.font = h_font
    for group in EXTRACTION_GROUPS:
        total = len(group["fields"])
        found = sum(
            1 for f in group["fields"]
            if all_results.get(f["id"]) and all_results.get(f["id"]) not in ["null", "None", ""]
        )
        ws2.append([group["name"], total, found,
                    f"%{found/total*100:.0f}", method_log.get(group["id"], "—")])
    for col, w in [("A", 38), ("B", 12), ("C", 10), ("D", 10), ("E", 24)]:
        ws2.column_dimensions[col].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# =============================================================================
# STREAMLIT ARAYÜZÜ
# =============================================================================
st.set_page_config(page_title="Banka Analiz Pipeline V11", layout="wide")
st.title("🏦 Banka Analiz Pipeline V11")
st.caption(
    f"Groq LLM | Docling PDF→Markdown | "
    f"{len(EXTRACTION_GROUPS)} grup | "
    f"{sum(len(g['fields']) for g in EXTRACTION_GROUPS)} alan"
)

total_fields_count = sum(len(g["fields"]) for g in EXTRACTION_GROUPS)

# ── Sidebar ──────────────────────────────────────────────────────────────────
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
    st.subheader("ℹ️ Kapsam")
    st.metric("Toplam Alan", total_fields_count)
    st.metric("Grup Sayısı", len(EXTRACTION_GROUPS))
    st.caption(
        "Ücretsiz Groq:\n"
        "• 30 req/dk\n"
        "• Qwen/Llama 70b dahil\n"
        "• Rate limit için bekleme var"
    )

# ── Ana Alan ─────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "📄 PDF Rapor(lar) Yükle",
    type="pdf",
    accept_multiple_files=True,
    help="Birden fazla rapor seçebilirsiniz."
)

if uploaded_files:
    st.info(
        f"**{len(uploaded_files)} rapor** yüklendi: "
        f"{', '.join(f.name for f in uploaded_files)}"
    )

    if not groq_api_key:
        st.error("❌ Groq API key gerekli.")
        st.stop()

    if st.button("🚀 Tümünü Analiz Et", type="primary"):
        llm = get_groq_llm(groq_api_key, selected_model)
        all_reports = []

        overall_progress = st.progress(0, f"0/{len(uploaded_files)} rapor işlendi")

        for file_idx, uploaded_file in enumerate(uploaded_files):
            st.divider()
            st.subheader(f"📄 {uploaded_file.name}")

            with st.status(f"İşleniyor: {uploaded_file.name}", expanded=True) as status:
                st.write("📄 PDF → Markdown (Docling)...")
                st.write(f"🤖 Groq ({selected_model}) ile {len(EXTRACTION_GROUPS)} grup çıkarılıyor...")

                report = process_report(
                    pdf_bytes=uploaded_file.getvalue(),
                    filename=uploaded_file.name,
                    llm=llm,
                )

                if report["errors"]:
                    status.update(
                        label=f"⚠️ {uploaded_file.name} — hatalar var",
                        state="error"
                    )
                    for err in report["errors"]:
                        st.error(err)
                else:
                    all_results = report["all_results"]
                    total_found = sum(
                        1 for v in all_results.values()
                        if v and v not in ["null", "None", ""]
                    )
                    status.update(
                        label=(
                            f"✅ {report['banka_adi']} | {report['donem']} "
                            f"— {total_found}/{total_fields_count} alan | "
                            f"{report['elapsed_sec']}sn"
                        ),
                        state="complete"
                    )

            all_reports.append(report)

            # Özet metrikler
            all_res = report["all_results"]
            total_found = sum(
                1 for v in all_res.values()
                if v and v not in ["null", "None", ""]
            )
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Banka", report["banka_adi"])
            c2.metric("Dönem", report["donem"])
            c3.metric("Bulunan Alan", f"{total_found}/{total_fields_count}")
            c4.metric("Süre", f"{report['elapsed_sec']}sn")

            # Grup detayları
            for group in EXTRACTION_GROUPS:
                found = sum(
                    1 for f in group["fields"]
                    if all_res.get(f["id"]) and all_res.get(f["id"]) not in ["null", "None", ""]
                )
                method = report["method_log"].get(group["id"], "—")
                with st.expander(f"📊 {group['name']} — {found}/{len(group['fields'])}"):
                    cols = st.columns(3)
                    for i, field in enumerate(group["fields"]):
                        val = all_res.get(field["id"])
                        display = val if val and val not in ["null", "None", ""] else None
                        with cols[i % 3]:
                            st.metric(
                                f"{field['label'][:30]} ({field['col']})",
                                display if display else "⚠️ —"
                            )

            # Raporlar arası bekleme (Groq rate limit)
            if file_idx < len(uploaded_files) - 1:
                with st.spinner("⏳ Groq rate limit bekleniyor (3sn)..."):
                    time.sleep(3)

            overall_progress.progress(
                (file_idx + 1) / len(uploaded_files),
                f"{file_idx + 1}/{len(uploaded_files)} rapor işlendi"
            )

        # ── Excel İndir ───────────────────────────────────────────────────────
        st.divider()
        if len(all_reports) == 1:
            excel_buf = create_excel(all_reports[0])
            r = all_reports[0]
            st.success(
                f"✅ **{r['banka_adi']}** tamamlandı — "
                f"{sum(1 for v in r['all_results'].values() if v and v not in ['null','None',''])}"
                f"/{total_fields_count} alan"
            )
            st.download_button(
                "📥 Excel İndir",
                excel_buf,
                file_name=f"{r['banka_adi']}_{r['donem']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # Çoklu rapor → ayrı Excel'ler + karşılaştırma
            st.success(f"✅ **{len(all_reports)} rapor** tamamlandı")
            c_dl = st.columns(len(all_reports))
            for i, r in enumerate(all_reports):
                excel_buf = create_excel(r)
                found = sum(
                    1 for v in r["all_results"].values()
                    if v and v not in ["null", "None", ""]
                )
                c_dl[i].download_button(
                    f"📥 {r['banka_adi'][:20]}",
                    excel_buf,
                    file_name=f"{r['banka_adi']}_{r['donem']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{i}"
                )

            # Karşılaştırma tablosu
            st.subheader("📊 Raporlar Arası Karşılaştırma")
            header_cols = st.columns([2] + [1] * len(all_reports))
            header_cols[0].markdown("**Alan**")
            for i, r in enumerate(all_reports):
                header_cols[i + 1].markdown(f"**{r['donem']}**")

            key_fields = [
                "nakit_tp", "nakit_yp",
                "varliklar_top_tp", "varliklar_top_yp",
                "krediler_tp", "krediler_yp",
                "yukumluluk_top_tp", "ozkaynak_tp",
                "faiz_gel", "faiz_gid", "donem_net_kz",
            ]
            all_field_labels = {
                f["id"]: f["label"]
                for g in EXTRACTION_GROUPS for f in g["fields"]
            }
            for fid in key_fields:
                if fid not in all_field_labels:
                    continue
                row_cols = st.columns([2] + [1] * len(all_reports))
                row_cols[0].caption(all_field_labels[fid][:45])
                for i, r in enumerate(all_reports):
                    val = r["all_results"].get(fid) or "—"
                    row_cols[i + 1].metric("", val)
