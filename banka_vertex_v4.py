import streamlit as st
import base64
import json
import re
import tempfile
import os
import time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pdfplumber
from pdf2image import convert_from_path
import anthropic

# Vertex AI authentication — gcloud auth application-default login ile otomatik
PROJECT_ID = os.getenv("GOOGLE_CLOUD_PROJECT", "")
REGION = "us-east5"  # Claude modelleri bu region'da

# =============================================================================
# HEDEF VERİ YAPISI
# =============================================================================
EXTRACTION_GROUPS = [
    {
        "id": "aktif",
        "name": "AKTİF (VARLIKLAR)",
        "table_hint": "KONSOLİDE OLMAYAN BİLANÇO - VARLIKLAR bölümü. Cari dönem TP ve YP sütunları.",
        "keywords": ["VARLIKLAR TOPLAMI", "FİNANSAL VARLIKLAR", "Nakit ve Nakit Benzerleri"],
        "fields": [
            {"id": "nakit_tp",              "label": "Nakit ve Nakit Benzerleri",                                                      "col": "TP"},
            {"id": "nakit_yp",              "label": "Nakit ve Nakit Benzerleri",                                                      "col": "YP"},
            {"id": "gercege_kzara_tp",      "label": "Gerçeğe Uygun Değer Farkı K/Z'a Yansıtılan Finansal Varlıklar",                "col": "TP"},
            {"id": "gercege_kzara_yp",      "label": "Gerçeğe Uygun Değer Farkı K/Z'a Yansıtılan Finansal Varlıklar",                "col": "YP"},
            {"id": "gercege_diger_tp",      "label": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar", "col": "TP"},
            {"id": "gercege_diger_yp",      "label": "Gerçeğe Uygun Değer Farkı Diğer Kapsamlı Gelire Yansıtılan Finansal Varlıklar", "col": "YP"},
            {"id": "itfa_tp",               "label": "İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar (2.4)",                       "col": "TP"},
            {"id": "itfa_yp",               "label": "İtfa Edilmiş Maliyetle Ölçülen Finansal Varlıklar (2.4)",                       "col": "YP"},
            {"id": "krediler_tp",           "label": "Krediler",                                                                       "col": "TP"},
            {"id": "krediler_yp",           "label": "Krediler",                                                                       "col": "YP"},
            {"id": "beklenen_zarar_tp",     "label": "Beklenen Zarar Karşılıkları (-)",                                                "col": "TP"},
            {"id": "beklenen_zarar_yp",     "label": "Beklenen Zarar Karşılıkları (-)",                                                "col": "YP"},
            {"id": "satis_amacli_tp",       "label": "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar",  "col": "TP"},
            {"id": "satis_amacli_yp",       "label": "Satış Amaçlı Elde Tutulan ve Durdurulan Faaliyetlere İlişkin Duran Varlıklar",  "col": "YP"},
            {"id": "ortaklik_tp",           "label": "Ortaklık Yatırımları",                                                           "col": "TP"},
            {"id": "ortaklik_yp",           "label": "Ortaklık Yatırımları",                                                           "col": "YP"},
            {"id": "maddi_duran_tp",        "label": "Maddi Duran Varlıklar (Net)",                                                    "col": "TP"},
            {"id": "maddi_duran_yp",        "label": "Maddi Duran Varlıklar (Net)",                                                    "col": "YP"},
            {"id": "maddi_olmayan_tp",      "label": "Maddi Olmayan Duran Varlıklar (Net)",                                            "col": "TP"},
            {"id": "maddi_olmayan_yp",      "label": "Maddi Olmayan Duran Varlıklar (Net)",                                            "col": "YP"},
            {"id": "yatirim_gay_tp",        "label": "Yatırım Amaçlı Gayrimenkuller (Net)",                                            "col": "TP"},
            {"id": "yatirim_gay_yp",        "label": "Yatırım Amaçlı Gayrimenkuller (Net)",                                            "col": "YP"},
            {"id": "cari_vergi_v_tp",       "label": "Cari Vergi Varlığı",                                                             "col": "TP"},
            {"id": "cari_vergi_v_yp",       "label": "Cari Vergi Varlığı",                                                             "col": "YP"},
            {"id": "ertelenmis_v_tp",       "label": "Ertelenmiş Vergi Varlığı",                                                       "col": "TP"},
            {"id": "ertelenmis_v_yp",       "label": "Ertelenmiş Vergi Varlığı",                                                       "col": "YP"},
            {"id": "diger_aktif_tp",        "label": "Diğer Aktifler (Net)",                                                           "col": "TP"},
            {"id": "diger_aktif_yp",        "label": "Diğer Aktifler (Net)",                                                           "col": "YP"},
            {"id": "varliklar_top_tp",      "label": "VARLIKLAR TOPLAMI",                                                              "col": "TP"},
            {"id": "varliklar_top_yp",      "label": "VARLIKLAR TOPLAMI",                                                              "col": "YP"},
        ]
    },
    {
        "id": "pasif",
        "name": "PASİF (YÜKÜMLÜLÜKLER)",
        "table_hint": "KONSOLİDE OLMAYAN BİLANÇO - YÜKÜMLÜLÜKLER bölümü. Cari dönem TP ve YP sütunları.",
        "keywords": ["YÜKÜMLÜLÜKLER TOPLAMI", "MEVDUAT", "ÖZKAYNAKLAR", "Ödenmiş Sermaye"],
        "fields": [
            {"id": "mevduat_tp",            "label": "Mevduat",                                 "col": "TP"},
            {"id": "mevduat_yp",            "label": "Mevduat",                                 "col": "YP"},
            {"id": "alinan_kred_tp",        "label": "Alınan Krediler",                         "col": "TP"},
            {"id": "alinan_kred_yp",        "label": "Alınan Krediler",                         "col": "YP"},
            {"id": "para_piy_bor_tp",       "label": "Para Piyasalarına Borçlar",               "col": "TP"},
            {"id": "para_piy_bor_yp",       "label": "Para Piyasalarına Borçlar",               "col": "YP"},
            {"id": "ihrac_menkul_tp",       "label": "İhraç Edilen Menkul Kıymetler (Net)",     "col": "TP"},
            {"id": "ihrac_menkul_yp",       "label": "İhraç Edilen Menkul Kıymetler (Net)",     "col": "YP"},
            {"id": "fonlar_tp",             "label": "Fonlar",                                  "col": "TP"},
            {"id": "fonlar_yp",             "label": "Fonlar",                                  "col": "YP"},
            {"id": "cari_vergi_b_tp",       "label": "Cari Vergi Borcu",                        "col": "TP"},
            {"id": "cari_vergi_b_yp",       "label": "Cari Vergi Borcu",                        "col": "YP"},
            {"id": "ertelenmis_b_tp",       "label": "Ertelenmiş Vergi Borcu",                  "col": "TP"},
            {"id": "ertelenmis_b_yp",       "label": "Ertelenmiş Vergi Borcu",                  "col": "YP"},
            {"id": "ozkaynak_tp",           "label": "Özkaynaklar",                             "col": "TP"},
            {"id": "ozkaynak_yp",           "label": "Özkaynaklar",                             "col": "YP"},
            {"id": "odenmis_ser_tp",        "label": "Ödenmiş Sermaye",                         "col": "TP"},
            {"id": "odenmis_ser_yp",        "label": "Ödenmiş Sermaye",                         "col": "YP"},
            {"id": "donem_net_kar_tp",      "label": "Dönem Net Kar veya Zararı",               "col": "TP"},
            {"id": "donem_net_kar_yp",      "label": "Dönem Net Kar veya Zararı",               "col": "YP"},
            {"id": "yukumluluk_top_tp",     "label": "YÜKÜMLÜLÜKLER TOPLAMI",                   "col": "TP"},
            {"id": "yukumluluk_top_yp",     "label": "YÜKÜMLÜLÜKLER TOPLAMI",                   "col": "YP"},
        ]
    },
    {
        "id": "gelir",
        "name": "GELİR TABLOSU",
        "table_hint": "KONSOLİDE OLMAYAN KAR VEYA ZARAR TABLOSU. Sadece Cari Dönem sütunu (1 Ocak - rapor tarihi).",
        "keywords": ["FAİZ GELİRLERİ", "DÖNEM NET KARI", "PERSONEL GİDERLERİ", "NET FAALIYET"],
        "fields": [
            {"id": "faiz_gel",              "label": "Faiz Gelirleri",                                              "col": "Cari Dönem"},
            {"id": "faiz_gid",              "label": "Faiz Giderleri (-)",                                          "col": "Cari Dönem"},
            {"id": "net_ucret_kom",         "label": "Net Ücret ve Komisyon Gelirleri/Giderleri",                   "col": "Cari Dönem"},
            {"id": "alinan_ucret",          "label": "Alınan Ücret ve Komisyonlar",                                 "col": "Cari Dönem"},
            {"id": "verilen_ucret",         "label": "Verilen Ücret ve Komisyonlar (-)",                            "col": "Cari Dönem"},
            {"id": "kar_sinif_olmayan",     "label": "Kar veya Zararda Yeniden Sınıflandırılmayacaklar",            "col": "Cari Dönem"},
            {"id": "kar_sinif_olan",        "label": "Kar veya Zararda Yeniden Sınıflandırılacaklar",               "col": "Cari Dönem"},
            {"id": "temettü",              "label": "Temettü Gelirleri",                                           "col": "Cari Dönem"},
            {"id": "ticari_kz",             "label": "Ticari Kar/Zarar (Net)",                                      "col": "Cari Dönem"},
            {"id": "diger_faal_gel",        "label": "Diğer Faaliyet Gelirleri",                                    "col": "Cari Dönem"},
            {"id": "beklenen_zarar_gid",    "label": "Beklenen Zarar Karşılıkları (-)",                             "col": "Cari Dönem"},
            {"id": "diger_karsilik_gid",    "label": "Diğer Karşılık Giderleri (-)",                               "col": "Cari Dönem"},
            {"id": "personel_gid",          "label": "Personel Giderleri (-)",                                      "col": "Cari Dönem"},
            {"id": "diger_faal_gid",        "label": "Diğer Faaliyet Giderleri (-)",                               "col": "Cari Dönem"},
            {"id": "net_faal_kz",           "label": "Net Faaliyet Karı/Zararı",                                    "col": "Cari Dönem"},
            {"id": "surd_vergi_oncesi",     "label": "Sürdürülen Faaliyetler Vergi Öncesi K/Z",                    "col": "Cari Dönem"},
            {"id": "surd_vergi",            "label": "Sürdürülen Faaliyetler Vergi Karşılığı",                     "col": "Cari Dönem"},
            {"id": "donem_net_kz",          "label": "Dönem Net Karı/Zararı",                                       "col": "Cari Dönem"},
        ]
    },
    {
        "id": "vergi",
        "name": "VERGİ DETAYI",
        "table_hint": "Sürdürülen faaliyetler vergi karşılığına ilişkin dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Cari Vergi Karşılığı", "Ertelenmiş Vergi Gider", "Ertelenmiş Vergi Gelir"],
        "fields": [
            {"id": "ozkaynak_ortaklik",     "label": "Özkaynak Yöntemi Uygulanan Ortaklıklardan Kar/Zarar",        "col": "Cari Dönem"},
            {"id": "cari_vergi_karsiliği",  "label": "Cari Vergi Karşılığı",                                       "col": "Cari Dönem"},
            {"id": "ertelenmis_net",        "label": "Ertelenmiş Vergi Net",                                        "col": "Cari Dönem"},
            {"id": "ertelenmis_gider",      "label": "Ertelenmiş Vergi Gider Etkisi (+)",                           "col": "Cari Dönem"},
            {"id": "ertelenmis_gelir",      "label": "Ertelenmiş Vergi Gelir Etkisi (-)",                           "col": "Cari Dönem"},
        ]
    },
    {
        "id": "faiz_detay",
        "name": "FAİZ GELİR/GİDER DETAYI",
        "table_hint": "Faiz gelirlerine ilişkin açıklamalar dipnot tablosu. TL ve YP sütunları ayrı.",
        "keywords": ["Kredilerden Alınan Faiz", "Zorunlu Karşılıklardan", "Mevduata Verilen Faizler"],
        "fields": [
            {"id": "kred_faiz_gel_tl",      "label": "Kredilerden Alınan Faiz Gelirleri",                          "col": "TL"},
            {"id": "kred_faiz_gel_yp",      "label": "Kredilerden Alınan Faiz Gelirleri",                          "col": "YP"},
            {"id": "zorunlu_kars_faiz",     "label": "Zorunlu Karşılıklardan Alınan Faizler",                      "col": "Cari Dönem"},
            {"id": "bankalar_faiz",         "label": "Bankalardan Alınan Faizler",                                  "col": "Cari Dönem"},
            {"id": "para_piy_faiz_gel",     "label": "Para Piyasası İşlemlerinden Alınan Faizler",                 "col": "Cari Dönem"},
            {"id": "menkul_faiz",           "label": "Menkul Değerlerden Alınan Faizler",                          "col": "Cari Dönem"},
            {"id": "fin_kira_faiz",         "label": "Finansal Kiralama Faiz Gelirleri",                           "col": "Cari Dönem"},
            {"id": "diger_faiz_gel",        "label": "Diğer Faiz Gelirleri",                                       "col": "Cari Dönem"},
            {"id": "mevduat_faiz_gid",      "label": "Mevduata Verilen Faizler",                                   "col": "Cari Dönem"},
            {"id": "kred_faiz_gid",         "label": "Kullanılan Kredilere Verilen Faizler",                       "col": "Cari Dönem"},
            {"id": "para_piy_faiz_gid",     "label": "Para Piyasası İşlemlerine Verilen Faizler",                  "col": "Cari Dönem"},
            {"id": "ihrac_faiz_gid",        "label": "İhraç Edilen Menkul Kıymetlere Verilen Faizler",             "col": "Cari Dönem"},
            {"id": "kira_faiz_gid",         "label": "Kiralama Faiz Giderleri",                                    "col": "Cari Dönem"},
            {"id": "diger_faiz_gid",        "label": "Diğer Faiz Giderleri",                                       "col": "Cari Dönem"},
        ]
    },
    {
        "id": "ort_faiz",
        "name": "ORTALAMA FAİZ ORANLARI (%)",
        "table_hint": "Parasal finansal araçlara uygulanan ortalama faiz oranları tablosu. TL ve YP (USD) sütunları.",
        "keywords": ["ortalama faiz", "Verilen Krediler", "Bankalar Mevduatı", "İhraç Edilen Menkul"],
        "fields": [
            {"id": "vkred_tl",              "label": "Verilen Krediler",                        "col": "TL"},
            {"id": "bank_mev_tl",           "label": "Bankalar Mevduatı",                       "col": "TL"},
            {"id": "diger_mev_tl",          "label": "Diğer Mevduat",                           "col": "TL"},
            {"id": "ihrac_menkul_oran_tl",  "label": "İhraç Edilen Menkul Değerler (Net)",      "col": "TL"},
            {"id": "vkred_yp",              "label": "Verilen Krediler",                        "col": "YP (USD)"},
            {"id": "bank_mev_yp",           "label": "Bankalar Mevduatı",                       "col": "YP (USD)"},
            {"id": "diger_mev_yp",          "label": "Diğer Mevduat",                           "col": "YP (USD)"},
            {"id": "ihrac_menkul_oran_yp",  "label": "İhraç Edilen Menkul Değerler (Net)",      "col": "YP (USD)"},
        ]
    },
    {
        "id": "ticari",
        "name": "TİCARİ KAR/ZARAR DETAYI",
        "table_hint": "Ticari kar/zarara ilişkin açıklamalar dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Sermaye Piyasası İşlemleri Kârı", "Türev Finansal İşlemlerden", "Kambiyo İşlemlerinden"],
        "fields": [
            {"id": "ticari_kar",            "label": "Ticari Kar",                              "col": "Cari Dönem"},
            {"id": "ser_piy_kar",           "label": "Sermaye Piyasası İşlemleri Karı",         "col": "Cari Dönem"},
            {"id": "turev_kar",             "label": "Türev Finansal İşlemlerden Kar",           "col": "Cari Dönem"},
            {"id": "kambiyo_kar",           "label": "Kambiyo İşlemlerinden Kar",               "col": "Cari Dönem"},
            {"id": "ticari_zarar",          "label": "Ticari Zarar",                            "col": "Cari Dönem"},
            {"id": "ser_piy_zarar",         "label": "Sermaye Piyasası İşlemleri Zararı",       "col": "Cari Dönem"},
            {"id": "turev_zarar",           "label": "Türev Finansal İşlemlerden Zarar",         "col": "Cari Dönem"},
            {"id": "kambiyo_zarar",         "label": "Kambiyo İşlemlerinden Zarar",             "col": "Cari Dönem"},
        ]
    },
    {
        "id": "bkzk",
        "name": "BEKLENEN KREDİ ZARAR KARŞILIKLARI",
        "table_hint": "Beklenen zarar karşılıkları dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Birinci Aşama", "İkinci Aşama", "Temerrüt", "Beklenen Kredi Zarar"],
        "fields": [
            {"id": "birinci_asama",         "label": "12 Aylık Beklenen Zarar Karşılığı (Birinci Aşama)",  "col": "Cari Dönem"},
            {"id": "ikinci_asama",          "label": "Kredi Riskinde Önemli Artış (İkinci Aşama)",         "col": "Cari Dönem"},
            {"id": "ucuncu_asama",          "label": "Temerrüt (Üçüncü Aşama)",                            "col": "Cari Dönem"},
        ]
    },
    {
        "id": "faal_gider",
        "name": "DİĞER FAALİYET GİDERLERİ DETAYI",
        "table_hint": "Diğer faaliyet giderlerine ilişkin bilgiler dipnot tablosu. Cari dönem sütunu.",
        "keywords": ["Kıdem Tazminat", "Amortisman", "TFRS 16", "Reklam ve İlan", "Bakım ve Onarım"],
        "fields": [
            {"id": "kidem_taz",             "label": "Kıdem Tazminat Karşılığı",                           "col": "Cari Dönem"},
            {"id": "amortisman",            "label": "Amortisman Giderleri",                               "col": "Cari Dönem"},
            {"id": "diger_isletme",         "label": "Diğer İşletme Giderleri",                            "col": "Cari Dönem"},
            {"id": "tfrs16_kira",           "label": "TFRS 16 İstisnalarına İlişkin Kiralama Giderleri",   "col": "Cari Dönem"},
            {"id": "bakim_onarim",          "label": "Bakım ve Onarım Giderleri",                          "col": "Cari Dönem"},
            {"id": "reklam_ilan",           "label": "Reklam ve İlan Giderleri",                           "col": "Cari Dönem"},
            {"id": "diger_giderler_1",      "label": "Diğer Giderler (1)",                                 "col": "Cari Dönem"},
            {"id": "aktif_satis_zarar",     "label": "Aktiflerin Satışından Doğan Zararlar",               "col": "Cari Dönem"},
            {"id": "diger_2",               "label": "Diğer (2)",                                          "col": "Cari Dönem"},
        ]
    },
    {
        "id": "kredi_kalite",
        "name": "KREDİ KALİTE TABLOLARI",
        "table_hint": "Standart nitelikli/yakın izlemedeki krediler ve sorunlu kredi grupları dipnot tabloları.",
        "keywords": ["Standart Nitelikli", "Yakın İzlemedeki", "Tahsil İmkânı Sınırlı", "Tahsili Şüpheli"],
        "fields": [
            {"id": "std_nitelikli",         "label": "Standart Nitelikli Krediler",                                                    "col": "Cari Dönem"},
            {"id": "yakin_yap_disinda",     "label": "Yakın İzlemedeki - Yeniden Yapılandırma Kapsamında Yer Almayanlar",              "col": "Cari Dönem"},
            {"id": "yakin_sozlesme",        "label": "Yakın İzlemedeki - Sözleşme Koşullarında Değişiklik",                           "col": "Cari Dönem"},
            {"id": "yakin_yeniden_fin",     "label": "Yakın İzlemedeki - Yeniden Finansman",                                          "col": "Cari Dönem"},
            {"id": "konut_kred",            "label": "Konut Kredileri (Tüketici+Personel)",                                           "col": "Cari Dönem"},
            {"id": "tasit_kred",            "label": "Taşıt Kredileri (Tüketici+Personel)",                                           "col": "Cari Dönem"},
            {"id": "ihtiyac_kred",          "label": "İhtiyaç Kredileri (Tüketici+Personel)",                                         "col": "Cari Dönem"},
            {"id": "diger_kred_tipi",       "label": "Diğer (Tüketici+Personel)",                                                     "col": "Cari Dönem"},
            {"id": "kredi_karti",           "label": "Kredi Kartları",                                                                "col": "Cari Dönem"},
            {"id": "kmh",                   "label": "KMH",                                                                           "col": "Cari Dönem"},
            {"id": "gr3_bakiye",            "label": "III. Grup Tahsil İmkânı Sınırlı - Dönem Sonu Bakiyesi",                        "col": "Cari Dönem"},
            {"id": "gr3_karsilik",          "label": "III. Grup Tahsil İmkânı Sınırlı - Karşılık (-)",                               "col": "Cari Dönem"},
            {"id": "gr3_net",               "label": "III. Grup Tahsil İmkânı Sınırlı - Bilançodaki Net Bakiyesi",                   "col": "Cari Dönem"},
            {"id": "gr4_bakiye",            "label": "IV. Grup Tahsili Şüpheli - Dönem Sonu Bakiyesi",                               "col": "Cari Dönem"},
            {"id": "gr4_karsilik",          "label": "IV. Grup Tahsili Şüpheli - Karşılık (-)",                                      "col": "Cari Dönem"},
            {"id": "gr4_net",               "label": "IV. Grup Tahsili Şüpheli - Bilançodaki Net Bakiyesi",                          "col": "Cari Dönem"},
            {"id": "gr5_bakiye",            "label": "V. Grup Zarar Niteliğindeki - Dönem Sonu Bakiyesi",                            "col": "Cari Dönem"},
            {"id": "gr5_karsilik",          "label": "V. Grup Zarar Niteliğindeki - Karşılık (-)",                                   "col": "Cari Dönem"},
            {"id": "gr5_net",               "label": "V. Grup Zarar Niteliğindeki - Bilançodaki Net Bakiyesi",                       "col": "Cari Dönem"},
        ]
    },
]


# =============================================================================
# GÜVENLİ JSON PARSER
# Claude çoğu zaman cevabını ```json ... ``` bloğu içinde döndürür.
# Bu fonksiyon markdown tırnaklarını temizleyip güvenli parse yapar.
# =============================================================================
def get_client(project_id: str) -> anthropic.AnthropicVertex:
    return anthropic.AnthropicVertex(
        project_id=project_id,
        region=REGION,
    )


# =============================================================================
# GÜVENLİ JSON PARSER
# Claude çoğu zaman cevabını ```json ... ``` bloğu içinde döndürür.
# Parantez içi negatif değerleri de düzeltir: (1.234) → -1234
# =============================================================================
def safe_parse_json(raw: str) -> dict:
    try:
        # Markdown kod bloklarını temizle
        clean = re.sub(r'```(?:json)?', '', raw).strip()
        clean = clean.replace('```', '').strip()

        # JSON bloğunu bul
        match = re.search(r'\{.*\}', clean, re.DOTALL)
        if match:
            result = json.loads(match.group())
        else:
            result = json.loads(clean)

        # Parantez içi değerleri negatife çevir: "(1.234.567)" → "-1.234.567"
        for k, v in result.items():
            if isinstance(v, str) and v.startswith('(') and v.endswith(')'):
                result[k] = '-' + v[1:-1]

        return result

    except Exception as e:
        print(f"JSON Parse Hatası: {e}\nHam metin: {raw[:200]}")
        return {}


# =============================================================================
# ADIM 1: PDFPLUMBER — METİN TABANLARI EXTRACTION (ÜCRETSİZ)
# =============================================================================
def extract_pages_text(pdf_bytes: bytes) -> dict[int, str]:
    """Her sayfanın metnini döndürür: {sayfa_no: metin}"""
    pages = {}
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        with pdfplumber.open(tmp_path) as pdf:
            for i, page in enumerate(pdf.pages):
                pages[i] = page.extract_text() or ""
    finally:
        os.remove(tmp_path)
    return pages


def find_relevant_pages(pages_text: dict, keywords: list[str], max_pages: int = 5) -> list[int]:
    """Anahtar kelimelere göre en alakalı sayfaları bulur."""
    scored = []
    for idx, text in pages_text.items():
        upper = text.upper()
        score = sum(1 for kw in keywords if kw.upper() in upper)
        if score > 0:
            scored.append((idx, score))
    scored.sort(key=lambda x: -x[1])
    return [idx for idx, _ in scored[:max_pages]]


def build_text_context(pages_text: dict, page_indices: list[int]) -> str:
    """Seçili sayfaların metnini birleştirir."""
    parts = []
    for idx in page_indices:
        text = pages_text.get(idx, "")
        if text.strip():
            parts.append(f"--- SAYFA {idx+1} ---\n{text}")
    return "\n\n".join(parts)


# =============================================================================
# ADIM 2: CLAUDE API — METİN TABANLARI (UCUZ)
# =============================================================================
def extract_with_text(
    client: anthropic.Anthropic,
    group: dict,
    context_text: str,
    banka_adi: str,
    donem: str
) -> tuple[dict, bool]:
    """
    Metin tabanlı extraction.
    Döndürür: (results_dict, success_bool)
    Başarı: tüm alanların %70'i bulunursa True
    """
    fields_desc = "\n".join([
        f'- "{f["id"]}": {f["label"]} ({f["col"]} sütunu)'
        for f in group["fields"]
    ])

    prompt = f"""Sen kıdemli bir banka denetçisisin.
Banka: {banka_adi} | Dönem: {donem}
Tablo: {group["table_hint"]}

## TABLO SÜTUN YAPISI (soldan sağa):
Bilanço tablosunda sütunlar şu sıradadır:
[Satır No] | [Açıklama] | [Dipnot] | Cari TP | Cari YP | Cari Toplam | Önceki TP | Önceki YP | Önceki Toplam

SADECE "Cari TP" (4. sütun) ve "Cari YP" (5. sütun) değerlerini al.
"Önceki" sütunları (7., 8., 9.) KESİNLİKLE ALMA — bunlar geçen yılın verileridir.

## NEGATİF DEĞERLER:
Tabloda parantez içindeki değerler negatiftir: (834,735) → "-834,735" olarak yaz.

## ÖRNEK (few-shot):
Tablo satırı: "2.5 | Beklenen Zarar Karşılıkları (-) | | 834,735 | 77,789 | 912,524 | 684,698 | 68,995 | 753,693"
Doğru cevap: beklenen_zarar_tp = "834,735", beklenen_zarar_yp = "77,789"
YANLIŞ cevap: beklenen_zarar_tp = "684,698" (bu Önceki Dönem TP'dir!)

## ALMAM GEREKEN DEĞERLER:
{fields_desc}

## KRİTİK KURALLAR:
1. SADECE Cari Dönem (ilk 3 sayısal sütun). Önceki dönem = KESİNLİKLE ALMA.
2. TP ve YP sütunlarını karıştırma — TP her zaman YP'nin solundadır.
3. Değeri net göremiyorsan null yaz. Asla tahmin etme.
4. Sayıları tabloda nasılsa birebir al (örn: "5,138,035").
5. Parantez içi = negatif: "(1.234)" → "-1.234"
6. SADECE JSON döndür, başka hiçbir şey yazma:
{{"field_id": "deger_veya_null"}}

## TABLO METNİ:
{context_text[:6000]}"""

    # Retry mekanizması — Vertex bazen 500 dönebilir
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-haiku-4-5@20251001",
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}]
            )
            break
        except Exception as e:
            if attempt == 2:
                print(f"Haiku API hatası (3 deneme): {e}")
                return {}, False
            time.sleep(30 + 2 ** attempt)  # 1s, 2s, 4s

    raw = response.content[0].text.strip()
    results = safe_parse_json(raw)
    if results:
        found = sum(1 for v in results.values() if v and v != "null" and v is not None)
        success = found >= len(group["fields"]) * 0.7
        return results, success
    return {}, False


# =============================================================================
# ADIM 3: CLAUDE API — GÖRÜNTÜ TABANLARI FALLBACK (GÜVENİLİR)
# =============================================================================
def pdf_page_to_image(pdf_bytes: bytes, page_index: int):
    """Sadece tek bir sayfayı görüntüye çevirir — bellek dostu."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        images = convert_from_path(
            tmp_path,
            dpi=180,
            first_page=page_index + 1,  # pdf2image 1-indexed
            last_page=page_index + 1
        )
        return images[0] if images else None
    finally:
        os.remove(tmp_path)


def pdf_pages_to_images(pdf_bytes: bytes, page_indices: list[int]) -> list:
    """Sadece belirtilen sayfaları görüntüye çevirir."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        images = []
        for idx in page_indices[:5]:  # Max 5 sayfa
            pages = convert_from_path(
                tmp_path,
                dpi=180,
                first_page=idx + 1,
                last_page=idx + 1
            )
            if pages:
                images.append(pages[0])
        return images
    finally:
        os.remove(tmp_path)


def image_to_base64(image) -> str:
    buf = BytesIO()
    image.save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8")


def extract_with_vision(
    client,
    group: dict,
    pdf_bytes: bytes,
    page_indices: list[int],
    banka_adi: str,
    donem: str
) -> dict:
    """Görüntü tabanlı extraction — fallback. Sadece ilgili sayfaları lazy load eder."""
    fields_desc = "\n".join([
        f'- "{f["id"]}": {f["label"]} ({f["col"]} sütunu)'
        for f in group["fields"]
    ])

    # Lazy load — sadece bu grup için ilgili sayfaları render et
    page_images = pdf_pages_to_images(pdf_bytes, page_indices)

    content = []
    for img in page_images[:5]:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png",
                       "data": image_to_base64(img)}
        })
    content.append({
        "type": "text",
        "text": f"""Sen kıdemli bir banka denetçisisin.
Banka: {banka_adi} | Dönem: {donem}
Tablo: {group["table_hint"]}

Tabloda sütun sırası (soldan sağa):
[Satır No] | [Açıklama] | [Dipnot] | Cari TP | Cari YP | Cari Toplam | Önceki TP | Önceki YP | Önceki Toplam

SADECE Cari TP (4. sütun) ve Cari YP (5. sütun) al. Önceki dönem sütunlarına bakma.
Parantez içi değerler negatiftir: (834,735) → "-834,735"

ALMAM GEREKEN DEĞERLER:
{fields_desc}

KURALLAR:
1. SADECE Cari Dönem. Önceki dönem = KESİNLİKLE ALMA.
2. TP ve YP karıştırma — TP her zaman YP'nin solundadır.
3. Bulamazsan null yaz. SADECE JSON döndür:
{{"field_id": "deger_veya_null"}}"""
    })

    # Retry mekanizması
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-5@20251001",
                max_tokens=1500,
                messages=[{"role": "user", "content": content}]
            )
            break
        except Exception as e:
            if attempt == 2:
                print(f"Sonnet Vision API hatası (3 deneme): {e}")
                return {}
            time.sleep(30 + 2 ** attempt)

    raw = response.content[0].text.strip()
    result = safe_parse_json(raw)
    return result if result else {}


def extract_metadata(client: anthropic.Anthropic, first_page_text: str) -> tuple[str, str]:
    response = client.messages.create(
        model="claude-haiku-4-5@20251001",
        max_tokens=150,
        messages=[{
            "role": "user",
            "content": f'Bu banka raporunun metninden bankanın tam resmi adını ve rapor dönemini bul.\nSADECE JSON: {{"banka_adi": "...", "donem": "..."}}\n\nMETİN:\n{first_page_text[:1500]}'
        }]
    )
    raw = response.content[0].text.strip()
    data = safe_parse_json(raw)
    if data:
        return data.get("banka_adi", "Bilinmeyen"), data.get("donem", "Bilinmiyor")
    return "Bilinmeyen Banka", "Bilinmiyor"


# =============================================================================
# EXCEL EXPORT
# =============================================================================
def create_excel(banka_adi: str, donem: str, all_results: dict,
                 method_log: dict) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Denetim Verileri"

    h_fill = PatternFill("solid", fgColor="1F4E79")
    h_font = Font(color="FFFFFF", bold=True, size=11)
    g_fill = PatternFill("solid", fgColor="D6E4F0")
    g_font = Font(bold=True, size=10)
    miss_fill = PatternFill("solid", fgColor="FFE0E0")
    vis_fill = PatternFill("solid", fgColor="FFF3CD")  # Sarı = vision ile bulundu
    thin = Side(style="thin", color="CCCCCC")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:E1")
    ws["A1"] = f"{banka_adi} — {donem}"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28
    ws.append([])

    ws.append(["Alan ID", "Açıklama", "Sütun", "Değer", "Kaynak"])
    for cell in ws[3]:
        cell.fill = h_fill
        cell.font = h_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = brd
    ws.row_dimensions[3].height = 20

    for group in EXTRACTION_GROUPS:
        ws.append([])
        method = method_log.get(group["id"], "—")
        ws.append([group["name"], "", "", "", f"Yöntem: {method}"])
        r = ws.max_row
        for cell in ws[r]:
            cell.fill = g_fill
            cell.font = g_font
            cell.border = brd

        for field in group["fields"]:
            val = all_results.get(field["id"])
            src = method_log.get(group["id"], "—")
            ws.append([field["id"], field["label"], field["col"],
                       val if val and val != "null" else "—", src])
            ri = ws.max_row
            for cell in ws[ri]:
                cell.border = brd
            if not val or val == "null":
                for cell in ws[ri]:
                    cell.fill = miss_fill
            elif src == "Vision":
                for cell in ws[ri]:
                    cell.fill = vis_fill

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 12

    # Özet sayfası
    ws2 = wb.create_sheet("Özet")
    ws2.append(["Grup", "Toplam Alan", "Bulunan", "Başarı %", "Yöntem"])
    for cell in ws2[1]:
        cell.fill = h_fill
        cell.font = h_font
    for group in EXTRACTION_GROUPS:
        total = len(group["fields"])
        found = sum(1 for f in group["fields"]
                    if all_results.get(f["id"]) and all_results.get(f["id"]) != "null")
        ws2.append([
            group["name"], total, found,
            f"%{found/total*100:.0f}",
            method_log.get(group["id"], "—")
        ])
    for col, w in [("A", 35), ("B", 12), ("C", 10), ("D", 10), ("E", 10)]:
        ws2.column_dimensions[col].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# =============================================================================
# STREAMLIT ARAYÜZÜ
# =============================================================================
st.set_page_config(page_title="Banka Denetim Analizi", layout="wide")
st.title("🏦 Banka Denetim Raporu Analizi")
st.caption("Hibrit sistem: pdfplumber (ücretsiz) → Claude Haiku metin → Claude Sonnet Vision (gerekirse) | Google Vertex AI")

with st.sidebar:
    st.subheader("☁️ Google Vertex AI")
    project_id = st.text_input(
        "Google Cloud Project ID",
        value=os.getenv("GOOGLE_CLOUD_PROJECT", PROJECT_ID),
        help="Google Cloud Console'dan proje ID'ni gir (örn: project-62aee5aa-8c4b-4416-9e7)"
    )
    if project_id:
        st.success("✅ Project ID girildi")
    else:
        st.warning("⚠️ Project ID gerekli")

    st.info("🔐 Authentication: gcloud auth application-default login ile yapıldı — ekstra key gerekmez.")

    st.divider()
    st.subheader("💰 Maliyet Tahmini")
    st.info(
        "**Sadece metin:** ~$0.01/rapor\n\n"
        "**Metin + Vision:** ~$0.05/rapor\n\n"
        "**Google $300 kredin var!**"
    )

    st.divider()
    total_fields = sum(len(g["fields"]) for g in EXTRACTION_GROUPS)
    st.metric("Toplam Veri Noktası", total_fields)
    st.metric("Grup Sayısı", len(EXTRACTION_GROUPS))

    st.divider()
    force_vision = st.toggle("Vision'ı zorla (test için)", value=False)

uploaded_file = st.file_uploader("📄 PDF Rapor Yükle", type="pdf")

if uploaded_file and project_id:
    if st.button("🚀 Analiz Et", type="primary"):
        client = get_client(project_id)
        pdf_bytes = uploaded_file.getvalue()
        all_results = {}
        method_log = {}
        vision_used = False

        progress = st.progress(0, "Başlatılıyor...")

        # 1. PDF → metin
        progress.progress(5, "📄 PDF metne dönüştürülüyor (pdfplumber)...")
        pages_text = extract_pages_text(pdf_bytes)
        st.success(f"✅ {len(pages_text)} sayfa okundu")

        # 2. Metadata
        progress.progress(8, "🤖 Banka adı ve dönem...")
        first_text = pages_text.get(0, "") + pages_text.get(1, "")
        banka_adi, donem = extract_metadata(client, first_text)
        st.info(f"🏦 **{banka_adi}** | 📅 **{donem}**")

        # 3. Her grup
        total_groups = len(EXTRACTION_GROUPS)
        vision_used = False

        for g_idx, group in enumerate(EXTRACTION_GROUPS):
            pct = 8 + int((g_idx / total_groups) * 87)
            progress.progress(pct, f"🔍 {group['name']}...")

            with st.status(f"📊 {group['name']}", expanded=False) as status:
                # İlgili sayfaları bul
                page_indices = find_relevant_pages(pages_text, group["keywords"])
                if not page_indices:
                    page_indices = list(range(min(4, len(pages_text))))

                context = build_text_context(pages_text, page_indices)
                st.caption(f"İlgili sayfalar: {[i+1 for i in page_indices]}")

                result = {}
                used_vision = False

                if not force_vision:
                    # ADIM 1: Metin ile dene (ucuz)
                    st.caption("1️⃣ Metin tabanlı deneniyor (Haiku)...")
                    result, success = extract_with_text(
                        client, group, context, banka_adi, donem
                    )

                    if success:
                        method_log[group["id"]] = "Metin (Haiku)"
                        found = sum(1 for v in result.values() if v and v != "null")
                        status.update(
                            label=f"✅ {group['name']} — {found}/{len(group['fields'])} | Metin",
                            state="complete"
                        )
                    else:
                        # ADIM 2: Vision fallback — lazy load, sadece bu sayfaları render et
                        st.caption("2️⃣ Metin yetersiz — Vision fallback (Sonnet)...")
                        result = extract_with_vision(
                            client, group, pdf_bytes, page_indices, banka_adi, donem
                        )
                        used_vision = True
                        vision_used = True
                        method_log[group["id"]] = "Vision (Sonnet)"
                        found = sum(1 for v in result.values() if v and v != "null")
                        status.update(
                            label=f"✅ {group['name']} — {found}/{len(group['fields'])} | Vision 🖼️",
                            state="complete"
                        )
                else:
                    # Force vision modu — lazy load
                    result = extract_with_vision(
                        client, group, pdf_bytes, page_indices, banka_adi, donem
                    )
                    method_log[group["id"]] = "Vision (Sonnet)"
                    found = sum(1 for v in result.values() if v and v != "null")
                    status.update(
                        label=f"✅ {group['name']} — {found}/{len(group['fields'])} | Vision",
                        state="complete"
                    )

                all_results.update(result)
                time.sleep(15)

        progress.progress(100, "✅ Tamamlandı!")

        # 4. Sonuçlar
        st.divider()
        total_found = sum(1 for v in all_results.values() if v and v != "null")
        total_fields = sum(len(g["fields"]) for g in EXTRACTION_GROUPS)
        text_groups = sum(1 for v in method_log.values() if "Metin" in v)
        vision_groups = sum(1 for v in method_log.values() if "Vision" in v)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Toplam Alan", total_fields)
        c2.metric("Bulunan", total_found)
        c3.metric("Başarı", f"%{total_found/total_fields*100:.0f}")
        c4.metric("Metin ile", f"{text_groups} grup")
        c5.metric("Vision ile", f"{vision_groups} grup")

        # Grup detayları
        for group in EXTRACTION_GROUPS:
            found = sum(1 for f in group["fields"]
                       if all_results.get(f["id"]) and all_results.get(f["id"]) != "null")
            method = method_log.get(group["id"], "—")
            icon = "🖼️" if "Vision" in method else "📝"
            with st.expander(f"{icon} {group['name']} — {found}/{len(group['fields'])} | {method}"):
                cols = st.columns(3)
                for i, field in enumerate(group["fields"]):
                    val = all_results.get(field["id"])
                    with cols[i % 3]:
                        st.metric(
                            f"{field['label'][:28]} ({field['col']})",
                            val if val and val != "null" else "⚠️ —"
                        )

        # Excel
        excel_buf = create_excel(banka_adi, donem, all_results, method_log)
        st.download_button(
            "📥 Excel İndir",
            excel_buf,
            file_name=f"{banka_adi}_{donem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif uploaded_file and not project_id:
    st.warning("⚠️ Sol sidebar'dan Google Cloud Project ID girin.")
