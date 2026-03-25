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
import anthropic

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
        # DEĞİŞİKLİK: Kar/Zarar tablosu + Kapsamlı Gelir tablosu birlikte aranıyor
        "table_hint": "KONSOLİDE OLMAYAN KAR VEYA ZARAR TABLOSU ve KAR VEYA ZARAR VE DİĞER KAPSAMLI GELİR TABLOSU. Sadece Cari Dönem sütunu.",
        "keywords": ["FAİZ GELİRLERİ", "DÖNEM NET KARI", "PERSONEL GİDERLERİ", "KAPSAMLI GELİR", "YENİDEN SINIFLANDIRILA"],
        "fields": [
            {"id": "faiz_gel",              "label": "Faiz Gelirleri",                                                          "col": "Cari Dönem"},
            {"id": "faiz_gid",              "label": "Faiz Giderleri (-)",                                                      "col": "Cari Dönem"},
            # DEĞİŞİKLİK: fin_kira_faiz ve mevduat_faiz_gid gelir tablosuna taşındı
            {"id": "fin_kira_faiz",         "label": "Finansal Kiralama Faiz Gelirleri",                                        "col": "Cari Dönem"},
            {"id": "mevduat_faiz_gid",      "label": "Mevduata Verilen Faizler",                                                "col": "Cari Dönem"},
            {"id": "net_ucret_kom",         "label": "Net Ücret ve Komisyon Gelirleri/Giderleri",                               "col": "Cari Dönem"},
            {"id": "alinan_ucret",          "label": "Alınan Ücret ve Komisyonlar",                                             "col": "Cari Dönem"},
            {"id": "verilen_ucret",         "label": "Verilen Ücret ve Komisyonlar (-)",                                        "col": "Cari Dönem"},
            {"id": "temettü",              "label": "Temettü Gelirleri",                                                       "col": "Cari Dönem"},
            {"id": "ticari_kz",             "label": "Ticari Kar/Zarar (Net)",                                                  "col": "Cari Dönem"},
            {"id": "diger_faal_gel",        "label": "Diğer Faaliyet Gelirleri",                                                "col": "Cari Dönem"},
            {"id": "beklenen_zarar_gid",    "label": "Beklenen Zarar Karşılıkları (-)",                                         "col": "Cari Dönem"},
            {"id": "diger_karsilik_gid",    "label": "Diğer Karşılık Giderleri (-)",                                           "col": "Cari Dönem"},
            {"id": "personel_gid",          "label": "Personel Giderleri (-)",                                                  "col": "Cari Dönem"},
            {"id": "diger_faal_gid",        "label": "Diğer Faaliyet Giderleri (-)",                                           "col": "Cari Dönem"},
            {"id": "net_faal_kz",           "label": "Net Faaliyet Karı/Zararı",                                                "col": "Cari Dönem"},
            {"id": "surd_vergi_oncesi",     "label": "Sürdürülen Faaliyetler Vergi Öncesi K/Z",                                "col": "Cari Dönem"},
            {"id": "surd_vergi",            "label": "Sürdürülen Faaliyetler Vergi Karşılığı",                                 "col": "Cari Dönem"},
            {"id": "donem_net_kz",          "label": "Dönem Net Karı/Zararı",                                                   "col": "Cari Dönem"},
            # DEĞİŞİKLİK: Kapsamlı Gelir tablosundan alınacak
            {"id": "kar_sinif_olmayan",     "label": "Kar veya Zararda Yeniden Sınıflandırılmayacaklar (2.1)",                  "col": "Cari Dönem"},
            {"id": "kar_sinif_olan",        "label": "Kar veya Zararda Yeniden Sınıflandırılacaklar (2.2)",                     "col": "Cari Dönem"},
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
        # DEĞİŞİKLİK: fin_kira_faiz ve mevduat_faiz_gid çıkarıldı (gelir tablosuna taşındı)
        "table_hint": "Faiz gelirlerine ilişkin açıklamalar dipnot tablosu. TL ve YP sütunları ayrı.",
        "keywords": ["Kredilerden Alınan Faiz", "Zorunlu Karşılıklardan", "Kullanılan Kredilere"],
        "fields": [
            {"id": "kred_faiz_gel_tl",      "label": "Kredilerden Alınan Faiz Gelirleri",                          "col": "TL"},
            {"id": "kred_faiz_gel_yp",      "label": "Kredilerden Alınan Faiz Gelirleri",                          "col": "YP"},
            {"id": "zorunlu_kars_faiz",     "label": "Zorunlu Karşılıklardan Alınan Faizler",                      "col": "Cari Dönem"},
            {"id": "bankalar_faiz",         "label": "Bankalardan Alınan Faizler",                                  "col": "Cari Dönem"},
            {"id": "para_piy_faiz_gel",     "label": "Para Piyasası İşlemlerinden Alınan Faizler",                 "col": "Cari Dönem"},
            {"id": "menkul_faiz",           "label": "Menkul Değerlerden Alınan Faizler",                          "col": "Cari Dönem"},
            {"id": "diger_faiz_gel",        "label": "Diğer Faiz Gelirleri",                                       "col": "Cari Dönem"},
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
        # DEĞİŞİKLİK: Tablo yapısı güncellendi — EURO, USD, Yen, TL sütunları var
        # Biz sadece TL ve USD (YP) sütunlarını alıyoruz
        "table_hint": "Parasal finansal araçlara uygulanan ortalama faiz oranları tablosu. Sütunlar: EURO | USD | Yen | TL. Sadece USD ve TL sütunlarını al.",
        "keywords": ["ortalama faiz", "Verilen krediler", "Bankalar mevduatı", "İhraç edilen menkul"],
        "fields": [
            # DEĞİŞİKLİK: Tablo yapısına göre düzeltildi
            {"id": "vkred_tl",              "label": "Verilen Krediler",                            "col": "TL"},
            {"id": "bank_mev_tl",           "label": "Bankalar Mevduatı",                           "col": "TL"},
            {"id": "diger_mev_tl",          "label": "Diğer Mevduat",                               "col": "TL"},
            {"id": "ihrac_menkul_oran_tl",  "label": "İhraç Edilen Menkul Değerler (Net)",          "col": "TL"},
            {"id": "vkred_yp",              "label": "Verilen Krediler",                            "col": "USD"},
            {"id": "bank_mev_yp",           "label": "Bankalar Mevduatı",                           "col": "USD"},
            {"id": "diger_mev_yp",          "label": "Diğer Mevduat",                               "col": "USD"},
            {"id": "ihrac_menkul_oran_yp",  "label": "İhraç Edilen Menkul Değerler (Net)",          "col": "USD"},
        ]
    },
    {
        "id": "ticari",
        "name": "TİCARİ KAR/ZARAR DETAYI",
        # DEĞİŞİKLİK: Kar ve Zarar ana başlıkları + alt kalemler ayrı ayrı alınıyor
        "table_hint": "Ticari kar/zarara ilişkin açıklamalar dipnot tablosu. Kar toplamı, Zarar toplamı ve her birinin alt kalemleri ayrı ayrı alınacak. Cari dönem sütunu.",
        "keywords": ["Sermaye Piyasası İşlemleri", "Türev Finansal İşlemlerden", "Kambiyo İşlemlerinden", "Ticari kar"],
        "fields": [
            # DEĞİŞİKLİK: ticari_kar = "Kar" toplamı (kalın satır), ser_piy_kar = alt kalem
            {"id": "ticari_kar",            "label": "Kar (Toplam - Kalın Satır)",                "col": "Cari Dönem"},
            {"id": "ser_piy_kar",           "label": "Sermaye piyasası işlemleri kârı",            "col": "Cari Dönem"},
            {"id": "turev_kar",             "label": "Türev finansal işlemlerden kâr",             "col": "Cari Dönem"},
            {"id": "kambiyo_kar",           "label": "Kambiyo işlemlerinden kar",                  "col": "Cari Dönem"},
            {"id": "ticari_zarar",          "label": "Zarar (-) (Toplam - Kalın Satır)",           "col": "Cari Dönem"},
            {"id": "ser_piy_zarar",         "label": "Sermaye piyasası işlemleri zararı",          "col": "Cari Dönem"},
            {"id": "turev_zarar",           "label": "Türev finansal işlemlerden zarar",           "col": "Cari Dönem"},
            {"id": "kambiyo_zarar",         "label": "Kambiyo işlemlerinden zarar",                "col": "Cari Dönem"},
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
        "table_hint": "Standart nitelikli/yakın izlemedeki krediler ve tüketici/personel kredileri dipnot tabloları.",
        "keywords": ["Standart Nitelikli", "Yakın İzlemedeki", "Tüketici Kredileri", "Personel Kredileri", "Bireysel Kredi Kartları", "Kredili Müstakriz"],
        "fields": [
            {"id": "std_nitelikli",             "label": "Standart Nitelikli Krediler (Toplam)",                                           "col": "Cari Dönem"},
            {"id": "yakin_yap_disinda",          "label": "Yakın İzlemedeki - Yeniden Yapılandırma Kapsamında Yer Almayanlar",              "col": "Cari Dönem"},
            {"id": "yakin_sozlesme",             "label": "Yakın İzlemedeki - Sözleşme Koşullarında Değişiklik",                           "col": "Cari Dönem"},
            {"id": "yakin_yeniden_fin",          "label": "Yakın İzlemedeki - Yeniden Finansman",                                          "col": "Cari Dönem"},
            {"id": "tuk_konut",                  "label": "Tüketici Kredileri - Konut",                                                    "col": "Toplam"},
            {"id": "tuk_tasit",                  "label": "Tüketici Kredileri - Taşıt",                                                    "col": "Toplam"},
            {"id": "tuk_ihtiyac",                "label": "Tüketici Kredileri - İhtiyaç/Tüketici",                                         "col": "Toplam"},
            {"id": "tuk_diger",                  "label": "Tüketici Kredileri - Diğer",                                                    "col": "Toplam"},
            {"id": "tuk_kredi_karti",            "label": "Bireysel Kredi Kartları (Toplam)",                                              "col": "Toplam"},
            {"id": "per_konut",                  "label": "Personel Kredileri - Konut",                                                    "col": "Toplam"},
            {"id": "per_tasit",                  "label": "Personel Kredileri - Taşıt",                                                    "col": "Toplam"},
            {"id": "per_ihtiyac",                "label": "Personel Kredileri - İhtiyaç/Tüketici",                                         "col": "Toplam"},
            {"id": "per_diger",                  "label": "Personel Kredileri - Diğer",                                                    "col": "Toplam"},
            {"id": "per_kredi_karti",            "label": "Personel Kredi Kartları (Toplam)",                                              "col": "Toplam"},
            {"id": "kmh",                        "label": "Kredili Mevduat Hesabı (KMH)",                                                  "col": "Toplam"},
            {"id": "kredili_must_tp",            "label": "Kredili Müstakriz Hesabı-TP (Gerçek Kişi)",                                     "col": "Toplam"},
            {"id": "kredili_must_yp",            "label": "Kredili Müstakriz Hesabı-YP (Gerçek Kişi)",                                     "col": "Toplam"},
        ]
    },
    {
        "id": "donuk_alacak",
        "name": "DONUK ALACAK HAREKETLERİ",
        "table_hint": "Toplam donuk alacak hareketlerine ilişkin bilgiler tablosu. III, IV, V. Grup sütunları.",
        "keywords": ["Toplam donuk alacak", "Tahsil imkânı sınırlı", "Tahsili şüpheli", "Zarar niteliğindeki", "Dönem sonu bakiyesi", "Bilançodaki net"],
        "fields": [
            {"id": "gr3_bakiye",   "label": "III. Grup Tahsil İmkânı Sınırlı - Dönem Sonu Bakiyesi",      "col": "III. Grup"},
            {"id": "gr3_karsilik", "label": "III. Grup Tahsil İmkânı Sınırlı - Karşılık (-)",              "col": "III. Grup"},
            {"id": "gr3_net",      "label": "III. Grup Tahsil İmkânı Sınırlı - Bilançodaki Net Bakiyesi",  "col": "III. Grup"},
            {"id": "gr4_bakiye",   "label": "IV. Grup Tahsili Şüpheli - Dönem Sonu Bakiyesi",              "col": "IV. Grup"},
            {"id": "gr4_karsilik", "label": "IV. Grup Tahsili Şüpheli - Karşılık (-)",                     "col": "IV. Grup"},
            {"id": "gr4_net",      "label": "IV. Grup Tahsili Şüpheli - Bilançodaki Net Bakiyesi",         "col": "IV. Grup"},
            {"id": "gr5_bakiye",   "label": "V. Grup Zarar Niteliğindeki - Dönem Sonu Bakiyesi",           "col": "V. Grup"},
            {"id": "gr5_karsilik", "label": "V. Grup Zarar Niteliğindeki - Karşılık (-)",                  "col": "V. Grup"},
            {"id": "gr5_net",      "label": "V. Grup Zarar Niteliğindeki - Bilançodaki Net Bakiyesi",      "col": "V. Grup"},
        ]
    },
]

# =============================================================================
# ÖZEL PROMPT'LAR — bazı gruplar için özel extraction talimatları
# =============================================================================
SPECIAL_PROMPTS = {
    "gelir": """
## ÖZEL TALİMAT — GELİR TABLOSU:
Bu grupta iki farklı tablodan veri alınacak:

1. "KAR VEYA ZARAR TABLOSU" → faiz_gel, faiz_gid, fin_kira_faiz, mevduat_faiz_gid,
   net_ucret_kom, alinan_ucret, verilen_ucret, temettü, ticari_kz, diger_faal_gel,
   beklenen_zarar_gid, diger_karsilik_gid, personel_gid, diger_faal_gid,
   net_faal_kz, surd_vergi_oncesi, surd_vergi, donem_net_kz

2. "KAR VEYA ZARAR VE DİĞER KAPSAMLI GELİR TABLOSU" → şu satırları al:
   - kar_sinif_olmayan: "2.1 Kar veya Zararda Yeniden Sınıflandırılmayacaklar" satırının değeri
   - kar_sinif_olan: "2.2 Kâr veya Zararda Yeniden Sınıflandırılacaklar" satırının değeri
   Bu iki değer ayrı bir tabloda, dönem karı/zararının altında yer alır.
""",
    "ticari": """
## ÖZEL TALİMAT — TİCARİ KAR/ZARAR:
Tabloda şu yapı var (büyük/küçük harf ve aksanlı harflere dikkat et):
- "Kar" → kalın yazılmış toplam satır → ticari_kar
  - "Sermaye piyasası işlemleri kârı" → ser_piy_kar
  - "Türev finansal işlemlerden kâr" → turev_kar
  - "Kambiyo işlemlerinden kar" → kambiyo_kar
- "Zarar (-)" → kalın yazılmış toplam satır → ticari_zarar
  - "Sermaye piyasası işlemleri zararı" → ser_piy_zarar
  - "Türev finansal işlemlerden zarar" → turev_zarar
  - "Kambiyo işlemlerinden zarar" → kambiyo_zarar
- "Net ticari kar / (zarar)" → bu satırı KESİNLİKLE ALMA

DİKKAT: Büyük/küçük harf farkı önemli değil, içerik olarak eşleştir.
ticari_kar ve ticari_zarar kalın yazılmış TOPLAM satırlarıdır, alt kalemler değil.
""",
    "kredi_kalite": """
## ÖZEL TALİMAT — KREDİ KALİTE TABLOLARI VE TÜKETİCİ KREDİLERİ:

Tabloda Tüketici Kredileri ve Personel Kredileri iki ayrı ana başlık altında verilir.
Her ana başlık altında TP, Dövize Endeksli, YP alt grupları olabilir.
Sütunlar: Kısa Vadeli | Orta ve Uzun Vadeli | Toplam
SADECE "Toplam" sütununu al.

DİKKAT — Bankaya göre farklılık var:
- Bazı bankalarda "İhtiyaç Kredisi" yazar → tuk_ihtiyac/per_ihtiyac
- Bazı bankalarda "Tüketici" yazar → tuk_ihtiyac/per_ihtiyac
- Her iki isim de aynı field'a eşlenir.

ALIŞ KURALLARI (her zaman Toplam sütunu):
- tuk_konut: "Tüketici Kredileri-TP/YP/Dövize" altındaki TÜM "Konut Kredisi" satırlarını topla
- tuk_tasit: "Tüketici Kredileri-TP/YP/Dövize" altındaki TÜM "Taşıt Kredisi" satırlarını topla
- tuk_ihtiyac: "Tüketici Kredileri-TP/YP/Dövize" altındaki TÜM "İhtiyaç Kredisi" VEYA "Tüketici" satırlarını topla
- tuk_diger: "Tüketici Kredileri-TP/YP/Dövize" altındaki TÜM "Diğer" satırlarını topla
- tuk_kredi_karti: "Bireysel Kredi Kartları-TP" + "Bireysel Kredi Kartları-YP" toplamı
- per_konut: "Personel Kredileri-TP/YP/Dövize" altındaki TÜM "Konut Kredisi" satırlarını topla
- per_tasit: "Personel Kredileri-TP/YP/Dövize" altındaki TÜM "Taşıt Kredisi" satırlarını topla
- per_ihtiyac: "Personel Kredileri-TP/YP/Dövize" altındaki TÜM "İhtiyaç Kredisi" VEYA "Tüketici" satırlarını topla
- per_diger: "Personel Kredileri-TP/YP/Dövize" altındaki TÜM "Diğer" satırlarını topla
- per_kredi_karti: "Personel Kredi Kartları-TP" + "Personel Kredi Kartları-YP" toplamı
- kmh: "Kredili Mevduat Hesabı" veya "KMH" satırı → Toplam sütunu
- kredili_must_tp: "Kredili Müstakriz Hesabı-TP (Gerçek Kişi)" → Toplam sütunu
- kredili_must_yp: "Kredili Müstakriz Hesabı-YP (Gerçek Kişi)" → Toplam sütunu

Eğer bir satır değeri boş veya sadece "-" ise "0" döndür.
Toplama yaparken sadece sayısal değerleri topla, "-" olanları sıfır say.
""",
    "donuk_alacak": """
## ÖZEL TALİMAT — DONUK ALACAK HAREKETLERİ:
Tablo başlığı: "Toplam donuk alacak hareketlerine ilişkin bilgiler"
Sütunlar: III. Grup (Tahsil imkânı sınırlı) | IV. Grup (Tahsili şüpheli) | V. Grup (Zarar niteliğindeki)

SADECE şu satırları al (diğer hareket satırlarını alma):
- gr3_bakiye: III. Grup → "Dönem sonu bakiyesi" satırı (kalın yazılmış)
- gr3_karsilik: III. Grup → "Karşılık (-)" satırı
- gr3_net: III. Grup → "Bilançodaki net bakiyesi" satırı (kalın yazılmış)
- gr4_bakiye: IV. Grup → "Dönem sonu bakiyesi" satırı (kalın yazılmış)
- gr4_karsilik: IV. Grup → "Karşılık (-)" satırı
- gr4_net: IV. Grup → "Bilançodaki net bakiyesi" satırı (kalın yazılmış)
- gr5_bakiye: V. Grup → "Dönem sonu bakiyesi" satırı (kalın yazılmış)
- gr5_karsilik: V. Grup → "Karşılık (-)" satırı
- gr5_net: V. Grup → "Bilançodaki net bakiyesi" satırı (kalın yazılmış)

KESİNLİKLE ALMA:
- "Önceki Dönem Sonu Bakiyesi" veya "Önceki dönem sonu bakiyesi" satırı — geçen yıl verisi
- "Dönem içinde intikal", "Dönem içinde tahsilat", "Satılan", "Kayıttan düşülen" gibi hareket satırları
""",
    "ort_faiz": """
## ÖZEL TALİMAT — ORTALAMA FAİZ ORANLARI:
Tablo başlığı: "Parasal finansal araçlara uygulanan ortalama faiz oranları (%)"
Sütun sırası: EURO | USD | Yen | TL

Sadece "Varlıklar" bölümündeki satırları al:
- vkred_tl: "Verilen krediler" satırı → TL sütunu
- bank_mev_tl: "Bankalar mevduatı" satırı → TL sütunu  
- diger_mev_tl: "Diğer mevduat" satırı → TL sütunu
- ihrac_menkul_oran_tl: "İhraç edilen menkul değerler" satırı → TL sütunu
- vkred_yp: "Verilen krediler" satırı → USD sütunu
- bank_mev_yp: "Bankalar mevduatı" satırı → USD sütunu
- diger_mev_yp: "Diğer mevduat" satırı → USD sütunu
- ihrac_menkul_oran_yp: "İhraç edilen menkul değerler" satırı → USD sütunu

Bu değerler yüzde (%) olarak gelir, örn: "46.00" veya "9.01"
"""
}


# =============================================================================
# CLIENT
# =============================================================================
def get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


# =============================================================================
# GÜVENLİ JSON PARSER
# =============================================================================
def safe_parse_json(raw: str) -> dict:
    try:
        clean = re.sub(r'```(?:json)?', '', raw).strip()
        clean = clean.replace('```', '').strip()
        match = re.search(r'\{.*\}', clean, re.DOTALL)
        if match:
            result = json.loads(match.group())
        else:
            result = json.loads(clean)
        for k, v in result.items():
            if isinstance(v, str) and v.startswith('(') and v.endswith(')'):
                result[k] = '-' + v[1:-1]
            # Tabloda "-" yazıyorsa 0 yaz — veri bulunmuş ama sıfır demek
            elif isinstance(v, str) and v.strip() in ['-', '—', '--', '–']:
                result[k] = "0"
            # null/None ise olduğu gibi bırak — veri bulunamadı demek
        return result
    except Exception as e:
        print(f"JSON Parse Hatası: {e}\nHam metin: {raw[:200]}")
        return {}


# =============================================================================
# PDF METİN ÇIKARIMI
# =============================================================================
def extract_pages_text(pdf_bytes: bytes) -> dict[int, str]:
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
    scored = []
    for idx, text in pages_text.items():
        upper = text.upper()
        score = sum(1 for kw in keywords if kw.upper() in upper)
        if score > 0:
            scored.append((idx, score))
    scored.sort(key=lambda x: -x[1])
    return [idx for idx, _ in scored[:max_pages]]


def build_text_context(pages_text: dict, page_indices: list[int]) -> str:
    parts = []
    for idx in page_indices:
        text = pages_text.get(idx, "")
        if text.strip():
            parts.append(f"--- SAYFA {idx+1} ---\n{text}")
    return "\n\n".join(parts)


# =============================================================================
# CLAUDE API — METİN TABANLARI EXTRACTION
# =============================================================================
def extract_with_text(
    client: anthropic.Anthropic,
    group: dict,
    context_text: str,
    banka_adi: str,
    donem: str,
    model: str = "claude-haiku-4-5-20251001"
) -> tuple[dict, bool]:

    fields_desc = "\n".join([
        f'- "{f["id"]}": {f["label"]} ({f["col"]} sütunu)'
        for f in group["fields"]
    ])

    # Gruba özel ek talimat varsa ekle
    special = SPECIAL_PROMPTS.get(group["id"], "")

    prompt = f"""Sen kıdemli bir banka denetçisisin.
Banka: {banka_adi} | Dönem: {donem}
Tablo: {group["table_hint"]}

## TABLO SÜTUN YAPISI (Bilanço için soldan sağa):
[Satır No] | [Açıklama] | [Dipnot] | Cari TP | Cari YP | Cari Toplam | Önceki TP | Önceki YP | Önceki Toplam

SADECE "Cari TP" (4. sütun) ve "Cari YP" (5. sütun) değerlerini al.
"Önceki" sütunları KESİNLİKLE ALMA — bunlar geçen yılın verileridir.

## NEGATİF DEĞERLER:
Tabloda parantez içindeki değerler negatiftir: (834,735) → "-834,735" olarak yaz.

## ÖRNEK (few-shot):
Tablo satırı: "2.5 | Beklenen Zarar Karşılıkları (-) | | 834,735 | 77,789 | 912,524 | 684,698 | 68,995 | 753,693"
Doğru cevap: beklenen_zarar_tp = "834,735", beklenen_zarar_yp = "77,789"
YANLIŞ cevap: beklenen_zarar_tp = "684,698" (bu Önceki Dönem TP'dir!)
{special}
## ALMAM GEREKEN DEĞERLER:
{fields_desc}

## KRİTİK KURALLAR:
1. SADECE Cari Dönem. Önceki dönem = KESİNLİKLE ALMA.
2. TP ve YP sütunlarını karıştırma.
3. Tabloda o satırda "-" veya "—" yazıyorsa "0" döndür. Sadece veriyi hiç bulamıyorsan null yaz.
4. Sayıları tabloda nasılsa birebir al (örn: "5,138,035").
5. Parantez içi = negatif: "(1.234)" → "-1.234"
6. SADECE JSON döndür:
{{"field_id": "deger_veya_null"}}

## TABLO METNİ:
{context_text[:6000]}"""

    for attempt in range(3):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}]
            )
            break
        except Exception as e:
            if attempt == 2:
                print(f"Haiku API hatası (3 deneme): {e}")
                return {}, False
            time.sleep(30 + 2 ** attempt)

    raw = response.content[0].text.strip()
    results = safe_parse_json(raw)
    if results:
        found = sum(1 for v in results.values() if v and v != "null" and v is not None)
        success = found >= len(group["fields"]) * 0.7
        return results, success
    return {}, False


def extract_metadata(client: anthropic.Anthropic, first_page_text: str, model: str = "claude-haiku-4-5-20251001") -> tuple[str, str]:
    response = client.messages.create(
        model=model,
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
def create_excel(banka_adi: str, donem: str, all_results: dict, method_log: dict) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Denetim Verileri"

    h_fill = PatternFill("solid", fgColor="1F4E79")
    h_font = Font(color="FFFFFF", bold=True, size=11)
    g_fill = PatternFill("solid", fgColor="D6E4F0")
    g_font = Font(bold=True, size=10)
    miss_fill = PatternFill("solid", fgColor="FFE0E0")
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

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 14

    ws2 = wb.create_sheet("Özet")
    ws2.append(["Grup", "Toplam Alan", "Bulunan", "Başarı %", "Yöntem"])
    for cell in ws2[1]:
        cell.fill = h_fill
        cell.font = h_font
    for group in EXTRACTION_GROUPS:
        total = len(group["fields"])
        found = sum(1 for f in group["fields"]
                    if all_results.get(f["id"]) and all_results.get(f["id"]) != "null")
        ws2.append([group["name"], total, found, f"%{found/total*100:.0f}",
                    method_log.get(group["id"], "—")])
    for col, w in [("A", 35), ("B", 12), ("C", 10), ("D", 10), ("E", 16)]:
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
st.caption("pdfplumber + Claude Haiku | Anthropic API")

with st.sidebar:
    st.subheader("🔑 Anthropic API Key")
    api_key = st.text_input(
        "API Key", type="password",
        value=os.getenv("ANTHROPIC_API_KEY", ""),
        help="console.anthropic.com → API Keys"
    )
    if api_key:
        st.success("✅ API key girildi")
    else:
        st.warning("⚠️ API key gerekli")
        st.markdown("[→ API Key Al](https://console.anthropic.com)")

    st.divider()
    model_sec = st.selectbox(
        "🤖 Model",
        options=["claude-haiku-4-5-20251001", "claude-sonnet-4-5-20250929"],
        index=0,
        help="Haiku: hızlı ve ucuz (~$0.01/rapor) | Sonnet: daha güçlü (~$0.08/rapor)"
    )
    st.divider()
    total_fields = sum(len(g["fields"]) for g in EXTRACTION_GROUPS)
    st.metric("Toplam Veri Noktası", total_fields)
    st.metric("Grup Sayısı", len(EXTRACTION_GROUPS))
    st.divider()
    st.caption("💰 Maliyet: ~$0.01-0.03/rapor")

uploaded_file = st.file_uploader("📄 PDF Rapor Yükle", type="pdf")

if uploaded_file and api_key:
    if st.button("🚀 Analiz Et", type="primary"):
        client = get_client(api_key)
        pdf_bytes = uploaded_file.getvalue()
        all_results = {}
        method_log = {}
        selected_model = model_sec

        progress = st.progress(0, "Başlatılıyor...")

        progress.progress(5, "📄 PDF metne dönüştürülüyor...")
        pages_text = extract_pages_text(pdf_bytes)
        st.success(f"✅ {len(pages_text)} sayfa okundu")

        progress.progress(8, "🤖 Banka adı ve dönem tespiti...")
        first_text = pages_text.get(0, "") + pages_text.get(1, "")
        banka_adi, donem = extract_metadata(client, first_text, selected_model)
        st.info(f"🏦 **{banka_adi}** | 📅 **{donem}**")

        total_groups = len(EXTRACTION_GROUPS)

        for g_idx, group in enumerate(EXTRACTION_GROUPS):
            pct = 8 + int((g_idx / total_groups) * 87)
            progress.progress(pct, f"🔍 {group['name']}...")

            with st.status(f"📊 {group['name']}", expanded=False) as status:
                page_indices = find_relevant_pages(pages_text, group["keywords"])
                if not page_indices:
                    page_indices = list(range(min(4, len(pages_text))))

                context = build_text_context(pages_text, page_indices)
                st.caption(f"İlgili sayfalar: {[i+1 for i in page_indices]}")

                st.caption("1️⃣ Metin tabanlı extraction (Haiku)...")
                result, success = extract_with_text(client, group, context, banka_adi, donem, selected_model)

                if success:
                    method_log[group["id"]] = f"Metin ({selected_model.split('-')[1].title()})"
                    found = sum(1 for v in result.values() if v and v != "null")
                    status.update(
                        label=f"✅ {group['name']} — {found}/{len(group['fields'])} | Metin",
                        state="complete"
                    )
                else:
                    st.caption("2️⃣ Daha geniş arama...")
                    all_pages = list(range(min(10, len(pages_text))))
                    wider_context = build_text_context(pages_text, all_pages)
                    result2, _ = extract_with_text(client, group, wider_context, banka_adi, donem, selected_model)
                    for k, v in result2.items():
                        if v and v != "null" and (not result.get(k) or result.get(k) == "null"):
                            result[k] = v
                    method_log[group["id"]] = "Metin Geniş (Haiku)"
                    found = sum(1 for v in result.values() if v and v != "null")
                    status.update(
                        label=f"✅ {group['name']} — {found}/{len(group['fields'])} | Geniş",
                        state="complete"
                    )

                all_results.update(result)
                time.sleep(15)

        progress.progress(100, "✅ Tamamlandı!")

        st.divider()
        total_found = sum(1 for v in all_results.values() if v and v != "null")

        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam Alan", total_fields)
        c2.metric("Bulunan", total_found)
        c3.metric("Başarı", f"%{total_found/total_fields*100:.0f}")

        for group in EXTRACTION_GROUPS:
            found = sum(1 for f in group["fields"]
                       if all_results.get(f["id"]) and all_results.get(f["id"]) != "null")
            method = method_log.get(group["id"], "—")
            with st.expander(f"📝 {group['name']} — {found}/{len(group['fields'])} | {method}"):
                cols = st.columns(3)
                for i, field in enumerate(group["fields"]):
                    val = all_results.get(field["id"])
                    with cols[i % 3]:
                        st.metric(
                            f"{field['label'][:30]} ({field['col']})",
                            val if val and val != "null" else "⚠️ —"
                        )

        excel_buf = create_excel(banka_adi, donem, all_results, method_log)
        st.download_button(
            "📥 Excel İndir",
            excel_buf,
            file_name=f"{banka_adi}_{donem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif uploaded_file and not api_key:
    st.warning("⚠️ Sol sidebar'dan Anthropic API key girin.")