# -*- coding: utf-8 -*-
"""
Quality Performance & Defect Monitoring Dashboard
---------------------------------------------------
Author: Built for Industrial Engineering internship/project presentation.

This Streamlit application ingests the TEKNOCEF daily quality-control export
(.xlsm), automatically cleans it, and renders an interactive quality /
defect-monitoring dashboard (KPIs, Pareto analysis, reference analysis,
time trend, critical-defect drill-down, filters, and auto-generated
insights). The interface is available in Turkish and English.

No result is hard-coded: every number on screen is (re)computed live from
whatever data is currently loaded, so the app keeps working if the source
file is refreshed with new rows/columns.
"""

import io
import os
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------------------
# LANGUAGE / TRANSLATIONS
# --------------------------------------------------------------------------------------
if "app_lang" not in st.session_state:
    st.session_state["app_lang"] = "tr"

T = {
    "tr": {
        "page_title": "Kalite Performansı & Hata İzleme Panosu",
        "dashboard_title": "📊 Kalite Performansı & Hata İzleme Panosu",
        "lang_label": "🌐 Dil / Language",
        "data_source_title": "⚙️ Veri Kaynağı",
        "upload_label": "Güncel bir kalite dosyası yükleyin (.xlsx / .xlsm)",
        "active_source": "📁 Aktif kaynak: **{source}**",
        "bundled_suffix": " (dahili dosya)",
        "no_default_error": (
            "Varsayılan veri seti bulunamadı ve herhangi bir dosya da yüklenmedi. "
            "Devam etmek için lütfen bir kalite-kontrol dosyası yükleyin."
        ),
        "clean_fail_error": "Yüklenen dosya okunamadı / temizlenemedi: {e}",
        "no_ref_date_error": (
            "Kaynak dosyada REFERANS / TARİH sütunları bulunamadı. "
            "Lütfen yüklediğiniz sayfanın sütun başlıklarını kontrol edin."
        ),
        "cleaning_log_title": "🧹 Veri temizleme günlüğü",
        "cleaning_log_none": "Herhangi bir temizleme işlemi gerekmedi.",
        "cleaning_log_summary": "Son veri seti: {n:,} geçerli kayıt · {m} hata tipi sütunu.",
        "filters_title": "🔍 Filtreler",
        "date_mode_label": "Tarih filtre modu",
        "date_mode_range": "Tarih aralığı",
        "date_mode_single": "Tek gün",
        "single_day_label": "Gün",
        "date_range_label": "Tarih aralığı",
        "reference_label": "Referans",
        "defect_type_label": "Hata tipi",
        "reset_button": "↺ Filtreleri sıfırla",
        "no_match_warning": "Seçili filtrelerle eşleşen kayıt yok. Lütfen filtreleri genişletin.",
        "header_caption": (
            "Veri aralığı: **{start} → {end}**  ·  {n} referans  ·  "
            "oluşturulma: {ts}"
        ),
        "kpi_total_inspections": "Toplam Kontrol",
        "kpi_total_defects": "Toplam Hata",
        "kpi_overall_nok": "Genel NOK %",
        "kpi_references": "Referans Sayısı",
        "kpi_defect_types": "Hata Tipi Sayısı",
        "kpi_caption": (
            "ℹ️ *Toplam Kontrol* = kontrol edilen toplam adet (KONTROL EDİLEN). "
            "*Toplam Hata* = tüm hata tiplerinin toplam adedi (Pareto tabanı). "
            "*Genel NOK %* = NOK adedi ÷ kontrol edilen adet."
        ),
        "section_pareto": "📈 Pareto Analizi — Hata Tipleri",
        "pareto_bar_name": "Hata Adedi",
        "pareto_cum_name": "Kümülatif %",
        "pareto_ref_line": "%80 referans çizgisi",
        "pareto_yaxis": "Hata Adedi",
        "pareto_yaxis2": "Kümülatif %",
        "pareto_caption": (
            "🔴 **{n} hata tipi** ({sample}{more}) seçili filtrede toplam hataların "
            "≈%80'ini oluşturuyor (dinamik olarak belirlendi — sabit kodlanmadı)."
        ),
        "section_reference": "🧩 Referans Analizi",
        "ref_vs_nok_title": "**Referans vs NOK %**",
        "heatmap_title": "**Referans × Hata Tipi Isı Haritası**",
        "top10_title": "**En Sık 10 Referans–Hata Kombinasyonu**",
        "col_reference": "Referans",
        "col_defect_type": "Hata Tipi",
        "col_quantity": "Adet",
        "heatmap_color_label": "Adet",
        "nok_pct_label": "NOK %",
        "section_time": "🗓️ Zaman Analizi — Günlük Hata Trendi",
        "time_xaxis": "Tarih",
        "time_yaxis": "Toplam Hata",
        "time_rolling_avg": "Hareketli ortalama",
        "col_cum_pct": "Kümülatif %",
        "col_main_refs": "Ana Etkilenen Referans(lar)",
        "section_detail": "📋 Filtrelenmiş Detay Veri",
        "download_csv": "⬇️ CSV İndir",
        "download_excel": "⬇️ Excel İndir",
        "section_insights": "💡 Öne Çıkan Bulgular",
        "insight_top_defect": (
            "**{defect}**, seçili veride en yüksek adede sahip tek hata tipi: "
            "**{qty:,.0f}** adet (tüm hataların %{pct:.1f}'i)."
        ),
        "insight_critical_share": (
            "**{n} / {total}** farklı hata tipi (hata tiplerinin %{pct:.0f}'i) "
            "toplam hata hacminin ≈%80'ini oluşturuyor — klasik Pareto (80-20) deseniyle uyumlu."
        ),
        "insight_worst_ref": (
            "**{ref}** referansı seçili veride en yüksek NOK oranına sahip: "
            "**%{pct:.2f}** ({nok:.0f} NOK / {insp:.0f} kontrol)."
        ),
        "insight_best_ref": (
            "**{ref}** referansı, kontrol edilen referanslar arasında en düşük NOK oranına sahip: **%{pct:.2f}**."
        ),
        "insight_top_combo": (
            "En sık görülen referans–hata kombinasyonu **{ref} / {defect}**, "
            "**{qty:,.0f}** adet ile."
        ),
        "insight_worst_day": (
            "Seçili aralıktaki en yüksek hata görülen tek gün **{date}**, "
            "**{qty:,.0f}** kayıtlı hata ile."
        ),
        "insight_trend": (
            "Seçili tarih aralığının ilk ve ikinci yarısı arasında günlük ortalama hata "
            "**%{pct:.1f}** oranında {direction}."
        ),
        "trend_increased": "arttı",
        "trend_decreased": "azaldı",
        "insight_overall_nok": (
            "Seçili veride genel NOK oranı **%{pct:.2f}** "
            "({nok:,.0f} NOK adedi / {insp:,.0f} kontrol edilen adet)."
        ),
        "insight_none": "Öne çıkan bir bulgu üretmek için seçili filtrede yeterli veri yok.",
        "insight_caption": (
            "Bulgular yalnızca o an filtrelenmiş veriden otomatik olarak üretilir — "
            "verinin *ne* gösterdiğini anlatır, *neden* olduğunu değil; herhangi bir kök neden çıkarımı yapılmaz."
        ),
        "footer_caption": (
            "Kalite Performansı & Hata İzleme Panosu · Streamlit, Pandas & Plotly ile geliştirilmiştir · "
            "Tüm rakamlar yüklenen dosyadan canlı olarak hesaplanır (sabit kodlanmış sonuç yoktur)."
        ),
        # cleaning-log message templates (used inside clean_data)
        "clean_removed_unnamed": "{n} adet boş/isimsiz şablon sütunu kaldırıldı.",
        "clean_removed_blank_rows": "Referansı veya tarihi olmayan {n} adet boş/şablon satır kaldırıldı.",
        "clean_normalized_refs": (
            "Referans kodu biçimlendirmesi normalize edildi (boşluk/büyük-küçük harf/tire biçimi): "
            "{raw} ham varyant → {norm} standart referans kodu."
        ),
        "clean_dropped_bad_dates": "Ayrıştırılamayan tarihe sahip {n} satır kaldırıldı.",
        "clean_dropped_defect_cols": (
            "Kullanılabilir hata adedi içermeyen {n} adet sayısal olmayan / boş sütun kaldırıldı."
        ),
        "clean_negative_ok": "{n} adet negatif 'OK' değeri 0 olarak düzeltildi (veri girişi hatası).",
        "clean_filled_insp": (
            "Eksik olan {n} adet 'KONTROL EDİLEN' (kontrol edilen adet) değeri OK + NOK kullanılarak dolduruldu."
        ),
        # --- new: time granularity & drill-down trend labels ---
        "granularity_label": "Zaman Aralığı",
        "granularity_daily": "Günlük",
        "granularity_weekly": "Haftalık",
        "granularity_monthly": "Aylık",
        "defect_trend_subheader": "🔎 Hata Tipi Bazında Trend",
        "defect_trend_select_label": "Hata tipi seçin",
        "reference_trend_subheader": "🔎 Referans Bazında Hata Trendi",
        "reference_trend_select_label": "Referans seçin",
        "trend_yaxis_generic": "Hata Adedi",
        # --- new: detail table column labels ---
        "col_date": "Tarih",
        "col_inspected": "Kontrol Edilen",
        "col_ok": "OK",
        "col_nok": "NOK",
        "col_nok_pct_short": "NOK %",
        "col_total_defects": "Toplam Hata",
        "detail_pct_note": (
            "ℹ️ NOK % sütunu ekranda yüzde olarak biçimlendirilmiştir; indirilen dosyalarda "
            "sayısal değer olarak yer alır (örn. 2.34 = %2.34)."
        ),
        # --- new: SPC p-chart ---
        "section_spc": "📐 İstatistiksel Proses Kontrol (p-Kontrol Grafiği)",
        "spc_intro": (
            "Her dönem için NOK oranı (p̂), o dönemin kontrol edilen adedine göre hesaplanan "
            "±3-sigma kontrol limitleriyle birlikte gösterilir. Limitler veriden dinamik olarak "
            "hesaplanır, sabit kodlanmamıştır."
        ),
        "spc_yaxis": "NOK %",
        "spc_series_actual": "Gerçekleşen NOK %",
        "spc_series_cl": "Merkez Çizgi (CL)",
        "spc_series_ucl": "Üst Kontrol Limiti (UCL)",
        "spc_series_lcl": "Alt Kontrol Limiti (LCL)",
        "spc_series_ooc": "Kontrol Dışı Nokta",
        "spc_caption": (
            "🔎 **{n} / {total}** dönem istatistiksel kontrol limitlerinin dışında kaldı. "
            "Bu, sürecin o dönemlerde 'özel neden' (special cause) varyasyonu gösterebileceğine "
            "işaret eder; kesin bir kök neden çıkarımı yapılmamıştır."
        ),
        "spc_insufficient": "P-kontrol grafiği için yeterli dönem verisi yok (en az 3 dönem gerekir).",
        # --- new: simple trend projection ---
        "section_forecast": "📉 Basit Trend Projeksiyonu",
        "forecast_intro": (
            "Seçilen son N güne dayanan basit doğrusal bir eğilim çizgisi, önümüzdeki 7 gün için "
            "ileriye doğru uzatılır. Bu; mevsimsellik, kök neden veya proses değişikliklerini "
            "dikkate almayan kaba bir projeksiyondur."
        ),
        "forecast_window_label": "Trend için kullanılacak son gün sayısı",
        "forecast_xaxis": "Tarih",
        "forecast_yaxis": "NOK %",
        "forecast_series_actual": "Gerçekleşen NOK %",
        "forecast_series_trend": "Doğrusal Trend (eğitim penceresi)",
        "forecast_series_projection": "Projeksiyon (önümüzdeki 7 gün)",
        "forecast_result": (
            "📈 Son **{n}** güne dayanan doğrusal eğilime göre, önümüzdeki 7 gün için beklenen "
            "ortalama NOK %: **{val:.2f}%** (eğilim yönü: {direction})."
        ),
        "forecast_trend_up": "artış",
        "forecast_trend_down": "azalış",
        "forecast_trend_flat": "yatay",
        "forecast_insufficient": "Trend projeksiyonu için yeterli günlük veri yok (en az 4 gün gerekir).",
    },
    "en": {
        "page_title": "Quality Performance & Defect Monitoring Dashboard",
        "dashboard_title": "📊 Quality Performance & Defect Monitoring Dashboard",
        "lang_label": "🌐 Dil / Language",
        "data_source_title": "⚙️ Data Source",
        "upload_label": "Upload an updated quality export (.xlsx / .xlsm)",
        "active_source": "📁 Active source: **{source}**",
        "bundled_suffix": " (bundled)",
        "no_default_error": (
            "No default dataset found and nothing uploaded. "
            "Please upload a quality-control export file to continue."
        ),
        "clean_fail_error": "Failed to read/clean the uploaded file: {e}",
        "no_ref_date_error": (
            "Could not locate REFERANS / TARİH columns in the source file. "
            "Please check the column headers of the uploaded sheet."
        ),
        "cleaning_log_title": "🧹 Data-cleaning log",
        "cleaning_log_none": "No cleaning actions were required.",
        "cleaning_log_summary": "Final dataset: {n:,} valid records · {m} defect-type columns.",
        "filters_title": "🔍 Filters",
        "date_mode_label": "Date filter mode",
        "date_mode_range": "Date range",
        "date_mode_single": "Single day",
        "single_day_label": "Day",
        "date_range_label": "Date range",
        "reference_label": "Reference",
        "defect_type_label": "Defect type",
        "reset_button": "↺ Reset filters",
        "no_match_warning": "No records match the current filter selection. Try widening the filters.",
        "header_caption": (
            "Data window: **{start} → {end}**  ·  {n} reference(s)  ·  generated {ts}"
        ),
        "kpi_total_inspections": "Total Inspections",
        "kpi_total_defects": "Total Defects",
        "kpi_overall_nok": "Overall NOK %",
        "kpi_references": "References",
        "kpi_defect_types": "Defect Types",
        "kpi_caption": (
            "ℹ️ *Total Inspections* = sum of units inspected (KONTROL EDİLEN). "
            "*Total Defects* = sum of individual defect-type occurrences (Pareto basis). "
            "*Overall NOK %* = NOK units ÷ inspected units."
        ),
        "section_pareto": "📈 Pareto Analysis — Defect Types",
        "pareto_bar_name": "Defect Quantity",
        "pareto_cum_name": "Cumulative %",
        "pareto_ref_line": "80% reference line",
        "pareto_yaxis": "Defect Quantity",
        "pareto_yaxis2": "Cumulative %",
        "pareto_caption": (
            "🔴 **{n} defect type(s)** ({sample}{more}) account for ≈80% of all recorded "
            "defects in the current filter selection (dynamically identified — not hard-coded)."
        ),
        "section_reference": "🧩 Reference Analysis",
        "ref_vs_nok_title": "**Reference vs NOK %**",
        "heatmap_title": "**Reference × Defect Type Heatmap**",
        "top10_title": "**Top 10 Reference–Defect Combinations**",
        "col_reference": "Reference",
        "col_defect_type": "Defect Type",
        "col_quantity": "Quantity",
        "heatmap_color_label": "Count",
        "nok_pct_label": "NOK %",
        "section_time": "🗓️ Time Analysis — Daily Defect Trend",
        "time_xaxis": "Date",
        "time_yaxis": "Total Defects",
        "time_rolling_avg": "Rolling avg",
        "col_cum_pct": "Cumulative %",
        "col_main_refs": "Main Affected Reference(s)",
        "section_detail": "📋 Detailed Filtered Data",
        "download_csv": "⬇️ Download CSV",
        "download_excel": "⬇️ Download Excel",
        "section_insights": "💡 Key Insights",
        "insight_top_defect": (
            "**{defect}** is the single largest defect type in the current selection, with "
            "**{qty:,.0f}** occurrences ({pct:.1f}% of all defects)."
        ),
        "insight_critical_share": (
            "**{n} of {total}** distinct defect types ({pct:.0f}% of defect types) account "
            "for ≈80% of total defect volume — consistent with a Pareto/80-20 pattern."
        ),
        "insight_worst_ref": (
            "Reference **{ref}** shows the highest NOK rate in the selection at **{pct:.2f}%** "
            "({nok:.0f} NOK / {insp:.0f} inspected)."
        ),
        "insight_best_ref": (
            "Reference **{ref}** shows the lowest NOK rate among inspected references at **{pct:.2f}%**."
        ),
        "insight_top_combo": (
            "The most frequent reference–defect combination is **{ref} / {defect}**, with "
            "**{qty:,.0f}** occurrences."
        ),
        "insight_worst_day": (
            "The single highest-defect day in the selection is **{date}** with **{qty:,.0f}** "
            "recorded defects."
        ),
        "insight_trend": (
            "Average daily defects {direction} by **{pct:.1f}%** between the first and second "
            "half of the selected date range."
        ),
        "trend_increased": "increased",
        "trend_decreased": "decreased",
        "insight_overall_nok": (
            "Overall NOK rate across the current selection is **{pct:.2f}%** ({nok:,.0f} NOK "
            "units out of {insp:,.0f} inspected)."
        ),
        "insight_none": "Not enough data in the current filter selection to generate insights.",
        "insight_caption": (
            "Insights are generated automatically from the currently filtered dataset only — "
            "they describe *what* the data shows, not *why* it happened, and no root cause is inferred."
        ),
        "footer_caption": (
            "Quality Performance & Defect Monitoring Dashboard · Built with Streamlit, Pandas & Plotly · "
            "All figures are computed live from the loaded file (no hard-coded results)."
        ),
        "clean_removed_unnamed": "Removed {n} empty/unnamed template column(s).",
        "clean_removed_blank_rows": "Removed {n} blank / template row(s) with no reference or date.",
        "clean_normalized_refs": (
            "Normalized reference-code formatting (spacing/case/dash style): "
            "{raw} raw variants → {norm} standardized reference codes."
        ),
        "clean_dropped_bad_dates": "Dropped {n} row(s) with an unparseable date.",
        "clean_dropped_defect_cols": (
            "Dropped {n} non-numeric / empty candidate column(s) that did not contain usable defect counts."
        ),
        "clean_negative_ok": "Corrected {n} negative 'OK' value(s) to 0 (data-entry error).",
        "clean_filled_insp": (
            "Filled {n} missing 'KONTROL EDİLEN' (inspected qty) value(s) using OK + NOK."
        ),
        # --- new: time granularity & drill-down trend labels ---
        "granularity_label": "Time Granularity",
        "granularity_daily": "Daily",
        "granularity_weekly": "Weekly",
        "granularity_monthly": "Monthly",
        "defect_trend_subheader": "🔎 Defect Type Trend",
        "defect_trend_select_label": "Select defect type",
        "reference_trend_subheader": "🔎 Reference Defect Trend",
        "reference_trend_select_label": "Select reference",
        "trend_yaxis_generic": "Defect Count",
        # --- new: detail table column labels ---
        "col_date": "Date",
        "col_inspected": "Inspected",
        "col_ok": "OK",
        "col_nok": "NOK",
        "col_nok_pct_short": "NOK %",
        "col_total_defects": "Total Defects",
        "detail_pct_note": (
            "ℹ️ The NOK % column is formatted as a percentage on screen; in downloaded files "
            "it is a plain numeric value (e.g. 2.34 = 2.34%)."
        ),
        # --- new: SPC p-chart ---
        "section_spc": "📐 Statistical Process Control (p-Chart)",
        "spc_intro": (
            "The NOK proportion (p̂) for each period is shown alongside ±3-sigma control limits "
            "computed from that period's inspected quantity. Limits are calculated dynamically "
            "from the data, not hard-coded."
        ),
        "spc_yaxis": "NOK %",
        "spc_series_actual": "Actual NOK %",
        "spc_series_cl": "Center Line (CL)",
        "spc_series_ucl": "Upper Control Limit (UCL)",
        "spc_series_lcl": "Lower Control Limit (LCL)",
        "spc_series_ooc": "Out-of-Control Point",
        "spc_caption": (
            "🔎 **{n} / {total}** period(s) fell outside the statistical control limits. "
            "This suggests the process may show 'special cause' variation in those periods; "
            "no definitive root cause is inferred."
        ),
        "spc_insufficient": "Not enough period data for a p-chart (at least 3 periods required).",
        # --- new: simple trend projection ---
        "section_forecast": "📉 Simple Trend Projection",
        "forecast_intro": (
            "A simple linear trend fitted on the last N days is extended forward 7 days. "
            "This is a rough projection that does not account for seasonality, root causes, "
            "or process changes."
        ),
        "forecast_window_label": "Number of recent days to use for the trend",
        "forecast_xaxis": "Date",
        "forecast_yaxis": "NOK %",
        "forecast_series_actual": "Actual NOK %",
        "forecast_series_trend": "Linear Trend (training window)",
        "forecast_series_projection": "Projection (next 7 days)",
        "forecast_result": (
            "📈 Based on a linear trend over the last **{n}** days, the expected average NOK % "
            "for the next 7 days is: **{val:.2f}%** (trend direction: {direction})."
        ),
        "forecast_trend_up": "increasing",
        "forecast_trend_down": "decreasing",
        "forecast_trend_flat": "flat",
        "forecast_insufficient": "Not enough daily data for a trend projection (at least 4 days required).",
    },
}


def tr(key, **kwargs):
    """Fetch a translated string for the currently selected language."""
    txt = T[st.session_state["app_lang"]][key]
    return txt.format(**kwargs) if kwargs else txt


# --------------------------------------------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------------------------------------------
st.set_page_config(
    page_title=tr("page_title"),
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "source_data.xlsm")

# --------------------------------------------------------------------------------------
# STYLE
# --------------------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main { background-color: #f6f8fb; }
    .block-container { padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1400px; }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e6e9ef;
        border-radius: 12px;
        padding: 14px 18px 10px 18px;
        box-shadow: 0 1px 4px rgba(16,24,40,0.05);
    }
    div[data-testid="stMetric"] * { color: #10233f !important; }
    div[data-testid="stMetricLabel"] { font-weight: 600; }
    div[data-testid="stMetricValue"] { font-weight: 700; }
    h1, h2, h3, h4, p, span, label, li { color: #10233f; }
    .stCaption, [data-testid="stCaptionContainer"] { color: #445 !important; }
    .section-header {
        background: linear-gradient(90deg, #10233f 0%, #1c3f6e 100%);
        color: white;
        padding: 10px 18px;
        border-radius: 10px;
        margin-top: 1.2rem;
        margin-bottom: 0.8rem;
        font-size: 1.05rem;
        font-weight: 600;
    }
    .insight-box {
        background: #ffffff;
        border-left: 4px solid #1c3f6e;
        border-radius: 8px;
        padding: 10px 16px;
        margin-bottom: 8px;
        box-shadow: 0 1px 3px rgba(16,24,40,0.05);
        color: #10233f !important;
    }
    .insight-box * { color: #10233f !important; }
    .badge-critical {
        background:#fdecea; color:#b3261e; padding:2px 10px; border-radius:12px;
        font-size:0.78rem; font-weight:600;
    }
    .caption-muted { color:#6b7688; font-size:0.85rem; }
    /* Force readable dark text inside dataframes / tables regardless of theme */
    div[data-testid="stDataFrame"] * { color: #10233f !important; }
    div[data-testid="stDataFrame"] { background: #ffffff; }
    /* Sidebar readability */
    section[data-testid="stSidebar"] * { color: #10233f !important; }
    section[data-testid="stSidebar"] { background: #f0f3f8; }
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)


def section_header(title: str):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------------------
# SIDEBAR — LANGUAGE SELECTOR (always first, so the rest of the UI renders translated)
# --------------------------------------------------------------------------------------
lang_options = {"Türkçe": "tr", "English": "en"}
lang_display = "Türkçe" if st.session_state["app_lang"] == "tr" else "English"
chosen = st.sidebar.selectbox(
    tr("lang_label"),
    options=list(lang_options.keys()),
    index=list(lang_options.keys()).index(lang_display),
)
if lang_options[chosen] != st.session_state["app_lang"]:
    st.session_state["app_lang"] = lang_options[chosen]
    st.rerun()

# --------------------------------------------------------------------------------------
# NAME NORMALIZATION HELPERS (Turkish-safe)
# --------------------------------------------------------------------------------------
TR_MAP = str.maketrans("İIıŞşĞğÜüÖöÇç", "iiisSgGuuOoCc")


def normalize_key(s: str) -> str:
    """Lowercase / trim / Turkish-char-safe key, used only to *match* known
    column names -- never used to alter displayed labels."""
    if s is None:
        return ""
    s = str(s).translate(TR_MAP)
    s = unicodedata.normalize("NFKD", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def normalize_reference(val) -> str:
    """Normalize obvious formatting inconsistencies in reference codes
    (int vs str, stray spaces around dashes, mixed case) without guessing
    at genuine typos in the underlying code itself."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    s = s.upper()
    s = re.sub(r"\s*-\s*", "-", s)   # "456 - B"  -> "456-B"
    s = re.sub(r"\s+", " ", s)       # collapse repeated spaces
    return s


# Canonical structural (meta) columns we expect to find in the export.
# Matching is done on a normalized key so trailing spaces / case differences
# in the source file don't break detection.
EXPECTED_META = {
    "referans": "REFERANS",
    "tarih": "TARİH",
    "hafta": "HAFTA",
    "kontrol edilen": "KONTROL_EDILEN",
    "tashih yapilacak": "TASHIH_YAPILACAK",
    "ok": "OK",
    "nok": "NOK",
    "ftt": "FTT",
}


# --------------------------------------------------------------------------------------
# DATA LOADING & CLEANING
# --------------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_raw(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=0, engine="openpyxl")


@st.cache_data(show_spinner=False)
def clean_data(file_bytes: bytes, lang: str):
    """Full cleaning pipeline. Returns (wide_df, long_df, defect_cols, meta_map, quality_log)."""
    L = T[lang]
    raw = load_raw(file_bytes)
    quality_log = []

    # ---- 1. Strip whitespace from column headers, drop fully-unnamed columns
    raw.columns = [str(c).strip() for c in raw.columns]
    unnamed = [c for c in raw.columns if c.lower().startswith("unnamed")]
    if unnamed:
        raw = raw.drop(columns=unnamed)
        quality_log.append(L["clean_removed_unnamed"].format(n=len(unnamed)))

    # ---- 2. Identify meta columns dynamically by normalized name match
    col_lookup = {normalize_key(c): c for c in raw.columns}
    meta_cols_found = {}
    for key, std_name in EXPECTED_META.items():
        if key in col_lookup:
            meta_cols_found[std_name] = col_lookup[key]

    ref_col = meta_cols_found.get("REFERANS")
    date_col = meta_cols_found.get("TARİH")

    if ref_col is None or date_col is None:
        raise ValueError(L["no_ref_date_error"])

    # ---- 3. Remove blank / template rows (no reference AND no date)
    before = len(raw)
    df = raw[raw[ref_col].notna() | raw[date_col].notna()].copy()
    df = df[df[ref_col].notna() & df[date_col].notna()].copy()
    removed = before - len(df)
    if removed:
        quality_log.append(L["clean_removed_blank_rows"].format(n=removed))

    # ---- 4. Normalize reference codes
    raw_unique = df[ref_col].astype(str).str.strip().nunique()
    df[ref_col] = df[ref_col].apply(normalize_reference)
    norm_unique = df[ref_col].nunique()
    if norm_unique < raw_unique:
        quality_log.append(L["clean_normalized_refs"].format(raw=raw_unique, norm=norm_unique))
    df = df.rename(columns={ref_col: "REFERANS"})

    # ---- 5. Parse date column
    df = df.rename(columns={date_col: "TARIH"})
    df["TARIH"] = pd.to_datetime(df["TARIH"], errors="coerce")
    bad_dates = df["TARIH"].isna().sum()
    if bad_dates:
        quality_log.append(L["clean_dropped_bad_dates"].format(n=bad_dates))
        df = df[df["TARIH"].notna()].copy()

    # ---- 6. Identify defect-type columns = everything that is NOT a meta column
    meta_display_cols = ["REFERANS", "TARIH"] + [
        v for k, v in meta_cols_found.items() if k not in ("REFERANS", "TARİH")
    ]
    all_cols = list(df.columns)
    defect_cols_raw = [c for c in all_cols if c not in meta_display_cols]

    # Keep only numeric-ish defect columns (defensive: coerce, drop columns
    # that are entirely non-numeric / entirely empty)
    defect_cols = []
    for c in defect_cols_raw:
        coerced = pd.to_numeric(df[c], errors="coerce")
        if coerced.notna().sum() > 0:
            df[c] = coerced
            defect_cols.append(c)
    dropped_defect_cols = [c for c in defect_cols_raw if c not in defect_cols]
    if dropped_defect_cols:
        quality_log.append(L["clean_dropped_defect_cols"].format(n=len(dropped_defect_cols)))

    # Clean defect column labels for display (strip stray whitespace only)
    rename_defects = {c: re.sub(r"\s+", " ", c).strip() for c in defect_cols}
    df = df.rename(columns=rename_defects)
    defect_cols = [rename_defects[c] for c in defect_cols]

    # ---- 7. Missing/invalid value handling
    df[defect_cols] = df[defect_cols].fillna(0)
    df[defect_cols] = df[defect_cols].clip(lower=0)  # a defect count cannot be negative

    ok_col = meta_cols_found.get("OK")
    nok_col = meta_cols_found.get("NOK")
    insp_col = meta_cols_found.get("KONTROL_EDILEN")

    for c in [ok_col, nok_col, insp_col]:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    neg_ok = 0
    if ok_col and ok_col in df.columns:
        neg_ok = (df[ok_col] < 0).sum()
        df[ok_col] = df[ok_col].clip(lower=0)
    if neg_ok:
        quality_log.append(L["clean_negative_ok"].format(n=neg_ok))

    if nok_col and nok_col in df.columns:
        df[nok_col] = df[nok_col].fillna(0).clip(lower=0)

    # Recompute / fill missing "KONTROL EDİLEN" (inspected qty) from OK+NOK where absent
    if insp_col and ok_col and nok_col:
        missing_insp = df[insp_col].isna().sum()
        computed = df[ok_col].fillna(0) + df[nok_col].fillna(0)
        df[insp_col] = df[insp_col].fillna(computed)
        # also fix rows where recorded value is 0/blank but OK+NOK > 0
        mismatch = (df[insp_col] < computed)
        df.loc[mismatch, insp_col] = computed[mismatch]
        if missing_insp:
            quality_log.append(L["clean_filled_insp"].format(n=missing_insp))
        df[insp_col] = df[insp_col].fillna(0)
    elif insp_col:
        df[insp_col] = pd.to_numeric(df[insp_col], errors="coerce").fillna(0)

    # ---- 8. Derived fields
    # (per-row total defect occurrences is intentionally not retained as a separate
    # column in the UI — NOK and the individual defect-type columns already cover it)
    if insp_col:
        with np.errstate(divide="ignore", invalid="ignore"):
            df["NOK_PCT"] = np.where(
                df[insp_col] > 0, (df[nok_col] / df[insp_col]) * 100, np.nan
            )
    df["WEEK"] = df["TARIH"].dt.isocalendar().week
    df["MONTH"] = df["TARIH"].dt.to_period("M").astype(str)
    df["DATE_ONLY"] = df["TARIH"].dt.date
    weekday_map = {
        0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun",
    }
    df["WEEKDAY"] = df["TARIH"].dt.dayofweek.map(weekday_map)

    # ---- 9. Detect an optional shift column (kept internally if present, no UI filter)
    shift_col = None
    for key in ["vardiya", "shift", "vardiya adi", "shift name"]:
        if key in col_lookup:
            shift_col = col_lookup[key]
            break
    if shift_col:
        df = df.rename(columns={shift_col: "SHIFT"})
        df["SHIFT"] = df["SHIFT"].astype(str).str.strip()

    # ---- 10. Build long (tidy) format for defect-level analysis
    id_vars = ["REFERANS", "TARIH", "DATE_ONLY", "WEEK", "MONTH", "WEEKDAY"]
    if insp_col:
        id_vars.append(insp_col)
    if nok_col:
        id_vars.append(nok_col)
    if "NOK_PCT" in df.columns:
        id_vars.append("NOK_PCT")
    if shift_col:
        id_vars.append("SHIFT")

    long_df = df.melt(
        id_vars=id_vars, value_vars=defect_cols, var_name="DEFECT_TYPE", value_name="DEFECT_COUNT"
    )
    long_df = long_df[long_df["DEFECT_COUNT"] > 0].copy()

    meta_map = {
        "insp_col": insp_col,
        "ok_col": ok_col,
        "nok_col": nok_col,
        "ftt_col": meta_cols_found.get("FTT"),
        "shift_col": "SHIFT" if shift_col else None,
    }

    return df, long_df, defect_cols, meta_map, quality_log


# --------------------------------------------------------------------------------------
# SIDEBAR — DATA SOURCE
# --------------------------------------------------------------------------------------
st.sidebar.title(tr("data_source_title"))
uploaded = st.sidebar.file_uploader(tr("upload_label"), type=["xlsx", "xlsm", "xls"])

if uploaded is not None:
    file_bytes = uploaded.read()
    source_label = uploaded.name
else:
    if not os.path.exists(DEFAULT_DATA_PATH):
        st.error(tr("no_default_error"))
        st.stop()
    with open(DEFAULT_DATA_PATH, "rb") as f:
        file_bytes = f.read()
    source_label = "TEKNOCEF_GÜNLÜK_GİDİŞAT-TEMMUZ_2026.xlsm" + tr("bundled_suffix")

st.sidebar.caption(tr("active_source", source=source_label))

# If a new/different file has been loaded, clear any stale filter values that
# belonged to the previous file (e.g. a date range or reference list that no
# longer matches the new dataset) so widgets never receive out-of-range values.
_source_fingerprint = f"{source_label}:{len(file_bytes)}"
if st.session_state.get("_active_source") != _source_fingerprint:
    for k in ["flt_date_mode", "flt_date", "flt_single_date", "flt_ref", "flt_defect"]:
        st.session_state.pop(k, None)
    st.session_state["_active_source"] = _source_fingerprint

try:
    df, long_df, defect_cols, meta_map, quality_log = clean_data(file_bytes, st.session_state["app_lang"])
except Exception as e:
    st.error(tr("clean_fail_error", e=e))
    st.stop()

insp_col = meta_map["insp_col"]
nok_col = meta_map["nok_col"]
ok_col = meta_map["ok_col"]
shift_col = meta_map["shift_col"]

with st.sidebar.expander(tr("cleaning_log_title"), expanded=False):
    if quality_log:
        for line in quality_log:
            st.write("• " + line)
    else:
        st.write(tr("cleaning_log_none"))
    st.caption(tr("cleaning_log_summary", n=len(df), m=len(defect_cols)))

# --------------------------------------------------------------------------------------
# SIDEBAR — FILTERS
# --------------------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.title(tr("filters_title"))

DEFAULT_MIN_DATE = df["TARIH"].min().date()
DEFAULT_MAX_DATE = df["TARIH"].max().date()
ALL_REFS = sorted(df["REFERANS"].dropna().unique().tolist())
ALL_DEFECTS = sorted(defect_cols)

# session-state defaults
defaults = {
    "flt_date_mode": tr("date_mode_range"),
    "flt_date": (DEFAULT_MIN_DATE, DEFAULT_MAX_DATE),
    "flt_single_date": DEFAULT_MAX_DATE,
    "flt_ref": [],
    "flt_defect": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def reset_filters():
    st.session_state["flt_date_mode"] = tr("date_mode_range")
    st.session_state["flt_date"] = (DEFAULT_MIN_DATE, DEFAULT_MAX_DATE)
    st.session_state["flt_single_date"] = DEFAULT_MAX_DATE
    st.session_state["flt_ref"] = []
    st.session_state["flt_defect"] = []


date_mode_options = [tr("date_mode_range"), tr("date_mode_single")]
# guard against a stale value (e.g. leftover from the other language) not being in options
if st.session_state.get("flt_date_mode") not in date_mode_options:
    st.session_state["flt_date_mode"] = date_mode_options[0]

date_mode = st.sidebar.radio(
    tr("date_mode_label"), date_mode_options, key="flt_date_mode", horizontal=True
)

if date_mode == tr("date_mode_single"):
    single_date = st.sidebar.date_input(
        tr("single_day_label"),
        min_value=DEFAULT_MIN_DATE,
        max_value=DEFAULT_MAX_DATE,
        key="flt_single_date",
    )
    d_start = d_end = single_date
else:
    date_range = st.sidebar.date_input(
        tr("date_range_label"),
        min_value=DEFAULT_MIN_DATE,
        max_value=DEFAULT_MAX_DATE,
        key="flt_date",
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        d_start, d_end = date_range
    elif isinstance(date_range, tuple) and len(date_range) == 1:
        # user has only picked the first day of the range so far
        d_start = d_end = date_range[0]
    else:
        d_start = d_end = date_range

sel_refs = st.sidebar.multiselect(tr("reference_label"), ALL_REFS, key="flt_ref")
sel_defects = st.sidebar.multiselect(tr("defect_type_label"), ALL_DEFECTS, key="flt_defect")

st.sidebar.button(tr("reset_button"), on_click=reset_filters)

# ---- Apply filters -----------------------------------------------------------------
mask_wide = (df["TARIH"].dt.date >= d_start) & (df["TARIH"].dt.date <= d_end)
if sel_refs:
    mask_wide &= df["REFERANS"].isin(sel_refs)
wide_f = df[mask_wide].copy()

mask_long = (long_df["TARIH"].dt.date >= d_start) & (long_df["TARIH"].dt.date <= d_end)
if sel_refs:
    mask_long &= long_df["REFERANS"].isin(sel_refs)
if sel_defects:
    mask_long &= long_df["DEFECT_TYPE"].isin(sel_defects)
long_f = long_df[mask_long].copy()

# If a defect-type filter is active, also restrict the wide table's records to
# rows that actually contain at least one of the selected defect types.
if sel_defects:
    matching_rows = long_df[
        (long_df["DEFECT_TYPE"].isin(sel_defects))
        & (long_df["TARIH"].dt.date >= d_start)
        & (long_df["TARIH"].dt.date <= d_end)
    ]
    if sel_refs:
        matching_rows = matching_rows[matching_rows["REFERANS"].isin(sel_refs)]
    key_cols = ["REFERANS", "TARIH"]
    wide_f = wide_f.merge(matching_rows[key_cols].drop_duplicates(), on=key_cols, how="inner")

if len(wide_f) == 0:
    st.warning(tr("no_match_warning"))
    st.stop()

# --------------------------------------------------------------------------------------
# HEADER
# --------------------------------------------------------------------------------------
st.title(tr("dashboard_title"))
st.caption(
    tr(
        "header_caption",
        start=d_start, end=d_end,
        n=wide_f["REFERANS"].nunique(),
        ts=f"{datetime.now():%Y-%m-%d %H:%M}",
    )
)

# --------------------------------------------------------------------------------------
# 1. KPI CARDS
# --------------------------------------------------------------------------------------
total_inspections = wide_f[insp_col].sum() if insp_col else np.nan
total_nok_units = wide_f[nok_col].sum() if nok_col else np.nan
overall_nok_pct = (total_nok_units / total_inspections * 100) if total_inspections else np.nan
total_defect_occurrences = long_f["DEFECT_COUNT"].sum()
n_references = wide_f["REFERANS"].nunique()
n_defect_types = long_f.loc[long_f["DEFECT_COUNT"] > 0, "DEFECT_TYPE"].nunique()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric(tr("kpi_total_inspections"), f"{total_inspections:,.0f}")
c2.metric(tr("kpi_total_defects"), f"{total_defect_occurrences:,.0f}")
c3.metric(tr("kpi_overall_nok"), f"{overall_nok_pct:,.2f}%")
c4.metric(tr("kpi_references"), f"{n_references:,}")
c5.metric(tr("kpi_defect_types"), f"{n_defect_types:,}")

st.caption(tr("kpi_caption"))

# --------------------------------------------------------------------------------------
# 2. PARETO ANALYSIS
# --------------------------------------------------------------------------------------
section_header(tr("section_pareto"))

pareto = (
    long_f.groupby("DEFECT_TYPE", as_index=False)["DEFECT_COUNT"]
    .sum()
    .sort_values("DEFECT_COUNT", ascending=False)
    .reset_index(drop=True)
)
pareto["CUM_PCT"] = pareto["DEFECT_COUNT"].cumsum() / pareto["DEFECT_COUNT"].sum() * 100
# dynamically identify critical defects (first set reaching ~80%)
critical_mask = pareto["CUM_PCT"] <= 80
if critical_mask.sum() < len(pareto):
    critical_mask.iloc[critical_mask.sum()] = True  # include first defect that crosses 80%
critical_defects = pareto[critical_mask]["DEFECT_TYPE"].tolist()

fig_pareto = go.Figure()
fig_pareto.add_bar(
    x=pareto["DEFECT_TYPE"],
    y=pareto["DEFECT_COUNT"],
    name=tr("pareto_bar_name"),
    marker_color=[
        "#c0392b" if d in critical_defects else "#7f9cc4" for d in pareto["DEFECT_TYPE"]
    ],
    text=pareto["DEFECT_COUNT"],
    textposition="outside",
)
fig_pareto.add_trace(
    go.Scatter(
        x=pareto["DEFECT_TYPE"],
        y=pareto["CUM_PCT"],
        name=tr("pareto_cum_name"),
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#10233f", width=2),
        marker=dict(size=6),
    )
)
fig_pareto.add_shape(
    type="line",
    x0=-0.5, x1=len(pareto) - 0.5, y0=80, y1=80,
    xref="x", yref="y2",
    line=dict(color="#e67e22", width=2, dash="dash"),
)
fig_pareto.add_annotation(
    x=len(pareto) - 1, y=83, yref="y2", text=tr("pareto_ref_line"),
    showarrow=False, font=dict(color="#e67e22", size=12),
)
fig_pareto.update_layout(
    yaxis=dict(title=tr("pareto_yaxis")),
    yaxis2=dict(title=tr("pareto_yaxis2"), overlaying="y", side="right", range=[0, 105]),
    xaxis=dict(title="", tickangle=-45),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    height=520,
    margin=dict(t=40, b=140),
    plot_bgcolor="white",
)
st.plotly_chart(fig_pareto, width='stretch')
st.caption(
    tr(
        "pareto_caption",
        n=len(critical_defects),
        sample=", ".join(critical_defects[:6]),
        more="…" if len(critical_defects) > 6 else "",
    )
)

# ---- Critical defects detail table (same analysis as the chart above, drilled down
# to each critical defect's main affected references) ----
crit_rows = []
for d in critical_defects:
    sub = long_f[long_f["DEFECT_TYPE"] == d]
    top_refs = (
        sub.groupby("REFERANS")["DEFECT_COUNT"].sum().sort_values(ascending=False).head(3)
    )
    ref_txt = ", ".join([f"{r} ({v:.0f})" for r, v in top_refs.items()])
    crit_rows.append({
        tr("col_defect_type"): d,
        tr("col_quantity"): int(sub["DEFECT_COUNT"].sum()),
        tr("col_cum_pct"): round(pareto.loc[pareto["DEFECT_TYPE"] == d, "CUM_PCT"].values[0], 1),
        tr("col_main_refs"): ref_txt,
    })
crit_table = pd.DataFrame(crit_rows)
st.dataframe(crit_table, hide_index=True, width='stretch')

# --------------------------------------------------------------------------------------
# 3. REFERENCE ANALYSIS
# --------------------------------------------------------------------------------------
section_header(tr("section_reference"))

ref_stats = (
    wide_f.groupby("REFERANS", as_index=False)
    .agg(INSPECTED=(insp_col, "sum"), NOK=(nok_col, "sum"))
)
ref_stats["NOK_PCT"] = np.where(
    ref_stats["INSPECTED"] > 0, ref_stats["NOK"] / ref_stats["INSPECTED"] * 100, np.nan
)
ref_stats = ref_stats.sort_values("NOK_PCT", ascending=False)

col_a, col_b = st.columns([1, 1])

with col_a:
    st.markdown(tr("ref_vs_nok_title"))
    fig_ref = px.bar(
        ref_stats.head(20),
        x="NOK_PCT", y="REFERANS", orientation="h",
        color="NOK_PCT", color_continuous_scale="Reds",
        labels={"NOK_PCT": tr("nok_pct_label"), "REFERANS": tr("col_reference")},
    )
    fig_ref.update_layout(
        height=520, yaxis=dict(categoryorder="total ascending"),
        coloraxis_showscale=False, plot_bgcolor="white",
    )
    st.plotly_chart(fig_ref, width='stretch')

with col_b:
    st.markdown(tr("heatmap_title"))
    heat = long_f.pivot_table(
        index="REFERANS", columns="DEFECT_TYPE", values="DEFECT_COUNT", aggfunc="sum", fill_value=0
    )
    # keep the view readable: top references & top defect types by volume
    top_refs_for_heat = wide_f.groupby("REFERANS")[insp_col].count().nlargest(20).index
    top_defects_for_heat = pareto["DEFECT_TYPE"].head(15).tolist()
    heat = heat.reindex(index=[r for r in top_refs_for_heat if r in heat.index],
                         columns=[d for d in top_defects_for_heat if d in heat.columns])
    fig_heat = px.imshow(
        heat, color_continuous_scale="OrRd", aspect="auto",
        labels=dict(color=tr("heatmap_color_label")),
    )
    fig_heat.update_layout(height=520, margin=dict(t=10))
    st.plotly_chart(fig_heat, width='stretch')

st.markdown(tr("top10_title"))
combo = (
    long_f.groupby(["REFERANS", "DEFECT_TYPE"], as_index=False)["DEFECT_COUNT"]
    .sum()
    .sort_values("DEFECT_COUNT", ascending=False)
    .head(10)
    .reset_index(drop=True)
)
combo.index = combo.index + 1
st.dataframe(
    combo.rename(columns={
        "REFERANS": tr("col_reference"),
        "DEFECT_TYPE": tr("col_defect_type"),
        "DEFECT_COUNT": tr("col_quantity"),
    }),
    width='stretch',
)

# --------------------------------------------------------------------------------------
# 4. TIME ANALYSIS
# --------------------------------------------------------------------------------------
section_header(tr("section_time"))


def aggregate_trend(data: pd.DataFrame, granularity: str) -> pd.DataFrame:
    """Group a long-format defect subset into a period trend at the requested granularity."""
    d = data.copy()
    if granularity == "weekly":
        d["PERIOD"] = (d["TARIH"] - pd.to_timedelta(d["TARIH"].dt.weekday, unit="D")).dt.date
    elif granularity == "monthly":
        d["PERIOD"] = d["TARIH"].dt.to_period("M").astype(str)
    else:
        d["PERIOD"] = d["TARIH"].dt.date
    return d.groupby("PERIOD", as_index=False)["DEFECT_COUNT"].sum().sort_values("PERIOD")


# daily aggregation is always computed (used later by Key Insights), independent of the
# granularity chosen for on-screen charts below
daily = long_f.groupby("DATE_ONLY", as_index=False)["DEFECT_COUNT"].sum().sort_values("DATE_ONLY")

granularity_options = {
    tr("granularity_daily"): "daily",
    tr("granularity_weekly"): "weekly",
    tr("granularity_monthly"): "monthly",
}
gran_choice_label = st.radio(
    tr("granularity_label"), list(granularity_options.keys()), horizontal=True, key="time_granularity"
)
gran_choice = granularity_options[gran_choice_label]

main_trend = aggregate_trend(long_f, gran_choice)
fig_time = px.line(main_trend, x="PERIOD", y="DEFECT_COUNT", markers=True,
                    labels={"PERIOD": tr("time_xaxis"), "DEFECT_COUNT": tr("time_yaxis")})
fig_time.update_traces(line_color="#1c3f6e")
if gran_choice == "daily" and len(main_trend) >= 3:
    main_trend["MA7"] = main_trend["DEFECT_COUNT"].rolling(window=min(7, len(main_trend)), min_periods=1).mean()
    fig_time.add_trace(go.Scatter(x=main_trend["PERIOD"], y=main_trend["MA7"], name=tr("time_rolling_avg"),
                                   line=dict(color="#e67e22", dash="dot")))
fig_time.update_layout(height=420, plot_bgcolor="white")
st.plotly_chart(fig_time, width='stretch')

# ---- 4b. Defect-type-specific trend -------------------------------------------------
st.markdown(f"#### {tr('defect_trend_subheader')}")
defects_present = pareto["DEFECT_TYPE"].tolist()
if defects_present:
    sel_defect_trend = st.selectbox(
        tr("defect_trend_select_label"), defects_present, index=0, key="defect_trend_select"
    )
    defect_sub = long_f[long_f["DEFECT_TYPE"] == sel_defect_trend]
    defect_trend = aggregate_trend(defect_sub, gran_choice)
    fig_defect_trend = px.line(
        defect_trend, x="PERIOD", y="DEFECT_COUNT", markers=True,
        labels={"PERIOD": tr("time_xaxis"), "DEFECT_COUNT": tr("trend_yaxis_generic")},
    )
    fig_defect_trend.update_traces(line_color="#c0392b")
    fig_defect_trend.update_layout(height=360, plot_bgcolor="white")
    st.plotly_chart(fig_defect_trend, width='stretch')

# ---- 4c. Reference-specific defect trend --------------------------------------------
st.markdown(f"#### {tr('reference_trend_subheader')}")
refs_present = sorted(long_f["REFERANS"].dropna().unique().tolist())
if refs_present:
    default_ref_idx = 0
    if not ref_stats.empty:
        top_vol_ref = ref_stats.sort_values("INSPECTED", ascending=False).iloc[0]["REFERANS"]
        if top_vol_ref in refs_present:
            default_ref_idx = refs_present.index(top_vol_ref)
    sel_ref_trend = st.selectbox(
        tr("reference_trend_select_label"), refs_present, index=default_ref_idx, key="reference_trend_select"
    )
    ref_sub = long_f[long_f["REFERANS"] == sel_ref_trend]
    ref_trend = aggregate_trend(ref_sub, gran_choice)
    fig_ref_trend = px.line(
        ref_trend, x="PERIOD", y="DEFECT_COUNT", markers=True,
        labels={"PERIOD": tr("time_xaxis"), "DEFECT_COUNT": tr("trend_yaxis_generic")},
    )
    fig_ref_trend.update_traces(line_color="#1c3f6e")
    fig_ref_trend.update_layout(height=360, plot_bgcolor="white")
    st.plotly_chart(fig_ref_trend, width='stretch')

# --------------------------------------------------------------------------------------
# 4d. SIMPLE TREND PROJECTION (linear extrapolation of daily NOK %, informational only)
# --------------------------------------------------------------------------------------
section_header(tr("section_forecast"))

daily_nok = (
    wide_f.groupby("DATE_ONLY", as_index=False)
    .agg(NOK=(nok_col, "sum"), INSPECTED=(insp_col, "sum"))
    .sort_values("DATE_ONLY")
)
daily_nok = daily_nok[daily_nok["INSPECTED"] > 0].reset_index(drop=True)
daily_nok["NOK_PCT"] = daily_nok["NOK"] / daily_nok["INSPECTED"] * 100

if len(daily_nok) >= 4:
    max_n = len(daily_nok)
    default_n = min(14, max_n)
    n_window = st.slider(
        tr("forecast_window_label"), min_value=4, max_value=max_n, value=default_n, key="forecast_window"
    )
    train = daily_nok.tail(n_window).reset_index(drop=True)
    x = np.arange(len(train))
    y = train["NOK_PCT"].to_numpy()
    slope, intercept = np.polyfit(x, y, 1)
    fitted = slope * x + intercept

    future_x = np.arange(len(train), len(train) + 7)
    future_y = np.clip(slope * future_x + intercept, 0, 100)
    last_date = pd.Timestamp(train["DATE_ONLY"].iloc[-1])
    future_dates = pd.date_range(last_date + pd.Timedelta(days=1), periods=7)
    avg_projection = float(future_y.mean())

    eps = 1e-6
    if slope > eps:
        direction = tr("forecast_trend_up")
    elif slope < -eps:
        direction = tr("forecast_trend_down")
    else:
        direction = tr("forecast_trend_flat")

    fig_fc = go.Figure()
    fig_fc.add_trace(go.Scatter(
        x=train["DATE_ONLY"], y=train["NOK_PCT"], mode="markers",
        name=tr("forecast_series_actual"), marker=dict(color="#1c3f6e", size=7),
    ))
    fig_fc.add_trace(go.Scatter(
        x=train["DATE_ONLY"], y=fitted, mode="lines",
        name=tr("forecast_series_trend"), line=dict(color="#7f9cc4", dash="dash"),
    ))
    fig_fc.add_trace(go.Scatter(
        x=future_dates, y=future_y, mode="lines+markers",
        name=tr("forecast_series_projection"), line=dict(color="#c0392b", dash="dot"),
    ))
    fig_fc.update_layout(
        height=400, plot_bgcolor="white",
        xaxis=dict(title=tr("forecast_xaxis")), yaxis=dict(title=tr("forecast_yaxis")),
    )
    st.plotly_chart(fig_fc, width='stretch')
    st.caption(tr("forecast_result", n=n_window, val=avg_projection, direction=direction))
    st.caption(tr("forecast_intro"))
else:
    st.info(tr("forecast_insufficient"))

# --------------------------------------------------------------------------------------
# 4e. STATISTICAL PROCESS CONTROL — p-CHART (control limits computed live, not hard-coded)
# --------------------------------------------------------------------------------------
section_header(tr("section_spc"))
st.caption(tr("spc_intro"))


def aggregate_period_nok(data: pd.DataFrame, granularity: str) -> pd.DataFrame:
    d = data.copy()
    if granularity == "weekly":
        d["PERIOD"] = (d["TARIH"] - pd.to_timedelta(d["TARIH"].dt.weekday, unit="D")).dt.date
    elif granularity == "monthly":
        d["PERIOD"] = d["TARIH"].dt.to_period("M").astype(str)
    else:
        d["PERIOD"] = d["TARIH"].dt.date
    g = d.groupby("PERIOD", as_index=False).agg(NOK=(nok_col, "sum"), INSPECTED=(insp_col, "sum"))
    return g.sort_values("PERIOD")


p_data = aggregate_period_nok(wide_f, gran_choice)
p_data = p_data[p_data["INSPECTED"] > 0].reset_index(drop=True)

if len(p_data) >= 3:
    p_bar = p_data["NOK"].sum() / p_data["INSPECTED"].sum()
    p_data["P_PCT"] = p_data["NOK"] / p_data["INSPECTED"] * 100
    sigma = np.sqrt(p_bar * (1 - p_bar) / p_data["INSPECTED"])
    p_data["UCL"] = np.minimum(100, (p_bar + 3 * sigma) * 100)
    p_data["LCL"] = np.maximum(0, (p_bar - 3 * sigma) * 100)
    p_data["OOC"] = (p_data["P_PCT"] > p_data["UCL"]) | (p_data["P_PCT"] < p_data["LCL"])

    fig_spc = go.Figure()
    fig_spc.add_trace(go.Scatter(
        x=p_data["PERIOD"], y=p_data["P_PCT"], mode="lines+markers",
        name=tr("spc_series_actual"), line=dict(color="#1c3f6e"),
    ))
    fig_spc.add_trace(go.Scatter(
        x=p_data["PERIOD"], y=[p_bar * 100] * len(p_data), mode="lines",
        name=tr("spc_series_cl"), line=dict(color="#27ae60", dash="dash"),
    ))
    fig_spc.add_trace(go.Scatter(
        x=p_data["PERIOD"], y=p_data["UCL"], mode="lines",
        name=tr("spc_series_ucl"), line=dict(color="#e67e22", dash="dot"),
    ))
    fig_spc.add_trace(go.Scatter(
        x=p_data["PERIOD"], y=p_data["LCL"], mode="lines",
        name=tr("spc_series_lcl"), line=dict(color="#e67e22", dash="dot"),
    ))
    ooc_points = p_data[p_data["OOC"]]
    if not ooc_points.empty:
        fig_spc.add_trace(go.Scatter(
            x=ooc_points["PERIOD"], y=ooc_points["P_PCT"], mode="markers",
            name=tr("spc_series_ooc"), marker=dict(color="#c0392b", size=12, symbol="x"),
        ))
    fig_spc.update_layout(
        height=440, plot_bgcolor="white",
        xaxis=dict(title=tr("time_xaxis")), yaxis=dict(title=tr("spc_yaxis")),
    )
    st.plotly_chart(fig_spc, width='stretch')
    st.caption(tr("spc_caption", n=int(p_data["OOC"].sum()), total=len(p_data)))
else:
    st.info(tr("spc_insufficient"))

# --------------------------------------------------------------------------------------
# 6. DETAILED FILTERED DATA + DOWNLOAD
# --------------------------------------------------------------------------------------
section_header(tr("section_detail"))

# Only keep defect-type columns that actually have at least one non-zero value
# within the current filter selection (drop all-zero columns dynamically).
nonzero_defect_cols = [c for c in defect_cols if c in wide_f.columns and wide_f[c].sum() > 0]

display_cols = ["REFERANS", "TARIH"]
if shift_col:
    display_cols.append("SHIFT")
display_cols += [insp_col, ok_col, nok_col, "NOK_PCT"] + nonzero_defect_cols
display_cols = [c for c in display_cols if c and c in wide_f.columns]

detail_view = wide_f[display_cols].sort_values("TARIH").reset_index(drop=True)

# Human-readable column headers (translated where the column is app-generated;
# original source headers like defect-type names are left untouched).
rename_map = {
    "REFERANS": tr("col_reference"),
    "TARIH": tr("col_date"),
    insp_col: tr("col_inspected"),
    ok_col: tr("col_ok"),
    nok_col: tr("col_nok"),
    "NOK_PCT": tr("col_nok_pct_short"),
}
rename_map = {k: v for k, v in rename_map.items() if k}
detail_view_display = detail_view.rename(columns=rename_map)
nok_pct_label = tr("col_nok_pct_short")

st.dataframe(
    detail_view_display,
    height=380,
    width='stretch',
    column_config={
        nok_pct_label: st.column_config.NumberColumn(nok_pct_label, format="%.2f%%")
    } if nok_pct_label in detail_view_display.columns else None,
)
st.caption(tr("detail_pct_note"))

dl1, dl2, _ = st.columns([1, 1, 4])
csv_bytes = detail_view_display.to_csv(index=False).encode("utf-8-sig")
dl1.download_button(tr("download_csv"), csv_bytes, file_name="filtered_quality_data.csv", mime="text/csv")

xlsx_buffer = io.BytesIO()
with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
    detail_view_display.to_excel(writer, index=False, sheet_name="Filtered Data")
dl2.download_button(
    tr("download_excel"),
    xlsx_buffer.getvalue(),
    file_name="filtered_quality_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# --------------------------------------------------------------------------------------
# 7. KEY INSIGHTS (auto-generated, data-driven only)
# --------------------------------------------------------------------------------------
section_header(tr("section_insights"))

insights = []

if not pareto.empty:
    top_defect = pareto.iloc[0]
    insights.append(tr(
        "insight_top_defect",
        defect=top_defect["DEFECT_TYPE"],
        qty=top_defect["DEFECT_COUNT"],
        pct=top_defect["DEFECT_COUNT"] / pareto["DEFECT_COUNT"].sum() * 100,
    ))
    insights.append(tr(
        "insight_critical_share",
        n=len(critical_defects), total=len(pareto),
        pct=len(critical_defects) / len(pareto) * 100,
    ))

if not ref_stats.empty and ref_stats["NOK_PCT"].notna().any():
    worst_ref = ref_stats.sort_values("NOK_PCT", ascending=False).iloc[0]
    best_ref = ref_stats[ref_stats["INSPECTED"] > 0].sort_values("NOK_PCT", ascending=True).iloc[0]
    insights.append(tr(
        "insight_worst_ref",
        ref=worst_ref["REFERANS"], pct=worst_ref["NOK_PCT"],
        nok=worst_ref["NOK"], insp=worst_ref["INSPECTED"],
    ))
    insights.append(tr("insight_best_ref", ref=best_ref["REFERANS"], pct=best_ref["NOK_PCT"]))

if not combo.empty:
    top_combo = combo.iloc[0]
    insights.append(tr(
        "insight_top_combo",
        ref=top_combo["REFERANS"], defect=top_combo["DEFECT_TYPE"], qty=top_combo["DEFECT_COUNT"],
    ))

if len(daily) >= 2:
    worst_day = daily.sort_values("DEFECT_COUNT", ascending=False).iloc[0]
    insights.append(tr("insight_worst_day", date=worst_day["DATE_ONLY"], qty=worst_day["DEFECT_COUNT"]))
    if len(daily) >= 4:
        first_half = daily.iloc[: len(daily) // 2]["DEFECT_COUNT"].mean()
        second_half = daily.iloc[len(daily) // 2:]["DEFECT_COUNT"].mean()
        if first_half > 0:
            change = (second_half - first_half) / first_half * 100
            direction = tr("trend_increased") if change > 0 else tr("trend_decreased")
            insights.append(tr("insight_trend", pct=abs(change), direction=direction))

if not np.isnan(overall_nok_pct):
    insights.append(tr("insight_overall_nok", pct=overall_nok_pct, nok=total_nok_units, insp=total_inspections))

if not insights:
    insights.append(tr("insight_none"))

for i, ins in enumerate(insights, 1):
    st.markdown(f'<div class="insight-box">{i}. {ins}</div>', unsafe_allow_html=True)

st.caption(tr("insight_caption"))

st.markdown("---")
st.caption(tr("footer_caption"))
