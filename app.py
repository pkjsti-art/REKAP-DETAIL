import streamlit as st
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill

# ==========================================
# UI/UX: TEMA DARK MODE (PROFESIONAL & BERSIH)
# ==========================================
st.set_page_config(page_title="Sistem Rekap Detail Produksi", layout="wide", initial_sidebar_state="collapsed")

custom_css = """
<style>
    /* Latar Belakang Utama Aplikasi (Dark Navy) */
    .stApp {
        background-color: #040814 !important;
        background-image: radial-gradient(circle at 50% 0%, #0a1931 0%, #040814 70%);
        color: #e0f7fa !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* Styling Judul dan Header */
    h1, h2, h3 {
        color: #00f3ff !important;
        text-shadow: 0 0
