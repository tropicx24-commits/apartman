# -*- coding: utf-8 -*-
"""
Apartman Aidat Sistemi (Mini v4) + Devir (Sakin Geçmişi) - UI Revamp

Masaüstü PySide6 uygulaması.

Özellikler:
- Daire ekle/güncelle/aktif-pasif
- Sakin geçmişi: Devir (taşınma) + geçmiş liste
- Tahakkuk (dönem borç yaz) + manuel geçmiş borç
- Excel'den borç içe aktar (daire yoksa otomatik oluşturur)
- Ödeme al (Banka/Elden) + peşin çok ay paylaştırma + Makbuz PDF
- Gider ekle/sil
- Duyuru ekle/sil + panoya kopyala
- Rapor: gelir/gider/net + borçlu/gecikmiş filtre + Excel export
- WhatsApp: güncel sakinin telefonuna wa.me linki açar

Not: WhatsApp otomatik göndermez, sadece tarayıcıda mesaj taslağı açar.
"""

import sys
import sqlite3
from pathlib import Path
from datetime import date, timedelta
import webbrowser
from urllib.parse import quote


from PySide6.QtCore import Qt, QDate
from PySide6.QtWidgets import (
    QApplication, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QComboBox, QGroupBox, QFileDialog, QSpinBox, QCheckBox,
    QTextEdit, QDateEdit, QSplitter, QHeaderView, QScrollArea, QDialog  # ⭐ QDialog ekle
)
from PySide6.QtGui import QColor

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


DB_PATH = Path("apartman_aidat.db")


# ✅ DÜZELTME 1: FIX ============================================
# odeme_detay tablosu oluşturmuyoruz (gereksiz ve kullanılmıyor)
# ============================================================

# ... [Burada yardımcı fonksiyonlar kalıyor - değiştirilmeyecek] ...
