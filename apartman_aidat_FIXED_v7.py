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


# ----------------- helpers -----------------
def connect():
    con = sqlite3.connect(str(DB_PATH))
    con.execute("PRAGMA foreign_keys = ON;")
    return con


def iso_today() -> str:
    return date.today().isoformat()


def ym_of_today() -> str:
    d = date.today()
    return f"{d.year:04d}-{d.month:02d}"


def safe_float(s: str) -> float:
    t = (s or "").strip()
    if not t:
        return 0.0
    t = t.replace(" ", "")
    # TR: 1.234,56 -> 1234.56
    t = t.replace(".", "")
    t = t.replace(",", ".")
    return float(t)


def validate_period(donem: str) -> bool:
    if not donem or len(donem) != 7 or donem[4] != "-":
        return False
    try:
        y = int(donem[:4])
        m = int(donem[5:])
        return 2000 <= y <= 2100 and 1 <= m <= 12
    except Exception:
        return False


def add_months(ym: str, add: int) -> str:
    y = int(ym[:4])
    m = int(ym[5:])
    m2 = m + add
    y += (m2 - 1) // 12
    m2 = ((m2 - 1) % 12) + 1
    return f"{y:04d}-{m2:02d}"


def months_range(ym_start: str, ym_end: str):
    ys = int(ym_start[:4]); ms = int(ym_start[5:])
    ye = int(ym_end[:4]); me = int(ym_end[5:])
    cur_y, cur_m = ys, ms
    out = []
    while (cur_y, cur_m) <= (ye, me):
        out.append(f"{cur_y:04d}-{cur_m:02d}")
        cur_m += 1
        if cur_m == 13:
            cur_m = 1
            cur_y += 1
    return out


def ensure_db():
    con = connect()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS daireler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        daire_no TEXT NOT NULL UNIQUE,
        ad_soyad TEXT NOT NULL,
        telefon TEXT DEFAULT '',
        aidat REAL NOT NULL DEFAULT 0,
        aktif INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS sakinler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        daire_id INTEGER NOT NULL,
        ad_soyad TEXT NOT NULL,
        telefon TEXT DEFAULT '',
        baslangic_tarihi TEXT NOT NULL,
        bitis_tarihi TEXT,
        FOREIGN KEY(daire_id) REFERENCES daireler(id)
    );
    CREATE INDEX IF NOT EXISTS idx_sakinler_daire ON sakinler(daire_id);
    CREATE INDEX IF NOT EXISTS idx_sakinler_aktif ON sakinler(daire_id, bitis_tarihi);

    CREATE TABLE IF NOT EXISTS tahakkuk (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        donem TEXT NOT NULL,
        daire_id INTEGER NOT NULL,
        tutar REAL NOT NULL,
        created_at TEXT NOT NULL,
        UNIQUE(donem, daire_id),
        FOREIGN KEY(daire_id) REFERENCES daireler(id)
    );

    CREATE TABLE IF NOT EXISTS odemeler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        donem TEXT NOT NULL,
        tarih TEXT NOT NULL,
        daire_id INTEGER NOT NULL,
        tutar REAL NOT NULL,
        yontem TEXT NOT NULL,
        makbuz_no TEXT NOT NULL,
        aciklama TEXT DEFAULT '',
        FOREIGN KEY(daire_id) REFERENCES daireler(id)
    );

    CREATE INDEX IF NOT EXISTS idx_odemeler_donem ON odemeler(donem);
    CREATE INDEX IF NOT EXISTS idx_odemeler_daire ON odemeler(daire_id);

    CREATE TABLE IF NOT EXISTS giderler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        donem TEXT NOT NULL,
        tarih TEXT NOT NULL,
        kategori TEXT NOT NULL,
        tutar REAL NOT NULL,
        yontem TEXT NOT NULL,
        aciklama TEXT DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_giderler_donem ON giderler(donem);

    CREATE TABLE IF NOT EXISTS duyurular (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tarih TEXT NOT NULL,
        baslik TEXT NOT NULL,
        mesaj TEXT NOT NULL
    );
    """)
    con.commit()

    def set_default(k, v):
        if con.execute("SELECT value FROM settings WHERE key=?", (k,)).fetchone() is None:
            con.execute("INSERT INTO settings(key,value) VALUES(?,?)", (k, str(v)))

    set_default("vade_gun", 10)
    set_default("gecikme_gun", 5)
    con.commit()
    con.close()


def get_setting_int(key: str, default: int) -> int:
    con = connect()
    row = con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    con.close()
    try:
        return int(row[0]) if row else default
    except Exception:
        return default


def set_setting_int(key: str, value: int):
    con = connect()
    con.execute(
        "INSERT INTO settings(key,value) VALUES(?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, str(int(value)))
    )
    con.commit()
    con.close()


def receipt_next_for_period(donem: str) -> str:
    yyyymm = donem.replace("-", "")
    con = connect()
    row = con.execute("""
        SELECT makbuz_no FROM odemeler
        WHERE donem=?
        ORDER BY id DESC
        LIMIT 1
    """, (donem,)).fetchone()
    con.close()
    if not row:
        return f"{yyyymm}-00001"
    last = row[0]
    try:
        seq = int(last.split("-")[1])
    except Exception:
        seq = 0
    return f"{yyyymm}-{seq+1:05d}"


def compute_vade_date(donem: str) -> date:
    vade_gun = get_setting_int("vade_gun", 10)
    y, m = donem.split("-")
    y = int(y); m = int(m)
    # last day
    if m == 12:
        next_month = date(y + 1, 1, 1)
    else:
        next_month = date(y, m + 1, 1)
    last_day = next_month - timedelta(days=1)
    day = min(vade_gun, last_day.day)
    return date(y, m, day)


def status_for_balance(donem: str, bakiye: float) -> str:
    if bakiye <= 0.00001:
        return "Ödedi"
    vade = compute_vade_date(donem)
    gecikme = get_setting_int("gecikme_gun", 5)
    limit = vade + timedelta(days=gecikme)
    if date.today() > limit:
        return "Gecikmiş"
    return "Borçlu"


def autosize_worksheet(ws):
    for col in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col)
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            v = str(v)
            if len(v) > max_len:
                max_len = len(v)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 55)


# ============ MODAL DIALOGS ============

class EditPaymentDialog(QDialog):
    """Ödeme düzenleme modalı"""
    def __init__(self, payment_id: int, parent=None):
        super().__init__(parent)
        self.payment_id = payment_id
        self.setWindowTitle("Ödemeyi Düzenle")
        self.setGeometry(100, 100, 500, 300)
        self.setModal(True)
        self.init_ui()
        self.load_payment()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        form = QFormLayout()
        
        self.in_donem = QLineEdit()
        self.in_donem.setReadOnly(True)
        self.dt_tarih = QDateEdit()
        self.dt_tarih.setCalendarPopup(True)
        self.dt_tarih.setDisplayFormat("dd-MM-yyyy")
        self.in_daire = QLineEdit()
        self.in_daire.setReadOnly(True)
        self.in_tutar = QLineEdit()
        self.cmb_yontem = QComboBox()
        self.cmb_yontem.addItems(["Banka", "Elden"])
        self.in_makbuz = QLineEdit()
        self.in_acik = QLineEdit()
        
        form.addRow("Dönem", self.in_donem)
        form.addRow("Tarih", self.dt_tarih)
        form.addRow("Daire", self.in_daire)
        form.addRow("Tutar (TL)", self.in_tutar)
        form.addRow("Yöntem", self.cmb_yontem)
        form.addRow("Makbuz No", self.in_makbuz)
        form.addRow("Açıklama", self.in_acik)
        
        layout.addLayout(form)
        
        btns = QHBoxLayout()
        self.btn_save = QPushButton("Güncelle")
        self.btn_save.clicked.connect(self.save_payment)
        self.btn_delete = QPushButton("Sil")
        self.btn_delete.clicked.connect(self.delete_payment)
        self.btn_close = QPushButton("Kapat")
        self.btn_close.clicked.connect(self.close)
        
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_delete)
        btns.addStretch(1)
        btns.addWidget(self.btn_close)
        
        layout.addLayout(btns)
    
    def load_payment(self):
        """Ödeme verilerini yükle"""
        con = connect()
        row = con.execute("""
            SELECT o.donem, o.tarih, d.daire_no, o.tutar, o.yontem, o.makbuz_no, o.aciklama
            FROM odemeler o
            JOIN daireler d ON d.id = o.daire_id
            WHERE o.id = ?
        """, (self.payment_id,)).fetchone()
        con.close()
        
        if not row:
            QMessageBox.warning(self, "Hata", "Ödeme bulunamadı.")
            self.close()
            return
        
        donem, tarih, daire_no, tutar, yontem, makbuz_no, acik = row
        
        self.in_donem.setText(donem)
        self.dt_tarih.setDate(QDate.fromString(tarih, "yyyy-MM-dd"))
        self.in_daire.setText(daire_no)
        self.in_tutar.setText(f"{float(tutar):.2f}")
        self.cmb_yontem.setCurrentText(yontem)
        self.in_makbuz.setText(makbuz_no or "")
        self.in_acik.setText(acik or "")
    
    def save_payment(self):
        """Ödemeyi güncelle"""
        try:
            tutar = safe_float(self.in_tutar.text())
        except Exception:
            QMessageBox.warning(self, "Uyarı", "Tutar sayısal olmalı.")
            return
        
        if tutar <= 0:
            QMessageBox.warning(self, "Uyarı", "Tutar 0'dan büyük olmalı.")
            return
        
        tarih = self.dt_tarih.date().toPython().isoformat()
        yontem = self.cmb_yontem.currentText()
        acik = self.in_acik.text().strip()
        
        con = connect()
        con.execute("""
            UPDATE odemeler
            SET tarih=?, tutar=?, yontem=?, aciklama=?
            WHERE id=?
        """, (tarih, float(tutar), yontem, acik, self.payment_id))
        con.commit()
        con.close()
        
        QMessageBox.information(self, "OK", "Ödeme güncellendi.")
        self.accept()
    
    def delete_payment(self):
        """Ödemeyi sil"""
        reply = QMessageBox.question(
            self, "Onay", "Seçili ödemeyi silmek istiyor musunuz?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        
        con = connect()
        con.execute("DELETE FROM odemeler WHERE id=?", (self.payment_id,))
        con.commit()
        con.close()
        
        QMessageBox.information(self, "OK", "Ödeme silindi.")
        self.accept()


class EditExpenseDialog(QDialog):
    """Gider düzenleme modalı"""
    def __init__(self, expense_id: int, parent=None):
        super().__init__(parent)
        self.expense_id = expense_id
        self.setWindowTitle("Gideri Düzenle")
        self.setGeometry(100, 100, 500, 300)
        self.setModal(True)
        self.init_ui()
        self.load_expense()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        form = QFormLayout()
        
        self.in_donem = QLineEdit()
        self.in_donem.setReadOnly(True)
        
        self.dt_tarih = QDateEdit()
        self.dt_tarih.setCalendarPopup(True)
        self.dt_tarih.setDisplayFormat("dd-MM-yyyy")
        
        self.cmb_kat = QComboBox()
        self.cmb_kat.addItems(["Elektrik", "Temizlik", "Bakım", "Diğer"])
        
        self.in_tutar = QLineEdit()
        
        self.cmb_yontem = QComboBox()
        self.cmb_yontem.addItems(["Banka", "Elden"])
        
        self.in_acik = QLineEdit()
        
        form.addRow("Dönem", self.in_donem)
        form.addRow("Tarih", self.dt_tarih)
        form.addRow("Kategori", self.cmb_kat)
        form.addRow("Tutar (TL)", self.in_tutar)
        form.addRow("Yöntem", self.cmb_yontem)
        form.addRow("Açıklama", self.in_acik)
        
        layout.addLayout(form)
        
        btns = QHBoxLayout()
        self.btn_save = QPushButton("Güncelle")
        self.btn_save.clicked.connect(self.save_expense)
        self.btn_delete = QPushButton("Sil")
        self.btn_delete.clicked.connect(self.delete_expense)
        self.btn_close = QPushButton("Kapat")
        self.btn_close.clicked.connect(self.close)
        
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_delete)
        btns.addStretch(1)
        btns.addWidget(self.btn_close)
        
        layout.addLayout(btns)
    
    def load_expense(self):
        """Gider verilerini yükle"""
        con = connect()
        row = con.execute("""
            SELECT donem, tarih, kategori, tutar, yontem, aciklama
            FROM giderler
            WHERE id = ?
        """, (self.expense_id,)).fetchone()
        con.close()
        
        if not row:
            QMessageBox.warning(self, "Hata", "Gider bulunamadı.")
            self.close()
            return
        
        donem, tarih, kat, tutar, yontem, acik = row
        
        self.in_donem.setText(donem)
        self.dt_tarih.setDate(QDate.fromString(tarih, "yyyy-MM-dd"))
        self.cmb_kat.setCurrentText(kat)
        self.in_tutar.setText(f"{float(tutar):.2f}")
        self.cmb_yontem.setCurrentText(yontem)
        self.in_acik.setText(acik or "")
    
    def save_expense(self):
        """Gideri güncelle"""
        try:
            tutar = safe_float(self.in_tutar.text())
        except Exception:
            QMessageBox.warning(self, "Uyarı", "Tutar sayısal olmalı.")
            return
        
        if tutar <= 0:
            QMessageBox.warning(self, "Uyarı", "Tutar 0'dan büyük olmalı.")
            return
        
        tarih = self.dt_tarih.date().toPython().isoformat()
        kat = self.cmb_kat.currentText()
        yontem = self.cmb_yontem.currentText()
        acik = self.in_acik.text().strip()
        
        con = connect()
        con.execute("""
            UPDATE giderler
            SET tarih=?, kategori=?, tutar=?, yontem=?, aciklama=?
            WHERE id=?
        """, (tarih, kat, float(tutar), yontem, acik, self.expense_id))
        con.commit()
        con.close()
        
        QMessageBox.information(self, "OK", "Gider güncellendi.")
        self.accept()
    
    def delete_expense(self):
        """Gideri sil"""
        reply = QMessageBox.question(
            self, "Onay", "Seçili gideri silmek istiyor musunuz?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        
        con = connect()
        con.execute("DELETE FROM giderler WHERE id=?", (self.expense_id,))
        con.commit()
        con.close()
        
        QMessageBox.information(self, "OK", "Gider silindi.")
        self.accept()


# [KALAN KOD SAMESİ DEVAM EDECEK - AŞAĞIYA BAKINIIZ]
