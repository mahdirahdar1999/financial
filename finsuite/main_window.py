# -*- coding: utf-8 -*-
import os, json
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtWidgets import (QMainWindow, QWidget, QFileDialog, QMessageBox, QLabel, QVBoxLayout, QHBoxLayout,
                             QGroupBox, QGridLayout, QPushButton, QCheckBox, QDateEdit, QComboBox, QStatusBar)
from .excel_builder import MODULES, build_excel, Options

RES_DIR = os.path.join(os.path.dirname(__file__), "resources")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FinSuite Pro — اکسل‌ساز حسابداری (فارسی)")
        self.resize(1100, 720)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        try:
            self.setStyleSheet(open(os.path.join(RES_DIR,"style.qss"),encoding="utf-8").read())
        except Exception: pass
        self._build_ui()

    def _build_ui(self):
        central = QWidget(self); self.setCentralWidget(central)
        root = QVBoxLayout(central); root.setContentsMargins(14,14,14,14); root.setSpacing(10)

        # Params
        box = QGroupBox("⚙️ پارامترهای اصلی"); grid = QGridLayout(box)
        grid.addWidget(QLabel("شروع سال مالی"),0,0)
        self.start = QDateEdit(calendarPopup=True); self.start.setDate(QDate.currentDate().addDays(-QDate.currentDate().dayOfYear()+1)); grid.addWidget(self.start,0,1)
        grid.addWidget(QLabel("پایان سال مالی"),0,2)
        self.end = QDateEdit(calendarPopup=True); self.end.setDate(QDate(self.start.date().year(),12,31)); grid.addWidget(self.end,0,3)
        grid.addWidget(QLabel("واحد پول گروه"),0,4)
        self.curr = QComboBox(); self.curr.addItems(["IRR","EUR","USD"]); grid.addWidget(self.curr,0,5)
        root.addWidget(box)

        # Modules
        mods = QGroupBox("📦 انتخاب ماژول‌ها"); mgrid = QGridLayout(mods)
        self.checks = {}
        for i,k in enumerate(MODULES.keys()):
            cb = QCheckBox(k); cb.setChecked(k not in ("README","CONFIG"))
            self.checks[k] = cb; mgrid.addWidget(cb, i//3, i%3)
        root.addWidget(mods,1)

        # Buttons
        row = QHBoxLayout()
        self.btn_all = QPushButton("انتخاب همه")
        self.btn_none = QPushButton("پاک‌سازی")
        self.btn_sme = QPushButton("Preset: SME ایران")
        self.btn_grp = QPushButton("Preset: گروه چندشرکتی")
        self.btn_save = QPushButton("ذخیره Preset…")
        self.btn_load = QPushButton("بارگذاری Preset…")
        self.btn_build = QPushButton("ساخت فایل اکسل")
        for b in (self.btn_all,self.btn_none,self.btn_sme,self.btn_grp,self.btn_save,self.btn_load,self.btn_build): row.addWidget(b)
        root.addLayout(row)

        self.status = QStatusBar(); self.setStatusBar(self.status); self.status.showMessage("آماده")

        # signals
        self.btn_all.clicked.connect(self._all)
        self.btn_none.clicked.connect(self._none)
        self.btn_sme.clicked.connect(self._sme)
        self.btn_grp.clicked.connect(self._grp)
        self.btn_save.clicked.connect(self._save)
        self.btn_load.clicked.connect(self._load)
        self.btn_build.clicked.connect(self._build)

    def _keys(self):
        keys = ["README","CONFIG"]
        for k,cb in self.checks.items():
            if cb.isChecked(): keys.append(k)
        seen=set(); out=[]
        for k in keys:
            if k not in seen: seen.add(k); out.append(k)
        return out

    def _opts(self):
        return Options(self.start.date().toString("yyyy-MM-dd"),
                       self.end.date().toString("yyyy-MM-dd"),
                       self.curr.currentText(), True)

    def _all(self):
        for cb in self.checks.values(): cb.setChecked(True)

    def _none(self):
        for cb in self.checks.values(): cb.setChecked(False)

    def _sme(self):
        self._none()
        for k in ["COA","MASTER_DATA","JOURNAL","GL","TRIAL","PL","BS","CASHFLOW","SALES_PURCH","VAT_RETURN","AP_AR","BANK_REC","INVENTORY","FIXED_ASSETS","PAYROLL","BUDGET","QUALITY_KPIs"]:
            self.checks[k].setChecked(True)

    def _grp(self):
        self._sme()
        for k in ["FX_RATES","ENTITIES","MAP_PL","MAP_BS","CONS_ELIMS","TB_GROUP_PRE","TB_GROUP_POST","CLOSE_CHECK"]:
            self.checks[k].setChecked(True)

    def _save(self):
        p,_ = QFileDialog.getSaveFileName(self,"ذخیره Preset","","JSON (*.json)"); 
        if not p: return
        json.dump({"selected":[k for k,cb in self.checks.items() if cb.isChecked()]}, open(p,"w",encoding="utf-8"), ensure_ascii=False, indent=2)
        QMessageBox.information(self,"ذخیره شد","Preset ذخیره شد.")

    def _load(self):
        p,_ = QFileDialog.getOpenFileName(self,"بارگذاری Preset","","JSON (*.json)"); 
        if not p: return
        try:
            sel=set(json.load(open(p,"r",encoding="utf-8")).get("selected",[]))
            for k,cb in self.checks.items(): cb.setChecked(k in sel)
            QMessageBox.information(self,"انجام شد","Preset بارگذاری شد.")
        except Exception as e:
            QMessageBox.critical(self,"خطا",str(e))

    def _build(self):
        p,_ = QFileDialog.getSaveFileName(self,"ذخیره فایل اکسل","FinSuite.xlsx","Excel (*.xlsx)")
        if not p: return
        try:
            build_excel(p, self._keys(), self._opts())
            QMessageBox.information(self,"موفق",f"فایل ساخته شد:\n{p}")
            self.status.showMessage(f"ساخته شد: {os.path.basename(p)}", 8000)
        except Exception as e:
            QMessageBox.critical(self,"خطا",f"ساخت فایل با خطا مواجه شد:\n{e}")
            self.status.showMessage("خطا در ساخت فایل", 8000)
