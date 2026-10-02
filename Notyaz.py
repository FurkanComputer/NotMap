import os
import sys
import pdfplumber
import datetime
from spellchecker import SpellChecker
import re
from PyQt5.QtCore import Qt,QTimer
from striprtf.striprtf import rtf_to_text
from PyQt5.QtGui import (
    QFont, QTextCharFormat, QTextListFormat, QTextImageFormat,
    QTextCursor, QTextBlockFormat,QTextDocument,QTextTableFormat
)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QComboBox, QFileDialog,
    QMessageBox, QTabWidget, QVBoxLayout, QHBoxLayout, QWidget,
    QFontComboBox, QColorDialog, QFrame, QToolButton, QLabel,
    QListWidget, QStatusBar, QPushButton, QAction, QButtonGroup,
    QSpinBox, QScrollArea,QInputDialog
)
try:
    import PyMuPDF
except ImportError:
    PyMuPDF = None
try:
    from docx import Document
except ImportError:
    Document = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
except ImportError:
    A4 = None
    canvas = None


class YazNot(QMainWindow):
    def __init__(self):
        super().__init__()

        self.TabNums = 0
        self.dark_mode = False

        self.setWindowTitle("Not Yaz ")
        self.resize(1500, 950)

        self.build_ui()
        self.apply_light_theme()
        self.NewFile()


    def check_license(self):
        valid_keys = ["ABC123-XYZ789", "DEF456-UVW000"]

        try:
            with open("license.key", "r", encoding="utf-8") as f:
                key = f.read().strip()
        except FileNotFoundError:
            # Dosya yoksa kullanıcıdan iste
            key, ok = QInputDialog.getText(self, "Lisans", "Lisans anahtarını girin:")
            if not ok:
                return False
            # Girilen anahtarı dosyaya kaydet
            with open("license.key", "w", encoding="utf-8") as f:
                f.write(key)

        return key in valid_keys

    # ==========================================================
    # ANA ARAYÜZ
    # ==========================================================
    def build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---------- ÜST BAŞLIK ----------
        self.titleBar = QFrame()
        self.titleBar.setObjectName("titleBar")
        title_layout = QHBoxLayout(self.titleBar)
        title_layout.setContentsMargins(14, 6, 14, 6)

        self.logo = QLabel("📝")
        self.logo.setObjectName("logo")

        self.appTitle = QLabel("Yaz Not")
        self.appTitle.setObjectName("appTitle")

        self.fileNameLabel = QLabel("Yeni Belge")
        self.fileNameLabel.setObjectName("fileNameLabel")

        title_layout.addWidget(self.logo)
        title_layout.addWidget(self.appTitle)
        title_layout.addSpacing(18)
        title_layout.addWidget(self.fileNameLabel)
        title_layout.addStretch()

        self.quickSave = self.smallButton("💾", "Kaydet")
        self.quickUndo = self.smallButton("↶", "Geri Al")
        self.quickRedo = self.smallButton("↷", "Yinele")

        self.quickSave.clicked.connect(self.saveAction)
        self.quickUndo.clicked.connect(self.undo)
        self.quickRedo.clicked.connect(self.redo)

        title_layout.addWidget(self.quickUndo)
        title_layout.addWidget(self.quickRedo)
        title_layout.addWidget(self.quickSave)

        root.addWidget(self.titleBar)

        # ---------- RIBBON SEKME BAŞLIKLARI ----------
        self.ribbonTabs = QTabWidget()
        self.ribbonTabs.setObjectName("ribbonTabs")
        self.ribbonTabs.setDocumentMode(True)
        self.ribbonTabs.setFixedHeight(175)

        

        self.create_home_ribbon()
        self.create_insert_ribbon()
        self.create_view_ribbon()

        root.addWidget(self.ribbonTabs)


        #tablo sekmesi ekle 
        self.ribbonTabs.addTab(
            self.ribbon_group(
                "Tablo",
                [

                    ("📊", "Tablo Ekle", lambda: self.insert_table(3, 3)),
                    ("➕", "Satır Ekle", lambda: self.add_row(self.get_current_table())),
                    ("➕", "Sütun Ekle", lambda: self.add_column(self.get_current_table())),
                    ("➖", "Satır Sil", lambda: self.remove_row(self.get_current_table(), 0)),
                    ("➖", "Sütun Sil", lambda: self.remove_column(self.get_current_table(), 0)),
                ]
            ),
            "Tablo"
        )

        self.ribbonTabs.addTab(
            self.ribbon_group(
                "Üstbilgi/AltBilgi",
                [
            ("⬆", "Üstbilgi", lambda: self.insert_header("Üstbilgi")),
            ("⬇", "Altbilgi", lambda: self.insert_footer("Altbilgi")),
            ("🔢", "Sayfa No", self.insert_page_number),
        ]
            ),
            "Üstbilgi/Altbilgi"
        )

        self.ribbonTabs.addTab(
            self.ribbon_group(
                "Yazım Denetimi",
                [
                    ("TR","Türkçe Kontrol", lambda: self.check_spelling("tr")),
                    ("🇬🇧", "İngilizce Kontrol", lambda: self.check_spelling("en")),
                ]
            ),
            "Denetim"
        )

        self.ribbonTabs.addTab(
            self.ribbon_group(
                "Tarih/Saat",
                [
                    ("📅", "Tarih", self.insert_date),
                    ("⏰", "Saat", self.insert_time),
                ]
            ),
            "Tarih/Saat"
        )
        self.ribbonTabs.addTab(
            self.ribbon_group(
                "Kaydetme",
                [
                    ("💾", "Sürüm Kaydet", self.save_version),
                ]
            ),
            "Kaydetme"
        )

        # ---------- ANA ALAN ----------
        body = QHBoxLayout()
        body.setContentsMargins(8, 8, 8, 8)
        body.setSpacing(8)

        # Sol panel
        self.leftPanel = QFrame()
        self.leftPanel.setObjectName("leftPanel")
        self.leftPanel.setFixedWidth(230)

        left_layout = QVBoxLayout(self.leftPanel)
        left_layout.setContentsMargins(10, 12, 10, 10)

        self.panelTitle = QLabel("📚 Belgeler")
        self.panelTitle.setObjectName("panelTitle")

        self.newDocButton = QPushButton("+  Yeni Belge")
        self.newDocButton.setObjectName("accentButton")
        self.newDocButton.clicked.connect(self.NewFile)

        self.tabList = QListWidget()
        self.tabList.currentRowChanged.connect(self.select_tab_from_list)

        left_layout.addWidget(self.panelTitle)
        left_layout.addWidget(self.newDocButton)
        left_layout.addSpacing(8)
        left_layout.addWidget(self.tabList)

        body.addWidget(self.leftPanel)

        # Editör alanı
        self.editorFrame = QFrame()
        self.editorFrame.setObjectName("editorFrame")

        editor_layout = QVBoxLayout(self.editorFrame)
        editor_layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("documentTabs")
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.tabs.tabCloseRequested.connect(self.CloseTheTab)

        editor_layout.addWidget(self.tabs)
        body.addWidget(self.editorFrame, 1)

        root.addLayout(body)

        # ---------- DURUM ÇUBUĞU ----------
        self.status = QStatusBar()
        self.status.setObjectName("statusBar")

        self.modeLabel = QLabel("Hazır")
        self.lineLabel = QLabel("Satır: 1")
        self.colLabel = QLabel("Sütun: 1")
        self.charLabel = QLabel("Karakter: 0")
        self.zoomLabel = QLabel("100%")

        self.status.addWidget(self.modeLabel)
        self.status.addPermanentWidget(self.lineLabel)
        self.status.addPermanentWidget(self.colLabel)
        self.status.addPermanentWidget(self.charLabel)
        self.status.addPermanentWidget(self.zoomLabel)

        self.setStatusBar(self.status)

    # ==========================================================
    # RIBBON
    # ==========================================================
    def create_home_ribbon(self):
        page = QWidget()
        main = QHBoxLayout(page)
        main.setContentsMargins(10, 8, 10, 8)
        main.setSpacing(6)

        # Pano
        main.addWidget(self.ribbon_group(
            "Pano",
            [
                ("📋", "Yapıştır", self.paste),
                ("✂", "Kes", self.cut),
                ("⧉", "Kopyala", self.copy),
            ]
        ))

        # Yazı tipi
        font_group = QFrame()
        font_group.setObjectName("ribbonGroup")
        fl = QVBoxLayout(font_group)
        fl.setContentsMargins(8, 5, 8, 5)

        ft = QLabel("Yazı Tipi")
        ft.setObjectName("groupTitle")

        top = QHBoxLayout()
        self.font_combo = QFontComboBox()
        self.font_combo.setFixedWidth(190)
        self.font_combo.currentFontChanged.connect(self.change_font_family)

        self.size_combo = QComboBox()
        self.size_combo.addItems(
            ["8", "9", "10", "11", "12", "14", "16", "18",
             "20", "22", "24", "28", "32", "36", "48", "72"]
        )
        self.size_combo.setCurrentText("12")
        self.size_combo.setFixedWidth(65)
        self.size_combo.currentTextChanged.connect(self.change_font_size)

        top.addWidget(self.font_combo)
        top.addWidget(self.size_combo)

        bottom = QHBoxLayout()
        self.boldBtn = self.ribbonTool("B", "Kalın")
        self.italicBtn = self.ribbonTool("I", "İtalik")
        self.underlineBtn = self.ribbonTool("U", "Altı Çizili")
        self.strikeBtn = self.ribbonTool("S", "Üstü Çizili")

        self.boldBtn.clicked.connect(self.make_bold)
        self.italicBtn.clicked.connect(self.make_italic)
        self.underlineBtn.clicked.connect(self.make_underline)
        self.strikeBtn.clicked.connect(self.make_strike)

        for b in (self.boldBtn, self.italicBtn, self.underlineBtn, self.strikeBtn):
            bottom.addWidget(b)

        fl.addWidget(ft)
        fl.addLayout(top)
        fl.addLayout(bottom)
        main.addWidget(font_group)

        # Renk
        main.addWidget(self.ribbon_group(
            "Renk",
            [
                ("A", "Yazı Rengi", self.Change_Text_Color),
                ("▰", "Vurgu", self.change_text_background_Color),
            ]
        ))

        # Paragraf
        para = QFrame()
        para.setObjectName("ribbonGroup")
        pl = QVBoxLayout(para)
        pl.setContentsMargins(8, 5, 8, 5)

        t = QLabel("Paragraf")
        t.setObjectName("groupTitle")

        grid = QHBoxLayout()

        self.alignLeftBtn = self.ribbonTool("≡", "Sola")
        self.alignCenterBtn = self.ribbonTool("≡", "Ortala")
        self.alignRightBtn = self.ribbonTool("≡", "Sağa")
        self.alignJustifyBtn = self.ribbonTool("≡", "İki Yana")

        self.alignLeftBtn.clicked.connect(lambda: self.set_alignment(Qt.AlignLeft))
        self.alignCenterBtn.clicked.connect(lambda: self.set_alignment(Qt.AlignCenter))
        self.alignRightBtn.clicked.connect(lambda: self.set_alignment(Qt.AlignRight))
        self.alignJustifyBtn.clicked.connect(lambda: self.set_alignment(Qt.AlignJustify))

        for b in (
            self.alignLeftBtn, self.alignCenterBtn,
            self.alignRightBtn, self.alignJustifyBtn
        ):
            grid.addWidget(b)

        lists = QHBoxLayout()
        bullet = self.ribbonTool("•", "Madde")
        number = self.ribbonTool("1.", "Numara")
        bullet.clicked.connect(self.insert_bullet_list)
        number.clicked.connect(self.insert_numbered_list)
        lists.addWidget(bullet)
        lists.addWidget(number)

        pl.addWidget(t)
        pl.addLayout(grid)
        pl.addLayout(lists)
        main.addWidget(para)

        # Düzenleme
        main.addWidget(self.ribbon_group(
            "Düzenleme",
            [
                ("↶", "Geri Al", self.undo),
                ("↷", "Yinele", self.redo),
                ("🔍", "Bul", self.find_text),
            ]
        ))

        main.addStretch()
        self.ribbonTabs.addTab(page, "Giriş")

    def create_insert_ribbon(self):
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(10, 8, 10, 8)

        layout.addWidget(self.ribbon_group(
            "Sayfalar",
            [
                ("📄", "Yeni Sayfa", self.new_page),
                ("↵", "Sayfa Sonu", self.page_break),
            ]
        ))

        layout.addWidget(self.ribbon_group(
            "Çizimler",
            [
                ("🖼", "Resim", self.insertImageFromFile),
                ("📋", "Panodan Resim", self.insertImageFromClipboard),
            ]
        ))

        layout.addWidget(self.ribbon_group(
            "Metin",
            [
                ("H1", "Başlık 1", lambda: self.apply_heading(1)),
                ("H2", "Başlık 2", lambda: self.apply_heading(2)),
                ("T", "Normal", self.apply_normal),
            ]
        ))

        layout.addStretch()
        self.ribbonTabs.addTab(page, "Ekle")



    def create_view_ribbon(self):
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(10, 8, 10, 8)

        layout.addWidget(self.ribbon_group(
            "Görünüm",
            [
                ("☀", "Açık Tema", self.apply_light_theme),
                ("🌙", "Koyu Tema", self.apply_dark_theme),
            ]
        ))

        zoom = QFrame()
        zoom.setObjectName("ribbonGroup")
        zl = QVBoxLayout(zoom)
        zl.setContentsMargins(10, 5, 10, 5)

        title = QLabel("Yakınlaştır")
        title.setObjectName("groupTitle")

        row = QHBoxLayout()
        minus = QPushButton("−")
        plus = QPushButton("+")
        self.zoomSpin = QSpinBox()
        self.zoomSpin.setRange(50, 200)
        self.zoomSpin.setValue(100)
        self.zoomSpin.setSuffix("%")

        minus.clicked.connect(lambda: self.change_zoom(-10))
        plus.clicked.connect(lambda: self.change_zoom(10))
        self.zoomSpin.valueChanged.connect(self.apply_zoom)

        row.addWidget(minus)
        row.addWidget(self.zoomSpin)
        row.addWidget(plus)

        zl.addWidget(title)
        zl.addLayout(row)
        layout.addWidget(zoom)

        layout.addStretch()
        self.ribbonTabs.addTab(page, "Görünüm")

    # ==========================================================
    # RIBBON YARDIMCILARI
    # ==========================================================
    def ribbon_group(self, title, buttons):
        frame = QFrame()
        frame.setObjectName("ribbonGroup")

        layout = QVBoxLayout(frame)
        layout.setContentsMargins(8, 5, 8, 5)

        label = QLabel(title)
        label.setObjectName("groupTitle")
        layout.addWidget(label)

        row = QHBoxLayout()
        for icon, text, func in buttons:
            row.addWidget(self.ribbonTool(icon, text, func))

        layout.addLayout(row)
        return frame

    def ribbonTool(self, icon, text, func=None):
        btn = QToolButton()
        btn.setText(f"{icon}\n{text}")
        btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        btn.setFixedSize(72, 72)

        # ribbon_group() fonksiyonundan gelen tıklama işlevini bağla.
        if func is not None:
            btn.clicked.connect(func)

        return btn

    def smallButton(self, text, tooltip):
        b = QPushButton(text)
        b.setToolTip(tooltip)
        b.setObjectName("smallButton")
        b.setFixedSize(38, 32)
        return b

    # ==========================================================
    # BELGE İŞLEMLERİ
    # ==========================================================
    def NewFile(self):
        self.TabNums += 1
        name = f"Yeni Belge {self.TabNums}"

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)

        editor = QTextEdit()
        editor.setObjectName("documentEditor")
        editor.setAcceptRichText(True)
        editor.setPlaceholderText(
            "Yaz Not'a hoş geldiniz!\n\n"
            "Belgenizi buraya yazmaya başlayabilirsiniz..."
        )

        editor.setCurrentFont(self.font_combo.currentFont())
        editor.setFontPointSize(12)

        editor.cursorPositionChanged.connect(self.updateStatus)
        editor.textChanged.connect(self.updateStatus)
        editor.installEventFilter(self)

        layout.addWidget(editor)

        index = self.tabs.addTab(page, name)
        self.tabs.setCurrentIndex(index)

        self.tabList.addItem(name)
        self.tabList.setCurrentRow(index)

        self.fileNameLabel.setText(name)
        self.updateStatus()

    def get_Active(self):
        page = self.tabs.currentWidget()
        if page:
            return page.findChild(QTextEdit)
        return None

    def on_tab_changed(self, index):
        if index < 0:
            return
        self.tabList.blockSignals(True)
        self.tabList.setCurrentRow(index)
        self.tabList.blockSignals(False)
        self.fileNameLabel.setText(self.tabs.tabText(index))
        self.updateStatus()

    def select_tab_from_list(self, index):
        if 0 <= index < self.tabs.count():
            self.tabs.setCurrentIndex(index)

    def CloseTheTab(self, index):
        widget = self.tabs.widget(index)
        if widget:
            widget.deleteLater()

        self.tabs.removeTab(index)

        if 0 <= index < self.tabList.count():
            self.tabList.takeItem(index)

        if self.tabs.count() == 0:
            self.NewFile()

    # ==========================================================
    # DOSYA
    # ==========================================================
    def OpenFile(self):
        path, _ = QFileDialog.getOpenFileName(
        self,
        "Dosya Aç",
        "",
        "Desteklenen Dosyalar (*.txt *.docx *.pdf *.rtf);;"
        "Metin Dosyaları (*.txt);;Word Dosyaları (*.docx);;"
        "PDF Dosyaları (*.pdf);;RTF Dosyaları (*.rtf)"
    )

        if not path:
            return

        editor = self.get_Active()
        if not editor:
                return

        try:
                if path.lower().endswith(".txt"):
                    with open(path, "r", encoding="utf-8") as f:
                        editor.setPlainText(f.read())

                elif path.lower().endswith(".docx") and Document:
                    doc = Document(path)
                    editor.setPlainText("\n".join(p.text for p in doc.paragraphs))

                elif path.lower().endswith(".pdf") and PyMuPDF:
                    doc = PyMuPDF.open(path)
                    text = ""
                    for page in doc:
                        text += page.get_text() + "\n"
                    editor.setPlainText(text)

                elif path.lower().endswith(".rtf"):
                    with open(path, "r", encoding="utf-8") as f:
                        rtf_content = f.read()
                        plain_text = rtf_to_text(rtf_content)
                        editor.setPlainText(plain_text)

                else:
                    QMessageBox.warning(self, "Hata", "Bu dosya formatı desteklenmiyor.")

                name = os.path.basename(path)
                self.tabs.setTabText(self.tabs.currentIndex(), name)
                self.tabList.currentItem().setText(name)
                editor.setProperty("filepath", path)
                self.fileNameLabel.setText(name)

        except Exception as e:
                QMessageBox.warning(self, "Hata", f"Dosya açılamadı:\n{e}")

    def saveAction(self):
        editor = self.get_Active()
        if not editor:
            return

        path = editor.property("filepath")
        if path:
            self.save_to_path(path, editor)
        else:
            self.SaveAsAction()

    def SaveAsAction(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Farklı Kaydet",
            "",
            "Metin Dosyası (*.txt);;Word Dosyası (*.docx);;PDF Dosyası (*.pdf);;RTF Dosyası(*.rtf)"
        )

        if not path:
            return

        editor = self.get_Active()
        if not editor:
            return

        self.save_to_path(path, editor)
        editor.setProperty("filepath", path)

        name = os.path.basename(path)
        self.tabs.setTabText(self.tabs.currentIndex(), name)
        if self.tabList.currentItem():
            self.tabList.currentItem().setText(name)
        self.fileNameLabel.setText(name)

    def save_to_path(self, path, editor):
        try:
            if path.lower().endswith(".txt"):
                with open(path, "w", encoding="utf-8") as f:
                    f.write(editor.toPlainText())

            elif path.lower().endswith(".docx") and Document:
                doc = Document()
                for block in editor.toPlainText().splitlines():
                    doc.add_paragraph(block)
                doc.save(path)

            elif path.lower().endswith(".pdf") and canvas and A4:
                c = canvas.Canvas(path, pagesize=A4)
                width, height = A4
                y = height - 50
                for line in editor.toPlainText().splitlines():
                    if y < 50:
                        c.showPage()
                        y = height - 50
                    c.drawString(50, y, line[:110])
                    y -= 16
                c.save()

            elif path.lower().endswith(".rtf"):
                with open(path, "w", encoding="utf-8") as f:
                    f.write(editor.toPlainText())

            else:
                QMessageBox.warning(
                    self,
                    "Hata",
                    "Bu dosya formatı desteklenmiyor veya gerekli kütüphane kurulu değil."
                )
                return

            self.modeLabel.setText("Kaydedildi")
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Kaydedilemedi:\n{e}")


    # ==========================================================
    # METİN BİÇİMLENDİRME
    # ==========================================================
    def make_bold(self):
        editor = self.get_Active()
        if editor:
            fmt = QTextCharFormat()
            current = editor.textCursor().charFormat().fontWeight()
            fmt.setFontWeight(QFont.Normal if current == QFont.Bold else QFont.Bold)
            editor.textCursor().mergeCharFormat(fmt)

    def make_italic(self):
        editor = self.get_Active()
        if editor:
            fmt = QTextCharFormat()
            fmt.setFontItalic(not editor.textCursor().charFormat().fontItalic())
            editor.textCursor().mergeCharFormat(fmt)

    def make_underline(self):
        editor = self.get_Active()
        if editor:
            fmt = QTextCharFormat()
            fmt.setFontUnderline(not editor.textCursor().charFormat().fontUnderline())
            editor.textCursor().mergeCharFormat(fmt)

    def make_strike(self):
        editor = self.get_Active()
        if editor:
            fmt = QTextCharFormat()
            current = editor.textCursor().charFormat().fontStrikeOut()
            fmt.setFontStrikeOut(not current)
            editor.textCursor().mergeCharFormat(fmt)

    def Change_Text_Color(self):
        editor = self.get_Active()
        if not editor:
            return

        color = QColorDialog.getColor(parent=self)
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setForeground(color)
            editor.textCursor().mergeCharFormat(fmt)

    def change_text_background_Color(self):
        editor = self.get_Active()
        if not editor:
            return

        color = QColorDialog.getColor(parent=self)
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setBackground(color)
            editor.textCursor().mergeCharFormat(fmt)

    def change_font_family(self, font):
        editor = self.get_Active()
        if editor:
            editor.setCurrentFont(font)

    def change_font_size(self, size):
        editor = self.get_Active()
        if editor:
            editor.setFontPointSize(float(size))

    def set_alignment(self, alignment):
        editor = self.get_Active()
        if editor:
            editor.setAlignment(alignment)

    def apply_heading(self, level):
        editor = self.get_Active()
        if not editor:
            return

        sizes = {1: 24, 2: 18}
        cursor = editor.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold)
        fmt.setFontPointSize(sizes[level])
        cursor.mergeCharFormat(fmt)

    def apply_normal(self):
        editor = self.get_Active()
        if editor:
            fmt = QTextCharFormat()
            fmt.setFontWeight(QFont.Normal)
            fmt.setFontPointSize(12)
            editor.textCursor().mergeCharFormat(fmt)

    def insert_bullet_list(self):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            fmt = QTextListFormat()
            fmt.setStyle(QTextListFormat.ListDisc)
            cursor.createList(fmt)

    def insert_numbered_list(self):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            fmt = QTextListFormat()
            fmt.setStyle(QTextListFormat.ListDecimal)
            cursor.createList(fmt)

    # ==========================================================
    # RESİM / EKLE
    # ==========================================================
    def insertImageFromFile(self):
        editor = self.get_Active()
        if not editor:
            return

        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Resim Seç",
            "",
            "Resim Dosyaları (*.png *.jpg *.jpeg *.bmp *.gif)"
        )

        if file_name:
            cursor = editor.textCursor()
            image_format = QTextImageFormat()
            image_format.setName(file_name)
            image_format.setWidth(450)
            cursor.insertImage(image_format)

    def insertImageFromClipboard(self):
        editor = self.get_Active()
        if editor:
            editor.paste()

    def new_page(self):
        editor = self.get_Active()
        if editor:
            editor.append("\n\n")
            self.modeLabel.setText("Yeni sayfa alanı eklendi")

    def page_break(self):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            cursor.insertText("\n\n")
            editor.append("\n──────────── SAYFA SONU ────────────\n\n")

    # ==========================================================
    # DÜZENLEME
    # ==========================================================
    def undo(self):
        editor = self.get_Active()
        if editor:
            editor.undo()

    def redo(self):
        editor = self.get_Active()
        if editor:
            editor.redo()

    def cut(self):
        editor = self.get_Active()
        if editor:
            editor.cut()

    def copy(self):
        editor = self.get_Active()
        if editor:
            editor.copy()

    def paste(self):
        editor = self.get_Active()
        if editor:
            editor.paste()

    def find_text(self):
        editor = self.get_Active()
        if not editor:
            return

        from PyQt5.QtWidgets import QInputDialog
        text, ok = QInputDialog.getText(self, "Bul", "Aranacak metin:")
        if ok and text:
            if not editor.find(text):
                QMessageBox.information(self, "Bul", "Metin bulunamadı.")

    # ==========================================================
    # GÖRÜNÜM
    # ==========================================================
    def change_zoom(self, amount):
        value = max(50, min(200, self.zoomSpin.value() + amount))
        self.zoomSpin.setValue(value)

    def apply_zoom(self, value):
        editor = self.get_Active()
        if not editor:
            return

        # QTextEdit'te zoom seviyesini font büyüklüğü üzerinden ayarla.
        zoom_ratio = value / 100

        cursor = editor.textCursor()
        fmt = cursor.charFormat()

        # Seçili metin varsa seçili metni büyüt/küçült.
        if cursor.hasSelection():
            base_size = fmt.fontPointSize()

            if base_size <= 0:
                base_size = 12

            fmt.setFontPointSize(base_size * zoom_ratio)
            cursor.mergeCharFormat(fmt)

        else:
        # Seçim yoksa tüm belge görünümünü değiştirmek için
        # QTextEdit'in zoomIn / zoomOut metodunu kullan.
            editor.zoomIn(0)

        self.zoomLabel.setText(f"{value}%")

    def apply_dark_theme(self):
        self.dark_mode = True
        self.setStyleSheet(self.dark_theme())

    def apply_light_theme(self):
        self.dark_mode = False
        self.setStyleSheet(self.light_theme())

    # ==========================================================
    # DURUM / EVENT
    # ==========================================================
    def updateStatus(self):
        editor = self.get_Active()
        if not editor:
            return

        cursor = editor.textCursor()
        self.lineLabel.setText(f"Satır: {cursor.blockNumber() + 1}")
        self.colLabel.setText(f"Sütun: {cursor.columnNumber() + 1}")
        self.charLabel.setText(f"Karakter: {len(editor.toPlainText())}")


    def eventFilter(self, obj, event):
        if isinstance(obj, QTextEdit) and event.type() == event.Wheel:
            if QApplication.keyboardModifiers() == Qt.ControlModifier:
                delta = 10 if event.angleDelta().y() > 0 else -10
                current = self.zoomSpin.value()
                self.zoomSpin.setValue(max(50, min(200, current + delta)))
                return True

        return super().eventFilter(obj, event)


    def insert_table(self, rows=2,columns=2):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            tableformat = QTextTableFormat()
            tableformat.setBorder(1)
            tableformat.setCellPadding(4)
            tableformat.setCellSpacing(2)
            cursor.insertTable(rows, columns, tableformat)



    def add_row(self,table,position=None):
        if table:
            if position is None:
                position = table.rows()
            table.insertRows(position,1)

    def add_column(self,table, position=None):
        if table:
            if position is None:
                position = table.columns()
            table.insertColumns(position, 1)

    def remove_row(self,table,position):
        if table and position < table.rows():
            table.removeRows(position,1)


    def remove_column(self, table, position):
        if table and position < table.columns():
            table.removeColumns(position, 1)

    def get_current_table(self):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            return cursor.currentTable()
        return None

    

    def remove_row_dialog(self, table):
        if table:
            row, ok = QInputDialog.getInt(self, "Satır Sil", "Satır numarası:", 0, 0, table.rows()-1)
            if ok:
                table.removeRows(row, 1)


    def insert_header(self, text="Üstbilgi"):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            cursor.movePosition(QTextCursor.Start)  # Belgenin başına git
            block_format = QTextBlockFormat()
            block_format.setAlignment(Qt.AlignCenter)
            cursor.insertBlock(block_format)
            cursor.insertText(text + "\n")
            self.modeLabel.setText("Üstbilgi eklendi")

    def insert_footer(self, text="Altbilgi"):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            cursor.movePosition(QTextCursor.End)  # Belgenin sonuna git
            block_format = QTextBlockFormat()
            block_format.setAlignment(Qt.AlignCenter)
            cursor.insertBlock(block_format)
            cursor.insertText("\n" + text)
            self.modeLabel.setText("Altbilgi eklendi")

    def insert_page_number(self):
        editor = self.get_Active()
        if editor:
            cursor = editor.textCursor()
            cursor.movePosition(QTextCursor.End)  # Belgenin sonuna git
            block_format = QTextBlockFormat()
            block_format.setAlignment(Qt.AlignRight)
            cursor.insertBlock(block_format)

            # Basit sayfa numarası (örnek: "Sayfa 1")
            page_count = self.tabs.count()
            current_index = self.tabs.currentIndex() + 1
            cursor.insertText(f"Sayfa {current_index}/{page_count}")
            self.modeLabel.setText("Sayfa numarası eklendi")


    def check_spelling(self, language="en"):
        editor = self.get_Active()
        if not editor:
            return

        text = editor.toPlainText()
        words = re.findall(r"\w+", text)

        # Dil seçimi
        if language == "tr":
            spell = SpellChecker(language=None)
            spell.word_frequency.load_text_file("turkish_words.txt")  # Türkçe sözlük dosyası ekle
        else:
            spell = SpellChecker(language="en")

        misspelled = spell.unknown(words)

        if not misspelled:
            QMessageBox.information(self, "Yazım Denetimi", "Hiç hata bulunmadı.")
            return

        corrections = []
        for word in misspelled:
            suggestion = spell.correction(word)
            corrections.append(f"{word} → {suggestion}")

        QMessageBox.information(self, "Yazım Denetimi", "\n".join(corrections))


    def insert_date(self):
        editor = self.get_Active()
        if editor:
            today = datetime.date.today().strftime("%d.%m.%Y")
            editor.insertPlainText(today)

    def insert_time(self):
        editor = self.get_Active()
        if editor:
            now = datetime.datetime.now().strftime("%H:%M:%S")
            editor.insertPlainText(now)



    def setup_autosave(self):
        self.autosave_timer = QTimer(self)
        self.autosave_timer.timeout.connect(self.autosave)
        self.autosave.timer.start(120000)


    def autosave(self):
        editor = self.get_Active()
        if not editor:
            return

        path = editor.property("filepath")
        if not path:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            path = f"autosave_{timestamp}.text"
            editor.setProperty("filepath",path)

        try:
            with open(path,"w",encoding="utf-8") as f:
                f.write(editor.toPlainText())
            self.modeLabel.setText("Otomatik kaydedildi")
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Otomatik kaydetme başarısız:\n{e}")

    def save_version(self):
        editor = self.get_Active()
        if not editor:
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        version_path = f"version_{timestamp}.txt"
        try:
            with open(version_path, "w", encoding="utf-8") as f:
                f.write(editor.toPlainText())
            self.modeLabel.setText(f"Sürüm kaydedildi: {version_path}")
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"Sürüm kaydedilemedi:\n{e}")
    # ==========================================================
    # TEMALAR
    # ==========================================================
    def light_theme(self):
        return """
        QMainWindow, QWidget {
            background: #F3F5F8;
            color: #202124;
            font-family: "Segoe UI";
        }

        #titleBar {
            background: #FFFFFF;
            border-bottom: 1px solid #D8DDE5;
        }

        #logo {
            font-size: 23px;
        }

        #appTitle {
            font-size: 22px;
            font-weight: 700;
            color: #1F4E79;
        }

        #fileNameLabel {
            color: #667085;
            font-size: 13px;
        }

        #smallButton {
            background: transparent;
            border: none;
            font-size: 18px;
            border-radius: 7px;
        }

        #smallButton:hover {
            background: #EAF2FF;
        }

        #ribbonTabs {
            background: #FFFFFF;
            border: none;
        }

        QTabBar::tab {
            background: transparent;
            padding: 8px 18px;
            border: none;
            color: #475467;
        }

        QTabBar::tab:selected {
            color: #185ABD;
            border-bottom: 3px solid #185ABD;
            font-weight: 600;
        }

        #ribbonGroup {
            background: #FFFFFF;
            border-right: 1px solid #E2E6EC;
            border-radius: 5px;
        }

        #groupTitle {
            color: #667085;
            font-size: 11px;
            font-weight: 600;
        }

        QToolButton {
            background: transparent;
            border: 1px solid transparent;
            color: #344054;
            font-size: 11px;
        }

        QToolButton:hover {
            background: #EAF2FF;
            border: 1px solid #C7DBFF;
        }

        QComboBox, QFontComboBox, QSpinBox {
            min-height: 24px;
            max-height: 24px;
            font-size: 9pt;
            padding: 2px 6px;
            background: #FFFFFF;
            border: 1px solid #C9D1DB;
            border-radius: 4px;
            color: #202124;
        }
        QComboBox:hover, QFontComboBox:hover, QSpinBox:hover {
            background: #F0F6FF;
            border: 1px solid #A0C4FF;
        }
        QComboBox::drop-down, QFontComboBox::drop-down {
            width: 18px;
            border-left: 1px solid #C0C0C0;
            background: #E0E0E0;
        }
        QComboBox::down-arrow, QFontComboBox::down-arrow {
            image: url(:/qt-project.org/styles/commonstyle/images/arrowdown.png);
            width: 10px;
            height: 10px;
        }

        #statusBar {
            background: #FFFFFF;
            color: #667085;
            border-top: 1px solid #D8DDE5;
        }
        """

    def dark_theme(self):
        return """
        QMainWindow, QWidget {
            background: #15171A;
            color: #E6E8EB;
            font-family: "Segoe UI";
        }

        #titleBar {
            background: #1D2024;
            border-bottom: 1px solid #30343A;
        }

        #logo {
            font-size: 23px;
        }

        #appTitle {
            font-size: 22px;
            font-weight: 700;
            color: #75A7FF;
        }

        #fileNameLabel {
            color: #98A2B3;
        }

        #smallButton {
            background: transparent;
            border: none;
            color: #E6E8EB;
            font-size: 18px;
            border-radius: 7px;
        }

        #smallButton:hover {
            background: #2B3139;
        }

        #ribbonTabs {
            background: #1D2024;
            border: none;
        }

        QTabBar::tab {
            background: transparent;
            padding: 8px 18px;
            border: none;
            color: #A8B0BB;
        }

        QTabBar::tab:selected {
            color: #75A7FF;
            border-bottom: 3px solid #75A7FF;
            font-weight: 600;
        }

        #ribbonGroup {
            background: #1D2024;
            border-right: 1px solid #30343A;
        }

        #groupTitle {
            color: #98A2B3;
            font-size: 11px;
            font-weight: 600;
        }

        QToolButton {
            background: transparent;
            border: 1px solid transparent;
            border-radius: 5px;
            color: #E6E8EB;
            font-size: 11px;
        }

        QToolButton:hover {
            background: #2B3139;
            border: 1px solid #424952;
        }

        QComboBox, QFontComboBox, QSpinBox {
            min-height: 24px;
            max-height: 24px;
            font-size: 9pt;
            padding: 2px 6px;
            background: #252A30;
            border: 1px solid #424952;
            border-radius: 4px;
            color: #E6E8EB;
        }
        QComboBox:hover, QFontComboBox:hover, QSpinBox:hover {
            background: #2B3139;
            border: 1px solid #75A7FF;
        }
        QComboBox::drop-down, QFontComboBox::drop-down {
            width: 18px;
            border-left: 1px solid #424952;
            background: #30343A;
        }
        QComboBox::down-arrow, QFontComboBox::down-arrow {
            image: url(:/qt-project.org/styles/commonstyle/images/arrowdown.png);
            width: 10px;
            height: 10px;
        }

        #statusBar {
            background: #1D2024;
            color: #98A2B3;
            border-top: 1px solid #30343A;
        }
        """

    def apply_light_theme(self):
        self.dark_mode = False
        self.setStyleSheet(self.light_theme())

    def apply_dark_theme(self):
        self.dark_mode = True
        self.setStyleSheet(self.dark_theme())



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = YazNot()
    if not window.check_license():
        sys.exit()
    window.show()
    sys.exit(app.exec_())