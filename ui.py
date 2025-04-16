import os
import json
from datetime import datetime, timedelta
from collections import defaultdict
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QPushButton, QLabel, QLineEdit, QFileDialog, QTreeWidget, 
                           QTreeWidgetItem, QMessageBox, QProgressBar, QComboBox, 
                           QFrame, QToolBar, QDialog, QTabWidget, QStyle, QSplashScreen, QGridLayout, QProgressDialog, QTextEdit)
from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QIcon, QAction, QColor, QPixmap, QFont
from PySide6.QtWidgets import QApplication

from scanner import FileScanner
from visualization import Visualization
from utils import format_size, parse_size, calculate_file_hash, get_file_categories, get_file_type_extensions

class ScanCache:
    def __init__(self, cache_file="scan_cache.json"):
        self.cache_file = cache_file
        self.cache = self.load_cache()
        
    def load_cache(self):
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception:
            return {}
            
    def save_cache(self):
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Fehler beim Speichern des Caches: {str(e)}")
            
    def get_cache_key(self, drive_path, years, file_types, owner_filter, size_filter):
        return f"{drive_path}_{years}_{str(file_types)}_{owner_filter}_{size_filter}"
        
    def get_cached_results(self, drive_path, years, file_types, owner_filter, size_filter):
        cache_key = self.get_cache_key(drive_path, years, file_types, owner_filter, size_filter)
        if cache_key in self.cache:
            cache_data = self.cache[cache_key]
            # Prüfe, ob der Cache noch gültig ist (maximal 24 Stunden alt)
            cache_time = datetime.fromisoformat(cache_data['timestamp'])
            if datetime.now() - cache_time < timedelta(hours=24):
                return cache_data['results']
        return None
        
    def cache_results(self, drive_path, years, file_types, owner_filter, size_filter, results):
        cache_key = self.get_cache_key(drive_path, years, file_types, owner_filter, size_filter)
        self.cache[cache_key] = {
            'timestamp': datetime.now().isoformat(),
            'results': results
        }
        self.save_cache()

class SplashScreen(QSplashScreen):
    def __init__(self):
        # Erstelle ein Pixmap für den Splashscreen
        pixmap = QPixmap(400, 200)
        pixmap.fill(Qt.GlobalColor.white)
        super().__init__(pixmap)
        
        # Layout erstellen
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Icon
        icon_label = QLabel(self)
        icon_label.setPixmap(QIcon("clean_drive.ico").pixmap(64, 64))
        icon_label.setGeometry(168, 20, 64, 64)
        
        # Titel
        title = QLabel("Laufwerk Bereiniger", self)
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setGeometry(0, 90, 400, 30)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: black;")
        
        # Version
        version = QLabel("Version 1.0", self)
        version.setFont(QFont("Arial", 10))
        version.setGeometry(0, 120, 400, 20)
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version.setStyleSheet("color: black;")
        
        # Fortschrittsbalken
        self.progress = QProgressBar(self)
        self.progress.setGeometry(50, 150, 300, 5)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #f1f3f4;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #4285f4;
                border-radius: 2px;
            }
        """)
        
        # Status-Text
        self.status = QLabel("Wird geladen...", self)
        self.status.setFont(QFont("Arial", 9))
        self.status.setGeometry(0, 160, 400, 20)
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status.setStyleSheet("color: black;")
        
        self.progress_value = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_progress)
        self.timer.start(30)  # Alle 30ms aktualisieren
        
    def update_progress(self):
        self.progress_value = min(100, self.progress_value + 2)
        self.progress.setValue(self.progress_value)
        
        if self.progress_value >= 100:
            self.timer.stop()
            
        # Aktualisiere den Status-Text basierend auf dem Fortschritt
        if self.progress_value < 30:
            self.status.setText("Initialisiere Anwendung...")
        elif self.progress_value < 60:
            self.status.setText("Lade Komponenten...")
        elif self.progress_value < 90:
            self.status.setText("Bereite Benutzeroberfläche vor...")
        else:
            self.status.setText("Starte Anwendung...")

class DriveCleanerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Laufwerk Bereiniger")
        self.setGeometry(100, 100, 1200, 800)
        
        # App-Icon setzen
        app_icon = QIcon("clean_drive.ico")
        self.setWindowIcon(app_icon)
        # Setze das Icon explizit für die Anwendung
        app = QApplication.instance()
        app.setWindowIcon(app_icon)
        
        # Cache initialisieren
        self.cache = ScanCache()
        
        self.setup_ui()
        self.setup_toolbar()
        self.apply_styles()
        
    def apply_styles(self):
        self.setStyleSheet("""
            * {
                color: black;
                background-color: white;
            }
            QMainWindow {
                background-color: #f5f5f5;
            }
            QToolBar {
                background-color: #ffffff;
                border-bottom: 1px solid #e0e0e0;
                padding: 5px;
            }
            QFrame {
                background-color: #ffffff;
                border: none;
            }
            QLineEdit {
                padding: 8px;
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 2px;
            }
            QPushButton {
                background-color: #4285f4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 2px;
            }
            QPushButton:hover {
                background-color: #3275e4;
            }
            QPushButton:pressed {
                background-color: #2265d4;
            }
            QTreeWidget {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 2px;
            }
            QTreeWidget::item {
                padding: 5px;
                color: black;
            }
            QTreeWidget::item:selected {
                background-color: #e8f0fe;
                color: black;
            }
            QTreeWidget::item:hover {
                background-color: #f8f9fa;
            }
            QTreeWidget QHeaderView::section {
                background-color: #f8f9fa;
                padding: 5px;
                border: none;
                border-right: 1px solid #e0e0e0;
                border-bottom: 1px solid #e0e0e0;
                color: black;
            }
            QProgressBar {
                border: none;
                background-color: #f1f3f4;
                border-radius: 2px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4285f4;
                border-radius: 2px;
            }
            QLabel {
                color: black;
            }
            QComboBox {
                padding: 8px;
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 2px;
                color: black;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #666;
                margin-right: 5px;
            }
            QDialog {
                background-color: white;
            }
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f8f9fa;
                border: 1px solid #e0e0e0;
                padding: 8px 16px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: none;
            }
            QTabBar::tab:hover {
                background-color: #e8f0fe;
            }
            QMessageBox {
                background-color: white;
            }
            QMessageBox QLabel {
                color: black;
            }
            QMessageBox QPushButton {
                background-color: #4285f4;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 2px;
            }
            QMessageBox QPushButton:hover {
                background-color: #3275e4;
            }
        """)

    def setup_toolbar(self):
        toolbar = QToolBar()
        toolbar.setIconSize(QSize(24, 24))
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        # Neu Button
        new_action = QAction("Neu", self)
        new_action.triggered.connect(self.new_scan)
        toolbar.addAction(new_action)

        # Speichern Button
        save_action = QAction("Speichern", self)
        save_action.triggered.connect(self.save_results)
        toolbar.addAction(save_action)

        # Öffnen Button
        open_action = QAction("Öffnen", self)
        open_action.triggered.connect(self.load_results)
        toolbar.addAction(open_action)

        # Excel Export Button
        excel_action = QAction("Excel Export", self)
        excel_action.triggered.connect(self.export_to_excel)
        toolbar.addAction(excel_action)

        toolbar.addSeparator()

        # Löschen Button
        delete_action = QAction("Löschen", self)
        delete_action.triggered.connect(self.delete_selected)
        toolbar.addAction(delete_action)

        # Massenlöschung Button
        mass_delete_action = QAction("Massenlöschung", self)
        mass_delete_action.triggered.connect(self.show_mass_delete_dialog)
        toolbar.addAction(mass_delete_action)

        toolbar.addSeparator()

        # Visualisierung Button
        visualize_action = QAction("Visualisieren", self)
        visualize_action.triggered.connect(self.visualize_data)
        toolbar.addAction(visualize_action)

        # Duplikate Button
        duplicates_action = QAction("Duplikate finden", self)
        duplicates_action.triggered.connect(self.find_duplicates)
        toolbar.addAction(duplicates_action)

        # Ungenutzte Dateien Button
        unused_action = QAction("Ungenutzte Dateien", self)
        unused_action.triggered.connect(self.find_unused_files)
        toolbar.addAction(unused_action)

        # Kategorisierung Button
        categorize_action = QAction("Kategorisieren", self)
        categorize_action.triggered.connect(self.show_categories)
        toolbar.addAction(categorize_action)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)

        # Laufwerk-Auswahl
        search_frame = QFrame()
        search_layout = QHBoxLayout(search_frame)
        search_layout.setContentsMargins(10, 10, 10, 10)
        
        laufwerk_label = QLabel("Laufwerk:")
        self.drive_input = QLineEdit()
        self.drive_input.setPlaceholderText("Laufwerk auswählen...")
        self.drive_input.setReadOnly(True)
        
        durchsuchen_button = QPushButton("📂 Durchsuchen")
        durchsuchen_button.clicked.connect(self.browse_drive)
        
        search_layout.addWidget(laufwerk_label)
        search_layout.addWidget(self.drive_input, 1)
        search_layout.addWidget(durchsuchen_button)

        # Filter-Bereich
        filter_frame = QFrame()
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(10, 10, 10, 10)
        
        # Alter Filter
        self.years_input = QLineEdit()
        self.years_input.setPlaceholderText("Jahre")
        self.years_input.setFixedWidth(100)
        
        # Dateityp Filter
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItem("Alle Dateitypen")
        self.file_type_combo.addItem("Bilder (.jpg, .jpeg, .png, .gif)")
        self.file_type_combo.addItem("Office-Dokumente")
        self.file_type_combo.addItem("PDF-Dateien (.pdf)")
        self.file_type_combo.addItem("Videos (.mp4, .avi, .mov)")
        self.file_type_combo.addItem("Audio (.mp3, .wav, .ogg)")
        self.file_type_combo.addItem("Archive (.zip, .rar, .7z)")
        self.file_type_combo.addItem("Benutzerdefiniert")
        self.file_type_combo.setFixedWidth(200)
        
        # Tooltip für Office-Dokumente
        self.file_type_combo.setItemData(2, 
            "Enthält: Word (.doc, .docx)\nExcel (.xls, .xlsx)\n"
            "PowerPoint (.ppt, .pptx)\nAccess (.accdb, .mdb)", 
            Qt.ItemDataRole.ToolTipRole)
            
        # Benutzerdefinierter Dateityp Eingabefeld
        self.custom_file_types = QLineEdit()
        self.custom_file_types.setPlaceholderText("Dateitypen eingeben (z.B. .txt,.csv,.log)")
        self.custom_file_types.setToolTip("Mehrere Dateitypen mit Komma trennen (z.B. .txt,.csv,.log)")
        self.custom_file_types.setVisible(False)
        self.custom_file_types.setFixedWidth(200)
        
        # Verbinde ComboBox-Änderung mit Sichtbarkeit des Eingabefelds
        self.file_type_combo.currentTextChanged.connect(self.toggle_custom_file_types)
        
        # Größenfilter
        self.size_filter_combo = QComboBox()
        self.size_filter_combo.addItem("Alle Größen")
        self.size_filter_combo.addItem("Kleine Dateien (0-10MB)")
        self.size_filter_combo.addItem("Mittlere Dateien (10-100MB)")
        self.size_filter_combo.addItem("Große Dateien (>100MB)")
        self.size_filter_combo.setFixedWidth(150)
        
        # Ersteller Filter
        self.owner_input = QLineEdit()
        self.owner_input.setPlaceholderText("Ersteller suchen (z.B. 'leon admin' findet 'DESKTOP-123\\Leon-Admin')")
        self.owner_input.setFixedWidth(200)
        self.owner_input.setToolTip("Groß-/Kleinschreibung wird ignoriert.\nMehrere Suchbegriffe möglich (z.B. 'leon admin').\nAlle Begriffe müssen im Namen vorkommen.")
        
        # Füge die Filter zum Layout hinzu
        filter_layout.addWidget(QLabel("Älter als (Jahre):"))
        filter_layout.addWidget(self.years_input)
        filter_layout.addWidget(QLabel("Dateityp:"))
        filter_layout.addWidget(self.file_type_combo)
        filter_layout.addWidget(self.custom_file_types)
        filter_layout.addWidget(QLabel("Größe:"))
        filter_layout.addWidget(self.size_filter_combo)
        filter_layout.addWidget(QLabel("Ersteller:"))
        filter_layout.addWidget(self.owner_input)
        
        # Füge einen Stretch am Ende hinzu, um die Filter nach links zu drücken
        filter_layout.addStretch()

        # Scan-Button und Kontroll-Buttons in einem Frame
        control_frame = QFrame()
        control_layout = QHBoxLayout(control_frame)
        control_layout.setContentsMargins(0, 0, 0, 0)
        
        # Scan-Button
        self.scan_button = QPushButton("🔍 Scannen")
        self.scan_button.clicked.connect(self.start_scan)
        self.scan_button.setFixedHeight(40)
        
        # Pause-Button
        self.pause_button = QPushButton("⏸️ Pause")
        self.pause_button.clicked.connect(self.toggle_pause_scan)
        self.pause_button.setFixedHeight(40)
        self.pause_button.setEnabled(False)
        
        # Abbruch-Button
        self.abort_button = QPushButton("⏹️ Abbrechen")
        self.abort_button.clicked.connect(self.abort_scan)
        self.abort_button.setFixedHeight(40)
        self.abort_button.setEnabled(False)
        
        control_layout.addWidget(self.scan_button)
        control_layout.addWidget(self.pause_button)
        control_layout.addWidget(self.abort_button)

        # Progress Bar und Status
        progress_frame = QFrame()
        progress_layout = QVBoxLayout(progress_frame)
        progress_layout.setContentsMargins(0, 0, 0, 0)
        
        # Fortschrittsanzeige für die Dateisammlung
        self.collection_progress = QProgressBar()
        self.collection_progress.setVisible(False)
        self.collection_progress.setFixedHeight(5)
        self.collection_progress.setTextVisible(False)
        self.collection_progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #f1f3f4;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #ffd700;
                border-radius: 2px;
            }
        """)
        
        # Fortschrittsanzeige für die Verarbeitung
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setTextVisible(False)
        
        progress_layout.addWidget(self.collection_progress)
        progress_layout.addWidget(self.progress_bar)

        # Dateiliste
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabels(["Dateipfad", "Größe", "Datum", "Typ", "Ersteller"])
        self.file_tree.setColumnWidth(0, 500)
        self.file_tree.setColumnWidth(1, 100)
        self.file_tree.setColumnWidth(2, 150)
        self.file_tree.setColumnWidth(3, 100)
        self.file_tree.setColumnWidth(4, 200)
        self.file_tree.setAlternatingRowColors(True)
        self.file_tree.setSortingEnabled(True)
        self.file_tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self.file_tree.setSelectionBehavior(QTreeWidget.SelectionBehavior.SelectRows)
        self.file_tree.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.file_tree.setStyleSheet("""
            QTreeWidget {
                selection-background-color: #e8f0fe;
                selection-color: black;
            }
            QTreeWidget::item:selected {
                background-color: #e8f0fe;
                color: black;
            }
            QTreeWidget::item:hover {
                background-color: #f8f9fa;
            }
        """)
        # Verbinde Doppelklick mit Kopier-Funktion
        self.file_tree.itemDoubleClicked.connect(self.copy_path_to_clipboard)

        # Aktionsbuttons
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 0, 0, 0)
        
        delete_selected = QPushButton("🗑️ Ausgewählte löschen")
        delete_all = QPushButton("🗑️ Alle löschen")
        delete_selected.clicked.connect(self.delete_selected)
        delete_all.clicked.connect(self.delete_all)
        
        button_layout.addWidget(delete_selected)
        button_layout.addWidget(delete_all)

        # Statusanzeige hinzufügen
        self.status_label = QLabel("Bereit")
        self.status_label.setStyleSheet("color: black; padding: 5px;")
        
        # Footer mit Entwicklerinformationen
        footer_frame = QFrame()
        footer_frame.setStyleSheet("""
            QFrame {
                background-color: #f5f5f5;
                border-top: 1px solid #e0e0e0;
                padding: 5px;
            }
            QLabel {
                color: #666666;
                font-size: 9pt;
            }
        """)
        footer_layout = QHBoxLayout(footer_frame)
        footer_layout.setContentsMargins(10, 5, 10, 5)
        
        developer_info = QLabel("Entwickler: Leon Stolz")
        footer_layout.addWidget(developer_info)
        
        # Layout zusammenfügen
        layout.addWidget(search_frame)
        layout.addWidget(filter_frame)
        layout.addWidget(control_frame)  # Control-Frame mit Scan-, Pause- und Abbruch-Button
        layout.addWidget(progress_frame)
        layout.addWidget(self.file_tree)
        layout.addWidget(button_frame)
        layout.addWidget(self.status_label)
        layout.addWidget(footer_frame)

        self.scanner = None
        self.file_types = get_file_type_extensions()
        self.is_paused = False

    def toggle_custom_file_types(self, text):
        """Zeigt oder versteckt das Eingabefeld für benutzerdefinierte Dateitypen"""
        self.custom_file_types.setVisible(text == "Benutzerdefiniert")

    def get_selected_file_types(self):
        selected = self.file_type_combo.currentText()
        if selected == "Alle Dateitypen":
            return None
        elif selected == "Benutzerdefiniert":
            # Verarbeite benutzerdefinierte Dateitypen
            custom_types = self.custom_file_types.text().strip()
            if not custom_types:
                return None
            # Entferne Leerzeichen und konvertiere zu Kleinbuchstaben
            return [ext.strip().lower() for ext in custom_types.split(',')]
        else:
            return self.file_types.get(selected, [])

    def browse_drive(self):
        drive = QFileDialog.getExistingDirectory(self, "Laufwerk auswählen")
        if drive:
            self.drive_input.setText(drive)

    def toggle_pause_scan(self):
        """Pausiert oder setzt den Scan fort"""
        if not self.scanner:
            return
            
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.scanner.pause_scan = True
            self.pause_button.setText("▶️ Fortsetzen")
            self.status_label.setText("Scan pausiert")
            self.scan_button.setEnabled(False)  # Deaktiviere Scan-Button während Pause
        else:
            self.scanner.pause_scan = False
            self.pause_button.setText("⏸️ Pause")
            self.status_label.setText("Scan wird fortgesetzt...")
            self.scan_button.setEnabled(False)  # Bleibe deaktiviert während des Scans

    def abort_scan(self):
        """Bricht den Scan ab"""
        if not self.scanner:
            return
            
        reply = QMessageBox.question(
            self,
            "Scan abbrechen",
            "Möchten Sie den Scan wirklich abbrechen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.scanner.stop_scan = True
            self.scanner.pause_scan = False  # Stelle sicher, dass der Scan nicht pausiert ist
            self.status_label.setText("Scan wird abgebrochen...")
            
            # Deaktiviere alle Buttons während des Abbruchs
            self.pause_button.setEnabled(False)
            self.abort_button.setEnabled(False)
            self.scan_button.setEnabled(False)
            
            # Warte auf das Ende des Scans
            if self.scanner.isRunning():
                self.scanner.wait()
            
            # Setze UI zurück
            self.reset_scan_ui()

    def reset_scan_ui(self):
        """Setzt die UI nach einem abgebrochenen oder abgeschlossenen Scan zurück"""
        self.pause_button.setEnabled(False)
        self.abort_button.setEnabled(False)
        self.scan_button.setEnabled(True)
        self.pause_button.setText("⏸️ Pause")
        self.is_paused = False
        self.progress_bar.setVisible(False)
        self.collection_progress.setVisible(False)
        self.status_label.setText("Bereit")

    def start_scan(self):
        drive = self.drive_input.text()
        years = 0  # Standardwert auf 0 setzen
        try:
            years_text = self.years_input.text().strip()
            if years_text:  # Nur konvertieren wenn ein Wert eingegeben wurde
                years = int(years_text)
        except ValueError:
            QMessageBox.warning(self, "Fehler", "Bitte geben Sie eine gültige Jahresanzahl ein oder lassen Sie das Feld leer.")
            return
            
        if not drive:
            QMessageBox.warning(self, "Fehler", "Bitte wählen Sie ein Laufwerk aus.")
            return
            
        # Setze UI zurück
        self.file_tree.clear()
        self.collection_progress.setVisible(True)
        self.collection_progress.setValue(0)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Initialisiere Scan...")
        
        # Aktiviere Kontroll-Buttons
        self.pause_button.setEnabled(True)
        self.abort_button.setEnabled(True)
        self.scan_button.setEnabled(False)
        self.pause_button.setText("⏸️ Pause")
        self.is_paused = False
        
        file_types = self.get_selected_file_types()
        owner_filter = self.owner_input.text().strip() or None
        
        # Größenfilter ermitteln
        size_filter = None
        size_text = self.size_filter_combo.currentText()
        if size_text != "Alle Größen":
            size_filter = size_text.split(" (")[0]  # Extrahiere nur den Namen ohne Größenangabe
            
        # Prüfe Cache
        cached_results = self.cache.get_cached_results(drive, years, file_types, owner_filter, size_filter)
        if cached_results:
            self.status_label.setText("Lade Ergebnisse aus Cache...")
            for result in cached_results:
                self.add_file_to_tree(*result)
            self.status_label.setText("Cache geladen")
            self.reset_scan_ui()
            return
        
        # Erstelle neuen Scanner
        self.scanner = FileScanner(drive, years, file_types, owner_filter, size_filter)
        self.scanner.file_found.connect(self.add_file_to_tree)
        self.scanner.progress_update.connect(self.update_progress)
        self.scanner.scan_complete.connect(self.scan_completed)
        self.scanner.status_update.connect(self.update_status)
        self.scanner.collection_progress.connect(self.update_collection_progress)
        self.scanner.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_collection_progress(self, current, total):
        if total > 0:
            progress = int((current / total) * 100)
            self.collection_progress.setValue(progress)
            # Kürze die Statusmeldung
            status_text = f"Sammle Dateien... ({current:,} von geschätzt {total:,})"
            if len(status_text) > 100:
                status_text = f"Sammle... ({current:,}/{total:,})"
            self.status_label.setText(status_text)

    def scan_completed(self, total_size_gb, file_count):
        self.reset_scan_ui()
        
        # Konvertiere GB zurück zu Bytes für die Anzeige
        total_size_bytes = total_size_gb * 1024 * 1024 * 1024
        
        # Sammle Ergebnisse für den Cache
        results = []
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            results.append([
                item.text(0),  # Pfad
                item.text(1),  # Größe
                item.text(2),  # Datum
                item.text(3),  # Typ
                item.text(4)   # Ersteller
            ])
            
        # Cache die Ergebnisse
        drive = self.drive_input.text()
        years = self.years_input.text().strip() or None
        file_types = self.get_selected_file_types()
        owner_filter = self.owner_input.text().strip() or None
        size_filter = self.size_filter_combo.currentText().split(" (")[0] if self.size_filter_combo.currentText() != "Alle Größen" else None
        
        self.cache.cache_results(drive, years, file_types, owner_filter, size_filter, results)
        
        QMessageBox.information(self, "Scan abgeschlossen", 
                              f"Es wurden {file_count:,} Dateien mit einer Gesamtgröße von {format_size(total_size_bytes)} gefunden.")

    def add_file_to_tree(self, path, size, date, file_type, owner):
        item = QTreeWidgetItem(self.file_tree)
        item.setText(0, path)
        item.setText(1, size)
        item.setText(2, date)
        item.setText(3, file_type)
        item.setText(4, owner)

    def delete_selected(self):
        selected_items = self.file_tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Fehler", "Bitte wählen Sie Dateien zum Löschen aus.")
            return
            
        reply = QMessageBox.question(
            self,
            "Bestätigung",
            f"Möchten Sie {len(selected_items)} ausgewählte Dateien löschen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            for item in selected_items:
                try:
                    os.remove(item.text(0))
                    self.file_tree.takeTopLevelItem(self.file_tree.indexOfTopLevelItem(item))
                except Exception as e:
                    QMessageBox.warning(self, "Fehler", f"Fehler beim Löschen von {item.text(0)}: {str(e)}")

    def delete_all(self):
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Dateien zum Löschen vorhanden.")
            return
            
        reply = QMessageBox.question(
            self,
            "Bestätigung",
            "Möchten Sie wirklich alle Dateien löschen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            for i in range(self.file_tree.topLevelItemCount()):
                item = self.file_tree.topLevelItem(i)
                try:
                    os.remove(item.text(0))
                except Exception as e:
                    QMessageBox.warning(self, "Fehler", f"Fehler beim Löschen von {item.text(0)}: {str(e)}")
            
            self.file_tree.clear()

    def update_status(self, message):
        # Kürze lange Pfade in der Statusmeldung
        if len(message) > 100:
            parts = message.split(": ")
            if len(parts) > 1:
                path = parts[1]
                if len(path) > 50:
                    # Kürze den Pfad in der Mitte
                    shortened_path = path[:20] + "..." + path[-20:]
                    message = f"{parts[0]}: {shortened_path}"
        self.status_label.setText(message)

    def new_scan(self):
        reply = QMessageBox.question(
            self,
            "Neuer Scan",
            "Möchten Sie wirklich einen neuen Scan starten? Alle nicht gespeicherten Ergebnisse gehen verloren.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.file_tree.clear()
            self.drive_input.clear()
            self.years_input.clear()
            self.file_type_combo.setCurrentIndex(0)
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(False)
            self.status_label.setText("Bereit")
            if self.scanner and self.scanner.isRunning():
                self.scanner.stop_scan = True
                self.scanner.wait()
            self.scanner = None

    def save_results(self):
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Scan-Ergebnisse zum Speichern vorhanden.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Ergebnisse speichern",
            "",
            "JSON-Dateien (*.json)"
        )

        if file_path:
            try:
                data = {
                    "scan_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "drive_path": self.drive_input.text(),
                    "years": self.years_input.text(),
                    "file_type": self.file_type_combo.currentText(),
                    "files": []
                }

                for i in range(self.file_tree.topLevelItemCount()):
                    item = self.file_tree.topLevelItem(i)
                    data["files"].append({
                        "path": item.text(0),
                        "size": item.text(1),
                        "date": item.text(2),
                        "type": item.text(3),
                        "owner": item.text(4)
                    })

                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)

                self.status_label.setText("Ergebnisse wurden gespeichert")
            except Exception as e:
                QMessageBox.critical(self, "Fehler", f"Fehler beim Speichern: {str(e)}")

    def load_results(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Ergebnisse laden",
            "",
            "JSON-Dateien (*.json)"
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                self.file_tree.clear()
                self.drive_input.setText(data.get("drive_path", ""))
                self.years_input.setText(data.get("years", ""))
                
                file_type = data.get("file_type", "Alle Dateitypen")
                index = self.file_type_combo.findText(file_type)
                if index >= 0:
                    self.file_type_combo.setCurrentIndex(index)

                for file_info in data["files"]:
                    item = QTreeWidgetItem(self.file_tree)
                    item.setText(0, file_info["path"])
                    item.setText(1, file_info["size"])
                    item.setText(2, file_info["date"])
                    item.setText(3, file_info["type"])
                    item.setText(4, file_info["owner"])

                self.status_label.setText(f"Ergebnisse vom {data.get('scan_date', 'unbekannt')} geladen")
            except Exception as e:
                QMessageBox.critical(self, "Fehler", f"Fehler beim Laden: {str(e)}")

    def visualize_data(self):
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Daten zur Visualisierung vorhanden.")
            return

        visualization = Visualization(self.file_tree, format_size, parse_size)
        if not visualization.visualize_data():
            QMessageBox.warning(self, "Fehler", "Keine Daten zur Visualisierung vorhanden.")

    def find_duplicates(self):
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Dateien zum Analysieren vorhanden.")
            return

        self.status_label.setText("Suche nach Duplikaten...")
        
        # Erstelle Fortschrittsdialog
        progress_dialog = QProgressDialog(
            "Suche nach Duplikaten...", 
            "Abbrechen", 
            0, 
            self.file_tree.topLevelItemCount(), 
            self
        )
        progress_dialog.setWindowTitle("Duplikatsuche")
        progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        
        size_dict = defaultdict(list)
        hash_dict = defaultdict(list)
        
        # Erst nach Größe gruppieren
        for i in range(self.file_tree.topLevelItemCount()):
            if progress_dialog.wasCanceled():
                self.status_label.setText("Duplikatsuche abgebrochen")
                return
                
            item = self.file_tree.topLevelItem(i)
            filepath = item.text(0)
            try:
                size = os.path.getsize(filepath)
                size_dict[size].append(filepath)
            except Exception:
                continue
                
            progress_dialog.setValue(i + 1)
            QApplication.processEvents()  # Verhindert, dass die GUI einfriert

        # Dann nach Hash für Dateien gleicher Größe
        current_progress = 0
        total_files = sum(len(files) for files in size_dict.values() if len(files) > 1)
        progress_dialog.setMaximum(total_files)
        progress_dialog.setValue(0)
        
        for size, filepaths in size_dict.items():
            if len(filepaths) > 1:  # Nur Dateien mit gleicher Größe prüfen
                for filepath in filepaths:
                    if progress_dialog.wasCanceled():
                        self.status_label.setText("Duplikatsuche abgebrochen")
                        return
                        
                    try:
                        # Schnelle Hash-Berechnung für große Dateien
                        file_hash = calculate_file_hash(filepath, quick_mode=True)
                        if file_hash:
                            hash_dict[file_hash].append(filepath)
                    except Exception:
                        continue
                        
                    current_progress += 1
                    progress_dialog.setValue(current_progress)
                    QApplication.processEvents()

        progress_dialog.close()

        # Zeige Duplikate an
        duplicates = {k: v for k, v in hash_dict.items() if len(v) > 1}
        if duplicates:
            self.show_duplicates_dialog(duplicates)
        else:
            QMessageBox.information(self, "Ergebnis", "Keine Duplikate gefunden.")
        
        self.status_label.setText("Duplikatsuche abgeschlossen")

    def show_duplicates_dialog(self, duplicates):
        dialog = QDialog(self)
        dialog.setWindowTitle("Gefundene Duplikate")
        dialog.setMinimumSize(1000, 700)
        
        layout = QVBoxLayout(dialog)
        
        # Erstelle Tabs für verschiedene Ansichten
        tab_widget = QTabWidget()
        
        # Übersichts-Tab
        overview_tab = QWidget()
        overview_layout = QVBoxLayout(overview_tab)
        
        # Quick-Delete Button für alle Duplikate
        quick_delete_frame = QFrame()
        quick_delete_layout = QHBoxLayout(quick_delete_frame)
        
        total_duplicates = sum(len(files) - 1 for files in duplicates.values())
        quick_delete_button = QPushButton(f"🗑️ Alle Duplikate löschen ({total_duplicates} Dateien)")
        quick_delete_button.clicked.connect(lambda: self.quick_delete_duplicates(duplicates, dialog))
        quick_delete_layout.addWidget(quick_delete_button)
        quick_delete_layout.addStretch()
        
        overview_layout.addWidget(quick_delete_frame)
        
        # Übersichtstabelle
        overview_tree = QTreeWidget()
        overview_tree.setHeaderLabels(["Duplikatgruppe", "Anzahl", "Gesamtgröße"])
        overview_tree.setColumnWidth(0, 400)
        
        total_size = 0
        for hash_value, filepaths in duplicates.items():
            group_item = QTreeWidgetItem(overview_tree)
            group_item.setText(0, f"Duplikatgruppe {overview_tree.topLevelItemCount() + 1}")
            group_item.setText(1, str(len(filepaths)))
            
            # Berechne Gesamtgröße der Gruppe
            group_size = 0
            for filepath in filepaths:
                try:
                    stats = os.stat(filepath)
                    group_size += stats.st_size
                except Exception:
                    pass
            
            group_item.setText(2, format_size(group_size))
            total_size += group_size
        
        # Füge Gesamtsumme hinzu
        total_item = QTreeWidgetItem(overview_tree)
        total_item.setText(0, "Gesamt")
        total_item.setText(1, str(sum(len(files) for files in duplicates.values())))
        total_item.setText(2, format_size(total_size))
        total_item.setBackground(0, QColor("#f0f0f0"))
        total_item.setBackground(1, QColor("#f0f0f0"))
        total_item.setBackground(2, QColor("#f0f0f0"))
        
        overview_layout.addWidget(overview_tree)
        tab_widget.addTab(overview_tab, "Übersicht")
        
        # Details-Tab
        details_tab = QWidget()
        details_layout = QVBoxLayout(details_tab)
        
        # Details-Tabelle
        details_tree = QTreeWidget()
        details_tree.setHeaderLabels(["Dateipfad", "Größe", "Datum", "Aktionen"])
        details_tree.setColumnWidth(0, 500)
        
        for hash_value, filepaths in duplicates.items():
            group_item = QTreeWidgetItem(details_tree)
            group_item.setText(0, f"Duplikatgruppe ({len(filepaths)} Dateien)")
            
            for filepath in filepaths:
                item = QTreeWidgetItem(group_item)
                item.setText(0, filepath)
                try:
                    stats = os.stat(filepath)
                    item.setText(1, format_size(stats.st_size))
                    item.setText(2, datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S"))
                    
                    # Füge Aktionsbuttons hinzu
                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(0, 0, 0, 0)
                    
                    delete_button = QPushButton()
                    delete_button.setIcon(QIcon.fromTheme("edit-delete", QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))))
                    delete_button.setToolTip("Datei löschen")
                    delete_button.setFixedSize(30, 30)
                    delete_button.clicked.connect(lambda checked, path=filepath: self.delete_duplicate(path, item))
                    
                    open_button = QPushButton()
                    open_button.setIcon(QIcon.fromTheme("folder", QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))))
                    open_button.setToolTip("Ordner öffnen")
                    open_button.setFixedSize(30, 30)
                    open_button.clicked.connect(lambda checked, path=filepath: self.open_file_location(path))
                    
                    action_layout.addWidget(delete_button)
                    action_layout.addWidget(open_button)
                    action_layout.addStretch()
                    
                    details_tree.setItemWidget(item, 3, action_widget)
                except Exception:
                    item.setText(1, "N/A")
                    item.setText(2, "N/A")
        
        details_layout.addWidget(details_tree)
        tab_widget.addTab(details_tab, "Details")
        
        layout.addWidget(tab_widget)
        dialog.exec()

    def delete_duplicate(self, filepath, item):
        reply = QMessageBox.question(
            self,
            "Bestätigung",
            f"Möchten Sie die Datei {filepath} löschen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(filepath)
                item.parent().removeChild(item)
                self.status_label.setText(f"Datei gelöscht: {filepath}")
            except Exception as e:
                QMessageBox.warning(self, "Fehler", f"Fehler beim Löschen: {str(e)}")

    def open_file_location(self, filepath):
        try:
            os.startfile(os.path.dirname(filepath))
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Fehler beim Öffnen des Ordners: {str(e)}")

    def quick_delete_duplicates(self, duplicates, dialog):
        """Löscht schnell alle Duplikate und behält jeweils eine Kopie"""
        total_duplicates = sum(len(files) - 1 for files in duplicates.values())
        
        reply = QMessageBox.question(
            self,
            "Duplikate löschen",
            f"Möchten Sie alle {total_duplicates} Duplikate löschen?\n"
            "Es wird jeweils eine Kopie pro Duplikatgruppe behalten.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            progress_dialog = QProgressDialog(
                "Lösche Duplikate...", 
                "Abbrechen", 
                0, 
                total_duplicates, 
                self
            )
            progress_dialog.setWindowTitle("Lösche Duplikate")
            progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            
            deleted_count = 0
            failed_count = 0
            current_progress = 0
            
            for filepaths in duplicates.values():
                if progress_dialog.wasCanceled():
                    break
                    
                # Behalte die erste Datei und lösche den Rest
                for filepath in filepaths[1:]:
                    try:
                        os.remove(filepath)
                        deleted_count += 1
                    except Exception:
                        failed_count += 1
                    
                    current_progress += 1
                    progress_dialog.setValue(current_progress)
            
            progress_dialog.close()
            
            QMessageBox.information(
                self,
                "Löschvorgang abgeschlossen",
                f"Erfolgreich gelöscht: {deleted_count} Duplikate\n"
                f"Fehlgeschlagen: {failed_count} Duplikate\n\n"
                "Die ursprünglichen Dateien wurden behalten."
            )
            
            dialog.accept()

    def find_unused_files(self):
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Dateien zum Analysieren vorhanden.")
            return

        self.status_label.setText("Analysiere Dateizugriffe...")
        unused_files = []
        cutoff_date = datetime.now() - timedelta(days=180)  # 6 Monate

        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            filepath = item.text(0)
            try:
                stats = os.stat(filepath)
                last_access = datetime.fromtimestamp(stats.st_atime)
                last_modified = datetime.fromtimestamp(stats.st_mtime)
                
                last_used = max(last_access, last_modified)
                
                if last_used < cutoff_date:
                    unused_files.append((filepath, last_used, stats.st_size))
            except Exception as e:
                self.status_label.setText(f"Fehler bei {filepath}: {str(e)}")

        if unused_files:
            self.show_unused_files_dialog(unused_files)
        else:
            QMessageBox.information(self, "Ergebnis", "Keine ungenutzten Dateien gefunden.")
        
        self.status_label.setText("Analyse abgeschlossen")

    def show_unused_files_dialog(self, unused_files):
        dialog = QDialog(self)
        dialog.setWindowTitle("Ungenutzte Dateien (> 6 Monate)")
        dialog.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(dialog)
        tree = QTreeWidget()
        tree.setHeaderLabels(["Dateipfad", "Zuletzt genutzt", "Größe"])
        
        for filepath, last_used, size in sorted(unused_files, key=lambda x: x[1]):
            item = QTreeWidgetItem(tree)
            item.setText(0, filepath)
            item.setText(1, last_used.strftime("%Y-%m-%d %H:%M:%S"))
            item.setText(2, format_size(size))
        
        layout.addWidget(tree)
        dialog.exec()

    def show_categories(self):
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Dateien zum Kategorisieren vorhanden.")
            return

        self.status_label.setText("Kategorisiere Dateien...")
        
        categories = get_file_categories()
        category_stats = defaultdict(lambda: {"count": 0, "size": 0, "files": []})
        
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            filepath = item.text(0)
            size = parse_size(item.text(1))
            ext = Path(filepath).suffix.lower()
            
            found_category = "Sonstige"
            for category, extensions in categories.items():
                if ext in extensions:
                    found_category = category
                    break
            
            category_stats[found_category]["count"] += 1
            category_stats[found_category]["size"] += size
            category_stats[found_category]["files"].append({
                "path": filepath,
                "size": size,
                "date": item.text(2)
            })

        self.show_categories_dialog(category_stats)
        self.status_label.setText("Kategorisierung abgeschlossen")

    def show_categories_dialog(self, category_stats):
        dialog = QDialog(self)
        dialog.setWindowTitle("Dateikategorisierung")
        dialog.setMinimumSize(1000, 700)
        
        layout = QVBoxLayout(dialog)
        
        # Erstelle Tabs für verschiedene Ansichten
        tab_widget = QTabWidget()
        
        # Übersichts-Tab
        overview_tab = QWidget()
        overview_layout = QVBoxLayout(overview_tab)
        overview_tree = QTreeWidget()
        overview_tree.setHeaderLabels(["Kategorie", "Anzahl", "Gesamtgröße", "Aktionen"])
        
        for category, stats in category_stats.items():
            item = QTreeWidgetItem(overview_tree)
            item.setText(0, category)
            item.setText(1, str(stats["count"]))
            item.setText(2, format_size(stats["size"]))
            
            # Füge Lösch-Button hinzu
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(0, 0, 0, 0)
            
            delete_button = QPushButton("🗑️ Kategorie löschen")
            delete_button.clicked.connect(lambda checked, cat=category, files=stats["files"]: 
                self.delete_category(cat, files))
            
            action_layout.addWidget(delete_button)
            action_layout.addStretch()
            overview_tree.setItemWidget(item, 3, action_widget)
        
        overview_layout.addWidget(overview_tree)
        tab_widget.addTab(overview_tab, "Übersicht")
        
        # Details-Tab
        details_tab = QWidget()
        details_layout = QVBoxLayout(details_tab)
        details_tree = QTreeWidget()
        details_tree.setHeaderLabels(["Kategorie/Datei", "Größe", "Datum"])
        
        for category, stats in category_stats.items():
            category_item = QTreeWidgetItem(details_tree)
            category_item.setText(0, f"{category} ({stats['count']} Dateien)")
            category_item.setText(1, format_size(stats["size"]))
            
            for file_info in stats["files"]:
                file_item = QTreeWidgetItem(category_item)
                file_item.setText(0, file_info["path"])
                file_item.setText(1, format_size(file_info["size"]))
                file_item.setText(2, file_info["date"])
        
        details_layout.addWidget(details_tree)
        tab_widget.addTab(details_tab, "Details")
        
        layout.addWidget(tab_widget)
        dialog.exec()

    def delete_category(self, category, files):
        """Löscht alle Dateien einer Kategorie"""
        reply = QMessageBox.question(
            self,
            "Kategorie löschen",
            f"Möchten Sie wirklich alle Dateien der Kategorie '{category}' löschen?\n"
            f"Es werden {len(files)} Dateien gelöscht.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            deleted_count = 0
            failed_count = 0
            total_count = len(files)
            
            progress_dialog = QProgressDialog(
                "Lösche Dateien...", 
                "Abbrechen", 
                0, 
                total_count, 
                self
            )
            progress_dialog.setWindowTitle("Lösche Kategorie")
            progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            
            for i, file_info in enumerate(files):
                if progress_dialog.wasCanceled():
                    break
                    
                try:
                    os.remove(file_info["path"])
                    deleted_count += 1
                except Exception:
                    failed_count += 1
                    
                progress_dialog.setValue(i + 1)
            
            progress_dialog.close()
            
            QMessageBox.information(
                self,
                "Löschvorgang abgeschlossen",
                f"Erfolgreich gelöscht: {deleted_count} Dateien\n"
                f"Fehlgeschlagen: {failed_count} Dateien"
            )

    def copy_path_to_clipboard(self, item, column):
        """Kopiert den Dateipfad in die Zwischenablage"""
        path = item.text(0)  # Erste Spalte enthält den Pfad
        QApplication.clipboard().setText(path)
        self.status_label.setText(f"Pfad in Zwischenablage kopiert: {path[:50]}...")

    def format_size(self, size_bytes):
        """Formatiert eine Größe in Bytes in eine lesbare Form"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.2f} TB"

    def export_to_excel(self):
        """Exportiert die Ergebnisse in eine Excel-Datei"""
        if self.file_tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "Fehler", "Keine Daten zum Exportieren vorhanden.")
            return

        try:
            import pandas as pd
        except ImportError:
            reply = QMessageBox.question(
                self,
                "Modul fehlt",
                "Das Modul 'pandas' ist nicht installiert. Möchten Sie es jetzt installieren?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                import subprocess
                subprocess.check_call(["pip", "install", "pandas", "openpyxl"])
                import pandas as pd
            else:
                return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Excel-Datei speichern",
            "",
            "Excel-Dateien (*.xlsx)"
        )

        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'

            # Sammle die Daten
            data = []
            total_size_bytes = 0
            for i in range(self.file_tree.topLevelItemCount()):
                item = self.file_tree.topLevelItem(i)
                # Konvertiere die Größenangabe in Bytes
                size_str = item.text(1)
                try:
                    size_bytes = self.parse_size_to_bytes(size_str)
                except:
                    size_bytes = 0
                total_size_bytes += size_bytes
                
                data.append({
                    'Dateipfad': item.text(0),
                    'Größe': size_str,
                    'Datum': item.text(2),
                    'Typ': item.text(3),
                    'Ersteller': item.text(4)
                })

            # Erstelle DataFrame und exportiere
            df = pd.DataFrame(data)
            
            # Erstelle Excel-Writer
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            
            # Schreibe Haupttabelle
            df.to_excel(writer, sheet_name='Dateien', index=False)
            
            # Berechne Statistiken
            avg_size_bytes = total_size_bytes / len(data) if data else 0
            
            # Erstelle Zusammenfassungsblatt
            summary_data = {
                'Metrik': [
                    'Gesamtanzahl Dateien',
                    'Gesamtgröße',
                    'Durchschnittliche Dateigröße',
                    'Häufigste Dateitypen',
                    'Älteste Datei',
                    'Neueste Datei'
                ],
                'Wert': [
                    len(data),
                    self.format_size(total_size_bytes),
                    self.format_size(avg_size_bytes),
                    df['Typ'].value_counts().head(3).to_string() if len(data) > 0 else 'N/A',
                    df['Datum'].min() if len(data) > 0 else 'N/A',
                    df['Datum'].max() if len(data) > 0 else 'N/A'
                ]
            }
            
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Zusammenfassung', index=False)
            
            # Speichere und schließe
            writer.close()

            QMessageBox.information(
                self,
                "Export erfolgreich",
                f"Die Daten wurden erfolgreich nach {file_path} exportiert."
            )

    def parse_size_to_bytes(self, size_str):
        """Konvertiert einen formatierten Größenstring in Bytes"""
        try:
            # Extrahiere Zahl und Einheit
            parts = size_str.strip().split()
            if len(parts) != 2:
                return 0
                
            value = float(parts[0])
            unit = parts[1].upper()
            
            # Konvertiere in Bytes
            multipliers = {
                'B': 1,
                'KB': 1024,
                'MB': 1024 * 1024,
                'GB': 1024 * 1024 * 1024,
                'TB': 1024 * 1024 * 1024 * 1024
            }
            
            if unit in multipliers:
                return int(value * multipliers[unit])
            return 0
        except:
            return 0

    def show_mass_delete_dialog(self):
        """Zeigt den Dialog für Massenlöschung"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Massenlöschung")
        dialog.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(dialog)
        
        # Erklärungstext
        info_label = QLabel(
            "Fügen Sie hier die zu löschenden Dateipfade ein (einen pro Zeile).\n"
            "Sie können die Pfade auch aus Excel kopieren und hier einfügen."
        )
        layout.addWidget(info_label)
        
        # Textfeld für Pfade
        self.path_text = QTextEdit()
        self.path_text.setPlaceholderText("C:\\Pfad\\zur\\Datei1.txt\nC:\\Pfad\\zur\\Datei2.doc\n...")
        layout.addWidget(self.path_text)
        
        # Buttons
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        
        validate_button = QPushButton("Pfade überprüfen")
        validate_button.clicked.connect(lambda: self.validate_paths(dialog))
        
        delete_button = QPushButton("🗑️ Ausgewählte Dateien löschen")
        delete_button.clicked.connect(lambda: self.execute_mass_delete(dialog))
        
        button_layout.addWidget(validate_button)
        button_layout.addWidget(delete_button)
        
        layout.addWidget(button_frame)
        
        # Ergebnisliste
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["Dateipfad", "Status"])
        self.result_tree.setColumnWidth(0, 500)
        layout.addWidget(self.result_tree)
        
        dialog.exec()

    def validate_paths(self, dialog):
        """Überprüft die eingegebenen Pfade"""
        self.result_tree.clear()
        paths = self.path_text.toPlainText().strip().split('\n')
        paths = [p.strip() for p in paths if p.strip()]
        
        for path in paths:
            item = QTreeWidgetItem(self.result_tree)
            item.setText(0, path)
            
            if os.path.exists(path):
                if os.path.isfile(path):
                    item.setText(1, "✅ Datei existiert")
                    item.setBackground(1, QColor("#e8f5e9"))  # Hellgrün
                else:
                    item.setText(1, "❌ Ist ein Ordner")
                    item.setBackground(1, QColor("#ffebee"))  # Hellrot
            else:
                item.setText(1, "❌ Datei nicht gefunden")
                item.setBackground(1, QColor("#ffebee"))  # Hellrot

    def execute_mass_delete(self, dialog):
        """Führt die Massenlöschung durch"""
        valid_items = []
        for i in range(self.result_tree.topLevelItemCount()):
            item = self.result_tree.topLevelItem(i)
            if item.text(1).startswith("✅"):
                valid_items.append(item.text(0))
        
        if not valid_items:
            QMessageBox.warning(self, "Fehler", "Keine gültigen Dateien zum Löschen gefunden.")
            return
        
        reply = QMessageBox.question(
            self,
            "Massenlöschung bestätigen",
            f"Möchten Sie wirklich {len(valid_items)} Dateien löschen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            progress_dialog = QProgressDialog(
                "Lösche Dateien...", 
                "Abbrechen", 
                0, 
                len(valid_items), 
                dialog
            )
            progress_dialog.setWindowTitle("Massenlöschung")
            progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            
            deleted_count = 0
            failed_count = 0
            
            for i, path in enumerate(valid_items):
                if progress_dialog.wasCanceled():
                    break
                    
                try:
                    os.remove(path)
                    deleted_count += 1
                except Exception:
                    failed_count += 1
                
                progress_dialog.setValue(i + 1)
            
            progress_dialog.close()
            
            QMessageBox.information(
                self,
                "Löschvorgang abgeschlossen",
                f"Erfolgreich gelöscht: {deleted_count} Dateien\n"
                f"Fehlgeschlagen: {failed_count} Dateien"
            )
            
            # Aktualisiere die Statusanzeige
            self.validate_paths(dialog) 