import os
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import win32security
import win32api
import win32con
from PySide6.QtCore import QThread, Signal
from queue import Queue
import gc

class FileScanner(QThread):
    file_found = Signal(str, str, str, str, str)  # Pfad, Größe, Datum, Typ, Ersteller
    progress_update = Signal(int)  # Fortschritt in Prozent
    scan_complete = Signal(float, int)  # Gesamtgröße in GB, Anzahl Dateien
    status_update = Signal(str)  # Statusmeldungen
    collection_progress = Signal(int, int)  # Aktueller Fortschritt, Geschätzte Gesamtanzahl

    # Größenkategorien in Bytes
    SIZE_CATEGORIES = {
        "Kleine Dateien": (0, 10 * 1024 * 1024),  # 0-10MB
        "Mittlere Dateien": (10 * 1024 * 1024, 100 * 1024 * 1024),  # 10-100MB
        "Große Dateien": (100 * 1024 * 1024, float('inf'))  # >100MB
    }

    def __init__(self, drive_path, years, file_types=None, owner_filter=None, size_filter=None, max_workers=None):
        super().__init__()
        self.drive_path = drive_path
        self.years = years
        self.file_types = file_types or []
        self.owner_filter = owner_filter
        self.size_filter = size_filter
        self.stop_scan = False
        self.pause_scan = False  # Neue Variable für Pause-Zustand
        self.total_size = 0
        self.file_count = 0
        # Optimiere Worker-Anzahl basierend auf CPU-Kernen
        self.max_workers = max_workers or min(32, os.cpu_count() * 2)
        self.size_lock = Lock()
        self.processed_files = 0
        self.total_files = 0
        self.skip_paths = set()
        self.file_queue = Queue(maxsize=10000)  # Begrenzte Queue-Größe
        self.collection_complete = False
        self.MAX_FILES = 50000  # Maximale Anzahl der zu scannenden Dateien

    def check_admin_access(self, path):
        try:
            sd = win32security.GetFileSecurity(
                path, 
                win32security.OWNER_SECURITY_INFORMATION
            )
            if not os.access(path, os.R_OK):
                return False
                
            attributes = win32api.GetFileAttributes(path)
            if attributes & win32con.FILE_ATTRIBUTE_HIDDEN or \
               attributes & win32con.FILE_ATTRIBUTE_SYSTEM:
                return False
                
            return True
        except Exception:
            return False

    def get_file_owner(self, filepath):
        try:
            sd = win32security.GetFileSecurity(
                filepath, 
                win32security.OWNER_SECURITY_INFORMATION
            )
            owner_sid = sd.GetSecurityDescriptorOwner()
            name, domain, type = win32security.LookupAccountSid(None, owner_sid)
            return f"{domain}\\{name}"
        except Exception:
            return "Unbekannt"

    def check_size_filter(self, file_size):
        """Prüft, ob die Dateigröße dem ausgewählten Filter entspricht"""
        if not self.size_filter:
            return True
            
        min_size, max_size = self.SIZE_CATEGORIES[self.size_filter]
        return min_size <= file_size < max_size

    def process_file(self, file_path):
        try:
            if not os.path.exists(file_path):
                return None

            file_stat = os.stat(file_path)
            file_size = file_stat.st_size
            
            # Prüfe Größenfilter
            if not self.check_size_filter(file_size):
                return None
                
            file_date = datetime.fromtimestamp(file_stat.st_mtime)
            cutoff_date = datetime.now() - timedelta(days=self.years * 365)
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # Hole den Dateieigentümer
            owner = self.get_file_owner(file_path)
            
            # Prüfe Owner-Filter mit Wildcard
            if self.owner_filter:
                owner_matches = False
                search_terms = self.owner_filter.lower().split()
                owner_lower = owner.lower()
                owner_matches = all(term in owner_lower for term in search_terms)
                if not owner_matches:
                    return None

            if file_date < cutoff_date and (not self.file_types or file_ext in self.file_types):
                with self.size_lock:
                    self.total_size += file_size
                    self.file_count += 1
                return (
                    file_path,
                    self.format_size(file_size),
                    file_date.strftime("%Y-%m-%d %H:%M:%S"),
                    file_ext,
                    owner
                )
        except (PermissionError, OSError):
            pass
        except Exception as e:
            self.status_update.emit(f"Fehler bei {file_path}: {str(e)}")
        return None

    def scan_directory(self, directory):
        """Scanne ein Verzeichnis und alle seine Unterverzeichnisse rekursiv"""
        try:
            # Sammle zuerst alle Dateipfade
            all_files = []
            for root, _, files in os.walk(directory):
                for file in files:
                    all_files.append(os.path.join(root, file))
            
            # Verarbeite die Dateien in Chunks
            total_files = len(all_files)
            self.status_update.emit(f"Verarbeite {total_files:,} Dateien...")
            
            for i, filepath in enumerate(all_files):
                if self.stop_scan:
                    break
                    
                self.process_file(filepath)
                
                # Aktualisiere den Fortschritt
                progress = int((i + 1) / total_files * 100)
                self.progress_update.emit(progress)
                
        except Exception as e:
            self.status_update.emit(f"Fehler beim Scannen von {directory}: {str(e)}")

    def collect_files(self):
        """Sammelt Dateien in einem separaten Thread"""
        estimated_total = 0
        collected_count = 0
        
        def scan_directory(path):
            nonlocal estimated_total, collected_count
            try:
                if self.stop_scan or not self.check_admin_access(path):
                    if not self.check_admin_access(path):
                        self.skip_paths.add(path)
                        self.status_update.emit(f"Überspringe geschützten Ordner: {path}")
                    return

                while self.pause_scan and not self.stop_scan:
                    # Warte während der Pause
                    self.msleep(100)

                with os.scandir(path) as entries:
                    entry_list = list(entries)
                    estimated_total += len(entry_list)
                    
                    for entry in entry_list:
                        if self.stop_scan:
                            return
                            
                        while self.pause_scan and not self.stop_scan:
                            # Warte während der Pause
                            self.msleep(100)
                            
                        if collected_count >= self.MAX_FILES:
                            self.status_update.emit(f"Maximale Anzahl von {self.MAX_FILES:,} Dateien erreicht")
                            self.stop_scan = True
                            return
                            
                        try:
                            if entry.is_file():
                                self.file_queue.put(entry.path)
                                collected_count += 1
                                if collected_count % 100 == 0:
                                    self.collection_progress.emit(collected_count, min(estimated_total, self.MAX_FILES))
                            elif entry.is_dir() and entry.path not in self.skip_paths:
                                scan_directory(entry.path)
                        except PermissionError:
                            self.skip_paths.add(entry.path)
                            self.status_update.emit(f"Keine Berechtigung für: {entry.path}")
                        except Exception as e:
                            self.status_update.emit(f"Fehler beim Scannen von {entry.path}: {str(e)}")
                            
            except PermissionError:
                self.skip_paths.add(path)
                self.status_update.emit(f"Keine Berechtigung für: {path}")
            except Exception as e:
                self.status_update.emit(f"Fehler beim Scannen von {path}: {str(e)}")

        try:
            scan_directory(self.drive_path)
        finally:
            self.collection_complete = True

    def process_file_chunk(self, files):
        """Verarbeitet einen Chunk von Dateien"""
        results = []
        for file_path in files:
            if self.stop_scan:
                break
            result = self.process_file(file_path)
            if result:
                results.append(result)
        return results

    def run(self):
        try:
            self.status_update.emit("Sammle Dateien...")
            
            # Starte Dateisammlung in separatem Thread
            from threading import Thread
            collection_thread = Thread(target=self.collect_files)
            collection_thread.start()
            
            # Verarbeite Dateien in Chunks
            chunk_size = 100
            current_chunk = []
            processed_count = 0
            
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                while not (self.collection_complete and self.file_queue.empty()) and not self.stop_scan:
                    while self.pause_scan and not self.stop_scan:
                        # Warte während der Pause
                        self.msleep(100)
                        
                    try:
                        # Sammle Dateien für den aktuellen Chunk
                        while len(current_chunk) < chunk_size and not self.pause_scan:
                            try:
                                file_path = self.file_queue.get(timeout=0.1)
                                current_chunk.append(file_path)
                            except Queue.Empty:
                                if self.collection_complete:
                                    break
                        
                        if not current_chunk:
                            continue
                        
                        # Verarbeite den Chunk
                        future = executor.submit(self.process_file_chunk, current_chunk)
                        results = future.result()
                        
                        # Verarbeite die Ergebnisse
                        for result in results:
                            if result:
                                self.file_found.emit(*result)
                        
                        processed_count += len(current_chunk)
                        if self.total_files > 0:
                            progress = int((processed_count / self.total_files) * 100)
                            self.progress_update.emit(progress)
                        
                        current_chunk = []
                        
                        # Führe Garbage Collection durch
                        if processed_count % 1000 == 0:
                            gc.collect()
                            
                    except Exception as e:
                        self.status_update.emit(f"Fehler bei der Chunk-Verarbeitung: {str(e)}")
            
            collection_thread.join()
            
            if not self.stop_scan:
                self.status_update.emit("Scan abgeschlossen")
                # Konvertiere die Werte in kleinere Einheiten
                total_size_gb = self.total_size / (1024 * 1024 * 1024)  # Konvertiere zu GB
                self.scan_complete.emit(total_size_gb, self.file_count)
            
        except Exception as e:
            self.status_update.emit(f"Kritischer Fehler: {str(e)}")
            total_size_gb = self.total_size / (1024 * 1024 * 1024)  # Konvertiere zu GB
            self.scan_complete.emit(total_size_gb, self.file_count)
        finally:
            # Finale Garbage Collection
            gc.collect()

    def format_size(self, size):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.2f} {unit}"
            size /= 1024
        return f"{size:.2f} TB" 