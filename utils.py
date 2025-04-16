import os
import hashlib
from datetime import datetime
from pathlib import Path

def format_size(size):
    """Formatiert eine Dateigröße in Bytes in lesbare Form"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} TB"

def parse_size(size_str):
    """Konvertiert einen Größenstring (z.B. '1.23 MB') in Bytes"""
    value, unit = size_str.split()
    value = float(value)
    
    multipliers = {
        'B': 1,
        'KB': 1024,
        'MB': 1024 ** 2,
        'GB': 1024 ** 3,
        'TB': 1024 ** 4
    }
    
    return value * multipliers[unit]

def calculate_file_hash(filepath, quick_mode=False):
    """Berechnet den Hash einer Datei"""
    try:
        file_size = os.path.getsize(filepath)
        
        # Für große Dateien (>100MB) verwenden wir einen schnelleren Algorithmus
        if quick_mode and file_size > 100 * 1024 * 1024:  # 100MB
            # Nur die ersten und letzten 4MB der Datei hashen
            sample_size = 4 * 1024 * 1024  # 4MB
            with open(filepath, 'rb') as f:
                # Erste 4MB
                first_part = f.read(sample_size)
                # Letzte 4MB
                f.seek(-sample_size, 2)
                last_part = f.read(sample_size)
                # Kombiniere die Teile
                content = first_part + last_part
        else:
            # Für kleine Dateien den gesamten Inhalt hashen
            with open(filepath, 'rb') as f:
                content = f.read()
                
        return hashlib.md5(content).hexdigest()
    except Exception:
        return None

def get_file_categories():
    """Gibt die vordefinierten Dateikategorien zurück"""
    return {
        "Bilder": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"],
        "Office": [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".accdb", ".mdb", ".odt", ".ods", ".odp"],
        "PDF": [".pdf"],
        "Videos": [".mp4", ".avi", ".mov", ".wmv", ".flv", ".mkv", ".webm"],
        "Audio": [".mp3", ".wav", ".ogg", ".m4a", ".flac", ".aac"],
        "Archive": [".zip", ".rar", ".7z", ".tar", ".gz"],
        "Code": [".py", ".js", ".html", ".css", ".cpp", ".java", ".php"],
        "Datenbanken": [".db", ".sqlite", ".mdb", ".accdb"],
        "System": [".sys", ".dll", ".exe", ".bat", ".cmd"],
        "Sonstige": []
    }

def get_file_type_extensions():
    """Gibt die vordefinierten Dateityp-Gruppen zurück"""
    return {
        "Bilder (.jpg, .jpeg, .png, .gif)": [".jpg", ".jpeg", ".png", ".gif"],
        "Office-Dokumente": [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".accdb", ".mdb"],
        "PDF-Dateien (.pdf)": [".pdf"],
        "Videos (.mp4, .avi, .mov)": [".mp4", ".avi", ".mov"],
        "Audio (.mp3, .wav, .ogg)": [".mp3", ".wav", ".ogg"],
        "Archive (.zip, .rar, .7z)": [".zip", ".rar", ".7z"]
    } 
