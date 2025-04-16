import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from PySide6.QtCore import QTimer
from ui import DriveCleanerApp, SplashScreen
import ctypes

def main():
    app = QApplication(sys.argv)
    
    # Windows-spezifische App-ID setzen
    myappid = 'edeka.drivecleaner.1.0'
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except:
        pass
    
    # Setze das Icon für die gesamte Anwendung
    icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "clean_drive.ico"))
    app_icon = QIcon(icon_path)
    app.setWindowIcon(app_icon)
    
    # Zeige Splashscreen
    splash = SplashScreen()
    splash.show()
    
    # Erstelle Hauptfenster
    window = DriveCleanerApp()
    
    # Schließe Splashscreen und zeige Hauptfenster nach 3 Sekunden
    QTimer.singleShot(3000, lambda: (splash.close(), window.show()))
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 