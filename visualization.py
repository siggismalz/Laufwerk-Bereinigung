import matplotlib.pyplot as plt
from collections import defaultdict

class Visualization:
    def __init__(self, file_tree, format_size, parse_size):
        self.file_tree = file_tree
        self.format_size = format_size
        self.parse_size = parse_size

    def visualize_data(self):
        if self.file_tree.topLevelItemCount() == 0:
            return False
        # Sammle Daten nach Dateitypen
        file_types = {}
        total_size = 0
        
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            file_type = item.text(3)
            size_str = item.text(1)
            
            size = self.parse_size(size_str)
            
            if file_type in file_types:
                file_types[file_type] += size
            else:
                file_types[file_type] = size
            total_size += size

        # Sortiere Dateitypen nach Größe
        sorted_types = sorted(file_types.items(), key=lambda x: x[1], reverse=True)
        
        # Gruppiere kleine Dateitypen (weniger als 1% der Gesamtgröße)
        threshold = total_size * 0.01
        main_types = []
        other_size = 0
        
        for file_type, size in sorted_types:
            if size >= threshold:
                main_types.append((file_type, size))
            else:
                other_size += size
        
        if other_size > 0:
            main_types.append(("Sonstige", other_size))

        # Erstelle Kreisdiagramm
        plt.figure(figsize=(12, 8))
        
        # Explodiere das größte Segment
        explode = [0.1 if i == 0 else 0 for i in range(len(main_types))]
        
        # Erstelle das Diagramm
        wedges, texts, autotexts = plt.pie(
            [size for _, size in main_types],
            labels=[f"{ft}\n({self.format_size(size)})" for ft, size in main_types],
            autopct='%1.1f%%',
            explode=explode,
            shadow=True,
            startangle=90,
            textprops={'fontsize': 10}
        )
        
        # Verbessere die Darstellung der Prozentangaben
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        # Füge Titel und Legende hinzu
        plt.title("Verteilung der Dateitypen nach Größe", fontsize=14, pad=20)
        plt.legend(
            wedges,
            [f"{ft} ({self.format_size(size)})" for ft, size in main_types],
            title="Dateitypen",
            loc="center left",
            bbox_to_anchor=(1, 0, 0.5, 1)
        )
        
        # Setze das Seitenverhältnis auf gleich
        plt.axis('equal')
        
        # Passe das Layout an
        plt.tight_layout()
        
        # Zeige das Diagramm
        plt.show()
        return True 
