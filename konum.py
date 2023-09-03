import sys
import geocoder
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget


class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Açık Adresten Konum")

        # Ana widget
        central_widget = QWidget()

        # Ana düzenleyici
        main_layout = QVBoxLayout(central_widget)
        main_layout.setAlignment(Qt.AlignCenter)
        main_layout.setContentsMargins(50, 50, 50, 50)
        # Buton ve labeller arasında boşluk bırakma
        main_layout.setSpacing(20)

        # Üst düzenleyici
        top_layout = QVBoxLayout()

        # Label 1
        label1 = QLabel(
            "Excel dosyasındaki açık adres bilgilerini konum bilgisine dönüştürür", self)
        label1.setStyleSheet("color: black; font-size: 20px;")
        label1.setAlignment(Qt.AlignCenter)
        top_layout.addWidget(label1)

        # Buton
        button = QPushButton("Programı Çalıştır", self)
        button.setStyleSheet(
            "background-color: orange; color: black; font-size: 16px;")
        button.clicked.connect(self.run_geocode)
        top_layout.addWidget(button)

        # Label 2
        label2 = QLabel(
            "Programın tüm verileri işlemesi 30-50 dakika sürebilir. Çalışması bittikten sonra Excel tablonuzun güncel hali otomatik olarak açılacaktır.", self)
        label2.setStyleSheet("color: red; font-size: 16px;")
        label2.setAlignment(Qt.AlignCenter)
        # MEtin düzenleme
        label2.setWordWrap(True)
        label2.setFont(QFont("Arial", 12))  # Yazı tipi ve boyutu ayarlama
        top_layout.addWidget(label2)

        # Üst düzenleyiciyi ana düzenleyiciye ekle
        main_layout.addLayout(top_layout)

        # Ana widget ayarları
        self.setCentralWidget(central_widget)

        # Pencere boyutu ayarla
        self.resize(600, 500)

    def run_geocode(self):
        # Excel dosya yolu
        file_path = 'Adres.xlsx'

        try:
            print('Program Başlatılıyor')
            # Excel dosyasını oku
            df = pd.read_excel(file_path)
            addresses = df['Adres']

            # Listeler
            enlem_liste = []
            boylam_liste = []
            google_map_linkler = []

            # Açık adres --> Konum
            for adres in addresses:
                g = geocoder.arcgis(adres)
                if g.ok:
                    enlem_liste.append(g.lat)
                    boylam_liste.append(g.lng)
                    google_map_link = f"https://www.google.com/maps?q={g.lat},{g.lng}"
                    google_map_linkler.append(google_map_link)
                else:
                    enlem_liste.append(None)
                    boylam_liste.append(None)
                    google_map_linkler.append(None)

            # Enlem,Boylam ve google link

            df['Enlem'] = enlem_liste

            df['Boylam'] = boylam_liste

            df['Google Maps Linki'] = google_map_linkler

            # Enlem boylam bilgisini tek bir kolona ekle

            df['Enlem,Boylam'] = df['Enlem'].astype(str) + ',' + df['Boylam'].astype(str)

            df = df.drop(['Enlem', 'Boylam'], axis=1)

            # Excel tablosunu güncelleeme
            df.to_excel(file_path, index=False)

            # Excel tablosunu açmaa
            import os
            os.startfile(file_path)

            print("Geocode işlemi tamamlandı.")
        except Exception as e:
            print(f'Hata: {str(e)}')


# Uygulama oluşturma
app = QApplication(sys.argv)

# Pencere oluşturma
window = AppWindow()

# Pencereyi gösterme
window.show()

# Uygulamayı çalıştırma
sys.exit(app.exec())
