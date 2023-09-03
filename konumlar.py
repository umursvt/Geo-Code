import pandas as pd
import requests


def adrestenKonum():
    # Dosya Okuma
    dosya_adi = "Migros.xlsx"
    kolon_index = 4
    data = pd.read_excel(dosya_adi)
    kolon = data.iloc[:, kolon_index]
    return kolon


def konumAPI(adres):
    # ChatGPT'nin düzenlediği API bölümü
    url = f"https://api.opencagedata.com/geocode/v1/json"
    params = {
        "q": adres,
        "key": "596e5df334ad4347a73246a92e820005"
    }
    response = requests.get(url, params=params)
    data = response.json()

    if "results" in data and len(data["results"]) > 0:
        enlem = float(data["results"][0]["geometry"]["lat"])
        boylam = float(data["results"][0]["geometry"]["lng"])
        return enlem, boylam
    else:
        return None


# Adresleri konum bilgilerine çevirin
adresler = adrestenKonum()
konumlar = []

for adres in adresler:
    konum = konumAPI(adres)
    if konum:
        konumlar.append(konum)
    else:
        konumlar.append((None, None))

# Konumları ekrana yazdırın
for adres, konum in zip(adresler, konumlar):
    print(f"Adres: {adres} - Enlem: {konum[0]}, Boylam: {konum[1]}")
