"""
Python ile photoshopta çalışma
Türkiye Futbol Federasyonunun resmi web sitesinden çektiği puan durumunu
photoshopta metin katmanlarına ekler. 10'ar 10'ar eklettim.
Bunun sebebi 1920x1080 gibi çalışmalarda 20 takım fazla oluyor.
2 ye böldü

Photoshop Açık olması gerekmektedir.
Metin katmanlarınında oluşturulmuş olması gerekmektedir.

"""
import win32com.client # windows applicationlar ile çalışma modülü
import requests # bağlantı modülü
from bs4 import BeautifulSoup # veri çekme modülü
from tqdm import tqdm  # progress modülü

#Bu sözlükte yer alan reklam, a.ş gibi yazıları siler.
def duzenle_takim_adi(takim_ismi):
    duzenleme_sozlugu = {
        "A.Ş.": "",
        "FUTBOL KULÜBÜ": "",
        "MONDİHOME": "",
        "YUKATEL": "",
        "BITEXEN": "",
        "RAMS": "",
        "ATAKAŞ": "",
        "TÜMOSAN": "",
        "YILPORT": "",
        "VAVACARS": "",
        "SİLTAŞ YAPI": "",
        "CORENDON": "",
        "EMS YAPI": "",
        "FUTBOL": "",
    }
#sözlükten sildiğimiz değerleri ayarlıyoruz ve çekilen verinin başındaki 1. 2. gibi sayıları kaldırıyoruz.
    for anahtar, deger in duzenleme_sozlugu.items():
        takim_ismi = takim_ismi.replace(anahtar, deger)

    takim_ismi = ''.join(i for i in takim_ismi if not i.isdigit() and i != '.')

    return takim_ismi.strip()

#url çekiyoruz
url = "https://www.tff.org/default.aspx?pageID=198"

#url içinde html yi parçalıyoruz
response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Puanlar, Averaj ve Oynanan Maç Sayısı için bilgileri çekme
puanlar = soup.find_all('span', id=lambda x: x and x.endswith('Label3'))
averaj = soup.find_all('span', id=lambda x: x and x.endswith('Label5'))
oynanan_mac_sayisi = soup.find_all('span', id=lambda x: x and x.endswith('lblOyun'))

# Photoshop ile çalışma
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"Düzenlemekistediğiniz psd adresi.psd")
doc = psApp.Application.ActiveDocument
# Metin katmanlarını seçme photoshoptaki metin layer isimleri burası önemli
puanText = doc.ArtLayers["puanText"].TextItem
averajText = doc.ArtLayers["averajText"].TextItem
oynananText = doc.ArtLayers["oynananText"].TextItem
takimText = doc.ArtLayers["takimText"].TextItem
puanText2 = doc.ArtLayers["puanText2"].TextItem
averajText2 = doc.ArtLayers["averajText2"].TextItem
oynananText2 = doc.ArtLayers["oynananText2"].TextItem
takimText2 = doc.ArtLayers["takimText2"].TextItem

# Eğer metin katmanlarının içi doluysa boşaltıyor.
puanText.contents = ""
averajText.contents = ""
oynananText.contents = ""
takimText.contents = ""
puanText2.contents = ""
averajText2.contents = ""
oynananText2.contents = ""
takimText2.contents = ""

# Her bir takım için puanlar, averaj ve oynanan maç sayısı bilgilerini metin katmanlarına ekleyerek alt alta yazma
for i, puan in tqdm(enumerate(puanlar), total=len(puanlar), desc="Puan Durumu Ekleniyor"):
    takim_ismi = duzenle_takim_adi(puan.find_previous('a').getText())
    
    if i < 10:
        takimText.contents += f"{takim_ismi.strip()}\r"
        puanText.contents += f"{puan.text.strip()}\r"
        averajText.contents += f"{averaj[i].text.strip()}\r"
        oynananText.contents += f"{oynanan_mac_sayisi[i].text.strip()}\r"
    else:
        takimText2.contents += f"{takim_ismi.strip()}\r"
        puanText2.contents += f"{puan.text.strip()}\r"
        averajText2.contents += f"{averaj[i].text.strip()}\r"
        oynananText2.contents += f"{oynanan_mac_sayisi[i].text.strip()}\r"
