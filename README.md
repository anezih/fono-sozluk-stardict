# ESKİ

Bu Powershell betiği yerine [https://github.com/anezih/FonoSozlukNet](https://github.com/anezih/FonoSozlukNet) reposunda yer alan uygulamayı kullanabilirsiniz.

# Nedir?
Fono Yayınları tarafından çıkartılan sözlüklerin CD verilerini TSV, Babylon GLS ve Stardict Textual formatlarına çeviren bir Powershell betiği. Bu üç çıktı tipi StarDict Editor aracılığıyla StarDict formatına dönüştürülebilir. Ayrıca, Textual format PyGlossary aracılığıyla programın desteklediği herhangi bir çıktı formatına dönüştürülebilir. Örneğin Kobo `dicthtml.zip` veya Kindle MOBI formatı.

# Kullanımı
Sözlük veri dosyalarını `SozlukDosyalari` klasörüne kopyalayın. İsteğe bağlı olarak Hunspell sözlüklerini de `Hunspell` klasörüne koyun. Belirtilen klasörlerdeki açıklamaları takip edin.

Betiği çalıştırmak için Powershell 7 gereklidir. Bu klasörde uçbirimi başlatın. Parametreleri görmek için: `.\Fono-Stardict.ps1 -?` yazın. 

`Fono-Stardict.ps1 [-Dil] <string> [-GLS] [-Textual] [-TSV] [-Hunspell] [<CommonParameters>]`

`-Dil` için kabul edilen 6 parametreden (EN, FR, ES, DE, RU, IT) birini girin. Örneğin: `-Dil EN`. Şu anda sadece `EN` (İngilizce), `FR` (Fransızca) ve `IT` (İtalyanca) dönüşümü yapılabiliyor. Diğer dillerin veri dosyaları elinizde varsa bu dillerin desteklenmesi için yardım edebilirsiniz.

`-TSV`, `-GLS`, `-Textual` dönüşüm formatlarından en az birini ekleyin. Birden fazla seceçenek geçebilirsiniz.

`-Hunspell` seçeneğini sözcüklerin farklı formlarının üretilmesini istiyorsanız kullanın.

Örneğin Fransızcayı hem Textual formata hem TSV formatına çevirmek ve Hunspell ile farklı sözcük sonlarını eklemek için:

```
.\Fono-Stardict.ps1 -Dil FR -Textual -TSV -Hunspell
```

# Ekran görüntüsü
![Sözlüklerin GoldenDict üzerindeki görünümü](/goruntu/goldendict_ornek_en_fr_it.png)
*Sözlüklerin GoldenDict üzerindeki görünümü*
