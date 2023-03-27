Bu klasöre İngilizce, Fransızca ve Türkçe için olan Hunspell sözlük dosyalarını koyun. Bu sayede sözlükte arama yapılırken arama yapılan sözcük farklı bir formda olsa dahi kök sözcük getirilir. Örneğin, `yapıtlarını` araması `yapıt` sonucunu getirecektir.

Her Hunspell sözlüğü iki ayrı dosyadan oluşur: \*.aff ve \*.dic dosyaları. Dosya isimlerinin aşağıdaki gibi olmasına özen gösterin.

```
en_US.dic
fr_FR.dic
tr_TR.dic
```
Sözlükler için kaynak:

https://github.com/titoBouzout/Dictionaries

# Not
- Sözcüklerin çekimlenmiş hallerinin üretilmesinde (özellikle Türkçe için) Hunspell kullanmak tam da beklenen sonuçları veremeyebilir.
- WordForms sınıfını içeren HunspellWordForms.dll kütüphanesinin [kaynak kodu](https://github.com/anezih/HunspellWordForms)
