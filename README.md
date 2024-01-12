# Dashboards-and-Sales-Reports
Merhaba, İlgili excel ve arkaplanında kullandığım yapıları Anlaşılır Ekonomi Youtube kanalından öğrenerek yaptım. 

## Genel Satış Raporu (Rapor 1)
İlgili genel satış rapounda ise raporlar için verilerimizden otomatik olarak çekebileceğimiz bir chech box yapısı kullandım. 3 farklı ürün yani Raf, Kitaplık, Masa ve Hepsi için. Ayrıca Her bölgeye ait aylık satış raporu da seçilen ürüne göre otomatik olarak değişmektedir. Ayrıca ilgili tabloda ortalama üstü, altını işaretleyebileceğimiz bir makro bulunmaktadır. Açılır listeden ise ilgili bölgeyi seçip sonrasında aşağıda yer alan tablolara veri otomatik olarak gelmektedir. Ayrıca her sayfada yer alan Satış raporlarına solda bulunan butonlar(aslında şekil) ile hızlıca geçiş sağlayabiliyoruz.

## Bölge Satış Raporu (Rapor2)
İlgili Bölge satış raporunda ise seçilen bölgede aylık olarak hangi ürün kaç adet ve ne kadar satıldığıyla ilgili otomatik tablomuz bulunmaktadır. Ayrıca her ürün için ve toplamları için bölgelere ait grafikler de yer almaktadır. Sağ tarafta yer alan ilgili butonlar ile tutar ve adet grafikleri arasında geçiş yapabilirsiniz.

## Şube Satış Raporu (Rapor 3)
Şube satış raporunda ise SUMIFS yapısı ile şubelerdeki verileri çekiyoruz. Sonrasında ise Selection Change dediğimiz click ile seçtiğimiz yere giden yapıyı yazıyoruz. Aşağpıda yer alan kodun açıklaması da tam olarak 9. satırda yer alan  12. ve 35. sütunlar arasındaki tabloda seçtiğimzi ilçenin X12 alanına yazdırılmasını göstermektedir. Ayrıca ilgili seçilen şehrin, ürünün verilerini de tabloya aktarmaktadır.

```
If Target.Row > 12 And Target.Row < 35 And Target.Column = 9 Then Range("X12") = Target.Value
```
## Temsilci Satış Raporu (Rapor4)
İlgili Temsilci satış raporunda ise istenilen bölgenin, şubenin ve temsilcinin seçilip otomatik olarak tabloya ve grafiklere aktaran bir yapı bulunmaktadır. Ayrıca selection change yapısına benzeyen bir modül üstüne yazdığımız yapıyı içerkmektedir. 12. satır 16. sütuna yazdırmasını sağlayan ve Range in mouseoverr şeklinde tanımlandığı bir yapı oluşturuyoruz. Bu arkada oluşan yapı da gizli olarak saklanacak bu gizli saklanan yapı ise aşağıda kurduğumuz tablo yapısını otomatik olarak değiştirecetkir.

```
Public Function mouseover(mouseoverr As Range)

Sheets("Temsilci Satış Raporu").Cells(12, 16) = mouseoverr

End Function
```

## Hedef Satış Raporu (Rapor 5)
İlgili Hedef satış raporunda da ilgili kişilerin seçilip otomatik olarak speedometerları değiştirdiği bir yapı kurmuş olduk. Buradaki bana göre diğer sistemlerden farklı yaptığımız şey aslında oluşturduğumuz ibrenin isminin değiştirilip kaçıncı açıda %100 e ulaştığını bulup makro kısmında da bunu tanımlamak. Diğer işlemler daha önceki raporlarda gerçekleştirdiğimiz gerek linked cell kısmı olsun gerekse ListFillRange'ler ile oluşturduğumuz Name Manager(Ad Yöneticisi) kısmındaki verileri çekmek gibi unsurlar bulunuyor. Burada ilgili ibreyi hareket ettirecek yapı aşağıda yer almaktadır. Bu yapıyı hem listbox'ın içine hem de WorkSheet_Change yapımızın içine yazıp sağlıklı şekilde ibre hareketi sağlayabiliriz. Ayrıca değişen ibre arka planda ekran update sorunu yarattığı içinde aşağıda yazdığım örnek kod olan ekran güncellenmesini kapatıp sonrasında ilgili ibrenizin hareket etmesini sağlayabilirsiniz.

```
Application.ScreenUpdating = False

ActiveSheet.Shapes.Range(Array("ibre1")).Select
Selection.ShapeRange.Rotation = Range("F24").Value * 191
ActiveCell.Select
```

## Sonuç
Sonuç olarak bu yapılarda diğer İş Zekası tooları ile kolayca yapılabilir ancak veri çekme açısından dinamik yapılarıyla excel onlara karşı daha üstün görselleştirme adına daha zayıf olduğunu düşünmekteyim. Ayırca formül ve kod yapıları verilerimizle daha net bir şekilde yeni yapılar sağlamamızı kolaylaştııyor. Son zamanlarda onlarca bootcamp bulunmakta ancak ücretleri bazılarının ateş pahası malesef. Buradaki amacım çeşitli YouTube ve ücretsiz içerikleri kullanarak. İş Analizi, Veri Analizi ve İş Zekası alanlarında kendimi geliştirmektedir. 
Ayrıca ilgili excellerin adım adım halleri ve her excel için yaptığım yapıları daha detaylı not aldığım bir son pdf elimde bulunmaktadır. İsteyenler rahatça ulaşabilir.
İyi günler dilerim.

