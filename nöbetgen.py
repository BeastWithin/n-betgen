#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import calendar,random,xlwt,xlrd,os
import PySimpleGUI as sg


yıl,ay=2020,6 #hangi tarih için nöbet hazırlanacağı
üyeler={ # üyeler ve mazeret günlerini içeren dict
         #'ELİF ÖZCAN':{},
         #'YÜKSEL AKGÜNEŞ':{},
         #'BEDİRHAN SÜZEN':{},
         #'SİNAN CENGİZ':{1,2,3,4,5,6,7,8,9,10},
         #'ŞERİFE ÖZSU':{},
         #'DEMET HALLAÇ':{},
         #'ÖZGEN BAKTIR KARADAŞ':{},
         #'PINAR EKMEKÇİOĞLU':{},
         }

buay=calendar.monthrange(yıl,ay) #ayın 1'i haftanın hangi günü ve ayın kaç gün olduğunu döndüren işlev
_aysözlük={i+1:0 for i in range(buay[1])} #programın sonunda günlere üye tutacak üye adlarının yazacağı takvim.
nöbetgünleri=("PSÇ","Pe","Cu","Ct","Pa") #gün kümelerinin demeti. aynı öneme sahip olan pazartesi salı çarşamba günleri aynı kümeye alındı.
günindeks={0:"PSÇ",1:"PSÇ",2:"PSÇ",3:"Pe",4:"Cu",5:"Ct",6:"Pa"}
dizindekiXLSler=[dosya for dosya in os.listdir(os.getcwd()) if ".xls" in dosya if "~" not in dosya]

def büyükHarfli(ad):
 ad=ad.replace("i","İ")
 return ad.upper()

class VeriTabanı:
 """global değişkenlerin kolay ulaşılabilmesi için class variable olarak atanması"""
 çlşDiz=os.getcwd() 
 def __init__(self):
  self.işlenenXLS=set()
  self.db={}
  self.çekilenVeri={}
  self.çıktı={"Üye Kontrol Çıktısı":"",
                 #"Sonuç":"",
                 }
  self.eşleşmeyenÜye=set()
  self.eşleşmeyenGünler={}
  self.işlenenGünler={}
  self.ek=0  #üyelere verilecek ek nöbet sayısı
  self.aralık=2 #üyenin ne kadar aralıklı nöbet alacağı
  self.aysözlük={i+1:0 for i in range(buay[1])} #programın sonunda günlere üye tutacak üye adlarının yazacağı takvim.

 
 
 
 def çıktıEkle(self,key,value):
  self.çıktı[key]=value
 def çıktıDök(self):
  for i in VT.çıktı:
   return str(i+"\n"+VT.çıktı[i])
 def istatistikDök(self):
  d=""
  for üye in VT.db:
   d+=str(üye)+"\n"
   for gün in VT.db[üye]:
    d+=str(str(gün)+": "+str(VT.db[üye][gün]))+"\t"
   d+="\n"
  return d
 def çıktıDökdüzenli(self,ç):
  p=""
  ç=list(ç)
  ç.sort()
  for i in ç:
   p+=str(i)+"\n"
  return p
 def çıktıDökSonuç(self,s): #aysözlük çıktısını GUI de görebilmek için
  ç=""
  for k in s: ç+=str(k)+":\t"+str(s[k])+"\n"
  return ç
 
 
 
VT=VeriTabanı()


#db={"sinan":{"PSÇ":45,"Pe":35,"Cu":48,"Ct":56,"Pa":32},
    #"pınar":{"PSÇ":31,"Pe":41,"Cu":41,"Ct":46,"Pa":33},
    #"demet":{"PSÇ":15,"Pe":35,"Cu":44,"Ct":51,"Pa":38},
    #"gürkan":{"PSÇ":15,"Pe":35,"Cu":44,"Ct":51,"Pa":38},
    #}



class günlerinsayısı: #belirtilmiş mazeret günleri çıkarıldığında kalan takvimde alabileceği günleri verir
 def __init__(self,kişi,ay,yıl):
  self.ay=int(ay)
  self.yıl=int(yıl)
  self.kişi=kişi
  self.üyeler=üyeler.get(kişi)
  self.sözlük={günküme:[] for günküme in nöbetgünleri}
  self.yürüt()
 def sözlük(self):
  return {self.PSÇ}
 def yürüt(self):
  for gün in aysözlük:
   if aysözlük[gün]==False:
    if gün not in self.üyeler:
     hg=aydakigünkümesi(gün)
     self.sözlük[hg].append(gün)
     
def EnAzGünKümesiniBulma():
 enaz=min([i for a in db for i in db[a].values()]) #üyelerin içinde en az değeri olan gün kümesini bulur
 for üye in db:#en az değeri olan kümenin sahibi üyeyi ve kümeyi bulmak için, üyeleri sıralar
   for gün in db[üye]: #üyelerin günlerini sırala
    if db[üye][gün]==enaz: #ilgili gün kümesini bulur
     sayı=db[üye].pop(gün) #ilgili üyenin ilgili gün kümesini "sayı" ya kaydedip siler
     return [üye,gün,sayı]



class işle: #belirtilen günkümesini en az tutmuş üyeyi bulup,o üyeyi takvimde ilgili güne yazar, o güne 1 ekler
 

 def __init__(self,
              gün,
              aralık,
              db,
              aysözlük,
              ekgünsayısı,
              ):
  self.ek=ekgünsayısı
  self.üye=0
  if gün:
   self.günküme=self.aydakigünkümesi(gün)
   self.üyelistesi=self.EnazGünSay(self.günküme,db)
   self.nöbetyaz(gün,aralık,aysözlük)
  else:
   return print("Gün sırası belirtilmeli")
  
 def EnazGünSay(self,günküme,db): #üyeler arasında belirtilen günün en az kim tarafından tutulduğunu ve ne kadar tutulduğunu döndürür.
  sıralı=[(db[üye][gün],üye) for üye in db for gün in db[üye] if gün==günküme] #min() fonksiyonu ilk sıradaki girdiye göre sıralar, bunlar eşitse ikinciye göre sıralamaya devam eder
  sıralı.sort() #sort küçükten büyüğe sıralıyor.
  return sıralı
 
 def DB1Arttır(self,üye,günküme): #nöbet yazılan üyenin ilgili gün için nöbet sayısını 1 arttırır.
  VT.db[üye][günküme]+=1
 
 def üyekontrol(self,üye,gün,aralık,ek,aysözlük):
  """Sırasıyla;
  Belirtilen gün üyenin mazeret günü mü,
  üyenin aydaki alacağı en çok nöbet sayısını geçiyor mu,
  belirtilen aralık kadar ileri ve gerideki günlerde nöbeti var mı,
  sorularını boolean olarak yanıtlar."""
  çıktı=VT.çıktı
  ç="Üye Kontrol Çıktısı"
  
  def günaralıkkontrol(aysözlük=aysözlük):
   for i in [i+gün for i in range(-1*aralık,aralık+1) if len(aysözlük)>=i+gün>0]:#verilen günün, belirtilen aralık kadar öncesinden sonrasına kadar nöbeti varmı diye sorgulayan fonksiyon
    if aysözlük[i]==üye:
     return False  
  def nöbetalmasayısıkontrol(aysözlük=aysözlük):#üyenin ay içindeki nöbet sayısı, aydaki gün sayısının üye sayısına bölümüne eşit mi?
   return int(len(aysözlük)/len(üyeler))+ek==[i for i in aysözlük.values()].count(üye) 
  if gün in üyeler.get(üye): #üye için mazeret günü mü? MAZERET KONTROL
   çıktı[ç]+=str(gün)+". gün için, "+str(üye)+" nin mazeret günü\n"
   return False
  elif nöbetalmasayısıkontrol():
   çıktı[ç]+=str(gün)+". gün için, "+str(üye)+" en fazla gün sayısına ulaşmış\n"
   return False
  elif günaralıkkontrol()==False:
   çıktı[ç]+=str(gün)+". gün için, "+str(aralık)+" gün içinde "+str(üye)+" yazılmış\n"
   return False
  else:
    return True
      
 def nöbetyaz(self,gün,aralık,aysözlük): #aysözlüke bulunan üyeyi ilgili güne yazmak için 
  for i in self.üyelistesi:
   üye=i[1]
   if self.üyekontrol(üye,gün,VT.aralık,VT.ek,aysözlük):
    aysözlük[gün]=üye
    self.DB1Arttır(üye,self.günküme)
    break



 def aydakigünkümesi(self,günsırası): #ayın gününün haftanın hangi günü olduğunu getirir. ayın 1'i salı günü gibi.
  return günindeks[calendar.weekday(yıl,ay,günsırası)]   
 
 def __repr__(self):# işle() komutunun çıktısını belirmek için
  if self.üye:
   return self.üye
  else:
   return "Başarısız"



def rastgeleİşle(aysözlük):#ay içinden rastgele seçip işleyen
 çıktı=""
 liste=[i for i in aysözlük if not aysözlük[i]] #üye atanmamış günleri süzmek için
 while liste:
  i=random.choice(liste)
  işle(i,VT.aralık,VT.db,aysözlük,VT.ek)
  liste.remove(i)
 for a in üyeler:
  çıktı+=a+" "+str([i for i in aysözlük.values()].count(a))+"\n"
 çıktı+="Boş günlerin sayısı"+" "+str([i for i in aysözlük.values()].count(0))+"\n"
 VT.çıktıEkle("Üyelere atanan gün sayıları",çıktı)


class okuveyaz:

 def XLSoku(VT,dosya):
  çekilenveri={}
  açılanXLS = xlrd.open_workbook(os.path.join(VT.çlşDiz,dosya))
  açılanXLS = açılanXLS.sheet_by_index(0) # xls deki ilk sayfaya odaklanmak
  VT.işlenenXLS.add(dosya) #hangi xls lerin işlendiğine sonra bakabilmek için 
  for satırnumarası in range(açılanXLS.nrows):#ilk satırı atlayarak satırları ele almak
   if any(açılanXLS.row_values(satırnumarası)): #boş satırları geçmek
    tarih,üye="",""      
    for i in açılanXLS.row_values(satırnumarası):
     if type(i) is float: #liste öğesi tarih mi sorgusu
      tarih=xlrd.xldate_as_tuple(i,0) #xldate, exceldeki tarih damgasını tarihe çeviriyor.
      tarih=tarih[2::-1]# tarihin sırasını düzeltme
     elif type(i) is str: #liste öğesi kişi adı mı sorgusu
      if i.split().__len__()>1: #"birden fazla kelime ise kişi adıdır" mantığı
       if not üye: #ilk saptadığı adı alması için
        üye=büyükHarfli(i) #o tarihe yazılı üyeyi alma
    if tarih and üye:
     çekilenveri[tarih] = üye # xls deki veriyi sözlüğe aktarma
  return çekilenveri
 def vtYaz(VT,çekilenveri):
  for i in çekilenveri: #excellerden çekilen veriyi, üye bazlı sayıp, db sözlüğüne işliyor.
   işlGünKüm=günindeks[calendar.weekday(i[2],i[1],i[0])] #çekilen tarih verisini günkümesine çevirir
   işlÜye=çekilenveri[i] #çekilen üye verisini kayıt etme
   try:#günümüzde olmayan üyeleri db ye yazarken hata vermesin
    VT.db[işlÜye][işlGünKüm]+=1 
    VT.işlenenGünler[i]=işlÜye
   except:
    VT.eşleşmeyenGünler[i]=işlÜye
    VT.eşleşmeyenÜye.add(işlÜye)
    pass
 




def çalıştır(aysözlük=VT.aysözlük):
 global VT
 os.chdir(VT.çlşDiz) 
 VT=VeriTabanı() #çalıştır fonksiyonu her çalıştırıldığında yeni tarama yapması için yeni örnek üretiyor.
 #!!!!!!!!!!!!!!!! ama şimdi çalışmıyor
 VT.db={üye:{gk:0 for gk in günindeks.values()} for üye in üyeler} 
 çıktı=""
 for dosya in dizindekiXLSler: VT.çekilenVeri.update(okuveyaz.XLSoku(VT,dosya)) # geçmiş tarih-nöbetçi bilgisini dizindeki XLS lerden VT.çekilenVeri ye kaydeder
 okuveyaz.vtYaz(VT,VT.çekilenVeri)
 rastgeleİşle(VT.aysözlük)
 VT.ek=1
 rastgeleİşle(VT.aysözlük)
 #for i in aysözlük: çıktı+=str(i)+" "+str(aysözlük[i])+"\n"
 #VT.çıktı["Sonuç"]=çıktı
 
def kaydet():
 xlyaz()


def xlyaz(ay=ay,
          yıl=yıl,
          sz=VT.aysözlük,
          ünvan="Ecz.",
          başlık=("Tarih","Gün","Ünvan","Nöbetçi Adı","Yardımcı Personel"),
          ):

 style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
 style1 = xlwt.easyxf(num_format_str="dd/mm/yyyy")
 date_format = xlwt.XFStyle()
 date_format.num_format_str = 'dd/mm/yyyy'  #ÇALIŞMIYOR bakılacak
 
 wb = xlwt.Workbook()
 ws = wb.add_sheet(str(ay)+str(yıl))
 
 def sütunOluştur(ws,sütunNo,liste,stil):
  for n,girdi in enumerate(liste):
   ws.write(sütunNo,n,girdi,stil)

 #sütunOluştur(ws,0,sz,style1)
 #sütunOluştur(ws,0,sz,style1)
 
 #sütunOluştur(ws,0,sz.values(),style1)
 __günAdı={0:"Pazartesi",1:"Salı",2:"Çarşamba",3:"Perşembe",4:"Cuma",5:"Cumartesi",6:"Pazar"}
  
 def satırOluştur(ws,satırNo,liste,):
  for n,girdi in enumerate(liste):
   ws.write(satırNo,n,girdi)
  
 satırOluştur(ws,0,başlık)
 for g in sz:
  satırOluştur(ws,g,(calendar.datetime.date(yıl,ay,g),__günAdı[calendar.datetime.date(yıl,ay,g).weekday()],ünvan,sz[g],))

 #ws.write(0, 0, 1234.56, style0)
 #ws.write(1, 0, datetime.now(), style1)
 #ws.write(2, 0, 1)
 #ws.write(2, 1, 1)
 #ws.write(2, 2, xlwt.Formula("A3+B3"))
 ws.col(0).set_style(date_format)
 dosyaadı=str(yıl)[2:4]+"{0:0=2d}".format(ay)+".xls"
 wb.save(dosyaadı) 
 
class XLSyaz: #işlemiyor
 """XLS dosyalarını yazan sınıf"""
 __günAdı={0:"Pazartesi",1:"Salı",2:"Çarşamba",3:"Perşembe",4:"Cuma",5:"Cumartesi",6:"Pazar"}
 
 def __init__(self,
              ay=ay,
              yıl=yıl,
              sz=VT.aysözlük,
              ünvan="Ecz.",
              başlık=("Tarih","Gün","Ünvan","Nöbetçi Adı","Yardımcı Personel"),
              ):
  #style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
  #style1 = xlwt.easyxf(num_format_str="dd/mm/yyyy")
  #self.date_format = xlwt.XFStyle()
  #self.date_format.num_format_str = 'dd/mm/yyyy'  #ÇALIŞMIYOR bakılacak
  #ws.col(0).set_style(self.date_format)
  self.dosyaadı=str(yıl)[2:4]+"{0:0=2d}".format(ay)+".xls"
  self.başlık=başlık
  self.wb = xlwt.Workbook()
  self.ws = self.wb.add_sheet(str(ay)+str(yıl))
  self.sayfaOluştur(self.ws)
  self.kaydet
  
 def sütunOluştur(self,sütunNo,liste,stil): #KULLANIMDA DEĞİL. tarih formatına çözüm olarak yapmıştım olmadı.
  for n,girdi in enumerate(liste):
   self.ws.write(sütunNo,n,girdi,stil)

 #sütunOluştur(ws,0,sz,style1)
 #sütunOluştur(ws,0,sz,style1)
 #sütunOluştur(ws,0,sz.values(),style1)
 
 def satırOluştur(self,satırNo,liste,):
  for n,girdi in enumerate(liste):
   self.ws.write(satırNo,n,girdi)
  
 def sayfaOluştur(self,ws):
  self.satırOluştur(ws,0,self.başlık)
  for g in sz:
   self.satırOluştur(ws,g,(calendar.datetime.date(yıl,ay,g),XLSyaz.__günAdı[calendar.datetime.date(yıl,ay,g).weekday()],ünvan,sz[g],))
 
 def kaydet(self):
  wb.save(dosyaadı)  
  
 #ws.write(0, 0, 1234.56, style0)
 #ws.write(1, 0, datetime.now(), style1)
 #ws.write(2, 0, 1)
 #ws.write(2, 1, 1)
 #ws.write(2, 2, xlwt.Formula("A3+B3"))
 
 
class GUI:
 """Arayüz sınıfı"""
 def __init__(self,
              sg=sg,
              ay=ay,
              yıl=yıl,
              çıktı=VT.çıktı,
              ek=VT.ek,
              ekle="",
              sil="",
              ):
  global üyeler
  sg.theme('DarkAmber')   
  self.aralık=VT.aralık
  self.ay=ay
  self.yıl=yıl
  self.çıktı=çıktı
  self.ek=ek
  self.sonXLS=max(dizindekiXLSler)
  sonüyeler=okuveyaz.XLSoku(VT,self.sonXLS)#son listeyi belirlemek için
  üyeler={i:{} for i in sonüyeler.values()}#son listedeki üyeleri belirlemek için
  self.window = sg.Window('NöbetGen', self.layout())
  self.eventLoop()
 
 def üyeTablosu(self,üyeler,çerçevebaşlığı):#
  üyetablosu="[sg.T('Üyeler, {} adlı listeden çekildi.')],".format(self.sonXLS)
  for üye in üyeler:
   a="[sg.I(key='{}',default_text='{}'),sg.I(key='mazeret{}',default_text='{}'),sg.Btn('Sil',key='silüye'),],".format(üye,üye,üye,üyeler[üye],üye)
   üyetablosu+=a
  üyetablosu+="[sg.B('ekle')],[sg.B('Mazeret Kaydet')]"
  
  return üyetablosu 
 def layout(self):
  üyelersekme=eval(self.üyeTablosu(üyeler,"Üyeler"))
  ayarlarsekme=[
   [sg.T('Önceki nöbet listelerinin olduğu dizin:'),sg.I(default_text=VT.çlşDiz, key="çlşDiz"),sg.FolderBrowse(tooltip="Klasör seçme penceresi açılır",button_text="Klasör Aç")],
   [sg.I(key="yıl",default_text=str(yıl), size=(4,4),), sg.T(" yılı "), sg.I(key="ay",default_text=str(ay), size=(2,4),), sg.T(" ayı için nöbet oluşturulacak."),], 
   [sg.I(key="aralık",default_text=str(self.aralık), size=(1,4),), sg.T(' gün aralıkla nöbet verilir')], 
   [sg.T('Verilebilecek ek nöbet sayısı: '), sg.I(key="ek",default_text=str(self.ek), size=(1,4),)],
   ]
  işlenentablolartab= [[sg.Multiline(size=(30,20),key="işlenentablo",autoscroll=True)]]
  çıktıtab= [[sg.Multiline(size=(30,20),key="çıktı",autoscroll=True)]]
  sonuçtab= [[sg.Multiline(size=(30,20),key="sonuç",autoscroll=False)]]
  eşleşmeyentab= [[sg.Multiline(size=(30,20),key="eşleşmeyen",autoscroll=True)]] 
  istatistik=[[sg.T(VT.db,key="istatistik",size=(90,20))]]
  sekmegurubu=[sg.TabGroup(
    [[
    sg.Tab('AYARLAR', ayarlarsekme), 
    sg.Tab('ÜYELER', üyelersekme),
    sg.Tab('İŞLENEN TABLOLAR', işlenentablolartab),
    sg.Tab('ÇIKTI', çıktıtab), 
    sg.Tab('SONUÇ', sonuçtab), 
    sg.Tab('EŞLEŞMEYEN', eşleşmeyentab),
    sg.Tab('İSTATİSTİK', istatistik),
    
    ]]
  )]
  eylembutonları=[sg.B("Üret"), sg.B("Kaydet"),
                   #sg.B("Çıkış"),
                   ]
 
  layout = [sekmegurubu,eylembutonları]
  return layout
 
 def eventLoop(self):
  def yenile():
   window["çıktı"].update(VT.çıktıDök())
   window["sonuç"].update(VT.çıktıDökSonuç(VT.aysözlük))
   window["eşleşmeyen"].update(VT.eşleşmeyenÜye)
   window["işlenentablo"].update(VT.çıktıDökdüzenli(VT.işlenenXLS))
   window["istatistik"].update(VT.istatistikDök())   
   
  window=self.window
  while True:
   event, value = window.read()
   print(event)
   print(value)
   if event in (None,'Çıkış'):
    return event, window.close()
   if event=="Üret":
    çalıştır()
    yenile()
   if event=="Kaydet":  kaydet()
   if event=="Mazeret Kaydet":
    for i in value:
     if type(i)==str:
      if "mazeret" in i:
       üyeler[i.strip("mazeret")]=eval(value[i])
   
   ay=value["ay"]
   yıl=value["yıl"]
   aralık=value["aralık"]
   ek=value["ek"]
   
   
   VT.çlşDiz=value["çlşDiz"]



if __name__ == "__main__":
 GUI()
