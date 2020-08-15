#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import calendar,random,xlwt,xlrd,os
import PySimpleGUI as sg


çlşDiz=""
yıl,ay=2020,6 #hangi tarih için nöbet hazırlanaca
üyeler={ # üyeler ve mazeret günlerini içeren dict
         'ELİF ÖZCAN':{},
         'YÜKSEL AKGÜNEŞ':{},
         'BEDİRHAN SÜZEN':{},
         'SİNAN CENGİZ':{1,2,3,4,5,6,7,8,9,10},
         'ŞERİFE ÖZSU':{},
         'DEMET HALLAÇ':{},
         'ÖZGEN BAKTIR KARADAŞ':{},
         'PINAR EKMEKÇİOĞLU':{},
         }
buay=calendar.monthrange(yıl,ay) #ayın 1'i haftanın hangi günü ve ayın kaç gün olduğunu döndüren işlev
aysözlük={i+1:0 for i in range(buay[1])} #programın sonunda günlere üye tutacak üye adlarının yazacağı takvim.
nöbetgünleri=("PSÇ","Pe","Cu","Ct","Pa") #gün kümelerinin demeti. aynı öneme sahip olan pazartesi salı çarşamba günleri aynı kümeye alındı.




def büyükHarfli(ad):
 ad=ad.replace("i","İ")
 return ad.upper()

class VT:
 """global değişkenlerin kolay ulaşılabilmesi için class variable olarak atanması"""
 işlenenXLS=set()
 db={}
 çekilenVeri={}
 çıktı={"Üye Kontrol Çıktısı":"",}
 eşleşmeyenÜye=set()
 eşleşmeyenGünler={}
 işlenenGünler={}
 ek=0  #üyelere verilecek ek nöbet sayısı
 aralık=2 #üyenin ne kadar aralıklı nöbet alacağı

 def çıktıEkle(self,key,value):
  self.çıktı[key]=value
 def çıktıDök():
  for i in VT.çıktı:
   print(i+"\n"+VT.çıktı[i])
 
 


#for i in üyeler: #mazeretgün den get error vermemesi için eksik üyeler eklemek
 #if i not in mazeretgün:
  #mazeretgün[i]=[]
  
#üyedengüne={i:{i:0 for i in nöbetgünleri} for i in mazeretgün}
#gündenüyeye={i:{i:0 for i in mazeretgün} for i in nöbetgünleri}
#db={"sinan":{"PSÇ":45,"Pe":35,"Cu":48,"Ct":56,"Pa":32},
    #"pınar":{"PSÇ":31,"Pe":41,"Cu":41,"Ct":46,"Pa":33},
    #"demet":{"PSÇ":15,"Pe":35,"Cu":44,"Ct":51,"Pa":38},
    #"gürkan":{"PSÇ":15,"Pe":35,"Cu":44,"Ct":51,"Pa":38},
    #}
günindeks={0:"PSÇ",1:"PSÇ",2:"PSÇ",3:"Pe",4:"Cu",5:"Ct",6:"Pa"}



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
 

 def __init__(self,gün=0, ):
  self.ek=VT.ek

  if gün:
   self.günküme=self.aydakigünkümesi(gün)
   self.üyelistesi=self.EnazGünSay(self.günküme,VT.db)
   self.nöbetyaz(gün,VT.aralık)
  
  
 def EnazGünSay(self,günküme,db): #üyeler arasında belirtilen günün en az kim tarafından tutulduğunu ve ne kadar tutulduğunu döndürür.
  sıralı=[(db[üye][gün],üye) for üye in db for gün in db[üye] if gün==günküme] #min() fonksiyonu ilk sıradaki girdiye göre sıralar, bunlar eşitse ikinciye göre sıralamaya devam eder
  sıralı.sort() #sort küçükten büyüğe sıralıyor.
  return sıralı
 
 def DB1Arttır(self,üye,günküme): #nöbet yazılan üyenin ilgili gün için nöbet sayısını 1 arttırır.
  VT.db[üye][günküme]+=1
 
 def üyekontrol(self,üye,gün,aralık,ek):
  """Sırasıyla;
  Belirtilen gün üyenin mazeret günü mü,
  üyenin aydaki alacağı en çok nöbet sayısını geçiyor mu,
  belirtilen aralık kadar ileri ve gerideki günlerde nöbeti var mı,
  sorularını boolean olarak yanıtlar."""
  çıktı=VT.çıktı
  ç="Üye Kontrol Çıktısı"
  
  def günaralıkkontrol():
   for i in [i+gün for i in range(-1*aralık,aralık+1) if len(aysözlük)>=i+gün>0]:#verilen günün, belirtilen aralık kadar öncesinden sonrasına kadar nöbeti varmı diye sorgulayan fonksiyon
    if aysözlük[i]==üye:
     return False  
  def nöbetalmasayısıkontrol():#üyenin ay içindeki nöbet sayısı, aydaki gün sayısının üye sayısına bölümüne eşit mi?
   return int(len(aysözlük)/len(üyeler))+ek==[i for i in aysözlük.values()].count(üye) 
  
  if gün in üyeler.get(üye): #üye için mazeret günü mü?
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
      
 def nöbetyaz(self,gün,aralık): #aysözlüke bulunan üyeyi ilgili güne yazmak için 
  for i in self.üyelistesi:
   üye=i[1]
   if self.üyekontrol(üye,gün,VT.aralık,VT.ek):
    aysözlük[gün]=üye
    self.üye=üye
    self.DB1Arttır(üye,self.günküme)
    break


   
 def aydakigünkümesi(self,günsırası): #ayın gününün haftanın hangi günü olduğunu getirir. ayın 1'i salı günü gibi.
  return günindeks[calendar.weekday(yıl,ay,günsırası)]   
 
 def __repr__(self):# işle() komutunun çıktısını belirmek için
  if self.üye:
   return self.üye
  else:
   return "Başarısız"


    
   

def rastgeleİşle():#ay içinden rastgele seçip işleyen
 çıktı=""
 liste=[i for i in aysözlük if not aysözlük[i]] #üye atanmamış günleri süzmek için
 while liste:
  i=random.choice(liste)
  işle(gün=i,)
  liste.remove(i)
 for a in üyeler:
  çıktı+=a+" "+str([i for i in aysözlük.values()].count(a))+"\n"
 çıktı+="Boş günlerin sayısı"+" "+str([i for i in aysözlük.values()].count(0))+"\n"
 VT().çıktıEkle("Üyelere atanan gün sayıları",çıktı)


def okuveyaz(çlşDiz=çlşDiz):
 #global db # daha sonra global kaldırıacak class kullanılacak
 VT.db={üye:{gk:0 for gk in günindeks.values()} for üye in üyeler}
 if not çlşDiz: çlşDiz=os.getcwd()
 for dosya in os.listdir(): #programın bulunduğu dizindeki dosyaların lisetelenmesi
  if ".xls" in dosya: #xls lerin süzülmesi
   if "~" not in dosya:# gizli kurtarma dosyalarını almasın diye
    açılanXLS = xlrd.open_workbook(os.path.join(çlşDiz,dosya))
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
       VT.çekilenVeri[tarih] = üye # xls deki veriyi sözlüğe aktarma
 for i in VT.çekilenVeri: #excellerden çekilen veriyi, üye bazlı sayıp, db sözlüğüne işliyor.
  işlGünKüm=günindeks[calendar.weekday(i[2],i[1],i[0])] #çekilen tarih verisini günkümesine çevirir
  işlÜye=VT.çekilenVeri[i] #çekilen üye verisini kayıt etme
  try:#günümüzde olmayan üyeleri db ye yazarken hata vermesin
   VT.db[işlÜye][işlGünKüm]+=1 
   VT.işlenenGünler[i]=işlÜye
  except:
   VT.eşleşmeyenGünler[i]=işlÜye
   VT.eşleşmeyenÜye.add(işlÜye)
   pass
 



 

def çalıştır():
 çıktı=""
 okuveyaz(çlşDiz=çlşDiz)
 rastgeleİşle()
 VT.ek=1
 rastgeleİşle()
 for i in aysözlük: çıktı+=str(i)+" "+str(aysözlük[i])+"\n"
 VT.çıktı["Sonuç"]=çıktı

def GUI():
 sg.theme('DarkAmber')  
 aralık=VT.aralık
 çıktı=VT.çıktı
 ek=VT.ek
 global ay
 global yıl
 global çlşDiz
 value=""
 
 
 layout = [  [sg.Text('Önceki nöbet listelerinin olduğu dizin:'),sg.Input(),sg.FolderBrowse(key="çlşDiz",tooltip="Klasör seçme penceresi açılır",button_text="Klasör Aç")],
             [sg.InputText(key="yıl",default_text=str(yıl),),sg.InputText(key="ay",default_text=str(ay))], 
             [sg.Text(str(aralık)+' gün aralıkla nöbet verilir')], 
             [sg.Text('Verilebilecek ek nöbet sayısı: '+str(ek))],
             [sg.Text("Üyeler:"),sg.Listbox(üyeler,size=(20,5)),sg.Button(button_text="sil")],
             [sg.Button("Yarat"), sg.Cancel()],
             [sg.Multiline(size=(30,20),key="çıktı",autoscroll=True), sg.Multiline(size=(30,20),key="sonuç",autoscroll=True),sg.Multiline(size=(30,20),key="eşleşmeyen",autoscroll=True)]
             ]
 def tabloGUI(başlık=("başlık_1","başlık_2"),satırsayısı=10,tabloadı="tbl"): #tablo GUIsi
  başlık =  [[sg.Text('  ')] + [sg.Text(h, size=(14,1)) for h in başlık]]
  satırlar = [[sg.Input(key=tabloadı+str(stn)+str(satır),size=(15,1), pad=(0,0)) for stn in range(len(başlık))] for satır in range(satırsayısı)]
  tablo = başlık+satırlar
  return tablo
 
 
 # Create the Window
 window = sg.Window('NöbetGen', layout)
 # Event Loop to process "events"
 while True:
  event, value = window.read()
  if event in (None, 'Cancel'):
   break
  if event=="Yarat": çalıştır()
  ay=value["ay"]
  yıl=value["yıl"]
  çlşDiz=value["çlşDiz"]
  window["çıktı"].update(VT.çıktı)
  window["sonuç"].update(VT.çıktı["Sonuç"])
  window["eşleşmeyen"].update(VT.eşleşmeyenÜye)

  

 return event, window.close()


def xlyaz(ay=ay,yıl=yıl,sz=aysözlük,ünvan="Ecz."):

 style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
 style1 = xlwt.easyxf(num_format_str="dd/mm/yyyy")
 date_format = xlwt.XFStyle()
 date_format.num_format_str = 'dd/mm/yyyy'
 
 wb = xlwt.Workbook()
 ws = wb.add_sheet(str(ay)+str(yıl))
 
 başlık=("Tarih","Gün","Ünvan","Nöbetçi Adı","Yardımcı Personel")
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
 
 wb.save(str(ay)+'.'+str(yıl)+".xls") 
