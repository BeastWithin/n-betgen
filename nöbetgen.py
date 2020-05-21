
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import calendar,random,xlrd,os

class değişkenler:
  

yıl=2019
ay=1
ek=0 #üyelere verilecek ek nöbet sayısı
aralık=2 #üyenin ne kadar aralıklı nöbet alacağı
üyeler={'Pınar Bilgin', 'Yüksel Akgüneş', 'Gürkan Emre', 'Sinan Cengiz', 'Tayyar Uysal', 'Demet Hallaç', 'Özgen Baktır Karadaş', 'Pınar Ekmekçioğlu'}
buay=calendar.monthrange(yıl,ay) #ayın 1 haftanın hangi günü ve ayın kaç gün olduğunu döndüren işlev
aysözlük={i+1:0 for i in range(buay[1])} #programın sonunda günlere üye tutacak üye adlarının yazacağı takvim.
#aysöz2={günindeks[calendar.monthcalendar(yıl,ay)[0].index(1)]}
nöbetgünleri=("PSÇ","Pe","Cu","Ct","Pa") #gün kümelerinin demeti. aynı öneme sahip olan pazartesi 1salı çarşamba günleri aynı kümeye alındı.
mazeretgün={"Sinan Cengiz":[16,17,18,19],
             } # üyelerin mazeret günlerini içeren dict

for i in üyeler: #mazeretgün den get error vermemesi için eksik üyeler eklemek
 if i not in mazeretgün:
  mazeretgün[i]=[]
  
üyedengüne={i:{i:0 for i in nöbetgünleri} for i in mazeretgün}
gündenüyeye={i:{i:0 for i in mazeretgün} for i in nöbetgünleri}
db={"sinan":{"PSÇ":45,"Pe":35,"Cu":48,"Ct":56,"Pa":32},
    "pınar":{"PSÇ":31,"Pe":41,"Cu":41,"Ct":46,"Pa":33},
    "demet":{"PSÇ":15,"Pe":35,"Cu":44,"Ct":51,"Pa":38},
    "gürkan":{"PSÇ":15,"Pe":35,"Cu":44,"Ct":51,"Pa":38},
    }
günindeks={0:"PSÇ",1:"PSÇ",2:"PSÇ",3:"Pe",4:"Cu",5:"Ct",6:"Pa"}

class günlerinsayısı: #belirtilmiş mazeret günleri çıkarıldığında kalan takvimde alabileceği günleri verir
 def __init__(self,kişi,ay,yıl):
  self.ay=int(ay)
  self.yıl=int(yıl)
  self.kişi=kişi
  self.mazeretgün=mazeretgün.get(kişi)
  self.sözlük={günküme:[] for günküme in nöbetgünleri}
  self.yürüt()
 def sözlük(self):
  return {self.PSÇ}
 def yürüt(self):
  for gün in aysözlük:
   if aysözlük[gün]==False:
    if gün not in self.mazeretgün:
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
 
 def __init__(self,gün=0, aralık=aralık,ek=ek):
  
  self.üye=0
  if gün:
   self.günküme=self.aydakigünkümesi(gün)
   self.üyelistesi=self.EnazGünSay(self.günküme)
   self.nöbetyaz(gün,aralık)
  
  
 def EnazGünSay(self,günküme): #üyeler arasında belirtilen günün en az kim tarafından tutulduğunu ve ne kadar tutulduğunu döndürür.
  sıralı=[(db[üye][gün],üye) for üye in db for gün in db[üye] if gün==günküme] #min() fonksiyonu ilk sıradaki girdiye göre sıralar, bunlar eşitse ikinciye göre sıralamaya devam eder
  sıralı.sort() #sort küçükten büyüğe sıralıyor.
  return sıralı
 
 def DB1Arttır(self,üye,günküme): #nöbet yazılan üyenin ilgili gün için nöbet sayısını 1 arttırır.
  db[üye][günküme]+=1
 
 def üyekontrol(self,üye,gün,aralık,):
  """Sırasıyla;
  Belirtilen gün üyenin mazeret günü mü,
  üyenin aydaki alacağı en çok nöbet sayısını geçiyor mu,
  belirtilen aralık kadar ileri ve gerideki günlerde nöbeti var mı,
  sorularını boolean olarak yanıtlar."""
  def günaralıkkontrol():
   for i in [i+gün for i in range(-1*aralık,aralık+1) if len(aysözlük)>=i+gün>0]:#verilen günün, belirtilen aralık kadar öncesinden sonrasına kadar nöbeti varmı diye sorgulayan fonksiyon
    if aysözlük[i]==üye:
     return False  
  def nöbetalmasayısıkontrol():#üyenin ay içindeki nöbet sayısı, aydaki gün sayısının üye sayısına bölümüne eşit mi?
   print("ek="+str(ek))
   return int(len(aysözlük)/len(üyeler))+ek==[i for i in aysözlük.values()].count(üye) 
  
  if gün in mazeretgün.get(üye): #üye için mazeret günü mü?
   print(str(gün)+" için, "+str(üye)+" nin mazeret günü")
   return False
  elif nöbetalmasayısıkontrol():
   print(str(gün)+" için, "+str(üye)+" en fazla gün sayısına ulaşmış")
   return False
  elif günaralıkkontrol()==False:
   print(str(gün)+" için, "+str(aralık)+" gün içinde "+str(üye)+" yazılmış")
   return False
  else:
    return True
   
   
 def nöbetyaz(self,gün,aralık): #aysözlüke bulunan üyeyi ilgili güne yazmak için 
  for i in self.üyelistesi:
   üye=i[1]
   if self.üyekontrol(üye,gün,aralık):
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


    
   

def çalıştır():#ay içinden rastgele seçip işleyen
 liste=[i for i in aysözlük if not aysözlük[i]] #üye atanmamış günleri süzmek için
 while liste:
  i=random.choice(liste)
  işle(gün=i,aralık=aralık)
  liste.remove(i)
 for a in üyeler:
  print(a+" "+str([i for i in aysözlük.values()].count(a))) 
 print("Boş günlerin sayısı"+str([i for i in aysözlük.values()].count(0)))
  

işlenenexcel=set()

def okuveyaz():
 global db# daha sonra global kaldırıacak class kullanılacak
 db={üye:{gk:0 for gk in günindeks.values()} for üye in üyeler}
 global işlenenexcel
 sözlük = {}
 cwd=os.getcwd()
 for dosya in os.listdir(): #programın bulunduğu dizindeki dosyaların lisetelenmesi
  if ".xls" in dosya: #xls lerin süzülmesi
   if "~" not in dosya:# gizli kurtarma dosyalarını almasın diye
    işlenenexcel.add(dosya) #hangi xls lerin işlendiğine sonra bakabilmek için
    file_handle = xlrd.open_workbook(os.path.join(cwd,dosya))
    sheet = file_handle.sheet_by_index(0) # xls deki ilk sayfaya odaklanmak
    
    for i in range(1,sheet.nrows):#ilk satırı atlayarak satırları ele almak
     line=sheet.row_values(i)
     t=xlrd.xldate_as_tuple(line[0],0) #xldate, exceldeki tarih damgasını tarihe çeviriyor.
     tarih=t[2::-1]# tarihin sırasını düzeltme
     üye=line[3]
     sözlük[tarih] = üye # xls deki veriyi sözlüğe aktarma
 for i in sözlük: #excellerden çekilen veriyi, üye bazlı sayıp, db sözlüğüne işliyor.
  try: #günümüzde olmayan üyeleri db ye yazarken hata vermesin
   db[sözlük[i]][günindeks[calendar.weekday(i[2],i[1],i[0])]]+=1 
  except:
   pass


 
def deneme():
 okuveyaz()
 çalıştır()
 global ek
 ek=1
 çalıştır()
 for i in aysözlük: print(str(i)+" "+aysözlük[i])
  
