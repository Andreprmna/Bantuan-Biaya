import xlrd #          untuk membaca file excel
import xlwt #          untuk menulis file excel
#-------------------------------------------Baca file---------------------------------------------------------------------
def read_file():
    workbook = xlrd.open_workbook('Mahasiswa.xls') #membaca file Mahasiswa.xls
    worksheet = workbook.sheet_by_index(0)  #membaca sheet dengan index 0
    penghasilan, pengeluaran = [], []
    for i in range(0,101) :
        penghasilan.append(worksheet.cell(i,1).value)#memasukkan data penghasilan pada excel kedalam array
        pengeluaran.append(worksheet.cell(i,2).value)#memasukkan data pengeluaran pada excel kedalam array
    return(penghasilan,pengeluaran)

#---------------------------------Fuzzyfikasi--------------------------------------------------------------------------------
def sedikit(x,a,b): #fungsi untuk menghitung keanggotaan pada Penghasilan dan Pengeluaran dengan bagian sedikit/kecil
    hasil = 0
    if(x>b):
        hasil = 0
    elif(x<=a):
        hasil = 1
    elif(x>a) and (x<=b):
        hasil = (b-x)/(b-a)
    return hasil

def sedang(x,a,b,c,d): #fungsi untuk menghitung keanggotaan pada Penghasilan dan Pengeluaran dengan bagian sedang/cukup
    hasil =0
    if (x<= a) or (x>d):
        hasil = 0
    elif(x>b) and (x<=c):
        hasil = 1
    elif(x>a) and (x<b): #untuk menghitung segitiga banyak
        hasil = (x-a)/(b-a)
    elif(x>c) and (x<=d): #untuk menghitung segitiga sedikit
        hasil = (d-x)/(d-c)
    return hasil

def banyak(x,c,d): #fungsi untuk menghitung keanggotaan pada Penghasilan dan Pengeluaran dengan bagian banyak/tinggi
    hasil=0
    if (x<= c):
        hasil = 0
    elif(x>d):
        hasil = 1
    elif(x>c) and (x<=d):
        hasil = (x-c)/(d-c)
    return hasil

def keanggotaan_penghasilan(x): # fungsi keanggotaan Penghasilan dengan di bagi menjadi 3 bagian sedikit, sedang, banyak
    aI = 5.70 #I = income /penghasilan
    bI = 8.50 
    cI = 15.70
    dI = 17.90

    sedikitt = sedikit(x,aI,bI)
    sedangg = sedang(x,aI,bI,cI,dI)
    banyakk = banyak(x,cI,dI)
    return [sedikitt,sedangg,banyakk]

def keanggotaan_pengeluaran(y): # fungsi keanggotaan Pengeluaran rate dibagi menjadi 3 bagian rendah, cukup, tinggi
    aO = 3.90 # O = outcome/pengeluaran 
    bO = 6.40
    cO = 8.20
    dO = 9.90

    tinggi = banyak(y,cO,dO)
    cukup = sedang(y,aO,bO,cO,dO)
    kecil = sedikit(y,aO,bO)
    return [tinggi,cukup,kecil]
#---------------------------------------------------------------rule-----------------------------------------------------------------
def output_inferensi(): #rule dalam menentukan inferensi
    pengeluaran = ['rendah','cukup','tinggi']
    penghasilan = ['banyak','sedang','sedikit']
    total = []
    if(pengeluaran[0]=='rendah') and (penghasilan[0]=='banyak'):
        k1 = 'ditolak'
    if(pengeluaran[0]=='rendah') and (penghasilan[1]=='sedang'):
        k2 = 'ditolak'
    if(pengeluaran[0]=='rendah') and (penghasilan[2]=='sedikit'):
        k3 = 'dipertimbangkan'
    if(pengeluaran[1]=='cukup') and (penghasilan[0]=='banyak'):
        y1 = 'ditolak'
    if(pengeluaran[1]=='cukup') and (penghasilan[1]=='sedang'):
        y2 = 'dipertimbangkan'
    if(pengeluaran[1]=='cukup') and (penghasilan[2]=='sedikit'):
        y3= 'diterima'

    if(pengeluaran[2]=='tinggi') and (penghasilan[0]=='banyak'):
        j1 = 'dipertimbangkan'
    if(pengeluaran[2]=='tinggi') and (penghasilan[1]=='sedang'):
        j2 = 'diterima'
    if(pengeluaran[2]=='tinggi') and (penghasilan[2]=='sedikit'):
        j3 = 'diterima'
        total.append(k1)
        total.append(k2)
        total.append(k3)
        total.append(y1)
        total.append(y2)
        total.append(y3)
        total.append(j1)
        total.append(j2)
        total.append(j3)
    return total

#|                                 fuzzy rule                                           |
#---------------------------------------------------------------------------------------#
#Pengeluaran\Penghasilan |       Banyak     |      Sedang     |     Sedikit             |#
#---------------------------------------------------------------------------------------#
#          rendah        |   ditolak        |      ditolak    |     dipertimbangkan    |#
#          cukup         |   ditolak        | dipertimbangkan |     diterima           |#
#          tinggi        |   dipertimbangkan|   diterima      |     diterima           |#
#--------------------------------------------------------------------------------------|#



# ------------------------------------------------------Inferensi---------------------------------------------------------------------
def inferensi_Value(pend,peng): #fungsi menentuka nilai inferensi yang nilainya akan digunakan pada fungsi defuzzyficatin dengan metode sugeno
    hasil = []
    income = pend
    outcome = peng

    for i in income:
        for j in outcome:
            if(i<j):
                hasil.append(i) #menyimpan i karena lebih kecil
            else:
                hasil.append(j) # menyimpan j karena j lebih kecil dari i
    ba1 = hasil[0]
    ba2 = hasil[1]
    ba3 = hasil[3]
    ditolak = (max(ba1,ba2,ba3))#ambil nilai max dari inferensi ditolak

    bu1 = hasil[2]
    bu2 = hasil[4]
    bu3 = hasil[6]
    dipertimbangkan= (max(bu1,bu2,bu3))#ambil nilai max dari inferensi dipertimbangkan

    bi1 = hasil[5]
    bi2 = hasil[7]
    bi3 = hasil[8]
    diterima = (max(bi1,bi2,bi3)) #mengambil nilai max dari inferensi diterima
    return (diterima,dipertimbangkan,ditolak) # mendapatkan hasil inferensi diterima, dipertimbangkan, dan ditolak

#keanggotaan fuzzy model sugeno
#
#       ditolak  dipertimbangkan    diterima
#  1 |     |           |               |
#    |     |           |               |
#0.5 |     |           |               |
#    |     |           |               |
#  0 |_____|___________|_______________|____
#        25           50               75  
# 
#-------------------------------------------------------------defuzzyfication--------------------------------------
def defuzzyfication (a,b,c): #defuzzyfication sugeno
        Total = (((a*25) + (b*50) + (c*75)) / (a+b+c))
        return Total
#-------------------------------------------------------------Save File--------------------------------------------------    
def simpan(tampung): #fungsi untuk menyimpan hasil outputan ke dalam format xls denga nama file Bantuan.xls
        storage = []
        storage = tampung
        ExcelBaru = xlwt.Workbook()
        SheetBaru = ExcelBaru.add_sheet('Bantuan')
        for i in range(0,len(storage)):
            SheetBaru.write(i, 0, (storage[i]))
        ExcelBaru.save('Bantuan.xls')
#------------------------------------------MAIN PROGRAM----------------------------------------------------
income,outcome = read_file()
total = output_inferensi()
print(total)
jawaban, menampung = [], []
for i in range(1,len(income)):
    hasil_penghasilan = keanggotaan_penghasilan(income[i])
    hasil_pengeluaran = keanggotaan_pengeluaran(outcome[i])
    j,k,l = inferensi_Value(hasil_penghasilan,hasil_pengeluaran)
    hasil_defuzzy = defuzzyfication(j,k,l)
    jawaban.append([hasil_defuzzy,i])
    jawaban.sort(reverse=True) #melakukan sorting dengan nilai yang terbesar hingga terkecil pada rate penerima bantuan
for x in range(0,20):
    menampung.append(jawaban[x][1])
    print('No penerima Bantuan Registrasi :',jawaban[x][1],'Rekomendasi : ',jawaban[x][0])
menampung.sort() #melakukan sorting nilai terkecil hingga terbesar pada no id penerima bantuan
print('Seleksi Suksesss dan sudah terurut dari kecil ke besar pada file')
simpan(menampung)

