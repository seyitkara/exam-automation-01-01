import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from tkinter import Tk, simpledialog

def create_meeting_schedule(konu1, konu2, konu3, output_file):
    # Excel dosyalarını oku
    df1 = pd.read_excel(konu1)
    df2 = pd.read_excel(konu2)
    df3 = pd.read_excel(konu3)

    # Her DataFrame'e dosya adını ekle
    df1['Dosya'] = konu1
    df2['Dosya'] = konu2
    df3['Dosya'] = konu3

    # Tüm DataFrame'leri birleştir ve ortak isimleri bul
    combined_df = pd.concat([df1, df2, df3])
    common_names = combined_df['İsim'].value_counts()[combined_df['İsim'].value_counts() > 1].index.tolist()
    non_conflicting_names = combined_df['İsim'].value_counts()[combined_df['İsim'].value_counts() == 1].index.tolist()
    
    # Kullanıcıdan Başlangıç tarihi ve saati al
    root = Tk()
    root.withdraw()  # Pencereyi gizle
    start_time_str = simpledialog.askstring(title="Başlangıç Zamanı",
                                           prompt="Toplantıların başlangıç tarihini ve saatini YYYY-AA-GG Saat:Dakika formatında girin (örneğin: 2023-11-24 10:00):")
    root.destroy()  # Pencereyi kapat

    # Girilen değeri datetime nesnesine dönüştür
    start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M")

    # Yeni bir Excel dosyası oluştur
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['İsim', 'Dosya', 'Toplantı Saati'])  # Ortak isimler sütunu
    sheet.append([''])  # Boş satır
    
    # Kullanılan isimleri takip etmek için bir set
    used_names = {name: 0 for name in combined_df['İsim'].unique()}
    #used_names = {}

    meeting_schedule = []  # Toplantı programını tutacak liste
    
    # Ortak isimler için toplantı saatleri
    for name in combined_df['İsim'].unique():
        # İsme göre filtrele
        filtered_df = combined_df[combined_df['İsim'] == name]
        
        # Eğer isim daha önce kullanılmadıysa, aynı saat atama
        if name not in used_names:
            for index, row in filtered_df.iterrows():
                sheet.append([name, row['Dosya'], start_time.strftime("%Y-%m-%d %H:%M")])
            used_names[name] = 1  # İsmi kullandık, 1. seans
        else:
            # İsim daha önce kullanıldıysa, bir saat ileri atama
            for index, row in filtered_df.iterrows():
                sheet.append([name, row['Dosya'], (start_time + timedelta(hours=1)).strftime("%Y-%m-%d %H:%M")])
            used_names[name] += 1  # İsim tekrar kullanıldı, seans sayısını artır
        
    # Çakışmayan isimler için toplantı saatleri
    for name in non_conflicting_names:
        # İsme göre filtrele
        filtered_df = combined_df[combined_df['İsim'] == name]

        # Çakışma olup olmadığını kontrol et
        if len(filtered_df) == 1:
            row = filtered_df.iloc[0]
            #sheet.append([name, row['Dosya'], start_time.strftime("%Y-%m-%d %H:%M")])
            # Farklı isimler için saatleri değiştirmiyoruz, aynı saatte katılacaklar
            meeting_schedule.append([name, row['Dosya'], start_time.strftime("%Y-%m-%d %H:%M")])

    # Toplantı programını yazdır
    for entry in meeting_schedule:
        sheet.append(entry)
        
    # Her satıra bir grup ataması yap
    combined_df['Grup'] = combined_df['İsim'] + '_' + combined_df['Dosya']
    # Gruplara göre toplantı saatleri
    groups = combined_df['Grup'].unique()
    for group in groups:
        filtered_df = combined_df[combined_df['Grup'] == group]
        for index, row in filtered_df.iterrows():
            sheet.append([row['İsim'], row['Dosya'], start_time.strftime("%Y-%m-%d %H:%M")])
        start_time += timedelta(hours=1)

    workbook.save(output_file)

# Dosya isimlerini değiştir
konu1 = "konu1.xlsx"
konu2 = "konu2.xlsx"
konu3 = "konu3.xlsx"
output_file = "toplantı_takvimi.xlsx"

#create_meeting_schedule(konu1, konu2, konu3, output_file)
#print("herşey başarılı")