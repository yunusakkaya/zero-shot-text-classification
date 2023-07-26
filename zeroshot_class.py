# kütüphaneler
import pandas as pd
from transformers import pipeline
import xlsxwriter as xw

# sınaflandırıcı
classifier = pipeline("zero-shot-classification",model="facebook/bart-large-mnli")

# veri çekme
df= pd.read_excel(r"New Microsoft Excel Worksheet.xlsx")
df.to_csv(r"C:/Users/yunus/Desktop/ıvır zıvır\cmc_speechdata.csv", index = None, header=True)

# doldurmak üzere boş excel dosyası oluşturma
workbook = xw.Workbook('cmc_labeled_data.xlsx')
worksheet = workbook.add_worksheet()

# excel dosyasına başlıkları koyma
worksheet.write(0,0,"Id")
worksheet.write(0,1,"Text")
worksheet.write(0,2,"Label")

scores_sum=0

# sınıflandırma işlemi
for i in range(100):
    
    # müşteri verilerinde gez
    sequence_to_classify = df["CustomerText"][i]
    
    # dağıtalacak etiketler
    candidate_labels = ["Bayi","Bilgi","Değişim","İade","Garanti","Kampanya","Kurulum","Servis","Yetkili Çağrısı","Şikayet", "Diğer"]
    
    # sonuçları al
    results = classifier(sequence_to_classify, candidate_labels)
    
    # sonuçlar içinde etiket verisini al
    label= results['labels'][(results['scores'].index(max(results['scores'])))]
    scores_sum= scores_sum + max(results['scores'])
    score= scores_sum / (i+1)
    # indekse göre etiketleri excel dosyasına yazdır
    worksheet.write(i+1,0,df["Id"][i])
    worksheet.write(i+1,1,df["CustomerText"][i])
    worksheet.write(i+1,2,label)
    print("Score:",score)

# dosyayı kapat    
workbook.close()    
    
    
               
        
