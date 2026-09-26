import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import datetime
import pytesseract
from PIL import Image
import re
import shutil

st.set_page_config(page_title="Ferçim Rapor", page_icon="🏭", layout="centered")
st.title("🏭 Ferçim Niğde - Otomatik Rapor")

rapor_tarihi = st.date_input("🗓️ Rapor Tarihi", datetime.date.today())
formatli_tarih = rapor_tarihi.strftime("%d.%m.%Y")
gun_ismi = str(rapor_tarihi.day)

uploaded_files = st.file_uploader("📸 SAP Ekran Görüntülerini Seçin", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

def parse_sap_image(image_file):
    # Fotoğrafı metne çevir
    img = Image.open(image_file)
    text = pytesseract.image_to_string(img, lang='tur+eng')
    return text.upper()

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} görüntü yüklendi. OCR analizi başlıyor...")
    
    if st.button("🚀 Raporu Oluştur"):
        with st.spinner('Fotoğraflar okunuyor ve veriler ayıklanıyor...'):
            try:
                # Tüm yüklenen fotoğraflardaki metinleri birleştir
                tum_metin = ""
                for file in uploaded_files:
                    tum_metin += parse_sap_image(file)
                
                # Şablonu bozmamak için orijinali koruyarak açıyoruz
                template_path = "Niğde Günlük Rapor 2026_09_24.xlsx"
                wb = load_workbook(template_path)
                ws = wb.worksheets[0]
                
                ws.title = gun_ismi
                ws['A2'] = "" 
                ws['B2'] = formatli_tarih 
                
                # --- DİNAMİK VERİ AYIKLAMA (Örnek Mantık) ---
                # Fotoğrafta "KLINKER GRI" kelimesi geçiyorsa çalıştığını kanıtlar
                if "KLINKER" in tum_metin or "KLİNKER" in tum_metin:
                    ws['C3'] = "Görselden Okundu!" # Geçici kontrol metni
                    ws['D3'] = len(tum_metin) # Okunan toplam karakter sayısını yazar (dinamik olduğunu kanıtlar)
                else:
                    ws['C3'] = "Klinker bulunamadı"

                # Sizin tarafınızdan belirtilen hammadde düzeltmesi: Kil yerine Marn kullanımı
                if "MARN" in tum_metin:
                    ws['C15'] = "Marn Okundu"

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                dosya_adi = f"Fercim_Nigde_Gunluk_Rapor_{formatli_tarih}.xlsx"
                st.download_button(
                    label=f"📥 İndir: {dosya_adi}",
                    data=output,
                    file_name=dosya_adi,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"⚠️ Sistemsel Hata: {e}")
