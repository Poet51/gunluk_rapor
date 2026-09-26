import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import datetime

# Sayfa ayarları
st.set_page_config(page_title="Ferçim Rapor", page_icon="🏭", layout="centered")

st.title("🏭 Ferçim Niğde - Otomatik Rapor")
st.markdown("SAP MM ekran görüntülerini yükleyin ve rapor tarihini seçin.")

# Tarih Seçici (Otomatik olarak bugünün tarihini getirir)
rapor_tarihi = st.date_input("🗓️ Rapor Tarihi", datetime.date.today())
formatli_tarih = rapor_tarihi.strftime("%d.%m.%Y")

# Dosya yükleme alanı
uploaded_files = st.file_uploader("📸 SAP Ekran Görüntülerini Seçin", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} görüntü yüklendi. İşleme hazır.")
    
    if st.button("🚀 Raporu Oluştur"):
        with st.spinner('Veriler okunuyor ve Excel hazırlanıyor...'):
            try:
                # Şablonu yükle
                template_path = "Niğde Günlük Rapor 2026_09_24.xlsx"
                wb = load_workbook(template_path)
                ws = wb['24'] 
                
                # 1. TARİH GÜNCELLEMESİ (Excel içindeki hücreyi günceller)
                # Not: Logonun altındaki tarih hücresi A2 olarak varsayılmıştır. 
                # Eğer Excel'de tarih başka hücredeyse (örneğin A3), burayı değiştirebiliriz.
                ws['A2'] = formatli_tarih 
                
                # 2. VERİ GÜNCELLEMESİ (Sabit test verileri kaldırıldı, OCR entegrasyonu buraya gelecek)
                # İleriki aşamada pytesseract veya bulut API ile okunan veriler değişken olarak buraya atanacak.
                # Örnek: ws['C3'] = okunan_klinker_db_stok
                
                # Yeni dosyayı belleğe kaydet
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"🎉 {formatli_tarih} tarihli Excel Raporu Başarıyla Dolduruldu!")
                
                # 3. DİNAMİK DOSYA İSMİ İLE İNDİRME BUTONU
                dosya_adi = f"Fercim_Nigde_Gunluk_Rapor_{formatli_tarih}.xlsx"
                st.download_button(
                    label=f"📥 İndir: {dosya_adi}",
                    data=output,
                    file_name=dosya_adi,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"⚠️ Hata: Excel dosyası bulunamadı veya işlenemedi. Detay: {e}")
