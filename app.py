import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

# Sayfa ayarları
st.set_page_config(page_title="Ferçim Rapor", page_icon="🏭", layout="centered")

st.title("🏭 Ferçim Niğde - Otomatik Rapor")
st.markdown("SAP MM ekran görüntülerini yükleyin, klinker, yakıt ve çimento verileriniz Excel'e işlensin.")

# Dosya yükleme alanı
uploaded_files = st.file_uploader("SAP Ekran Görüntülerini Seçin", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} görüntü yüklendi. İşleme hazır.")
    
    if st.button("🚀 Raporu Oluştur"):
        with st.spinner('Veriler okunuyor ve Excel hazırlanıyor...'):
            try:
                # Kendi Excel şablonunu okutuyoruz
                template_path = "Niğde Günlük Rapor 2026_09_24.xlsx"
                wb = load_workbook(template_path)
                ws = wb['24'] # 24 isimli sayfaya bağlanıyoruz
                
                # --- OCR ve Veri İşleme Aşaması (Şu an prototip çalışması) ---
                # İlerleyen aşamada fotoğraftan okunan veriler otomatik buraya aktarılacak
                ws['C3'] = 96094.233 # Örnek: Klinker DB Stok
                ws['D3'] = 2418      # Örnek: Klinker Üretim
                ws['E12'] = 125.736  # Örnek: Petrokok Sarf
                
                # Yeni dosyayı belleğe kaydet
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("🎉 Excel Raporu Başarıyla Dolduruldu!")
                
                # İndirme butonu
                st.download_button(
                    label="📥 Güncel Excel Raporunu İndir",
                    data=output,
                    file_name="Guncel_Nigde_Rapor.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"⚠️ Hata: Excel dosyası bulunamadı. Lütfen deponuzdaki dosya adının doğru olduğundan emin olun. Detay: {e}")
