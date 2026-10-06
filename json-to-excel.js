const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const INPUT_JSON = path.join(__dirname, 'words.json');
const OUTPUT_EXCEL = path.join(__dirname, 'lugat.xlsx');

try {
    console.log('📖 words.json okunuyor...');
    if (!fs.existsSync(INPUT_JSON)) {
        throw new Error('words.json dosyası bulunamadı!');
    }

    const rawData = JSON.parse(fs.readFileSync(INPUT_JSON, 'utf-8'));

    // Excel sütun başlıklarına geri dönüştür
    const formattedData = rawData.map(item => ({
        kelime_adi: item.kelime || '',
        kelimenin_tasnifi: item.tasnif || '',
        kelime_hakkinda_bilgi: item.bilgi || '',
        ilk_gectigi_kaynak: item.kaynak || '',
        ikame_edildigi_kelimeler: item.ikame || '',
        nuans_metni: item.nuans || '',
        ornek_cumleler: item.ornek || '',
        mustak_kelimeler: item.mustak || ''
    }));

    // Yeni bir çalışma kitabı (workbook) ve sayfa oluştur
    const worksheet = XLSX.utils.json_to_sheet(formattedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Lugat");

    // Excel dosyasına kaydet
    XLSX.writeFile(workbook, OUTPUT_EXCEL);

    console.log(`✅ İşlem tamam! Toplam ${formattedData.length} kelime lugat.xlsx dosyasına aktarıldı.`);
} catch (error) {
    console.error('❌ Dönüştürme hatası:', error.message);
}