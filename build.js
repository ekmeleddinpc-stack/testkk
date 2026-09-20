const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const EXCEL_FILE = path.join(__dirname, 'lugat.xlsx');
const OUTPUT_JSON = path.join(__dirname, 'words.json');

try {
    console.log('📖 lugat.xlsx okunuyor...');
    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Excel verisini JSON nesnelerine çevir
    const rawData = XLSX.utils.sheet_to_json(sheet);

    // Başlıkları eşle ve boşlukları temizle
    const formattedData = rawData.map(row => ({
        kelime: String(row.kelime_adi || '').trim(),
        tasnif: String(row.kelimenin_tasnifi || '').trim(),
        bilgi: String(row.kelime_hakkinda_bilgi || '').trim(),
        kaynak: String(row.ilk_gectigi_kaynak || '').trim(),
        ikame: String(row.ikame_edildigi_kelimeler || '').trim(),
        nuans: String(row.nuans_metni || '').trim(),
        ornek: String(row.ornek_cumleler || '').trim(),
        mustak: String(row.mustak_kelimeler || '').trim()
    })).filter(item => item.kelime !== '');

    // Türkçe alfabetik sıralama yap
    formattedData.sort((a, b) => a.kelime.localeCompare(b.kelime, 'tr'));

    // words.json olarak yaz
    fs.writeFileSync(OUTPUT_JSON, JSON.stringify(formattedData, null, 2), 'utf-8');
    
    console.log(`✅ İşlem tamam! Toplam ${formattedData.length} kelime words.json dosyasına aktarıldı.`);
} catch (error) {
    console.error('❌ Derleme hatası:', error.message);
}