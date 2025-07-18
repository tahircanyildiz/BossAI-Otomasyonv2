const { chromium } = require('playwright');
const XLSX = require('xlsx');
require('dotenv').config(); // .env dosyasından çevre değişkenlerini yükle

// Excel dosyasından soruları oku
const workbook = XLSX.readFile('sorular.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

(async () => {
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext();
    const page = await context.newPage();

    await page.goto('https://app.sertelvida.com.tr/');

    // E-posta gir (çevre değişkenlerinden)
    await page.fill('#email', "tahir.yildiz@siriusaitech.com");
    await page.click('button:has-text("Devam")');
    await page.waitForTimeout(1000);

    // Şifre gir (çevre değişkenlerinden)
    await page.fill('#password', "Ss123456");
    await page.click('button:has-text("Giriş")');

    // Sayfa yüklensin 
    await page.waitForTimeout(5000);

    // 1. satır başlıksa oradan başlama
    for (let i = 0; i < data.length; i++) {
        const question = data[i][0];
        if (!question) continue;

        console.log(`Soru ${i} gönderiliyor: ${question}`);

        try {
            // API yanıtını beklemek için bir Promise oluştur
            const apiResponsePromise = page.waitForResponse(
                response => response.url().includes('https://api.sertelvida.com.tr/ai/0.0.1/ask/'),
                { timeout: 180000 } // 3 dakikaya kadar bekle
            );

            // Soru input alanını bul ve soruyu gönder
            await page.fill('textarea', question);
            await page.press('textarea', 'Enter');

            // API yanıtını bekle
            console.log(`Soru ${i} için API yanıtı bekleniyor...`);
            await apiResponsePromise;
            console.log(`Soru ${i} için API yanıtı alındı, bir sonraki soruya geçiliyor...`);

            const cevaplar = await page.$$('p.text-sm.whitespace-pre-wrap');

            const sonCevap = await cevaplar[cevaplar.length - 1].textContent();

            console.log("Cevap:", sonCevap);
            data[i][1] = sonCevap;

            // Cevabın DOM'a yansıması için kısa bir bekleme
            await page.waitForTimeout(2000);

        } catch (error) {
            console.error(`Soru ${i} için hata:`, error.message);
        }
    }

    console.log("Tüm sorular gönderildi.");
    const updatedSheet = XLSX.utils.aoa_to_sheet(data);
    workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;
    XLSX.writeFile(workbook, 'sorular.xlsx');
    await browser.close();
})();
