const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const Store = require('electron-store');

// Ayarları saklayacak store oluştur
const store = new Store();

let mainWindow;
let botProcess = null;

function createWindow() {
  // Ana pencereyi oluştur
  mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: path.join(__dirname, 'logo.avif')
  });

  // Ana HTML dosyasını yükle
  mainWindow.loadFile('index.html');

  // Geliştirme aracını aç (geliştirme sırasında kullanışlı)
  // mainWindow.webContents.openDevTools();

  // Pencere kapatıldığında olayı yakala
  mainWindow.on('closed', function () {
    mainWindow = null;
    if (botProcess) {
      botProcess.kill();
      botProcess = null;
    }
  });
}

// Electron hazır olduğunda pencereyi oluştur
app.whenReady().then(createWindow);

// Tüm pencereler kapatıldığında uygulamayı kapat (macOS hariç)
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', function () {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// Excel dosyasını seçme
ipcMain.handle('select-excel', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Dosyaları', extensions: ['xlsx', 'xls'] }
    ]
  });

  if (!result.canceled && result.filePaths.length > 0) {
    store.set('excelPath', result.filePaths[0]);
    return result.filePaths[0];
  }
  return null;
});

// Botu başlat
ipcMain.handle('start-bot', async (event, data) => {
  // Daha önce çalışan bir bot varsa durdur
  if (botProcess) {
    botProcess.kill();
    botProcess = null;
  }

  // Ayarları kaydet
  store.set('url', data.url);
  store.set('email', data.email);
  store.set('password', data.password);

  // Bot için geçici bir script oluştur
 const botScript = `
const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');

(async () => {
    try {
        // Excel dosyasından soruları oku
        const workbook = XLSX.readFile('${data.excelPath.replace(/\\/g, '\\\\')}');
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const allData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        // İlk satır header olarak ayarla, eğer yoksa oluştur
        let headers = ['SORULAR', 'GİTMESİ GEREKEN RAPOR', 'GİTTİĞİ RAPOR', 'EŞLEŞTİ Mİ?', 'CEVAP'];
        
        // Eğer ilk satır header değilse, header'ı ekle
        if (allData.length === 0 || allData[0][0] !== 'SORULAR') {
            allData.unshift(headers);
        }
        
        // Veri satırlarını al (header'ı atla)
        const dataRows = allData.slice(1);

        const browser = await chromium.launch({ headless: false });
        const context = await browser.newContext();
        const page = await context.newPage();

        await page.goto('${data.url}');

        // E-posta gir
        await page.fill('#email', '${data.email}');
        await page.click('button:has-text("Devam")');
        await page.waitForTimeout(1000);

        // Şifre gir
        await page.fill('#password', '${data.password}');
        await page.click('button:has-text("Giriş")');

        // Sayfa yüklensin 
        await page.waitForTimeout(5000);
        await page.getByRole('button', { name: 'Sohbeti Temizle' }).click();
        await page.getByRole('button', { name: 'Sohbeti Temizle' }).nth(1).click();
        
        for (let i = 0; i < dataRows.length; i++) {
            const question = dataRows[i][0];
            if (!question) continue;

            console.log(\`Soru \${i+1} gönderiliyor: \${question}\`);

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
                process.send({ type: 'log', message: \`Soru \${i+1} için API yanıtı bekleniyor...\` });
                await apiResponsePromise;
                process.send({ type: 'log', message: \`Soru \${i+1} için API yanıtı alındı, bir sonraki soruya geçiliyor...\` });

                // Cevabın DOM'a yansıması için biraz daha fazla bekleme
                await page.waitForTimeout(3000);
                
                const cevaplar = await page.$$('p.text-sm.whitespace-pre-wrap');
                const sonCevap = await cevaplar[cevaplar.length - 1].textContent();

                console.log("Cevap:", sonCevap);
                dataRows[i][4] = sonCevap; // CEVAP kolonu

                if(i == 0){
                    await page.getByRole('button', { name: '?' }).first().click();
                    const locator = page.locator('.text-xs.text-purple-700');
                    dataRows[i][2] = await locator.textContent(); // GİTTİĞİ RAPOR kolonu
                    console.log("Gittiği Rapor :", dataRows[i][2]);
                } else { 
                    await page.getByRole('button', { name: '?' }).nth(i).click();
                    const locator = page.locator('.text-xs.text-purple-700');
                    dataRows[i][2] = await locator.textContent(); // GİTTİĞİ RAPOR kolonu
                    console.log("Gittiği Rapor : ", dataRows[i][2]);
                }

                // EŞLEŞTİ Mİ kontrolü
                if (dataRows[i][2] && dataRows[i][2].includes(dataRows[i][1])) {
                    dataRows[i][3] = "EVET";
                } else {
                    dataRows[i][3] = "HAYIR";
                }
                
                await page.getByRole('button', { name: 'Kapat' }).click();
                
                // Kısa bir bekleme
                await page.waitForTimeout(2000);
                
            } catch (error) {
                process.send({ type: 'error', message: \`Soru \${i+1} için hata: \${error.message}\` });
            }
        }

        process.send({ type: 'log', message: "Tüm sorular gönderildi." });
        
        // Header + veri satırlarını birleştir
        const finalData = [headers, ...dataRows];
        
        // Cevapları Excel'e kaydet
        const updatedSheet = XLSX.utils.aoa_to_sheet(finalData);
        workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;
        XLSX.writeFile(workbook, '${data.excelPath.replace(/\\/g, '\\\\')}');
        
        process.send({ type: 'complete' });
        await browser.close();
        
    } catch (error) {
        process.send({ type: 'error', message: \`Genel hata: \${error.message}\` });
    }
})();
`;

  const tempScriptPath = path.join(app.getPath('temp'), 'bot-script.js');
  fs.writeFileSync(tempScriptPath, botScript);

  // Bot çalıştırılacak dizin ayarları - gerekli modülleri bulabilmesi için proje dizini kullanılacak
  const botEnv = { ...process.env, NODE_PATH: path.join(__dirname, 'node_modules') };

  // Botu başlat - projemizdeki node_modules klasörünü kullanarak
  botProcess = spawn('node', [tempScriptPath], {
    stdio: ['pipe', 'pipe', 'pipe', 'ipc'],
    env: botEnv,
    cwd: __dirname // Çalışma dizini olarak proje dizinini kullan
  });

  // Bot çıktılarını yakala
  botProcess.stdout.on('data', (data) => {
    console.log(`stdout: ${data}`);
    mainWindow.webContents.send('bot-log', data.toString());
  });

  botProcess.stderr.on('data', (data) => {
    console.error(`stderr: ${data}`);
    mainWindow.webContents.send('bot-error', data.toString());
  });

  // Mesaj olaylarını dinle
  botProcess.on('message', (message) => {
    if (message.type === 'log') {
      mainWindow.webContents.send('bot-log', message.message);
    } else if (message.type === 'error') {
      mainWindow.webContents.send('bot-error', message.message);
    } else if (message.type === 'answer') {
      mainWindow.webContents.send('bot-answer', { index: message.index, answer: message.answer });
    } else if (message.type === 'complete') {
      mainWindow.webContents.send('bot-complete');
    }
  });

  botProcess.on('close', (code) => {
    console.log(`Bot process exited with code ${code}`);
    mainWindow.webContents.send('bot-stopped', code);
    botProcess = null;
  });

  return { success: true };
});

// Botu durdur
ipcMain.handle('stop-bot', async () => {
  if (botProcess) {
    botProcess.kill();
    botProcess = null;
    return { success: true };
  }
  return { success: false, message: 'Bot çalışmıyor' };
});
