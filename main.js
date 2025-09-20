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
    const workbook = XLSX.readFile('${data.excelPath.replace(/\\/g, '\\\\')}');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const allData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

   // let headers = ['SORULAR', 'GİTMESİ GEREKEN RAPOR', 'GİTTİĞİ RAPOR', 'EŞLEŞTİ Mİ?', 'CEVAP']; //  sertel vida için
        let headers = ['SORULAR','GİTTİĞİ RAPOR', 'CEVAP']; // güney için

    if (allData.length === 0 || allData[0][0] !== 'SORULAR') {
      allData.unshift(headers);
    }
    const dataRows = allData.slice(1);

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext();
    const page = await context.newPage();

    await page.goto('${data.url}');
    await page.fill('#email', '${data.email}');
    await page.click('button:has-text("Devam")');
    await page.waitForTimeout(1000);
    await page.fill('#password', '${data.password}');
    await page.click('button:has-text("Giriş")');
    await page.waitForTimeout(5000);

    // Global timeout: 90s
    page.setDefaultTimeout(90000);

    await page.getByRole('button', { name: 'Sohbeti Temizle' }).click();
    await page.getByRole('button', { name: 'Sohbeti Temizle' }).nth(1).click();

    for (let i = 0; i < dataRows.length; i++) {
      const question = dataRows[i][0];
      if (!question) continue;

      console.log(\`Soru \${i+1} gönderiliyor: \${question}\`);

      try {
        const apiResponsePromise = page.waitForResponse(
        //  response => response.url().includes('https://api.sertelvida.com.tr/ai/0.0.1/ask/'),
                    response => response.url().includes('https://api.bossai.app/ai/0.0.1/ask/'),
          { timeout: 180000 }
        );

        await page.fill('textarea', question);
        await page.press('textarea', 'Enter');

        process.send({ type: 'log', message: \`Soru \${i+1} için API yanıtı bekleniyor...\` });
        await apiResponsePromise;
        process.send({ type: 'log', message: \`Soru \${i+1} için API yanıtı alındı.\` });

        await page.waitForTimeout(2000);

        const cevaplar = await page.$$('p.text-sm.whitespace-pre-wrap');
        if (cevaplar.length) {
          const sonCevap = await cevaplar[cevaplar.length - 1].textContent();
          dataRows[i][2] = sonCevap;
        } else {
          dataRows[i][2] = "CEVAP BULUNAMADI";
        }

        // '?' butonunu güvenli şekilde tıkla
        try {
          const qCount = await page.getByRole('button', { name: '?' }).count();
          if (qCount > i) {
            await page.getByRole('button', { name: '?' }).nth(i).click({ timeout: 30000 });
          } else if (qCount > 0) {
            await page.getByRole('button', { name: '?' }).nth(qCount - 1).click({ timeout: 30000 });
          } else {
            dataRows[i][1] = "SORU İŞARETİ BULUNAMADI";
            process.send({ type: 'log', message: \`Soru \${i+1}: '?' butonu bulunamadı.\` });
          }
        } catch (qErr) {
          process.send({ type: 'error', message: \`Soru \${i+1} için '?' tıklama hatası: \${qErr.message}\` });
        }

       try {
  const locator = page.locator('.text-xs.text-purple-700');
  await locator.first().waitFor({ timeout: 15000 }).catch(() => {});

  let rawText = await locator.first().textContent().catch(() => "GİTTİĞİ RAPOR ALINAMADI");

  // Temizleme işlemleri
  if (rawText) {
    rawText = rawText.trim();                     // baştaki/sondaki boşlukları sil
    rawText = rawText.replace(/^📁\s*/, "");      // başındaki 📁 ve boşlukları sil
    rawText = rawText.replace(/\s+/g, " ");       // fazla boşlukları tek boşluk yap
  }

  dataRows[i][1] = rawText || "GİTTİĞİ RAPOR ALINAMADI";

} catch {
  dataRows[i][1] = "GİTTİĞİ RAPOR HATASI";
}


        // // Eşleşti mi?
        // if (dataRows[i][2] && dataRows[i][2].includes(dataRows[i][1] || '')) {
        //   dataRows[i][3] = "EVET";
        // } else {
        //   dataRows[i][3] = "HAYIR";
        // }

        // "Kapat" butonunu güvenli şekilde kapat
        try {
          let closed = false;
          const kapatCount = await page.getByRole('button', { name: 'Kapat' }).count().catch(()=>0);
          if (kapatCount > 0) {
            await page.getByRole('button', { name: 'Kapat' }).first().click({ timeout: 30000 });
            closed = true;
          }
          if (!closed) {
            await page.keyboard.press('Escape').catch(()=>{});
            await page.waitForTimeout(500);
            const stillKapat = await page.getByRole('button', { name: 'Kapat' }).count().catch(()=>0);
            if (stillKapat === 0) closed = true;
          }
          if (!closed) {
            const closeSelectors = [
              'button[aria-label="close"]',
              'button[aria-label="kapat"]',
              'button[title="Kapat"]',
              '.modal .close',
              '.dialog .close',
              'button:has-text("×")',
              'button:has-text("Close")'
            ];
            for (const sel of closeSelectors) {
              const el = await page.$(sel);
              if (el) {
                await page.click(sel).catch(()=>{});
                closed = true;
                break;
              }
            }
          }
          if (!closed) {
            process.send({ type: 'log', message: \`Soru \${i+1}: 'Kapat' butonu bulunamadı veya kapatılamadı. Devam ediliyor.\` });
          }
        } catch (closeErr) {
          process.send({ type: 'error', message: \`Soru \${i+1} için kapatma hatası: \${closeErr.message}\` });
        }

      } catch (error) {
        process.send({ type: 'error', message: \`Soru \${i+1} için genel hata: \${error.message}\` });
      }

      // Her sorudan sonra Excel'e yaz
      try {
        const finalData = [headers, ...dataRows];
        const updatedSheet = XLSX.utils.aoa_to_sheet(finalData);
        const sheetName = "BOT_CEVAPLAR";

        if (workbook.SheetNames.includes(sheetName)) {
          delete workbook.Sheets[sheetName];
          workbook.SheetNames = workbook.SheetNames.filter(n => n !== sheetName);
        }

        workbook.SheetNames.push(sheetName);
        workbook.Sheets[sheetName] = updatedSheet;

        XLSX.writeFile(workbook, '${data.excelPath.replace(/\\/g, '\\\\')}');
        process.send({ type: 'log', message: \`Soru \${i+1} sonrası BOT_CEVAPLAR sayfasına kaydedildi.\` });
      } catch (saveError) {
        process.send({ type: 'error', message: \`Excel kaydedilemedi: \${saveError.message}\` });
      }
    }

    process.send({ type: 'complete' });
    await browser.close();

  } catch (error) {
    process.send({ type: 'error', message: \`Genel hata: \${error.message}\` });
  }
})();
`;

  const tempScriptPath = path.join(app.getPath('temp'), 'bot-script.js');
  fs.writeFileSync(tempScriptPath, botScript);

  const botEnv = { ...process.env, NODE_PATH: path.join(__dirname, 'node_modules') };

  botProcess = spawn('node', [tempScriptPath], {
    stdio: ['pipe', 'pipe', 'pipe', 'ipc'],
    env: botEnv,
    cwd: __dirname
  });

  botProcess.stdout.on('data', (data) => {
    console.log(`stdout: ${data}`);
    mainWindow.webContents.send('bot-log', data.toString());
  });

  botProcess.stderr.on('data', (data) => {
    console.error(`stderr: ${data}`);
    mainWindow.webContents.send('bot-error', data.toString());
  });

  botProcess.on('message', (message) => {
    if (message.type === 'log') {
      mainWindow.webContents.send('bot-log', message.message);
    } else if (message.type === 'error') {
      mainWindow.webContents.send('bot-error', message.message);
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
