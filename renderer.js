// Form alanları
const urlInput = document.getElementById('url');
const emailInput = document.getElementById('email');
const passwordInput = document.getElementById('password');
const excelPathInput = document.getElementById('excel-path');

// Butonlar
const selectExcelBtn = document.getElementById('select-excel');
const startBotBtn = document.getElementById('start-bot');
const stopBotBtn = document.getElementById('stop-bot');
const clearLogsBtn = document.getElementById('clear-logs');

// Çıktı ve ilerleme alanları
const outputEl = document.getElementById('output');
const progressBarEl = document.getElementById('progress-bar');
const progressTextEl = document.getElementById('progress-text');
const progressPercentEl = document.getElementById('progress-percent');
const lastAnswerEl = document.getElementById('last-answer');

// Durum değişkenleri
let isRunning = false;
let excelPath = '';
let totalQuestions = 0;
let answeredQuestions = 0;

// Form validation
function validateForm() {
  const url = urlInput.value.trim();
  const email = emailInput.value.trim();
  const password = passwordInput.value.trim();
  
  if (url && email && password && excelPath) {
    startBotBtn.disabled = false;
  } else {
    startBotBtn.disabled = true;
  }
}

// Excel dosyasını seçme
selectExcelBtn.addEventListener('click', async () => {
  const filePath = await window.electronAPI.selectExcel();
  if (filePath) {
    excelPath = filePath;
    excelPathInput.value = filePath;
    validateForm();
  }
});

// Input değişikliklerini dinle
[urlInput, emailInput, passwordInput].forEach(input => {
  input.addEventListener('input', validateForm);
});

// Log ekleme
function addLog(message, isError = false) {
  const logLine = document.createElement('div');
  logLine.className = `output-line${isError ? ' error' : ''}`;
  logLine.textContent = message;
  outputEl.appendChild(logLine);
  outputEl.scrollTop = outputEl.scrollHeight;
}

// İlerleme güncelleme
function updateProgress() {
  if (totalQuestions === 0) return;
  
  const percent = Math.round((answeredQuestions / totalQuestions) * 100);
  progressBarEl.style.width = `${percent}%`;
  progressPercentEl.textContent = `${percent}%`;
  progressTextEl.textContent = `İşleniyor: ${answeredQuestions} / ${totalQuestions}`;
}

// Botu başlatma
startBotBtn.addEventListener('click', async () => {
  const url = urlInput.value.trim();
  const email = emailInput.value.trim();
  const password = passwordInput.value.trim();
  
  if (!url || !email || !password || !excelPath) {
    addLog('Lütfen tüm alanları doldurun.', true);
    return;
  }
  
  try {
    isRunning = true;
    startBotBtn.disabled = true;
    stopBotBtn.disabled = false;
    
    // Çıktıyı temizle
    outputEl.innerHTML = '';
    lastAnswerEl.textContent = 'Henüz cevap yok';
    
    // İlerlemeyi sıfırla
    answeredQuestions = 0;
    totalQuestions = 0;
    updateProgress();
    
    addLog('Bot başlatılıyor...');
    
    const result = await window.electronAPI.startBot({
      url,
      email,
      password,
      excelPath
    });
    
    if (result.success) {
      addLog('Bot başlatıldı.');
    } else {
      addLog(`Bot başlatılamadı: ${result.message}`, true);
      isRunning = false;
      startBotBtn.disabled = false;
      stopBotBtn.disabled = true;
    }
  } catch (error) {
    addLog(`Hata: ${error.message}`, true);
    isRunning = false;
    startBotBtn.disabled = false;
    stopBotBtn.disabled = true;
  }
});

// Botu durdurma
stopBotBtn.addEventListener('click', async () => {
  try {
    const result = await window.electronAPI.stopBot();
    
    if (result.success) {
      addLog('Bot durduruldu.');
    } else {
      addLog(`Bot durdurulamadı: ${result.message}`, true);
    }
    
    isRunning = false;
    startBotBtn.disabled = false;
    stopBotBtn.disabled = true;
  } catch (error) {
    addLog(`Hata: ${error.message}`, true);
  }
});

// Log'ları temizleme
clearLogsBtn.addEventListener('click', () => {
  outputEl.innerHTML = '';
});

// Bot log mesajlarını dinleme
window.electronAPI.onBotLog((message) => {
  addLog(message);
  
  // Excel toplam soru sayısını tahmin et
  if (message.includes('gönderiliyor:') && totalQuestions === 0) {
    const match = message.match(/Soru (\d+) gönderiliyor/);
    if (match && match[1]) {
      // İlk soru numarası ile yaklaşık toplam sayı hesapla
      // Bu kesin değil, sadece ilerleme göstergesi için
      totalQuestions = parseInt(match[1]) + 10;
    }
  }
});

// Bot hata mesajlarını dinleme
window.electronAPI.onBotError((message) => {
  addLog(message, true);
});

// Bot cevaplarını dinleme
window.electronAPI.onBotAnswer((data) => {
  answeredQuestions++;
  updateProgress();
  
  // Son cevabı güncelle
  lastAnswerEl.textContent = data.answer;
  
  addLog(`Soru ${data.index} için cevap alındı.`);
});

// Bot tamamlandı mesajını dinleme
window.electronAPI.onBotComplete(() => {
  addLog('İşlem tamamlandı!');
  progressTextEl.textContent = 'Tamamlandı';
  progressBarEl.style.width = '100%';
  progressPercentEl.textContent = '100%';
  
  isRunning = false;
  startBotBtn.disabled = false;
  stopBotBtn.disabled = true;
});

// Bot durdu mesajını dinleme
window.electronAPI.onBotStopped((code) => {
  addLog(`Bot sonlandı (Çıkış kodu: ${code})`);
  
  isRunning = false;
  startBotBtn.disabled = false;
  stopBotBtn.disabled = true;
});
