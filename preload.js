const { contextBridge, ipcRenderer } = require('electron');

// Ana süreç ile renderer süreç arasında güvenli bir API expose et
contextBridge.exposeInMainWorld('electronAPI', {
  // Excel dosyası seçme
  selectExcel: () => ipcRenderer.invoke('select-excel'),
  
  // Botu başlatma
  startBot: (data) => ipcRenderer.invoke('start-bot', data),
  
  // Botu durdurma
  stopBot: () => ipcRenderer.invoke('stop-bot'),
  
  // Bot log mesajlarını dinleme
  onBotLog: (callback) => ipcRenderer.on('bot-log', (_, message) => callback(message)),
  
  // Bot hata mesajlarını dinleme
  onBotError: (callback) => ipcRenderer.on('bot-error', (_, message) => callback(message)),
  
  // Bot cevaplarını dinleme
  onBotAnswer: (callback) => ipcRenderer.on('bot-answer', (_, data) => callback(data)),
  
  // Bot tamamlandı mesajını dinleme
  onBotComplete: (callback) => ipcRenderer.on('bot-complete', () => callback()),
  
  // Bot durdu mesajını dinleme
  onBotStopped: (callback) => ipcRenderer.on('bot-stopped', (_, code) => callback(code))
});
