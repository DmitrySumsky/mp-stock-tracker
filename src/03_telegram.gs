// ════════════════════════════════════════════════════════════
// МОДУЛЬ: Telegram — мульти-чат рассылка
// ════════════════════════════════════════════════════════════

/**
 * Отправляет сообщение во все чаты, указанные в Панели управления.
 * Поддерживает Markdown-разметку.
 * @param {string} text — текст сообщения
 */
function sendTelegram(text) {
  const tg = loadTelegramConfig();
  if (!tg.token || tg.chatIds.length === 0) {
    console.warn('Telegram не настроен: проверьте Панель управления (B2, B3)');
    return;
  }

  const url = `https://api.telegram.org/bot${tg.token}/sendMessage`;

  tg.chatIds.forEach(chatId => {
    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        payload: { chat_id: chatId, text: text, parse_mode: 'Markdown' },
        muteHttpExceptions: true
      });
    } catch (e) {
      console.error(`Telegram → ${chatId}: ${e.message}`);
    }
  });
}
