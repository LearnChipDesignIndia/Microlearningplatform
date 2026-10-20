// ============================================================
// TELEGRAM MICROLEARN BOT — Complete Google Apps Script
// Paste this entire file into Extensions > Apps Script
// ============================================================

// ─── CONFIG ──────────────────────────────────────────────────
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  return config;
}

const TELEGRAM_API = (token) => `https://api.telegram.org/bot${token}`;

// ─── SHEET HELPERS ───────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getUserByChatId(chatId) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(chatId)) {
      return { rowIndex: i + 1, data: data[i] };
    }
  }
  return null;
}

function saveUser(user) {
  const sheet = getSheet('Users');
  const existing = getUserByChatId(user.chatId);
  const row = [
    user.chatId,
    user.firstName || '',
    user.lastName || '',
    user.username || '',
    user.languageCode || '',
    user.track || '',
    user.paymentStatus || 'FREE',
    user.currentDay || 1,
    user.registrationDate || new Date().toISOString(),
    user.lastMessageDate || new Date().toISOString(),
    user.source || 'organic',
    user.deviceInfo || '',
    user.totalMessages || 1,
    user.notes || ''
  ];
  if (existing) {
    sheet.getRange(existing.rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function logMessage(chatId, direction, messageText, status) {
  const sheet = getSheet('MessageLog');
  sheet.appendRow([
    new Date().toISOString(),
    chatId,
    direction,
    messageText.substring(0, 500),
    status
  ]);
}

function getTopicForDay(track, day) {
  const sheet = getSheet('Topics');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(day) && data[i][1] === track) {
      return {
        day: data[i][0],
        track: data[i][1],
        title: data[i][2],
        content: data[i][3],
        imageUrl: data[i][4],
        quizQuestion: data[i][5],
        quizAnswer: data[i][6]
      };
    }
  }
  return null;
}

// ─── TELEGRAM API HELPERS ────────────────────────────────────
function callTelegram(token, method, payload) {
  const url = `${TELEGRAM_API(token)}/${method}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      console.error(`Telegram API error [${method}]:`, JSON.stringify(result));
    }
    return result;
  } catch (e) {
    console.error(`Network error [${method}]:`, e.message);
    return { ok: false, error: e.message };
  }
}

function sendMessage(token, chatId, text, extra) {
  return callTelegram(token, 'sendMessage', {
    chat_id: chatId,
    text: text,
    parse_mode: 'Markdown',
    ...extra
  });
}

function sendPhoto(token, chatId, photoUrl, caption) {
  return callTelegram(token, 'sendPhoto', {
    chat_id: chatId,
    photo: photoUrl,
    caption: caption,
    parse_mode: 'Markdown'
  });
}

// ─── WEBHOOK SETUP ───────────────────────────────────────────
// Run this function ONCE manually from Apps Script editor
function setWebhook() {
  const config = getConfig();
  const scriptUrl = ScriptApp.getService().getUrl();
  const result = callTelegram(config.BOT_TOKEN, 'setWebhook', {
    url: scriptUrl,
    allowed_updates: ['message', 'callback_query']
  });
  console.log('Webhook set:', JSON.stringify(result));
  return result;
}


function debugIncomingWebhook() {
  // Paste a real update from Telegram here to simulate it locally
  const fakeUpdate = {
    postData: {
      contents: JSON.stringify({
        update_id: 123456,
        callback_query: {
          id: "test123",
          from: {
            id: 123456789,       // ← your TEST_CHAT_ID
            first_name: "Harshal",
            username: "harshal_test",
            language_code: "en"
          },
          message: {
            chat: { id: 123456789 },  // ← same
            text: "test"
          },
          data: "track_axi4"
        }
      })
    }
  };

  // Run doPost with the fake update
  const result = doPost(fakeUpdate);
  console.log('doPost result:', result.getContent());
}
// ─── MAIN WEBHOOK HANDLER ────────────────────────────────────
function doPost(e) {
  try {
    const config = getConfig();
    const token = config.BOT_TOKEN;
    const update = JSON.parse(e.postData.contents);

    console.log('Incoming update:', JSON.stringify(update));

    // Handle regular messages
    if (update.message) {
      handleMessage(token, update.message, config);
    }

    // Handle button taps
    if (update.callback_query) {
      handleCallbackQuery(token, update.callback_query, config);
    }

  } catch (err) {
    console.error('doPost error:', err.message, err.stack);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── MESSAGE HANDLER ─────────────────────────────────────────
function handleMessage(token, message, config) {
  const chatId = message.chat.id;
  const text = message.text || '';
  const from = message.from || {};

  // Collect rich user data
  const userInfo = {
    chatId: chatId,
    firstName: from.first_name || '',
    lastName: from.last_name || '',
    username: from.username || '',
    languageCode: from.language_code || '',
    deviceInfo: message.via_bot ? 'bot' : 'direct',
    lastMessageDate: new Date().toISOString(),
    source: 'organic'
  };

  const existingUser = getUserByChatId(chatId);
  const totalMessages = existingUser ? (Number(existingUser.data[12]) + 1) : 1;
  userInfo.totalMessages = totalMessages;

  logMessage(chatId, 'IN', text, 'received');

  if (text === '/start') {
    handleStart(token, chatId, userInfo, existingUser, config);

  } else if (text === '/status') {
    handleStatus(token, chatId, existingUser, config);

  } else if (text === '/today') {
    handleToday(token, chatId, existingUser, config);

  } else if (text === '/help') {
    handleHelp(token, chatId, config);

  } else if (text === '/tracks') {
    handleTrackSelection(token, chatId, config);

  } else {
    // Save message count even for unknown commands
    if (existingUser) {
      const updatedUser = buildUserObject(existingUser.data);
      updatedUser.totalMessages = totalMessages;
      updatedUser.lastMessageDate = new Date().toISOString();
      saveUser(updatedUser);
    }

    sendMessage(token, chatId,
      `I didn't understand that command. Type /help to see what I can do.`
    );
  }
}

// ─── /start HANDLER ──────────────────────────────────────────
function handleStart(token, chatId, userInfo, existingUser, config) {
  const platformName = config.PLATFORM_NAME || 'MicroLearn';
  const firstName = userInfo.firstName || 'there';

  // Save or update user with registration info
  const userData = {
    ...userInfo,
    paymentStatus: existingUser ? existingUser.data[6] : 'FREE',
    track: existingUser ? existingUser.data[5] : '',
    currentDay: existingUser ? existingUser.data[7] : 1,
    registrationDate: existingUser ? existingUser.data[8] : new Date().toISOString()
  };
  saveUser(userData);

  const welcomeText =
    `👋 *Welcome to ${platformName}, ${firstName}!*\n\n` +
    `Get a focused 5-minute technical lesson every morning — ` +
    `straight to your Telegram.\n\n` +
    `*Choose your learning track to get started:*`;

  sendMessage(token, chatId, welcomeText, {
    reply_markup: {
      inline_keyboard: [
        [{ text: '📚 AMBA AXI4 Protocol', callback_data: 'track_axi4' }],
        [{ text: '💾 RISC-V Base Integer ISA', callback_data: 'track_riscv' }],
        [{ text: '🔌 PCIe Gen5 Layering', callback_data: 'track_pcie' }],
        [{ text: '⚡ SystemVerilog Fundamentals', callback_data: 'track_sv' }]
      ]
    }
  });

  logMessage(chatId, 'OUT', 'welcome + track selection', 'sent');
  console.log(`✅ New user started: ${chatId} | ${userInfo.firstName} @${userInfo.username}`);
}

function handleCallbackQuery(token, callbackQuery, config) {
  const chatId = callbackQuery.message.chat.id;
  const data   = callbackQuery.data;
  const from   = callbackQuery.from || {};

  callTelegram(token, 'answerCallbackQuery', {
    callback_query_id: callbackQuery.id
  });

  if (data.startsWith('track_')) {
    // Instead of direct confirmation, send payment request
    const user = getUserByChatId(chatId);
    const userInfo = {
      chatId, firstName: from.first_name || '', lastName: from.last_name || '',
      username: from.username || '', languageCode: from.language_code || '',
      track: data, paymentStatus: user ? user.data[6] : 'FREE',
      currentDay: 1, registrationDate: user ? user.data[8] : new Date().toISOString(),
      lastMessageDate: new Date().toISOString(),
      source: 'organic', totalMessages: user ? Number(user.data[12]) + 1 : 1
    };
    saveUser(userInfo);
    sendPaymentRequest(token, chatId, data, config);

  } else if (data.startsWith('confirm_payment_')) {
    handlePaymentConfirmation(token, chatId, data, from, config);

  } else if (data === 'payment_reminder') {
    sendMessage(token, chatId,
      `No problem! We'll remind you again in 3 days.\n\n` +
      `When ready, just send /start to pick up where you left off. 👍`
    );

  } else if (data === 'restart_payment') {
    handleTrackSelection(token, chatId, config);

  } else if (data.startsWith('admin_approve_')) {
    adminApproveUser(token, chatId, data, config);

  } else if (data.startsWith('admin_reject_')) {
    adminRejectUser(token, chatId, data, config);

  } else if (data.startsWith('quiz_')) {
    handleQuizAnswer(token, chatId, data, from, config);

  } else if (data === 'change_track') {
    handleTrackSelection(token, chatId, config);
  }
}



// ─── TRACK SELECTION ─────────────────────────────────────────
function handleTrackChosen(token, chatId, trackData, from, config) {
  const trackMap = {
    'track_axi4':  { name: 'AMBA AXI4 Protocol',          emoji: '📚' },
    'track_riscv': { name: 'RISC-V Base Integer ISA',      emoji: '💾' },
    'track_pcie':  { name: 'PCIe Gen5 Layering',           emoji: '🔌' },
    'track_sv':    { name: 'SystemVerilog Fundamentals',   emoji: '⚡' }
  };

  const track = trackMap[trackData];
  if (!track) return;

  // Update user's track
  const existingUser = getUserByChatId(chatId);
  const userData = existingUser ? buildUserObject(existingUser.data) : {
    chatId: chatId,
    firstName: from.first_name || '',
    lastName: from.last_name || '',
    username: from.username || '',
    languageCode: from.language_code || '',
    registrationDate: new Date().toISOString()
  };

  userData.track = trackData;
  userData.currentDay = 1;
  userData.lastMessageDate = new Date().toISOString();
  saveUser(userData);

  const confirmText =
    `${track.emoji} *${track.name}* — track selected!\n\n` +
    `You'll receive your first lesson *tomorrow at 9:00 AM IST*.\n\n` +
    `📅 *Your 30-day journey starts now.*\n\n` +
    `*Commands you can use:*\n` +
    `/today — get today's lesson now\n` +
    `/status — see your progress\n` +
    `/tracks — switch to a different track\n` +
    `/help — all commands`;

  sendMessage(token, chatId, confirmText);
  logMessage(chatId, 'OUT', `track confirmed: ${trackData}`, 'sent');
  console.log(`✅ Track set: ${chatId} → ${trackData}`);
}

// ─── /today COMMAND ──────────────────────────────────────────
function handleToday(token, chatId, existingUser, config) {
  if (!existingUser || !existingUser.data[5]) {
    sendMessage(token, chatId,
      `You haven't selected a track yet. Use /start to choose one.`
    );
    return;
  }

  const track = existingUser.data[5];
  const currentDay = Number(existingUser.data[7]) || 1;
  const topic = getTopicForDay(track, currentDay);

  if (!topic) {
    sendMessage(token, chatId,
      `No lesson found for Day ${currentDay}. ` +
      `Please contact support: @${config.SUPPORT_USERNAME || 'admin'}`
    );
    return;
  }

  sendTopicMessage(token, chatId, topic, currentDay, config);
}

// ─── /status COMMAND ─────────────────────────────────────────
function handleStatus(token, chatId, existingUser, config) {
  if (!existingUser) {
    sendMessage(token, chatId, `You're not registered yet. Send /start to begin.`);
    return;
  }

  const d = existingUser.data;
  const track = d[5] || 'None selected';
  const currentDay = d[7] || 1;
  const regDate = d[8] ? new Date(d[8]).toLocaleDateString('en-IN') : 'Unknown';
  const totalMsgs = d[12] || 0;

  const trackNames = {
    'track_axi4':  '📚 AMBA AXI4 Protocol',
    'track_riscv': '💾 RISC-V ISA',
    'track_pcie':  '🔌 PCIe Gen5',
    'track_sv':    '⚡ SystemVerilog'
  };

  const statusText =
    `📊 *Your Learning Status*\n\n` +
    `👤 *Name:* ${d[1]} ${d[2]}\n` +
    `🎯 *Track:* ${trackNames[track] || track}\n` +
    `📅 *Day:* ${currentDay} of 30\n` +
    `📈 *Progress:* ${Math.round((currentDay / 30) * 100)}%\n` +
    `🗓 *Member since:* ${regDate}\n` +
    `💬 *Total interactions:* ${totalMsgs}\n\n` +
    `Keep going! You're building real expertise. 🚀`;

  sendMessage(token, chatId, statusText);
}

// ─── /help COMMAND ───────────────────────────────────────────
function handleHelp(token, chatId, config) {
  const helpText =
    `🤖 *Available Commands*\n\n` +
    `/start — welcome & pick your learning track\n` +
    `/today — get today's lesson right now\n` +
    `/status — check your progress\n` +
    `/tracks — switch to a different track\n` +
    `/help — show this menu\n\n` +
    `📬 *Lessons arrive daily at 9:00 AM IST.*\n\n` +
    `Support: @${config.SUPPORT_USERNAME || 'admin'}`;

  sendMessage(token, chatId, helpText);
}

// ─── /tracks COMMAND ─────────────────────────────────────────
function handleTrackSelection(token, chatId, config) {
  sendMessage(token, chatId,
    `🎯 *Choose your learning track:*\n\nPicking a new track resets your day counter to Day 1.`, {
    reply_markup: {
      inline_keyboard: [
        [{ text: '📚 AMBA AXI4 Protocol', callback_data: 'track_axi4' }],
        [{ text: '💾 RISC-V Base Integer ISA', callback_data: 'track_riscv' }],
        [{ text: '🔌 PCIe Gen5 Layering', callback_data: 'track_pcie' }],
        [{ text: '⚡ SystemVerilog Fundamentals', callback_data: 'track_sv' }]
      ]
    }
  });
}

// ─── SEND TOPIC MESSAGE ──────────────────────────────────────
function sendTopicMessage(token, chatId, topic, day, config) {
  const messageText =
    `📖 *Day ${day} — ${topic.title}*\n\n` +
    `${topic.content}\n\n` +
    `━━━━━━━━━━━━━━━━━━━━\n` +
    `🧠 *Quick Quiz:*\n${topic.quizQuestion}`;

  const quizKeyboard = topic.quizAnswer ? {
    reply_markup: {
      inline_keyboard: [
        [{ text: '💡 Reveal Answer', callback_data: `quiz_reveal_${day}` }],
        [{ text: '✅ Got it!', callback_data: `quiz_got_${day}` }]
      ]
    }
  } : {};

  // Send image if available, otherwise text only
  if (topic.imageUrl && topic.imageUrl.startsWith('http')) {
    sendPhoto(token, chatId, topic.imageUrl, messageText.substring(0, 1024));
    if (Object.keys(quizKeyboard).length > 0) {
      sendMessage(token, chatId, '👇 Quiz options:', quizKeyboard);
    }
  } else {
    sendMessage(token, chatId, messageText, quizKeyboard);
  }

  logMessage(chatId, 'OUT', `Day ${day}: ${topic.title}`, 'sent');
}

// ─── QUIZ ANSWER HANDLER ─────────────────────────────────────
function handleQuizAnswer(token, chatId, data, from, config) {
  const parts = data.split('_');
  const action = parts[1];
  const day = parts[2];

  const existingUser = getUserByChatId(chatId);
  if (!existingUser) return;

  const track = existingUser.data[5];
  const topic = getTopicForDay(track, day);
  if (!topic) return;

  if (action === 'reveal') {
    sendMessage(token, chatId,
      `💡 *Answer for Day ${day}:*\n\n${topic.quizAnswer}\n\n` +
      `Great work! See you tomorrow at 9 AM. 🎯`
    );
  } else if (action === 'got') {
    sendMessage(token, chatId,
      `✅ Excellent! You're crushing it.\n\nSee you tomorrow at 9 AM. 🚀`
    );
  }
}

// ─── DAILY TOPIC SENDER (runs on schedule) ───────────────────
function sendDailyTopics() {
  const config = getConfig();
  const token = config.BOT_TOKEN;
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();

  let sent = 0;
  let errors = 0;

  console.log(`📤 Daily send started — ${new Date().toISOString()}`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const chatId = row[0];
    const track = row[5];
    const currentDay = Number(row[7]) || 1;

    if (!chatId || !track) continue;
    if (currentDay > 30) {
      // Course complete — send completion message
      sendCompletionMessage(token, chatId, track, config);
      continue;
    }

    const topic = getTopicForDay(track, currentDay);
    if (!topic) {
      console.warn(`⚠️ No topic for chatId=${chatId} track=${track} day=${currentDay}`);
      errors++;
      continue;
    }

    try {
      sendTopicMessage(token, chatId, topic, currentDay, config);

      // Advance day counter
      sheet.getRange(i + 1, 8).setValue(currentDay + 1);
      sheet.getRange(i + 1, 10).setValue(new Date().toISOString());

      sent++;
      console.log(`✅ Sent Day ${currentDay} to ${chatId} | track: ${track}`);

      // Small delay to avoid Telegram rate limits
      Utilities.sleep(100);

    } catch (err) {
      console.error(`❌ Failed for ${chatId}:`, err.message);
      errors++;
    }
  }

  console.log(`📊 Daily send complete — sent: ${sent} | errors: ${errors}`);
  logMessage('SYSTEM', 'OUT', `Daily send: ${sent} sent, ${errors} errors`, 'complete');
}

// ─── COURSE COMPLETION MESSAGE ───────────────────────────────
function sendCompletionMessage(token, chatId, track, config) {
  const trackNames = {
    'track_axi4':  'AMBA AXI4 Protocol',
    'track_riscv': 'RISC-V ISA',
    'track_pcie':  'PCIe Gen5',
    'track_sv':    'SystemVerilog'
  };

  const msg =
    `🎉 *Congratulations! You've completed the 30-day ${trackNames[track] || track} track!*\n\n` +
    `You've built real, focused knowledge. That's a genuine achievement.\n\n` +
    `*What's next?*\n` +
    `/tracks — start a new learning track\n\n` +
    `Thank you for learning with us! 🙏`;

  sendMessage(token, chatId, msg);
}

// ─── UTILITY: Build user object from sheet row ───────────────
function buildUserObject(row) {
  return {
    chatId: row[0],
    firstName: row[1],
    lastName: row[2],
    username: row[3],
    languageCode: row[4],
    track: row[5],
    paymentStatus: row[6],
    currentDay: row[7],
    registrationDate: row[8],
    lastMessageDate: row[9],
    source: row[10],
    deviceInfo: row[11],
    totalMessages: row[12],
    notes: row[13]
  };
}

// ─── TEST FUNCTION (run manually to test) ────────────────────
function testSendToSelf() {
  const config = getConfig();
  // Add your own chatId to Config sheet as TEST_CHAT_ID
  const testChatId = config.TEST_CHAT_ID;
  if (!testChatId) {
    console.log('Add TEST_CHAT_ID to Config sheet first');
    return;
  }
  sendMessage(config.BOT_TOKEN, testChatId,
    '🧪 *Test message from your MicroLearn Bot!*\n\nIf you see this, your bot is working perfectly. ✅'
  );
}

// ============================================================
// TESTING FUNCTIONS — paste these at the bottom of Code.gs
// Run any of these directly from the Apps Script editor
// ============================================================

// ── TEST 1: Get your own Chat ID ─────────────────────────────
// Step 1: Send ANY message to your bot in Telegram
// Step 2: Run this function — it prints your chatId in the logs
function getRecentUpdates() {
  const config = getConfig();

  // First delete the webhook temporarily
  const deleteResult = callTelegram(config.BOT_TOKEN, 'deleteWebhook', {});
  //console.log('Webhook deleted:', JSON.parse(deleteResult).description);
  // Now fetch updates
  const url = `${TELEGRAM_API(config.BOT_TOKEN)}/getUpdates`;
  const response = UrlFetchApp.fetch(url);
  const result = JSON.parse(response.getContentText());
  console.log('Recent updates:');
  result.result.forEach(u => {
    if (u.message) {
      console.log(`chatId: ${u.message.chat.id} | name: ${u.message.from.first_name} | text: ${u.message.text}`);
    }
  });
  // Copy the chatId from logs → paste into Config sheet as TEST_CHAT_ID
}

// ── TEST 2: Send a test message to yourself ──────────────────
function testMessageToSelf() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }

  sendMessage(config.BOT_TOKEN, chatId,
    '✅ *Bot is working!*\n\nWebhook and message sending are both live.\n\nTry /start in the bot now.'
  );
  console.log('Test message sent to', chatId);
}

// ── TEST 3: Simulate /start for yourself ─────────────────────
function testStartCommand() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }

  const fakeMessage = {
    chat: { id: chatId },
    from: { first_name: 'Harshal', last_name: 'Test', username: 'harshal_test', language_code: 'en' },
    text: '/start'
  };
  handleMessage(config.BOT_TOKEN, fakeMessage, config);
  console.log('Simulated /start for', chatId);
}

// ── TEST 4: Simulate track selection ─────────────────────────
function testTrackSelection() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }

  const fakeCallback = {
    id: 'test_callback_123',
    message: { chat: { id: chatId } },
    from: { first_name: 'Harshal', username: 'harshal_test', language_code: 'en' },
    data: 'track_axi4'
  };
  handleCallbackQuery(config.BOT_TOKEN, fakeCallback, config);
  console.log('Simulated track selection: track_axi4');
}

// ── TEST 5: Send Day 1 lesson to yourself immediately ─────────
function testDayOneTopic() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }

  const topic = getTopicForDay('track_axi4', 1);
  if (!topic) { console.log('No topic found — check Topics sheet has data'); return; }

  sendTopicMessage(config.BOT_TOKEN, chatId, topic, 1, config);
  console.log('Sent Day 1 AXI4 topic to', chatId);
}

// ── TEST 6: Run the full daily send right now ─────────────────
// WARNING: This sends real messages to ALL users in the sheet
function testDailySendNow() {
  console.log('Starting manual daily send...');
  sendDailyTopics();
  console.log('Daily send complete — check MessageLog tab');
}

// ── TEST 7: Manually set a user's payment status to PAID ─────
// Use this to test the paid flow without real payment
function testSetUserPaid() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }

  const user = getUserByChatId(chatId);
  if (!user) { console.log('User not found — run testStartCommand first'); return; }

  const userData = buildUserObject(user.data);
  userData.paymentStatus = 'PAID';
  userData.currentDay = 1;
  saveUser(userData);

  sendMessage(config.BOT_TOKEN, chatId,
    '💳 *Payment verified!* (Test mode)\n\n' +
    'Your 30-day premium track is now active.\n' +
    'First lesson arrives tomorrow at 9:00 AM IST. 🎯'
  );
  console.log('Set PAID status for', chatId);
}

// ── TEST 8: Verify webhook is registered correctly ────────────
function testCheckWebhook() {
  const config = getConfig();
  const url = `${TELEGRAM_API(config.BOT_TOKEN)}/getWebhookInfo`;
  const response = UrlFetchApp.fetch(url);
  const result = JSON.parse(response.getContentText());
  console.log('Webhook info:', JSON.stringify(result.result, null, 2));
  // Should show your Apps Script URL and pending_update_count
}

// ── TEST 9: Send payment reminder to yourself ─────────────────
function testPaymentReminder() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }

  sendPaymentReminder(config.BOT_TOKEN, chatId, config);
  console.log('Payment reminder sent to', chatId);
}
// ============================================================
// PAYMENT SYSTEM — paste this into Code.gs
// Two modes: UPI link (free, manual confirm) OR Razorpay (webhook auto-confirm)
// ============================================================

// ── PAYMENT CONFIG (add these rows to Config sheet) ──────────
// UPI_ID            | yourname@upi
// UPI_NAME          | Your Name
// PAYMENT_AMOUNT    | 30
// RAZORPAY_KEY      | rzp_live_xxxx  (optional, for webhook mode)
// RAZORPAY_SECRET   | your_secret    (optional, for webhook mode)

// ── STEP 1: Send payment message after track selection ────────
// Call this INSTEAD of the plain confirmation in handleTrackChosen
function sendPaymentRequest(token, chatId, trackData, config) {
  const trackNames = {
    'track_axi4':  '📚 AMBA AXI4 Protocol',
    'track_riscv': '💾 RISC-V Base Integer ISA',
    'track_pcie':  '🔌 PCIe Gen5 Layering',
    'track_sv':    '⚡ SystemVerilog Fundamentals'
  };

  const upiId   = config.UPI_ID   || 'yourname@upi';
  const upiName = config.UPI_NAME || 'TechMicroLearn';
  const amount  = config.PAYMENT_AMOUNT || '30';

  // Build UPI deep link — works with GPay, PhonePe, Paytm, BHIM
  const upiLink =
    `upi://pay?pa=${upiId}&pn=${encodeURIComponent(upiName)}` +
    `&am=${amount}&cu=INR&tn=${encodeURIComponent('MicroLearn 30-day track')}`;

  // Fallback: GPay web link
  const gpayLink = `https://gpay.app.goo.gl/pay?pa=${upiId}&pn=${encodeURIComponent(upiName)}&am=${amount}&cu=INR`;

  const paymentText =
    `🎯 *${trackNames[trackData]}*\n\n` +
    `You're one step away from 30 days of focused technical learning.\n\n` +
    `💳 *₹${amount} — one-time for the full 30-day track*\n\n` +
    `*Pay via UPI:*\n` +
    `UPI ID: \`${upiId}\`\n` +
    `Amount: ₹${amount}\n` +
    `Name: ${upiName}\n\n` +
    `After paying, tap the button below to confirm. ✅`;

  sendMessage(token, chatId, paymentText, {
    reply_markup: {
      inline_keyboard: [
        [{ text: '📱 Open GPay', url: gpayLink }],
        [{ text: '✅ I have paid — confirm my access', callback_data: `confirm_payment_${trackData}` }],
        [{ text: '⏰ Remind me later', callback_data: 'payment_reminder' }]
      ]
    }
  });

  logMessage(chatId, 'OUT', `payment request: ${trackData}`, 'sent');
}

// ── STEP 2: Handle "I have paid" button tap ───────────────────
// This is the MANUAL confirmation flow (no webhook needed)
// Add this case inside handleCallbackQuery:
//   } else if (data.startsWith('confirm_payment_')) {
//       handlePaymentConfirmation(token, chatId, data, from, config);

function handlePaymentConfirmation(token, chatId, data, from, config) {
  const trackData = data.replace('confirm_payment_', '');

  // Update user status to PENDING_VERIFICATION
  const user = getUserByChatId(chatId);
  if (!user) return;

  const userData = buildUserObject(user.data);
  userData.paymentStatus = 'PENDING_VERIFICATION';
  userData.track = trackData;
  userData.lastMessageDate = new Date().toISOString();
  saveUser(userData);

  // Send pending message to user
  sendMessage(token, chatId,
    `⏳ *Payment received — verifying now.*\n\n` +
    `We'll confirm your access within a few minutes.\n\n` +
    `You'll get a message as soon as it's verified. 🔔`
  );

  // Alert YOU (the admin) via Telegram to verify manually
  const adminChatId = config.ADMIN_CHAT_ID;
  if (adminChatId) {
    const trackNames = {
      'track_axi4': 'AMBA AXI4', 'track_riscv': 'RISC-V',
      'track_pcie': 'PCIe Gen5', 'track_sv': 'SystemVerilog'
    };
    sendMessage(token, adminChatId,
      `💰 *New payment claim!*\n\n` +
      `👤 Name: ${userData.firstName} ${userData.lastName}\n` +
      `📱 Username: @${userData.username}\n` +
      `🎯 Track: ${trackNames[trackData] || trackData}\n` +
      `🔑 ChatId: \`${chatId}\`\n\n` +
      `Check UPI and approve:`, {
      reply_markup: {
        inline_keyboard: [
          [{ text: '✅ Approve — activate account', callback_data: `admin_approve_${chatId}` }],
          [{ text: '❌ Reject — refund needed',     callback_data: `admin_reject_${chatId}` }]
        ]
      }
    });
  }

  logMessage(chatId, 'OUT', 'payment pending verification', 'sent');
}

// ── STEP 3: Admin approves/rejects from their Telegram ────────
// Add these cases inside handleCallbackQuery:
//   } else if (data.startsWith('admin_approve_')) {
//       adminApproveUser(token, chatId, data, config);
//   } else if (data.startsWith('admin_reject_')) {
//       adminRejectUser(token, chatId, data, config);

function adminApproveUser(token, adminChatId, data, config) {
  const userChatId = data.replace('admin_approve_', '');

  const user = getUserByChatId(userChatId);
  if (!user) {
    sendMessage(token, adminChatId, `❌ User ${userChatId} not found in sheet.`);
    return;
  }

  // Activate the user
  const userData = buildUserObject(user.data);
  userData.paymentStatus = 'PAID';
  userData.currentDay = 1;
  userData.lastMessageDate = new Date().toISOString();
  saveUser(userData);

  // Notify the user — they're in!
  sendMessage(token, userChatId,
    `🎉 *Payment verified! Welcome to TechMicroLearn.*\n\n` +
    `Your 30-day ${userData.track === 'track_axi4' ? 'AMBA AXI4' : userData.track} track is now *active*.\n\n` +
    `📅 Your first lesson arrives *tomorrow at 9:00 AM IST.*\n\n` +
    `Use /today to get Day 1 right now if you're impatient. 🚀`
  );

  // Confirm to admin
  sendMessage(token, adminChatId,
    `✅ Approved! User ${userData.firstName} (@${userData.username}) is now active.`
  );

  logMessage(userChatId, 'SYSTEM', 'payment approved by admin', 'paid');
  console.log(`✅ Admin approved payment for ${userChatId}`);
}

function adminRejectUser(token, adminChatId, data, config) {
  const userChatId = data.replace('admin_reject_', '');

  const user = getUserByChatId(userChatId);
  if (!user) return;

  const userData = buildUserObject(user.data);
  userData.paymentStatus = 'REJECTED';
  saveUser(userData);

  sendMessage(token, userChatId,
    `❌ *We couldn't verify your payment.*\n\n` +
    `Please contact us: @${config.SUPPORT_USERNAME || 'admin'}\n\n` +
    `If you've paid, share your UPI transaction ID and we'll sort it out quickly. 🙏`
  );

  sendMessage(token, adminChatId,
    `Rejected. User ${userData.firstName} (@${userData.username}) has been notified.`
  );
}

// ── STEP 4: Payment reminder for FREE users ───────────────────
// Called by sendDailyTopics() for users with paymentStatus = 'FREE'
function sendPaymentReminder(token, chatId, config) {
  const user = getUserByChatId(chatId);
  if (!user) return;

  const userData = buildUserObject(user.data);
  const upiId = config.UPI_ID || 'yourname@upi';

  sendMessage(token, chatId,
    `👋 *Hey ${userData.firstName}!*\n\n` +
    `You signed up for TechMicroLearn but haven't unlocked your track yet.\n\n` +
    `🎯 *₹30 gets you 30 days of daily technical lessons.*\n` +
    `That's ₹1/day — less than a chai.\n\n` +
    `UPI: \`${upiId}\`\n\n` +
    `Tap below to pick your track and pay:`, {
    reply_markup: {
      inline_keyboard: [
        [{ text: '🚀 Pick a track and pay ₹30', callback_data: 'restart_payment' }]
      ]
    }
  });
}

// ── STEP 5: Razorpay webhook handler (OPTIONAL — auto mode) ───
// If you use Razorpay Payment Links, they POST to your webhook URL
// Add a second doPost route OR handle inside existing doPost
function handleRazorpayWebhook(data) {
  // Razorpay sends: event = 'payment.captured', payload.payment.entity
  if (data.event !== 'payment.captured') return;

  const payment = data.payload.payment.entity;
  const notes   = payment.notes || {};
  const chatId  = notes.telegram_chat_id; // You embed this when creating the payment link
  const track   = notes.track;

  if (!chatId) {
    console.error('Razorpay webhook: no telegram_chat_id in notes');
    return;
  }

  const config = getConfig();
  const user   = getUserByChatId(chatId);
  if (!user) return;

  const userData = buildUserObject(user.data);
  userData.paymentStatus = 'PAID';
  userData.currentDay    = 1;
  userData.track         = track || userData.track;
  userData.notes         = `razorpay_payment_id:${payment.id}`;
  saveUser(userData);

  sendMessage(config.BOT_TOKEN, chatId,
    `✅ *₹${payment.amount / 100} received via ${payment.method}!*\n\n` +
    `Your 30-day track is now *active*.\n` +
    `First lesson: tomorrow at 9:00 AM IST. 🚀\n\n` +
    `Use /today to start right now.`
  );

  logMessage(chatId, 'SYSTEM', `razorpay paid: ${payment.id}`, 'paid');
  console.log(`✅ Auto-confirmed Razorpay payment for chatId ${chatId}`);
}
/*
// ── UPDATED sendDailyTopics with paid/free routing ────────────
// Replace the original sendDailyTopics() with this version
function sendDailyTopicsWithPaymentCheck() {
  const config = getConfig();
  const token  = config.BOT_TOKEN;
  const sheet  = getSheet('Users');
  const data   = sheet.getDataRange().getValues();

  let sentPaid = 0, sentReminder = 0, errors = 0;
  console.log(`📤 Daily send started — ${new Date().toISOString()}`);

  for (let i = 1; i < data.length; i++) {
    const row          = data[i];
    const chatId       = row[0];
    const track        = row[5];
    const payStatus    = row[6];
    const currentDay   = Number(row[7]) || 1;

    if (!chatId) continue;

    try {
      if (payStatus === 'PAID') {
        // Paid user — send the actual lesson
        if (currentDay > 30) {
          sendCompletionMessage(token, chatId, track, config);
        } else {
          const topic = getTopicForDay(track, currentDay);
          if (topic) {
            sendTopicMessage(token, chatId, topic, currentDay, config);
            sheet.getRange(i + 1, 8).setValue(currentDay + 1);
            sheet.getRange(i + 1, 10).setValue(new Date().toISOString());
            sentPaid++;
          }
        }
      } else if (payStatus === 'FREE' || payStatus === 'PENDING_VERIFICATION') {
        // Free user — send payment nudge (only every 3 days, not every day)
        const lastMsg = row[9] ? new Date(row[9]) : null;
        const daysSinceLastMsg = lastMsg
          ? Math.floor((Date.now() - lastMsg.getTime()) / 86400000)
          : 999;

        if (daysSinceLastMsg >= 3) {
          sendPaymentReminder(token, chatId, config);
          sheet.getRange(i + 1, 10).setValue(new Date().toISOString());
          sentReminder++;
        }
      }

      Utilities.sleep(100);

    } catch (err) {
      console.error(`❌ Failed for ${chatId}:`, err.message);
      errors++;
    }
  }

  console.log(`📊 Done — paid lessons: ${sentPaid} | reminders: ${sentReminder} | errors: ${errors}`);
}
*/
function sendDailyTopicsWithPaymentCheck() {
  const config = getConfig();
  const token  = config.BOT_TOKEN;
  const sheet  = getSheet('Users');
  const data   = sheet.getDataRange().getValues();

  let sent = 0, skipped = 0, errors = 0;
  console.log(`📤 Daily send started — ${new Date().toISOString()}`);

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    const chatId     = row[0];
    const track      = row[5];
    const currentDay = Number(row[7]) || 1;

    // Skip empty rows or users with no track selected
    if (!chatId || !track) { skipped++; continue; }

    // Skip users who already finished all 30 days
    if (currentDay > 30) {
      sendCompletionMessage(token, chatId, track, config);
      skipped++;
      continue;
    }

    try {
      const topic = getTopicForDay(track, currentDay);

      if (!topic) {
        console.warn(`⚠️ No topic: chatId=${chatId} track=${track} day=${currentDay}`);
        skipped++;
        continue;
      }

      sendTopicMessage(token, chatId, topic, currentDay, config);

      // Advance day counter and update last message date
      sheet.getRange(i + 1, 8).setValue(currentDay + 1);
      sheet.getRange(i + 1, 10).setValue(new Date().toISOString());

      sent++;
      console.log(`✅ Day ${currentDay} sent → ${chatId} | ${track}`);

      // Small delay to stay within Telegram rate limits
      Utilities.sleep(150);

    } catch (err) {
      console.error(`❌ Failed for ${chatId}:`, err.message);
      errors++;
    }
  }

  console.log(`📊 Daily send done — sent: ${sent} | skipped: ${skipped} | errors: ${errors}`);
  logMessage('SYSTEM', 'OUT', `Daily send: ${sent} sent, ${skipped} skipped, ${errors} errors`, 'complete');
}
