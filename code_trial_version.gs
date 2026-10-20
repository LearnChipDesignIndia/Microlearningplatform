// ============================================================
// TRIAL PERIOD VERSION — NO PAYMENTS
// ============================================================

// ── UPDATED /start HANDLER (simplified) ──────────────────────
function handleStart(token, chatId, userInfo, existingUser, config) {
  const platformName = config.PLATFORM_NAME || 'TechMicroLearn';
  const firstName = userInfo.firstName || 'there';

  // Save user
  const userData = {
    ...userInfo,
    paymentStatus: 'TRIAL',
    track: existingUser ? existingUser.data[5] : '',
    currentDay: existingUser ? existingUser.data[7] : 1,
    registrationDate: existingUser ? existingUser.data[8] : new Date().toISOString()
  };
  saveUser(userData);

  const welcomeText =
    `👋 *Welcome to ${platformName}, ${firstName}!*\n\n` +
    `🎁 *FREE TRIAL — 30 days of focused technical learning*\n\n` +
    `Get a new 5-minute lesson every morning at 9:00 AM IST.\n\n` +
    `*Pick your learning track:*`;

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
  console.log(`✅ New user: ${chatId} | ${userInfo.firstName} @${userInfo.username}`);
}

// ── UPDATED CALLBACK QUERY (no payment logic) ────────────────
function handleCallbackQuery(token, callbackQuery, config) {
  const chatId = callbackQuery.message.chat.id;
  const data   = callbackQuery.data;
  const from   = callbackQuery.from || {};

  // Answer callback to remove spinner
  callTelegram(token, 'answerCallbackQuery', {
    callback_query_id: callbackQuery.id
  });

  // ── TRACK SELECTION ──────────────────────────────────────
  if (data.startsWith('track_')) {
    handleTrackChosenTrial(token, chatId, data, from, config);
  }

  // ── QUIZ: REVEAL ANSWER ──────────────────────────────────
  else if (data.startsWith('quiz_reveal_')) {
    const day = data.replace('quiz_reveal_', '');
    const user = getUserByChatId(chatId);
    if (!user) {
      sendMessage(token, chatId, `Please send /start first.`);
      return;
    }

    const track = user.data[5];
    const topic = getTopicForDay(track, day);
    if (!topic) {
      sendMessage(token, chatId, `Could not find answer for Day ${day}.`);
      return;
    }

    sendMessage(token, chatId,
      `💡 *Answer — Day ${day}:*\n\n` +
      `${topic.quizAnswer}\n\n` +
      `Great work! See you tomorrow at 9 AM. 🎯`
    );

    logMessage(chatId, 'OUT', `quiz answer revealed day ${day}`, 'sent');
  }

  // ── QUIZ: GOT IT ─────────────────────────────────────────
  else if (data.startsWith('quiz_got_')) {
    const day = data.replace('quiz_got_', '');
    sendMessage(token, chatId,
      `🔥 *Nailed it!*\n\n` +
      `You completed Day ${day}. Keep the momentum!\n\n` +
      `Next lesson tomorrow at 9:00 AM IST. 🚀`
    );

    logMessage(chatId, 'OUT', `quiz got it day ${day}`, 'sent');
  }

  // ── CHANGE TRACK ─────────────────────────────────────────
  else if (data === 'change_track') {
    handleTrackSelection(token, chatId, config);
  }

  else {
    console.warn(`Unknown callback: ${data}`);
  }
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

// ── TRACK CHOSEN — SEND DAY 1 IMMEDIATELY (no payment) ────────
function handleTrackChosenTrial(token, chatId, trackData, from, config) {
  const trackMap = {
    'track_axi4':  { name: 'AMBA AXI4 Protocol',          emoji: '📚' },
    'track_riscv': { name: 'RISC-V Base Integer ISA',      emoji: '💾' },
    'track_pcie':  { name: 'PCIe Gen5 Layering',           emoji: '🔌' },
    'track_sv':    { name: 'SystemVerilog Fundamentals',   emoji: '⚡' }
  };

  const track = trackMap[trackData];
  if (!track) return;

  // Update user
  const existingUser = getUserByChatId(chatId);
  const userData = existingUser ? buildUserObject(existingUser.data) : {
    chatId,
    firstName: from.first_name || '',
    lastName: from.last_name || '',
    username: from.username || '',
    languageCode: from.language_code || '',
    registrationDate: new Date().toISOString()
  };

  userData.track = trackData;
  userData.paymentStatus = 'TRIAL';
  userData.currentDay = 1;
  userData.lastMessageDate = new Date().toISOString();
  saveUser(userData);

  // Send confirmation
  sendMessage(token, chatId,
    `${track.emoji} *${track.name}* — selected!\n\n` +
    `🎁 *Your free trial starts now.*\n\n` +
    `Here's your Day 1 lesson 👇`
  );

  // Send Day 1 immediately
  const topic = getTopicForDay(trackData, 1);
  if (topic) {
    sendTopicMessage(token, chatId, topic, 1, config);
  } else {
    sendMessage(token, chatId,
      `⚠️ Day 1 not found. Check Topics sheet has data for Day 1.`
    );
  }

  logMessage(chatId, 'OUT', `trial started: ${trackData}`, 'sent');
  console.log(`✅ Trial user: ${chatId} → ${trackData}`);
}

function sendMessage(token, chatId, text, extra) {
  return callTelegram(token, 'sendMessage', {
    chat_id: chatId,
    text: text,
    parse_mode: 'Markdown',
    ...extra
  });
}
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
// ── SIMPLIFIED DAILY SENDER (trial only, no payment checks) ───
function sendDailyTopicsSimplified() {
  const config = getConfig();
  const token  = config.BOT_TOKEN;
  const sheet  = getSheet('Users');
  const data   = sheet.getDataRange().getValues();

  let sent = 0, skipped = 0, completed = 0, errors = 0;
  console.log(`📤 Daily send started — ${new Date().toISOString()}`);

  for (let i = 1; i < data.length; i++) {
    const row        = data[i];
    let chatId       = row[0];
    const track      = row[5];
    const currentDay = Number(row[7]) || 1;

    // Validate chatId
    if (!chatId || chatId === '') { skipped++; continue; }
    chatId = Number(String(chatId).trim());
    if (isNaN(chatId)) { skipped++; continue; }

    // Skip if no track
    if (!track) { skipped++; continue; }

    try {
      // Course complete
      if (currentDay > 30) {
        sendMessage(token, chatId,
          `🎉 *You completed 30 days!*\n\n` +
          `Congrats on building real expertise. 🏆\n\n` +
          `Ready for another track? Send /start`
        );
        completed++;
        continue;
      }

      // Get and send topic
      const topic = getTopicForDay(track, currentDay);
      if (!topic) {
        console.warn(`No topic: ${chatId} | ${track} | day ${currentDay}`);
        skipped++;
        continue;
      }

      sendTopicMessage(token, chatId, topic, currentDay, config);

      // Advance day
      sheet.getRange(i + 1, 8).setValue(currentDay + 1);
      sheet.getRange(i + 1, 10).setValue(new Date().toISOString());

      sent++;
      console.log(`✅ Day ${currentDay} → ${chatId} | ${track}`);

      Utilities.sleep(250);

    } catch (err) {
      console.error(`❌ Error for ${chatId}:`, err.message);
      errors++;
    }
  }

  console.log(`\n📊 Send complete:`);
  console.log(`   Sent: ${sent}`);
  console.log(`   Completed: ${completed}`);
  console.log(`   Skipped: ${skipped}`);
  console.log(`   Errors: ${errors}`);

  logMessage('SYSTEM', 'OUT', `Daily send: ${sent} sent, ${completed} completed`, 'done');
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

// ── TEST 8: Verify webhook is registered correctly ────────────
function testCheckWebhook() {
  const config = getConfig();
  const url = `${TELEGRAM_API(config.BOT_TOKEN)}/getWebhookInfo`;
  const response = UrlFetchApp.fetch(url);
  const result = JSON.parse(response.getContentText());
  console.log('Webhook info:', JSON.stringify(result.result, null, 2));
  // Should show your Apps Script URL and pending_update_count
}

// ── KEEP THESE — They still work as-is ────────────────────────
// handleStart() — updated above
// handleStatus() — no changes needed
// handleHelp() — no changes needed
// handleToday() — no changes needed
// handleTrackSelection() — no changes needed
// sendTopicMessage() — no changes needed
// handleQuizAnswer() — renamed to handleCallbackQuery quiz logic
// sendMessage() — no changes needed
// saveUser() — no changes needed
// getTopicForDay() — no changes needed
