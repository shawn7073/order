/** 從 GitHub 匯入 Claude 整理好的店家菜單（menus 資料夾） */
function importMenusFromGitHub() {
  const base = 'https://raw.githubusercontent.com/shawn7073/order/main/menus/';
  const files = JSON.parse(UrlFetchApp.fetch(base + 'index.json').getContentText()).filter(f => f !== 'index.json');
  files.forEach(f => {
    try {
      const r = UrlFetchApp.fetch(base + encodeURIComponent(f), { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) { Logger.log(f + ' 讀取失敗 ' + r.getResponseCode()); return; }
      Logger.log(f + '：' + importStore(JSON.parse(r.getContentText())));
    } catch (e) { Logger.log(f + ' 失敗：' + e.message); }
  });
}

/**
 * 揪團點餐機器人 — Google Apps Script 後端
 * --------------------------------------------------------------
 * 功能：
 *  - LINE 機器人 Webhook：群組開團、收單、誰沒付、結單；管理員私訊菜單照片 → AI 辨識建檔
 *  - 點餐頁（LIFF）API：讀取訂單、送出我的餐點、付款回報與核銷
 *  - 資料全部存在這份 Google 試算表
 *
 * 需要的「指令碼屬性」（專案設定 → 指令碼屬性）：
 *  CHANNEL_ACCESS_TOKEN  LINE Messaging API 的 Channel access token（必填）
 *  WEBHOOK_KEY           自訂一組密碼，Webhook 網址後面要加 ?k=這組密碼（必填）
 *  LIFF_ID               點餐頁的 LIFF ID（必填）
 *  LOGIN_CHANNEL_ID      LINE Login channel 的 Channel ID（必填，用來驗證點餐者身分）
 *  GEMINI_API_KEY        Google AI Studio 的 API key（用 Gemini 辨識菜單時必填）
 *  GEMINI_MODEL          選填，預設 gemini-flash-latest
 *  AI_PROVIDER           選填，gemini（預設）或 claude
 *  ANTHROPIC_API_KEY     選填，AI_PROVIDER=claude 時使用
 *  CLAUDE_MODEL          選填，預設 claude-haiku-5-5
 *  PAY_INFO              選填，付款說明，會顯示在點餐頁（例如：LINE Pay 轉給 王小美）
 *  ADMIN_USER_IDS        由機器人自動維護，不用手動填
 */

/* ============================== 設定與工具 ============================== */

const PROPS = PropertiesService.getScriptProperties();
function cfg(key, def) {
  const v = PROPS.getProperty(key);
  return v === null || v === '' ? def : v;
}

const SHEETS = {
  stores: ['storeId', 'name', 'phone', 'address', 'hours', 'contact', 'note', 'menuJson', 'createdAt', 'updatedAt', 'line', 'minOrder'],
  orders: ['orderId', 'groupId', 'storeId', 'storeName', 'status', 'creatorId', 'creatorName', 'deliveryFee', 'discount', 'createdAt', 'closedAt'],
  entries: ['entryId', 'orderId', 'userId', 'displayName', 'itemsJson', 'subtotal', 'payStatus', 'payMethod', 'updatedAt'],
  drafts: ['userId', 'draftJson', 'updatedAt'],
  wishes: ['createdAt', 'userId', 'name', 'text', 'status'],
  members: ['groupId', 'userId', 'name', 'lastSeen', 'left'],
  payqr: ['userId', 'name', 'fileId', 'updatedAt']
};

const STATUS_LABEL = { open: '收單中', closed: '已收單', done: '已結單' };
const PAY_LABEL = { unpaid: '未付款', claimed: '待確認', paid: '已收款' };
const METHOD_LABEL = { linepay: 'LINE Pay', cash: '現金', transfer: '轉帳' };

function book() {
  const id = cfg('SHEET_ID', '1Nz0M7vnxev8251NyOqsPw1UqaDSkp9cegW4R7Zv3fD8');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet(name) {
  const ss = book();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@'); // 全部存成文字，避免電話開頭的 0 被吃掉
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < SHEETS[name].length) {
    // 新增欄位時補上標題（舊資料不受影響）
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
  }
  return sh;
}

function readAll(name) {
  const sh = sheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = SHEETS[name];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map((row, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c] = row[j] === null ? '' : String(row[j]); });
    return o;
  });
}

function toRow(name, obj) {
  return SHEETS[name].map(c => (obj[c] === undefined || obj[c] === null ? '' : String(obj[c])));
}

function insertRow(name, obj) {
  const sh = sheet(name);
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, SHEETS[name].length).setNumberFormat('@').setValues([toRow(name, obj)]);
}

function updateRow(name, obj) {
  sheet(name).getRange(obj._row, 1, 1, SHEETS[name].length).setNumberFormat('@').setValues([toRow(name, obj)]);
}

function deleteRow(name, obj) {
  sheet(name).deleteRow(obj._row);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function now() { return new Date().toISOString(); }
function shortId() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function money(n) { return '$' + Math.round(n); }

function textOut(s) { return ContentService.createTextOutput(s); }
function jsonOut(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** 第一次使用時在編輯器執行一次：建立工作表並完成授權 */
function setup() {
  Object.keys(SHEETS).forEach(sheet);
  const def = book().getSheetByName('工作表1') || book().getSheetByName('Sheet1');
  if (def && book().getSheets().length > 1 && def.getLastRow() === 0) book().deleteSheet(def);
  if (!cfg('WEBHOOK_KEY', '')) PROPS.setProperty('WEBHOOK_KEY', shortId());
  Logger.log('完成。WEBHOOK_KEY = ' + cfg('WEBHOOK_KEY'));
}

/* ============================== 進入點 ============================== */

function doGet() {
  return textOut('揪團點餐機器人運作中');
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }

  // LINE Webhook（Apps Script 讀不到 header，所以用網址上的密碼驗證）
  if (Array.isArray(body.events)) {
    if (((e.parameter && e.parameter.k) || '') !== cfg('WEBHOOK_KEY', '__unset__')) return textOut('forbidden');
    body.events.forEach(ev => {
      try { handleEvent(ev); } catch (err) { console.error(err && err.stack || err); }
    });
    return textOut('ok');
  }

  // 點餐頁 API
  try {
    return jsonOut(Object.assign({ ok: true }, handleApi(body)));
  } catch (err) {
    console.error(err && err.stack || err);
    return jsonOut({ ok: false, error: String(err && err.message || err) });
  }
}

/* ============================== LINE 共用 ============================== */

function lineApi(path, payload) {
  return UrlFetchApp.fetch('https://api.line.me/v2/bot/' + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + cfg('CHANNEL_ACCESS_TOKEN', '') },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

/** 先用免費的回覆訊息；回覆權杖過期（例如辨識太久）時改用推播 */
function reply(token, messages, fallbackTo) {
  const msgs = messages.slice(0, 5);
  const r = lineApi('message/reply', { replyToken: token, messages: msgs });
  if (r.getResponseCode() === 200) return true;
  console.warn('reply failed: ' + r.getContentText());
  if (fallbackTo) lineApi('message/push', { to: fallbackTo, messages: msgs });
  return false;
}

function txt(text, quickItems) {
  const m = { type: 'text', text: String(text).slice(0, 4900) };
  if (quickItems && quickItems.length) {
    m.quickReply = {
      items: quickItems.slice(0, 13).map(q => ({
        type: 'action',
        action: { type: 'message', label: String(q.label).slice(0, 20), text: q.text }
      }))
    };
  }
  return m;
}

function liffUrl(orderId) {
  return 'https://liff.line.me/' + cfg('LIFF_ID', '') + '?order=' + encodeURIComponent(orderId);
}

function getDisplayName(src) {
  try {
    let path;
    if (src.type === 'group') path = 'group/' + src.groupId + '/member/' + src.userId;
    else if (src.type === 'room') path = 'room/' + src.roomId + '/member/' + src.userId;
    else path = 'profile/' + src.userId;
    const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/' + path, {
      headers: { Authorization: 'Bearer ' + cfg('CHANNEL_ACCESS_TOKEN', '') },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() === 200) return JSON.parse(r.getContentText()).displayName || '朋友';
  } catch (err) { /* ignore */ }
  return '朋友';
}

/* ============================== 管理員 ============================== */

function admins() {
  return cfg('ADMIN_USER_IDS', '').split(',').map(s => s.trim()).filter(Boolean);
}
function isAdmin(uid) { return admins().indexOf(uid) >= 0; }
function addAdmin(uid) {
  const list = admins();
  if (list.indexOf(uid) < 0) list.push(uid);
  PROPS.setProperty('ADMIN_USER_IDS', list.join(','));
}

/* ============================== 事件處理 ============================== */

const HELP_GROUP = [
  '🍱 揪團點餐指令',
  '・開團：選店家開一團',
  '・點餐：重新取得點餐連結',
  '・投票：大家投票選店，發起人輸入「開票」自動開團',
  '・誰沒點：列出還沒點餐的人',
  '・誰沒付：查看付款狀況',
  '・清單：合併後的下單清單（收單時也會自動列出）',
  '・催款：標記還沒付款的人提醒付款（開團的人或管理員）',
  '・運費 60／折扣 30：設定運費或折扣（開團的人或管理員）',
  '・收單：停止點餐',
  '・結單：全部收齊後結束這團',
  '・店家：列出已建檔店家與聯絡方式',
  '・許願 想要的功能：記下來，下次更新一起做',
  '（每週一 10:00 機器人會自動公布還沒付清的款項）'
].join('\n');

const HELP_ADMIN = [
  '🛠 管理員私訊指令',
  '・直接傳菜單照片：AI 辨識店家與菜單（同一家可一次選多張送出，會合併）',
  '・搜尋 店名：上網找菜單做成草稿，例如「搜尋 五十嵐 民生店」',
  '・確認：建檔（同名店家會更新）',
  '・取消：放棄目前草稿',
  '・預覽：再看一次草稿',
  '・直接打字：修改草稿，例如「珍奶 L 改 65」「電話是 02-1234-5678」',
  '・店家：列出店家',
  '・刪除店家 店名',
  '・門檻 店名 300：設定店家外送門檻',
  '・設定收款碼：傳你的收款 QR code，開團時給大家掃',
  '・邀請管理員：產生邀請碼給另一個人',
  '・許願 想要的功能：記下來，下次更新一起做'
].join('\n');

function handleEvent(ev) {
  if (ev.type === 'join') {
    reply(ev.replyToken, [txt('大家好！我是點餐小幫手。\n\n' + HELP_GROUP)]);
    return;
  }
  if (ev.type === 'follow') {
    reply(ev.replyToken, [txt('嗨！要建檔店家菜單請先成為管理員：\n・第一位使用者請輸入「設定管理員」\n・其他人請向管理員索取邀請碼，再輸入「管理員 邀請碼」')]);
    return;
  }
  if (ev.type === 'memberLeft' && ev.left && ev.left.members) {
    ev.left.members.forEach(mb => markMemberLeft(groupKey(ev.source), mb.userId));
    return;
  }
  if (ev.type !== 'message') return;

  const src = ev.source || {};
  const isGroup = src.type === 'group' || src.type === 'room';
  if (isGroup && src.userId) { try { touchMember(src); } catch (err) { console.warn(err); } }
  const msg = ev.message;

  if (msg.type === 'image') {
    if (isGroup) return;
    if (CacheService.getScriptCache().get('awaitpay_' + src.userId)) return handlePayQrImage(ev);
    if (isAdmin(src.userId)) handleMenuImage(ev);
    return;
  }
  if (msg.type !== 'text') return;
  const t = String(msg.text || '').trim();
  if (isGroup) handleGroupText(ev, t);
  else handleDmText(ev, t);
}

/* ---------- 群組 ---------- */

function groupKey(src) { return src.groupId || src.roomId; }

function activeOrder(groupId) {
  const list = readAll('orders').filter(o => o.groupId === groupId && o.status !== 'done');
  return list.length ? list[list.length - 1] : null;
}

function canManage(order, uid) { return !!order && (order.creatorId === uid || isAdmin(uid)); }

function handleGroupText(ev, t) {
  const src = ev.source;
  const gid = groupKey(src);
  const uid = src.userId;
  const tok = ev.replyToken;

  if (t === '說明' || t === '指令' || t === 'help') return reply(tok, [txt(HELP_GROUP)]);

  if (t === '店家') return reply(tok, [txt(storeListText())]);

  const wish = t.match(/^許願\s*([\s\S]+)$/);
  if (wish) return reply(tok, [txt(saveWish(src, wish[1]))]);

  if (t === '開團') {
    const cur = activeOrder(gid);
    if (cur) return reply(tok, [txt('目前已有一團「' + cur.storeName + '」（' + STATUS_LABEL[cur.status] + '），請先「結單」再開新團。\n點餐頁：' + liffUrl(cur.orderId))]);
    const stores = readAll('stores');
    if (!stores.length) return reply(tok, [txt('還沒有任何店家，請管理員先私訊我菜單照片建檔。')]);
    const lines = stores.map((s, i) => (i + 1) + '. ' + s.name);
    return reply(tok, [txt('要開哪一家？點下方按鈕，或輸入「開團 店名」\n\n' + lines.join('\n'),
      stores.slice(0, 13).map(s => ({ label: s.name, text: '開團 ' + s.name })))]);
  }

  let m = t.match(/^開團\s*(.+)$/);
  if (m) {
    const store = findStore(m[1]);
    if (!store) return reply(tok, [txt('找不到店家「' + m[1] + '」，輸入「店家」看清單。')]);
    return reply(tok, [openOrderMessage(gid, uid, getDisplayName(src), store)]);
  }

  // ----- 投票選店 -----
  m = t.match(/^(投票|揪團投票)\s*(.*)$/);
  if (m) {
    const cur = activeOrder(gid);
    if (cur) return reply(tok, [txt('目前已有一團「' + cur.storeName + '」，請先「結單」再投票。')]);
    let stores = readAll('stores');
    if (m[2].trim()) {
      const picked = m[2].trim().split(/[\s,，、]+/).map(findStore).filter(Boolean);
      if (picked.length) stores = picked.filter((s, i, a) => a.findIndex(x => x.storeId === s.storeId) === i);
    }
    if (stores.length < 2) return reply(tok, [txt('至少要有兩家店才能投票。可以輸入「投票 店名A 店名B 店名C」指定候選店家。')]);
    stores = stores.slice(0, 12);
    const poll = { by: uid, byName: getDisplayName(src), options: stores.map(s => s.name), votes: {}, names: {} };
    savePoll(gid, poll);
    return reply(tok, [txt('🗳 ' + poll.byName + ' 發起投票：今天喝哪家？\n點下方按鈕投票（可改票），' + poll.byName + ' 輸入「開票」就會自動開團。\n\n' +
      stores.map((s, i) => (i + 1) + '. ' + s.name).join('\n'),
      stores.map(s => ({ label: s.name, text: '我投 ' + s.name })).concat([{ label: '🔔 開票', text: '開票' }]))]);
  }

  m = t.match(/^我投\s*(.+)$/);
  if (m) {
    const poll = loadPoll(gid);
    if (!poll) return reply(tok, [txt('目前沒有進行中的投票，輸入「投票」發起。')]);
    const q = m[1].trim();
    const opt = poll.options.find(o => o === q) || poll.options.find(o => o.indexOf(q) >= 0);
    if (!opt) return reply(tok, [txt('「' + q + '」不在候選名單裡。')]);
    poll.votes[uid] = opt;
    poll.names[uid] = getDisplayName(src);
    savePoll(gid, poll);
    return reply(tok, [txt('✅ ' + poll.names[uid] + ' 投給 ' + opt + '\n' + pollTally(poll).map(x => x.name + ' ' + x.count + ' 票').join('｜'))]);
  }

  if (t === '開票' || t === '投票結果') {
    const poll = loadPoll(gid);
    if (!poll) return reply(tok, [txt('目前沒有進行中的投票。')]);
    const tally = pollTally(poll);
    if (t === '投票結果') return reply(tok, [txt('🗳 目前票數：\n' + tally.map(x => '・' + x.name + ' ' + x.count + ' 票').join('\n'))]);
    if (uid !== poll.by && !isAdmin(uid)) return reply(tok, [txt('只有發起投票的 ' + poll.byName + ' 或管理員可以開票。')]);
    if (!tally.length || !tally[0].count) return reply(tok, [txt('還沒有人投票喔！')]);
    const store = findStore(tally[0].name);
    CacheService.getScriptCache().remove('poll_' + gid);
    const head = txt('🎉 開票結果：\n' + tally.map(x => '・' + x.name + ' ' + x.count + ' 票').join('\n') + '\n\n由「' + store.name + '」勝出！');
    return reply(tok, [head, openOrderMessage(gid, poll.by, poll.byName, store)]);
  }

  const order = activeOrder(gid);
  const needOrder = ['點餐', '連結', '誰沒付', '誰沒點', '進度', '清單', '下單清單', '催款', '收單', '結單', '確定結單', '重新開放'];
  if (!order && (needOrder.indexOf(t) >= 0 || /^(運費|折扣)/.test(t))) {
    return reply(tok, [txt('目前沒有進行中的團，輸入「開團」開始。')]);
  }
  if (!order) return;

  if (t === '點餐' || t === '連結') {
    return reply(tok, [txt('「' + order.storeName + '」' + STATUS_LABEL[order.status] + '\n' + liffUrl(order.orderId))]);
  }

  if (t === '誰沒付' || t === '進度') return reply(tok, [txt(orderSummaryText(order))]);

  if (t === '清單' || t === '下單清單') return reply(tok, [txt(mergedListText(order))]);

  if (t === '誰沒點') {
    const names = notOrderedNames(order);
    if (!names.length) return reply(tok, [txt('目前已知的成員都點了 👍\n（機器人只認得在群組說過話或點過餐的人）')]);
    return reply(tok, [txt('🙋 還沒點「' + order.storeName + '」的人（' + names.length + '）：\n' + names.join('、') +
      '\n\n點餐：' + liffUrl(order.orderId) + '\n（名單來自在群組說過話或點過餐的人，可能不完整）')]);
  }

  if (t === '催款') {
    if (!canManage(order, uid)) return reply(tok, [txt('只有開團的人（' + order.creatorName + '）或管理員可以催款。')]);
    return reply(tok, [remindMessage(order)]);
  }

  m = t.match(/^(運費|折扣)\s*\$?\s*(\d+)$/);
  if (m) {
    if (!canManage(order, uid)) return reply(tok, [txt('只有開團的人或管理員可以設定' + m[1] + '。')]);
    withLock(() => {
      const o = readAll('orders').find(x => x.orderId === order.orderId);
      if (m[1] === '運費') o.deliveryFee = Number(m[2]); else o.discount = Number(m[2]);
      updateRow('orders', o);
    });
    const fresh = readAll('orders').find(x => x.orderId === order.orderId);
    return reply(tok, [txt('已設定' + m[1] + ' ' + money(m[2]) + '，會平均分攤給每個人。\n\n' + orderSummaryText(fresh))]);
  }

  if (t === '收單' || t === '結單' || t === '確定結單' || t === '重新開放') {
    if (!canManage(order, uid)) return reply(tok, [txt('只有開團的人（' + order.creatorName + '）或管理員可以' + t + '。')]);
    if (t === '結單') {
      const es = orderEntries(order.orderId);
      computeTotals(order, es);
      const owe = es.filter(e => e.payStatus !== 'paid');
      if (owe.length) {
        return reply(tok, [txt('⚠️ 還有人沒付清：\n' + owe.map(e => '・' + e.displayName + ' ' + money(e.amount) + (e.payStatus === 'claimed' ? '（已回報，待確認）' : '')).join('\n') +
          '\n\n確定要結單請輸入「確定結單」。沒付清的款項會出現在每週一 10:00 的欠款公告。')]);
      }
    }
    const status = t === '收單' ? 'closed' : (t === '結單' || t === '確定結單') ? 'done' : 'open';
    withLock(() => {
      const o = readAll('orders').find(x => x.orderId === order.orderId);
      o.status = status;
      if (status !== 'open') o.closedAt = now();
      updateRow('orders', o);
    });
    const fresh = readAll('orders').find(x => x.orderId === order.orderId);
    const head = status === 'closed' ? '🛑 已收單，不能再改餐點了。' : status === 'done' ? '✅ 已結單，辛苦了！' : '🔓 已重新開放點餐。';
    const msgs = [txt(head + '\n\n' + orderSummaryText(fresh))];
    if (status === 'closed') {
      msgs.push(txt(mergedListText(fresh)));
      const st = readAll('stores').find(x => x.storeId === fresh.storeId);
      const url = st ? storeOaUrl(st, storeOrderText(fresh)) : '';
      if (url) msgs.push(txt('📤 一鍵傳給店家（會打開店家的 LINE 聊天室並填好訂單，按送出就完成）：\n' + url));
    }
    return reply(tok, msgs);
  }
}

/* ---------- 私訊 ---------- */

function handleDmText(ev, t) {
  const uid = ev.source.userId;
  const tok = ev.replyToken;

  if (t === '我的ID') return reply(tok, [txt(uid)]);

  let w = t.match(/^許願\s*([\s\S]+)$/);
  if (w) return reply(tok, [txt(saveWish(ev.source, w[1]))]);

  if (t === '設定收款碼') {
    CacheService.getScriptCache().put('awaitpay_' + uid, '1', 600);
    return reply(tok, [txt('請在 10 分鐘內傳一張你的「收款」QR code 圖片（LINE Pay、街口或銀行轉帳都可以）。\n\n⚠️ 只能傳「收款碼」！千萬不要傳「付款碼／付款條碼」，那是讓店家掃了直接扣你錢用的。\n\n之後你開的團，點餐頁會顯示這張收款碼給團員掃。')]);
  }
  if (t === '刪除收款碼') {
    withLock(() => {
      const r = readAll('payqr').find(x => x.userId === uid);
      if (r) {
        if (r.fileId) { try { DriveApp.getFileById(r.fileId).setTrashed(true); } catch (e) { /* ignore */ } }
        deleteRow('payqr', r);
      }
    });
    return reply(tok, [txt('已刪除你的收款碼。')]);
  }

  if (t === '設定管理員') {
    if (isAdmin(uid)) return reply(tok, [txt('你已經是管理員了。\n\n' + HELP_ADMIN)]);
    if (admins().length) return reply(tok, [txt('已經有管理員了，請向管理員索取邀請碼，再輸入「管理員 邀請碼」。')]);
    addAdmin(uid);
    return reply(tok, [txt('✅ 你現在是管理員。\n\n' + HELP_ADMIN)]);
  }

  let m = t.match(/^管理員\s*(\d{6})$/);
  if (m) {
    const cache = CacheService.getScriptCache();
    if (cache.get('invite_' + m[1])) {
      cache.remove('invite_' + m[1]);
      addAdmin(uid);
      return reply(tok, [txt('✅ 你現在是管理員，可以傳菜單照片給我建檔了。\n\n' + HELP_ADMIN)]);
    }
    return reply(tok, [txt('邀請碼錯誤或已過期，請再向管理員索取。')]);
  }

  if (!isAdmin(uid)) {
    return reply(tok, [txt('要建檔菜單需要管理員權限：請向管理員索取邀請碼，輸入「管理員 邀請碼」。\n點餐請在群組輸入「開團」。')]);
  }

  // 以下為管理員
  if (t === '說明' || t === '指令' || t === 'help') return reply(tok, [txt(HELP_ADMIN)]);

  w = t.match(/^(設定)?門檻\s*(.+?)\s+\$?(\d+)$/);
  if (w) {
    const st = findStore(w[2]);
    if (!st) return reply(tok, [txt('找不到店家「' + w[2] + '」。')]);
    withLock(() => { const r = readAll('stores').find(x => x.storeId === st.storeId); r.minOrder = Number(w[3]); r.updatedAt = now(); updateRow('stores', r); });
    return reply(tok, [txt('已設定「' + st.name + '」外送門檻 ' + money(w[3]) + '。')]);
  }

  if (t === '邀請管理員') {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    CacheService.getScriptCache().put('invite_' + code, '1', 1800);
    return reply(tok, [txt('邀請碼：' + code + '（30 分鐘內有效，只能用一次）\n請對方加我好友後私訊：\n管理員 ' + code)]);
  }

  if (t === '店家') return reply(tok, [txt(storeListText())]);

  m = t.match(/^刪除店家\s*(.+)$/);
  if (m) {
    const s = findStore(m[1]);
    if (!s) return reply(tok, [txt('找不到店家「' + m[1] + '」。')]);
    withLock(() => deleteRow('stores', readAll('stores').find(x => x.storeId === s.storeId)));
    return reply(tok, [txt('已刪除「' + s.name + '」。')]);
  }

  m = t.match(/^搜尋\s*(.+)$/);
  if (m) {
    const q = m[1].trim();
    let data;
    try {
      data = aiSearchMenu(q);
    } catch (err) {
      console.error(err);
      const msg = String(err.message || err);
      if (msg.indexOf('429') >= 0) return reply(tok, [txt('搜尋功能目前用不了：Google 的免費額度不包含網路搜尋，需要開通付費方案。\n可以改傳菜單截圖，或把店名交給管理員用電腦匯入。')], uid);
      return reply(tok, [txt('搜尋失敗：' + msg.slice(0, 200) + '\n可以換個關鍵字再試，或改傳菜單截圖。')], uid);
    }
    if (!countItems(data)) return reply(tok, [txt('找不到「' + q + '」的菜單。可以加上分店或地區再搜尋一次，或改傳菜單截圖。')], uid);
    saveDraft(uid, data);
    return reply(tok, [txt('🔎 以下是網路上找到的資料，價格可能不是最新，請核對後再回「確認」。\n\n' + draftPreview(data))], uid);
  }

  const draft = getDraft(uid);

  if (t === '取消') {
    if (draft) withLock(() => deleteRow('drafts', draft.row));
    return reply(tok, [txt('已放棄草稿。')]);
  }

  if (t === '預覽') {
    if (!draft) return reply(tok, [txt('目前沒有草稿，請傳菜單照片。')]);
    return reply(tok, [txt(draftPreview(draft.data))]);
  }

  if (t === '確認' || t === '確認建檔') {
    if (!draft) return reply(tok, [txt('目前沒有草稿，請傳菜單照片。')]);
    const d = draft.data;
    if (!d.store.name) return reply(tok, [txt('還沒有店名，請打字告訴我，例如「店名是 五十嵐 民生店」。')]);
    const saved = withLock(() => {
      const existing = readAll('stores').find(s => s.name === d.store.name);
      const rec = existing || { storeId: shortId(), createdAt: now() };
      ['phone', 'address', 'hours', 'contact', 'line', 'minOrder', 'note'].forEach(k => {
        if (d.store[k] || !existing) rec[k] = d.store[k] || '';
      });
      rec.name = d.store.name;
      rec.menuJson = JSON.stringify({ categories: d.categories });
      rec.updatedAt = now();
      if (existing) updateRow('stores', rec); else insertRow('stores', rec);
      deleteRow('drafts', readAll('drafts').find(x => x.userId === uid));
      return { rec: rec, updated: !!existing };
    });
    return reply(tok, [txt('✅ 已' + (saved.updated ? '更新' : '建檔') + '「' + saved.rec.name + '」，共 ' + countItems(d) + ' 個品項。\n在群組輸入「開團 ' + saved.rec.name + '」就能開團。')]);
  }

  if (draft) {
    // 其他文字視為修改草稿
    const updated = aiReviseDraft(draft.data, t);
    saveDraft(uid, updated);
    return reply(tok, [txt('已修改。\n\n' + draftPreview(updated))], uid);
  }

  return reply(tok, [txt('傳菜單照片給我就能建檔，或輸入「說明」看所有指令。')]);
}

/* ============================== 菜單辨識 ============================== */

function getDraft(uid) {
  const row = readAll('drafts').find(d => d.userId === uid);
  if (!row) return null;
  try { return { row: row, data: normalizeMenu(JSON.parse(row.draftJson)) }; } catch (e) { return null; }
}

function saveDraft(uid, data) {
  withLock(() => {
    const row = readAll('drafts').find(d => d.userId === uid);
    const rec = row || { userId: uid };
    rec.draftJson = JSON.stringify(data);
    rec.updatedAt = now();
    if (row) updateRow('drafts', rec); else insertRow('drafts', rec);
  });
}

function fetchLineImage(messageId) {
  const r = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + messageId + '/content', {
    headers: { Authorization: 'Bearer ' + cfg('CHANNEL_ACCESS_TOKEN', '') },
    muteHttpExceptions: true
  });
  return r.getResponseCode() === 200 ? r.getBlob() : null;
}

/**
 * 一次傳多張照片時，LINE 會同時送來多個事件。
 * 第一張的執行程序負責處理：先等幾秒讓其他照片排進佇列，再整批交給 AI；
 * 其他執行程序只負責把照片排進佇列，避免同時改草稿互相覆蓋。
 */
function handleMenuImage(ev) {
  const uid = ev.source.userId;
  const cache = CacheService.getScriptCache();
  const qKey = 'imgq_' + uid, busyKey = 'imgbusy_' + uid;
  let isProcessor = false;
  withLock(() => {
    const q = JSON.parse(cache.get(qKey) || '[]');
    q.push(ev.message.id);
    cache.put(qKey, JSON.stringify(q), 600);
    if (!cache.get(busyKey)) { cache.put(busyKey, '1', 300); isProcessor = true; }
  });
  if (!isProcessor) return;

  Utilities.sleep(6000);
  const hadDraft = !!getDraft(uid);
  let total = 0, failed = 0, data = null;
  const errors = [];
  for (let round = 0; round < 10; round++) {
    let ids = [];
    withLock(() => {
      ids = JSON.parse(cache.get(qKey) || '[]');
      cache.remove(qKey);
      if (!ids.length) cache.remove(busyKey);
    });
    if (!ids.length) break;
    for (let i = 0; i < ids.length; i += 6) {
      const blobs = [];
      ids.slice(i, i + 6).forEach(id => { const b = fetchLineImage(id); if (b) blobs.push(b); else failed++; });
      if (!blobs.length) continue;
      const draft = getDraft(uid);
      try {
        data = aiReadMenu(blobs, draft ? draft.data : null);
        saveDraft(uid, data);
        total += blobs.length;
      } catch (err) {
        console.error(err);
        failed += blobs.length;
        errors.push(String(err.message || err).slice(0, 150));
      }
    }
  }
  if (!data) {
    return reply(ev.replyToken, [txt('辨識失敗' + (errors.length ? '：' + errors[0] : '') + '\n請稍後再傳一次，或換清楚一點的照片。')], uid);
  }
  let head = '已辨識 ' + total + ' 張照片' + (hadDraft ? '，並合併到原本的草稿' : '') + '。';
  if (failed) head += '\n⚠️ 有 ' + failed + ' 張沒有成功，可以重傳那幾張。';
  return reply(ev.replyToken, [txt(head + '\n\n' + draftPreview(data))], uid);
}

const MENU_SCHEMA_PROMPT = [
  '輸出格式（只輸出 JSON，不要任何說明文字或 markdown）：',
  '{"store":{"name":"店名（含分店）","phone":"電話，多支用、分隔","address":"地址","hours":"營業時間","line":"店家 LINE 官方帳號 ID（例如 @abc1234）或 LINE 加好友網址","minOrder":外送門檻金額數字或null,"contact":"其他聯絡方式，例如 Facebook、IG、官網、外送平台","note":"其他重要資訊，例如外送門檻、公休日"},',
  ' "categories":[{"name":"分類名稱","items":[{"name":"品名","price":數字或null,"sizes":[{"name":"M","price":50}],"options":[{"name":"加珍珠","price":10}],"note":""}]}]}',
  '規則：',
  '- 價格一律是新台幣整數。品項有多種尺寸或份量時填 sizes，price 填 null；只有一個價格時 sizes 為 []。',
  '- options 只放需要加價的加料或選項，並填加價金額；甜度、冰塊等免費選項不要放。',
  '- 若菜單上有「全品項可加購」的加料表，請把這些加料放進每個適用品項的 options。',
  '- 店家聯絡方式要仔細找，包含角落的小字、QR code 旁的文字。看到「LINE」「加入好友」「@」開頭的 ID 或 line.me 網址，一定要填進 line 欄位。',
  '- 看不清楚或菜單上沒有的欄位填空字串，不要自己猜。',
  '- 品名、分類保留菜單上的原本用字。'
].join('\n');

function aiReadMenu(blobs, existing) {
  const n = blobs.length;
  let prompt = '你是菜單辨識助手。' + (n > 1 ? '以下 ' + n + ' 張照片是同一家店的菜單（可能是外送 APP 的連續截圖，內容會有重疊，重複的品項只列一次）。' : '') +
    '請從菜單照片擷取店家資訊與完整菜單。\n' + MENU_SCHEMA_PROMPT;
  if (existing) {
    prompt += '\n\n以下是同一家店目前的草稿，請把新照片的內容合併進去（重複品項以新照片為準，店家資訊保留已有的、補上新的），輸出合併後的完整 JSON：\n' + JSON.stringify(existing);
  }
  return normalizeMenu(callAI(prompt, blobs));
}

function aiSearchMenu(query) {
  const prompt = '請用 Google 搜尋台灣的店家「' + query + '」，找出店家資訊和目前的完整菜單與價格' +
    '（優先採用官方網站、官方社群、Foodpanda、Uber Eats 等外送平台，以及最近一年內的資料）。\n' +
    '如果有多家同名分店，以最符合關鍵字的那家為主，店名寫完整（含分店）。\n' +
    '在 store.note 註明資料來源網站和大約日期。找不到菜單就讓 categories 為空陣列。\n' + MENU_SCHEMA_PROMPT;
  return normalizeMenu(callAI(prompt, null, { search: true }));
}

/** 供 Claude 在編輯器直接匯入整理好的菜單資料（同名店家會更新） */
function importStore(raw) {
  const d = normalizeMenu(raw);
  if (!d.store.name) throw new Error('缺少店名');
  return withLock(() => {
    const existing = readAll('stores').find(s => s.name === d.store.name);
    const rec = existing || { storeId: shortId(), createdAt: now() };
    ['phone', 'address', 'hours', 'contact', 'line', 'minOrder', 'note'].forEach(k => {
      if (d.store[k] || !existing) rec[k] = d.store[k] || '';
    });
    rec.name = d.store.name;
    rec.menuJson = JSON.stringify({ categories: d.categories });
    rec.updatedAt = now();
    if (existing) updateRow('stores', rec); else insertRow('stores', rec);
    return (existing ? '已更新 ' : '已新增 ') + rec.name + '，共 ' + countItems(d) + ' 個品項';
  });
}

function aiReviseDraft(data, instruction) {
  const prompt = '以下是菜單草稿 JSON，請依照使用者的修改指示更新，輸出更新後的完整 JSON。\n' +
    MENU_SCHEMA_PROMPT + '\n\n草稿：\n' + JSON.stringify(data) + '\n\n修改指示：' + instruction;
  return normalizeMenu(callAI(prompt, null));
}

function callAI(prompt, blob, opts) {
  opts = opts || {};
  const provider = cfg('AI_PROVIDER', 'gemini');
  let text;
  const blobs = !blob ? [] : (Array.isArray(blob) ? blob : [blob]);
  if (provider === 'claude') {
    if (opts.search) throw new Error('搜尋功能需要使用 Gemini');
    const content = [];
    blobs.forEach(b => content.push({ type: 'image', source: { type: 'base64', media_type: b.getContentType() || 'image/jpeg', data: Utilities.base64Encode(b.getBytes()) } }));
    content.push({ type: 'text', text: prompt });
    const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': cfg('ANTHROPIC_API_KEY', ''), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: cfg('CLAUDE_MODEL', 'claude-haiku-5-5'), max_tokens: 16000, messages: [{ role: 'user', content: content }] })
    });
    if (r.getResponseCode() !== 200) throw new Error('Claude API ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
    text = JSON.parse(r.getContentText()).content.filter(c => c.type === 'text').map(c => c.text).join('');
  } else {
    const parts = [];
    blobs.forEach(b => parts.push({ inline_data: { mime_type: b.getContentType() || 'image/jpeg', data: Utilities.base64Encode(b.getBytes()) } }));
    parts.push({ text: prompt });
    // 主模型忙碌（503/429）時自動重試，再改用備用模型
    const main = cfg('GEMINI_MODEL', 'gemini-flash-latest');
    const attempts = [main, main, 'gemini-flash-lite-latest', 'gemini-flash-lite-latest'];
    const body = { contents: [{ role: 'user', parts: parts }] };
    if (opts.search) {
      body.tools = [{ google_search: {} }]; // 搜尋模式不能同時指定 JSON 輸出格式
      body.generationConfig = { temperature: 0 };
    } else {
      body.generationConfig = { responseMimeType: 'application/json', temperature: 0 };
    }
    let r = null;
    for (let i = 0; i < attempts.length; i++) {
      r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + attempts[i] + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-goog-api-key': cfg('GEMINI_API_KEY', '') },
        payload: JSON.stringify(body)
      });
      const code = r.getResponseCode();
      if (code !== 200) console.warn('Gemini ' + attempts[i] + ' ' + code + ' ' + r.getContentText().slice(0, 300));
      if (code === 200 || [404, 429, 500, 503].indexOf(code) < 0) break;
      if (i < attempts.length - 1) Utilities.sleep(1500);
    }
    if (r.getResponseCode() !== 200) throw new Error('Gemini API ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
    const cand = (JSON.parse(r.getContentText()).candidates || [])[0];
    text = cand && cand.content && cand.content.parts ? cand.content.parts.map(p => p.text || '').join('') : '';
  }
  return parseJsonLoose(text);
}

function parseJsonLoose(text) {
  let s = String(text || '').trim().replace(/^\x60{3}(?:json)?\s*/i, '').replace(/\x60{3}\s*$/, '');
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a < 0 || b < a) throw new Error('AI 沒有回傳菜單資料');
  return JSON.parse(s.slice(a, b + 1));
}

function cleanPrice(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(String(v).replace(/[^\d.]/g, ''));
  return isFinite(n) && String(v).match(/\d/) ? Math.round(n) : null;
}

function normalizeMenu(raw) {
  raw = raw || {};
  const st = raw.store || {};
  const store = {};
  ['name', 'phone', 'address', 'hours', 'contact', 'line', 'note'].forEach(k => { store[k] = st[k] ? String(st[k]).trim() : ''; });
  store.minOrder = cleanPrice(st.minOrder) || '';
  store.line = normalizeLineId(store.line);
  const categories = (raw.categories || []).map(c => ({
    name: String(c.name || '其他').trim(),
    items: (c.items || []).filter(i => i && i.name).map(i => {
      const sizes = (i.sizes || []).filter(s => s && s.name).map(s => ({ name: String(s.name).trim(), price: cleanPrice(s.price) || 0 }));
      return {
        name: String(i.name).trim(),
        price: sizes.length ? null : (cleanPrice(i.price) || 0),
        sizes: sizes,
        options: (i.options || []).filter(o => o && o.name).map(o => ({ name: String(o.name).trim(), price: cleanPrice(o.price) || 0 })),
        note: i.note ? String(i.note).trim() : ''
      };
    })
  })).filter(c => c.items.length);
  return { store: store, categories: categories };
}

function countItems(d) { return d.categories.reduce((n, c) => n + c.items.length, 0); }

function itemPriceText(i) {
  let s = i.sizes.length ? i.sizes.map(z => z.name + money(z.price)).join(' / ') : money(i.price);
  if (i.options.length) s += '（加：' + i.options.map(o => o.name + '+' + o.price).join('、') + '）';
  return s;
}

function draftPreview(d) {
  const st = d.store;
  const head = ['📋 菜單草稿：' + (st.name || '（未辨識到店名）')];
  if (st.phone) head.push('☎ ' + st.phone);
  if (st.address) head.push('📍 ' + st.address);
  if (st.hours) head.push('🕘 ' + st.hours);
  if (st.line) head.push('🟢 LINE ' + st.line);
  if (st.minOrder) head.push('🛵 外送門檻 ' + money(st.minOrder));
  if (st.contact) head.push('💬 ' + st.contact);
  if (st.note) head.push('📝 ' + st.note);
  head.push('共 ' + d.categories.length + ' 類、' + countItems(d) + ' 個品項');
  let body = '';
  for (const c of d.categories) {
    const block = '\n【' + c.name + '】\n' + c.items.map(i => '・' + i.name + ' ' + itemPriceText(i)).join('\n');
    if ((head.join('\n') + body + block).length > 4300) { body += '\n…（其餘品項略，建檔後可在點餐頁看到全部）'; break; }
    body += block;
  }
  const tail = '\n\n👉 回「確認」建檔、「取消」放棄；要改就直接打字（例：珍奶 L 改 65）。還有下一頁就再傳照片，會自動合併。';
  return head.join('\n') + '\n' + body + tail;
}

/* ============================== 店家 ============================== */

function findStore(name) {
  const q = String(name).trim();
  const stores = readAll('stores');
  return stores.find(s => s.name === q) || stores.find(s => s.name.indexOf(q) >= 0) || null;
}

/** 整理 LINE ID：去掉「LINE:」等前綴；網址保持原樣 */
function normalizeLineId(v) {
  v = String(v || '').trim().replace(/^(LINE\s*(ID)?|line\s*(id)?)\s*[:：]?\s*/i, '');
  if (!v) return '';
  if (/^https?:\/\//i.test(v)) return v;
  return v.replace(/\s+/g, '');
}

/** LINE ID 轉成加好友網址 */
function lineAddUrl(v) {
  if (/^https?:\/\//i.test(v)) return v;
  const id = v.charAt(0) === '@' ? v : '@' + v;
  return 'https://line.me/R/ti/p/' + encodeURIComponent(id);
}

/** 給店家的訂單文字（備註不帶名字） */
function storeOrderText(order) {
  const list = mergedItems(orderEntries(order.orderId));
  const total = list.reduce((n, x) => n + x.qty, 0);
  return '您好，我想訂購（共 ' + total + ' 份）：\n' +
    list.map(x => x.label + ' ×' + x.qty + (x.notes.length ? '（' + x.notes.map(n => n.replace(/^[^：]*：/, '')).join('、') + '）' : '')).join('\n') +
    '\n謝謝！';
}

/** 開啟店家 LINE 官方帳號聊天室並預填訊息的網址；店家沒有 @ID 時回傳空字串 */
function storeOaUrl(store, text) {
  const v = String(store.line || '').trim();
  if (!v || /^https?:\/\//i.test(v)) return '';
  const id = v.charAt(0) === '@' ? v : '@' + v;
  return 'https://line.me/R/oaMessage/' + encodeURIComponent(id) + '/?' + encodeURIComponent(text);
}

function storeContactLine(s) {
  const parts = [];
  if (s.phone) parts.push('☎ ' + s.phone);
  if (s.address) parts.push('📍 ' + s.address);
  if (s.line) parts.push('🟢 LINE ' + s.line + '（加好友：' + lineAddUrl(s.line) + '）');
  if (s.contact) parts.push('💬 ' + s.contact);
  return parts.join('\n');
}

function storeListText() {
  const stores = readAll('stores');
  if (!stores.length) return '還沒有任何店家。管理員私訊我菜單照片就能建檔。';
  return '🏪 已建檔店家\n\n' + stores.map((s, i) => {
    const extra = [s.phone && '☎ ' + s.phone, s.hours && '🕘 ' + s.hours, s.line && '🟢 LINE ' + s.line, s.contact && '💬 ' + s.contact].filter(Boolean).join('\n   ');
    return (i + 1) + '. ' + s.name + (extra ? '\n   ' + extra : '');
  }).join('\n\n');
}

function storeMenu(store) {
  try { return JSON.parse(store.menuJson || '{}').categories || []; } catch (e) { return []; }
}

/* ============================== 金額計算 ============================== */

/** 依菜單重新計算價格，避免前端被竄改 */
function priceLine(categories, it) {
  let item = null;
  for (const c of categories) { item = c.items.find(i => i.name === it.name); if (item) break; }
  if (!item) throw new Error('找不到品項：' + it.name + '（菜單可能已更新，請重新整理）');
  let unit, size = '';
  if (item.sizes && item.sizes.length) {
    const s = item.sizes.find(z => z.name === it.size);
    if (!s) throw new Error(item.name + ' 請選擇尺寸');
    unit = num(s.price); size = s.name;
  } else {
    unit = num(item.price);
  }
  const opts = (it.options || []).map(n => {
    const o = (item.options || []).find(x => x.name === n);
    if (!o) throw new Error('找不到加料：' + n);
    unit += num(o.price);
    return o.name;
  });
  const qty = Math.max(1, Math.min(99, parseInt(it.qty, 10) || 1));
  return { name: item.name, size: size, options: opts, sugar: String(it.sugar || '').slice(0, 10), ice: String(it.ice || '').slice(0, 10),
    qty: qty, note: String(it.note || '').slice(0, 60), unit: unit, total: unit * qty };
}

/** 運費減折扣平均分攤，尾數依序分給前面的人，確保加總剛好等於總額 */
function computeTotals(order, entries) {
  const n = entries.length;
  const extra = num(order.deliveryFee) - num(order.discount);
  const base = n ? (extra >= 0 ? Math.floor(extra / n) : Math.ceil(extra / n)) : 0;
  let rem = n ? extra - base * n : 0;
  let itemsTotal = 0, paid = 0, claimed = 0;
  const unpaidNames = [];
  entries.forEach(en => {
    let share = base;
    if (rem > 0) { share++; rem--; } else if (rem < 0) { share--; rem++; }
    en.share = share;
    en.amount = num(en.subtotal) + share;
    itemsTotal += num(en.subtotal);
    if (en.payStatus === 'paid') paid += en.amount;
    else if (en.payStatus === 'claimed') claimed += en.amount;
    else unpaidNames.push(en.displayName);
  });
  const grand = itemsTotal + (n ? extra : 0);
  return { people: n, itemsTotal: itemsTotal, deliveryFee: num(order.deliveryFee), discount: num(order.discount), grand: grand, paid: paid, claimed: claimed, outstanding: grand - paid, unpaidNames: unpaidNames };
}

function orderEntries(orderId) {
  return readAll('entries').filter(e => e.orderId === orderId).map(e => {
    let items = [];
    try { items = JSON.parse(e.itemsJson || '[]'); } catch (x) { items = []; }
    return { entryId: e.entryId, userId: e.userId, displayName: e.displayName, items: items, subtotal: num(e.subtotal), payStatus: e.payStatus || 'unpaid', payMethod: e.payMethod || '', updatedAt: e.updatedAt };
  });
}

/** 把所有人的餐點合併成下單清單（同品項、同尺寸、同加料合併數量，備註標上是誰的） */
function mergedItems(entries) {
  const map = {}, keys = [];
  entries.forEach(e => (e.items || []).forEach(i => {
    const opts = (i.options || []).join('、');
    const si = [i.sugar, i.ice].filter(Boolean).join(' ');
    const key = i.name + '|' + (i.size || '') + '|' + opts + '|' + si;
    if (!map[key]) {
      map[key] = { label: i.name + (i.size ? ' ' + i.size : '') + (opts ? '（' + opts + '）' : '') + (si ? ' ' + si : ''), qty: 0, notes: [] };
      keys.push(key);
    }
    map[key].qty += num(i.qty);
    if (i.note) map[key].notes.push(e.displayName + '：' + i.note + (i.qty > 1 ? ' ×' + i.qty : ''));
  }));
  return keys.map(k => map[k]);
}

function mergedListText(order) {
  const list = mergedItems(orderEntries(order.orderId));
  if (!list.length) return '還沒有人點餐。';
  const total = list.reduce((n, x) => n + x.qty, 0);
  return '📋 下單清單：' + order.storeName + '（共 ' + total + ' 份）\n' +
    list.map(x => '・' + x.label + ' ×' + x.qty + (x.notes.length ? '\n   備註：' + x.notes.join('；') : '')).join('\n');
}

/** 用 LINE 的「提及」功能標記還沒付款的人 */
function remindMessage(order) {
  const entries = orderEntries(order.orderId);
  computeTotals(order, entries);
  const unpaid = entries.filter(e => e.payStatus === 'unpaid');
  if (!unpaid.length) return txt('🎉 大家都付了或已回報，不用催！');
  const subs = {};
  const lines = unpaid.map((e, i) => {
    subs['u' + i] = { type: 'mention', mentionee: { type: 'user', userId: e.userId } };
    return '{u' + i + '} ' + money(e.amount);
  });
  return {
    type: 'textV2',
    text: '💰 「' + order.storeName + '」還沒付款，請付給 ' + order.creatorName + '：\n' + lines.join('\n') +
      '\n\n付完請到點餐頁按「我付了」：\n' + liffUrl(order.orderId),
    substitution: subs
  };
}

/** 這個人上一次在同一家店點的餐點（給「跟上次一樣」用） */
function lastOrderItems(userId, storeId, excludeOrderId) {
  const orderStore = {};
  readAll('orders').forEach(o => { orderStore[o.orderId] = o.storeId; });
  const mine = readAll('entries').filter(e => e.userId === userId && e.orderId !== excludeOrderId && orderStore[e.orderId] === storeId);
  if (!mine.length) return [];
  try {
    return JSON.parse(mine[mine.length - 1].itemsJson || '[]')
      .map(i => ({ name: i.name, size: i.size || '', options: i.options || [], sugar: i.sugar || '', ice: i.ice || '', qty: i.qty, note: i.note || '' }));
  } catch (e) { return []; }
}

/* ============================== 群組成員名單 ============================== */

/** 記住在群組說過話的人（用快取避免每則訊息都寫試算表） */
function touchMember(src) {
  const gid = groupKey(src), uid = src.userId;
  const cache = CacheService.getScriptCache();
  const key = 'mem_' + gid.slice(-12) + '_' + uid.slice(-12);
  if (cache.get(key)) return;
  const name = getDisplayName(src);
  withLock(() => {
    const row = readAll('members').find(r => r.groupId === gid && r.userId === uid);
    const rec = row || { groupId: gid, userId: uid };
    rec.name = name; rec.lastSeen = now(); rec.left = '';
    if (row) updateRow('members', rec); else insertRow('members', rec);
  });
  cache.put(key, '1', 21600);
}

function markMemberLeft(gid, uid) {
  withLock(() => {
    const row = readAll('members').find(r => r.groupId === gid && r.userId === uid);
    if (row) { row.left = '1'; updateRow('members', row); }
  });
}

/** 已知成員（說過話的人 + 在這個群組點過餐的人）扣掉這團已點的人 */
function notOrderedNames(order) {
  const ordered = {};
  orderEntries(order.orderId).forEach(e => { ordered[e.userId] = true; });
  const known = {};
  readAll('members').filter(r => r.groupId === order.groupId).forEach(r => { known[r.userId] = r.left ? null : r.name; });
  const groupOrders = {};
  readAll('orders').filter(o => o.groupId === order.groupId).forEach(o => { groupOrders[o.orderId] = true; });
  readAll('entries').filter(e => groupOrders[e.orderId]).forEach(e => { if (!(e.userId in known)) known[e.userId] = e.displayName; });
  return Object.keys(known).filter(uid => known[uid] && !ordered[uid]).map(uid => known[uid]);
}

/* ============================== 開團共用 ============================== */

/** 建立一團並回傳要貼在群組的訊息 */
function openOrderMessage(gid, creatorId, creatorName, store) {
  const res = withLock(() => {
    const cur = activeOrder(gid);
    if (cur) return { exists: cur };
    const order = {
      orderId: shortId(), groupId: gid, storeId: store.storeId, storeName: store.name, status: 'open',
      creatorId: creatorId, creatorName: creatorName, deliveryFee: 0, discount: 0, createdAt: now(), closedAt: ''
    };
    insertRow('orders', order);
    return { order: order };
  });
  if (res.exists) return txt('目前已有一團「' + res.exists.storeName + '」，請先「結單」。');
  const o = res.order;
  const min = num(store.minOrder);
  return txt('🍱 開團：' + store.name + '\n由 ' + o.creatorName + ' 開團、收款\n' +
    storeContactLine(store) + (min ? '\n🛵 外送門檻 ' + money(min) : '') + '\n\n👉 點這裡點餐：\n' + liffUrl(o.orderId) +
    '\n\n點完後輸入「誰沒付」可查看付款狀況。');
}

/* ============================== 投票 ============================== */

function loadPoll(gid) {
  const v = CacheService.getScriptCache().get('poll_' + gid);
  try { return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function savePoll(gid, poll) { CacheService.getScriptCache().put('poll_' + gid, JSON.stringify(poll), 21600); }
function pollTally(poll) {
  const c = {};
  poll.options.forEach(o => { c[o] = 0; });
  Object.keys(poll.votes).forEach(u => { if (c[poll.votes[u]] !== undefined) c[poll.votes[u]]++; });
  return poll.options.map(o => ({ name: o, count: c[o] })).sort((a, b) => b.count - a.count);
}

/* ============================== 收款碼 ============================== */

function handlePayQrImage(ev) {
  const uid = ev.source.userId;
  CacheService.getScriptCache().remove('awaitpay_' + uid);
  const blob = fetchLineImage(ev.message.id);
  if (!blob) return reply(ev.replyToken, [txt('讀取圖片失敗，請重新輸入「設定收款碼」再傳一次。')]);
  const name = getDisplayName(ev.source);
  const file = DriveApp.createFile(blob.setName('收款碼_' + name + '.jpg'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  withLock(() => {
    const row = readAll('payqr').find(x => x.userId === uid);
    if (row && row.fileId) { try { DriveApp.getFileById(row.fileId).setTrashed(true); } catch (e) { /* ignore */ } }
    const rec = row || { userId: uid };
    rec.name = name; rec.fileId = file.getId(); rec.updatedAt = now();
    if (row) updateRow('payqr', rec); else insertRow('payqr', rec);
  });
  return reply(ev.replyToken, [txt('✅ 收款碼已設定！之後你開的團，點餐頁的「總表・付款」會顯示這張收款碼給大家掃。\n要換掉就再輸入「設定收款碼」，要移除輸入「刪除收款碼」。')]);
}

function payQrUrl(userId) {
  const row = readAll('payqr').find(x => x.userId === userId);
  return row && row.fileId ? 'https://drive.google.com/thumbnail?id=' + row.fileId + '&sz=w600' : '';
}

/* ============================== 許願 ============================== */

/** 把使用者想要的新功能記到試算表 wishes 分頁，之後交給 Claude 處理 */
function saveWish(src, text) {
  const name = getDisplayName(src);
  insertRow('wishes', { createdAt: now(), userId: src.userId, name: name, text: String(text).trim().slice(0, 1000), status: '待處理' });
  return '🌟 收到 ' + name + ' 的願望，已經記下來了！\n會在下次更新機器人時一起處理。';
}

/* ============================== 每週欠款公告 ============================== */

function shortDate(iso) {
  try { return Utilities.formatDate(new Date(iso), 'Asia/Taipei', 'M/d'); } catch (e) { return ''; }
}

/** 整理某個群組所有團（含已結單）裡還沒付清的款項；沒有欠款回傳 null */
function debtText(groupId) {
  const orders = readAll('orders').filter(o => o.groupId === groupId);
  if (!orders.length) return null;
  const allEntries = readAll('entries');
  const people = {}, order = [];
  orders.forEach(o => {
    const es = allEntries.filter(e => e.orderId === o.orderId).map(e => ({
      userId: e.userId, displayName: e.displayName, subtotal: num(e.subtotal), payStatus: e.payStatus || 'unpaid'
    }));
    computeTotals(o, es);
    es.forEach(e => {
      if (e.payStatus === 'paid' || e.amount <= 0) return;
      if (!people[e.userId]) { people[e.userId] = { name: e.displayName, total: 0, list: [] }; order.push(e.userId); }
      const p = people[e.userId];
      p.total += e.amount;
      p.list.push(o.storeName + '（' + shortDate(o.createdAt) + '，付給 ' + o.creatorName + '）' + money(e.amount) +
        (e.payStatus === 'claimed' ? ' 已回報待確認' : ''));
    });
  });
  if (!order.length) return null;
  return '📒 每週欠款公告\n以下款項還沒結清，付完請到該團點餐頁按「我付了」，或直接跟收款的人說一聲：\n\n' +
    order.map(id => '・' + people[id].name + ' 共 ' + money(people[id].total) + '\n   ' + people[id].list.join('\n   ')).join('\n');
}

/** 由排程觸發：每個有欠款的群組各發一則公告（沒有欠款就不發，節省訊息則數） */
function weeklyDebtReport() {
  const groups = {};
  readAll('orders').forEach(o => { if (o.groupId) groups[o.groupId] = true; });
  Object.keys(groups).forEach(gid => {
    const text = debtText(gid);
    if (!text) return;
    const r = lineApi('message/push', { to: gid, messages: [txt(text)] });
    if (r.getResponseCode() !== 200) console.warn('push ' + gid + ' ' + r.getContentText().slice(0, 200));
  });
}

/** 在編輯器執行一次：建立「每週一 10:00」的排程 */
function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklyDebtReport').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklyDebtReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(10).nearMinute(0).inTimezone('Asia/Taipei').create();
  Logger.log('已設定每週一 10:00 欠款公告');
}

function orderSummaryText(order) {
  const entries = orderEntries(order.orderId);
  const t = computeTotals(order, entries);
  const lines = ['🍱 ' + order.storeName + '（' + STATUS_LABEL[order.status] + '）'];
  if (!entries.length) {
    lines.push('還沒有人點餐。');
  } else {
    let fee = '';
    if (t.deliveryFee || t.discount) fee = '（含運費 ' + money(t.deliveryFee) + (t.discount ? '、折扣 -' + money(t.discount) : '') + '）';
    lines.push('共 ' + t.people + ' 人，總額 ' + money(t.grand) + fee);
    lines.push('已收 ' + money(t.paid) + '，還差 ' + money(t.outstanding));
    const st = readAll('stores').find(x => x.storeId === order.storeId);
    const min = st ? num(st.minOrder) : 0;
    if (min && t.itemsTotal < min) lines.push('🛵 外送門檻 ' + money(min) + '，餐點還差 ' + money(min - t.itemsTotal));
    const group = (st) => entries.filter(e => e.payStatus === st)
      .map(e => e.displayName + ' ' + money(e.amount) + (st === 'claimed' && e.payMethod ? '（' + (METHOD_LABEL[e.payMethod] || e.payMethod) + '）' : ''));
    const paid = group('paid'), claimed = group('claimed'), unpaid = group('unpaid');
    if (unpaid.length) lines.push('\n❌ 未付款\n' + unpaid.join('\n'));
    if (claimed.length) lines.push('\n🕐 已回報、待 ' + order.creatorName + ' 確認\n' + claimed.join('\n'));
    if (paid.length) lines.push('\n✅ 已收款\n' + paid.join('\n'));
    if (!unpaid.length && !claimed.length) lines.push('\n🎉 全部收齊！');
  }
  lines.push('\n點餐／付款頁：' + liffUrl(order.orderId));
  return lines.join('\n');
}

/* ============================== 點餐頁 API ============================== */

function verifyUser(idToken) {
  if (!idToken) throw new Error('請從 LINE 開啟點餐頁');
  const key = 'tok_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const r = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', muteHttpExceptions: true,
    payload: { id_token: idToken, client_id: cfg('LOGIN_CHANNEL_ID', '') }
  });
  if (r.getResponseCode() !== 200) throw new Error('登入已過期，請關閉頁面後從 LINE 重新開啟');
  const p = JSON.parse(r.getContentText());
  const u = { userId: p.sub, name: p.name || '朋友' };
  cache.put(key, JSON.stringify(u), 600);
  return u;
}

function handleApi(body) {
  const me = verifyUser(body.idToken);
  const action = body.action;
  if (action === 'getOrder') return getOrderView(body.orderId, me);
  if (action === 'saveMyEntry') { withLock(() => saveMyEntry(body, me)); return getOrderView(body.orderId, me); }
  if (action === 'setPay') { withLock(() => setPay(body, me)); return getOrderView(body.orderId, me); }
  if (action === 'setOrder') { withLock(() => setOrder(body, me)); return getOrderView(body.orderId, me); }
  throw new Error('未知的操作');
}

function loadOrder(orderId) {
  const o = readAll('orders').find(x => x.orderId === orderId);
  if (!o) throw new Error('找不到這一團，可能已被刪除');
  return o;
}

function getOrderView(orderId, me) {
  const order = loadOrder(orderId);
  const store = readAll('stores').find(s => s.storeId === order.storeId) || { name: order.storeName };
  const entries = orderEntries(orderId);
  const totals = computeTotals(order, entries);
  return {
    me: { userId: me.userId, name: me.name, isManager: canManage(order, me.userId), lastItems: lastOrderItems(me.userId, order.storeId, order.orderId) },
    order: { orderId: order.orderId, storeName: order.storeName, status: order.status, creatorId: order.creatorId, creatorName: order.creatorName, deliveryFee: num(order.deliveryFee), discount: num(order.discount), createdAt: order.createdAt },
    store: { name: store.name, phone: store.phone || '', address: store.address || '', hours: store.hours || '', contact: store.contact || '', line: store.line || '', minOrder: num(store.minOrder), lineUrl: store.line ? lineAddUrl(store.line) : '', note: store.note || '', categories: storeMenu(store) },
    entries: entries,
    merged: mergedItems(entries),
    notOrdered: notOrderedNames(order),
    storeOrderUrl: store.storeId ? storeOaUrl(store, storeOrderText(order)) : '',
    totals: totals,
    payInfo: cfg('PAY_INFO', ''),
    payQr: payQrUrl(order.creatorId)
  };
}

function saveMyEntry(body, me) {
  const order = loadOrder(body.orderId);
  if (order.status !== 'open') throw new Error('已經收單，不能再修改餐點');
  const store = readAll('stores').find(s => s.storeId === order.storeId);
  if (!store) throw new Error('找不到店家資料');
  const cats = storeMenu(store);
  const items = (body.items || []).slice(0, 50).map(it => priceLine(cats, it));
  const subtotal = items.reduce((s, i) => s + i.total, 0);
  const existing = readAll('entries').find(e => e.orderId === order.orderId && e.userId === me.userId);
  if (!items.length) { if (existing) deleteRow('entries', existing); return; }
  if (existing && existing.payStatus === 'paid') throw new Error('你的款項已確認收款，要修改請先請 ' + order.creatorName + ' 改回未付');
  const rec = existing || { entryId: shortId(), orderId: order.orderId, userId: me.userId, payStatus: 'unpaid', payMethod: '' };
  rec.displayName = me.name;
  rec.itemsJson = JSON.stringify(items);
  rec.subtotal = subtotal;
  if (existing && existing.payStatus === 'claimed' && num(existing.subtotal) !== subtotal) { rec.payStatus = 'unpaid'; rec.payMethod = ''; }
  rec.updatedAt = now();
  if (existing) updateRow('entries', rec); else insertRow('entries', rec);
}

function setPay(body, me) {
  const order = loadOrder(body.orderId);
  const target = body.userId || me.userId;
  const en = readAll('entries').find(e => e.orderId === order.orderId && e.userId === target);
  if (!en) throw new Error('找不到這筆訂單');
  const manager = canManage(order, me.userId);
  const self = target === me.userId;
  const to = body.status;
  if (['unpaid', 'claimed', 'paid'].indexOf(to) < 0) throw new Error('狀態錯誤');
  if (!manager) {
    if (!self) throw new Error('只能回報自己的付款');
    if (to === 'paid') throw new Error('收款要由 ' + order.creatorName + ' 確認');
    if (en.payStatus === 'paid') throw new Error('已確認收款，無法修改');
  }
  en.payStatus = to;
  if (to === 'claimed' || (to === 'paid' && body.method)) en.payMethod = METHOD_LABEL[body.method] ? body.method : (en.payMethod || '');
  if (to === 'unpaid') en.payMethod = '';
  en.updatedAt = now();
  updateRow('entries', en);
}

function setOrder(body, me) {
  const order = loadOrder(body.orderId);
  if (!canManage(order, me.userId)) throw new Error('只有開團的人或管理員可以修改');
  if (body.status && STATUS_LABEL[body.status]) {
    order.status = body.status;
    if (body.status !== 'open') order.closedAt = now();
  }
  if (body.deliveryFee !== undefined) order.deliveryFee = Math.max(0, Math.round(num(body.deliveryFee)));
  if (body.discount !== undefined) order.discount = Math.max(0, Math.round(num(body.discount)));
  updateRow('orders', order);
}
