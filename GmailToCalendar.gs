/***********************
 * Gmail (Primary) → Calendar (single deadline event) + Custom email reminders
 * + Done label to stop reminders but keep calendar record
 ***********************/

// === 配置区 ===
const MARKER = 'CreatedBy:GmailToCalendarV4';
const LABEL_ADDED = 'GTC_Added';        // 已创建日历事件（防重复建）
const LABEL_DONE  = 'GTC_处理完了';     // 你手动贴：处理完后停止提醒（保留事件）
const CALENDAR_ID = '';                 // 留空=默认主日历；如需指定日历填日历ID
const LABEL_OLD_IGNORED = 'GTC_旧邮件忽略';   // 老邮件不再处理
const PROP_INGEST_CUTOFF_MS = 'INGEST_CUTOFF_MS'; // 只处理此时间之后收到的邮件

// 只处理“像截止”的邮件（避免普通带日期邮件误入）
const KEYWORDS = ['截止', '期限', '締切', '締め切り', 'deadlin', 'due'];

// 提醒点：14天前，10天前，7天前，5天前，3天前，1天前，12小时前
const REMINDER_OFFSETS_MIN = [
  14 * 24 * 60,
  10 * 24 * 60,
  7  * 24 * 60,
  5  * 24 * 60,
  3  * 24 * 60,
  1  * 24 * 60,
  12 * 60
];

// 每小时跑一次提醒：窗口设 70 分钟，避免错过
const REMINDER_WINDOW_MIN = 70;

// 事件默认时长
const EVENT_DURATION_MIN = 15;

// 每次 ingest 最多处理线程数（防爆）
const MAX_THREADS_PER_RUN = 40;


// === 入口：一键安装触发器 ===
function installTriggers() {
  // 清理旧触发器（避免重复装）
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  initIngestCutoffIfMissing_();

  // 1) 每10分钟：Primary 收件箱 → 解析截止 → 建 1 个截止事件
  ScriptApp.newTrigger('ingestPrimaryInboxToCalendar')
    .timeBased()
    .everyMinutes(10)
    .create();

  // 2) 每小时：自定义提醒邮件（并尊重 Done 标签）
  ScriptApp.newTrigger('sendCustomReminderEmails')
    .timeBased()
    .everyHours(1)
    .create();

  // （可选）每天早上 8 点晨报：未来14天截止 + 未入日历的疑似截止邮件
  ScriptApp.newTrigger('sendMorningBriefing')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function initIngestCutoffIfMissing_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PROP_INGEST_CUTOFF_MS)) {
    props.setProperty(PROP_INGEST_CUTOFF_MS, String(Date.now())); // 默认从“现在开始”
  }
}

function setIngestCutoffToNow() {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_INGEST_CUTOFF_MS, String(Date.now()));
}

function getIngestCutoffMs_() {
  const v = PropertiesService.getScriptProperties().getProperty(PROP_INGEST_CUTOFF_MS);
  return v ? parseInt(v, 10) : 0;
}

/**
 * 只从 Gmail「主要收件箱 Primary」读取，并创建“单个截止事件”
 * - 线程若已贴 GTC_Added：跳过（防重复建）
 * - 线程若你已贴 GTC_处理完了：跳过（你都处理完了就不再建）
 */
function ingestPrimaryInboxToCalendar() {
  // 确保标签存在
  getOrCreateLabel_(LABEL_ADDED);
  getOrCreateLabel_(LABEL_DONE);

  initIngestCutoffIfMissing_();
  const cutoffMs = getIngestCutoffMs_();
  const afterStr = Utilities.formatDate(new Date(cutoffMs), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const oldIgnoredLabel = getOrCreateLabel_(LABEL_OLD_IGNORED);

  // Primary tab 对应 smart label：^smartlabel_personal
  const primaryOnly =
    'in:inbox ' +
    'label:^smartlabel_personal ' +
    '-label:^smartlabel_promo ' +
    '-label:^smartlabel_social ' +
    '-label:^smartlabel_updates ' +
    '-label:^smartlabel_forums ' +
    '-in:spam -in:trash -in:chats';

  const q = `${primaryOnly} after:${afterStr} ${buildKeywordQuery_()} -label:${LABEL_ADDED} -label:${LABEL_DONE} -label:${LABEL_OLD_IGNORED}`;
  const threads = GmailApp.search(q, 0, MAX_THREADS_PER_RUN);
  const cal = getCalendar_();
  const addedLabel = GmailApp.getUserLabelByName(LABEL_ADDED);

  threads.forEach(thread => {
    if (!thread.isInInbox()) return;

    const msg = thread.getMessages().slice(-1)[0];
    const msgDateMs = msg.getDate().getTime();
    if (msgDateMs < cutoffMs) {
      thread.addLabel(oldIgnoredLabel); // 标记为老邮件忽略，避免下次一直扫到
      return;
    }


    const subject = msg.getSubject() || '(no subject)';
    const body = msg.getPlainBody() || '';

    const parsed = parseDeadline_(subject, body);
    if (!parsed) return; // 解析不到明确时间：不建事件

    const start = parsed.start;
    const end = new Date(start.getTime() + EVENT_DURATION_MIN * 60000);

    const threadId = thread.getId();
    const gmailLink = `https://mail.google.com/mail/u/0/#all/${threadId}`;

    const title = `【截止】${subject}`;
    const desc = [
      MARKER,
      `Status: ACTIVE`,
      `ParsedDeadline: ${parsed.deadlineText}`,
      `From: ${msg.getFrom()}`,
      `MailDate: ${msg.getDate()}`,
      `ThreadId: ${threadId}`,
      `GmailThread: ${gmailLink}`,
      `SentRemindersMin: `,
      '',
      '--- Mail Snippet ---',
      clip_(body, 800)
    ].join('\n');

    const ev = cal.createEvent(title, start, end, { description: desc });
    ev.setDescription(`${ev.getDescription()}\nEventId: ${ev.getId()}`);

    // 标记线程已创建事件（防重复）
    thread.addLabel(addedLabel);
  });
}


/**
 * 自定义提醒（发邮件）
 * - 只扫描未来15天的脚本事件（最早提醒14天前）
 * - 若对应邮件线程已贴 “GTC_处理完了”：停止提醒，并把事件状态写成 DONE（留痕）
 */
function sendCustomReminderEmails() {
  const cal = getCalendar_();
  const now = new Date();
  const rangeEnd = new Date(now.getTime() + 15 * 24 * 60 * 60000);
  const events = cal.getEvents(now, rangeEnd);

  const me = Session.getEffectiveUser().getEmail();
  const doneLabel = GmailApp.getUserLabelByName(LABEL_DONE) || getOrCreateLabel_(LABEL_DONE);

  events.forEach(ev => {
    const desc = ev.getDescription() || '';
    if (!desc.includes(MARKER)) return;

    const start = ev.getStartTime();
    if (start.getTime() <= now.getTime()) return;

    // 1) 如果事件已标 DONE，直接跳过
    if (desc.includes('Status: DONE')) return;

    // 2) 如果对应线程被你标“处理完了”，则把事件写成 DONE 并跳过后续提醒
    const tid = extractLineValue_(desc, 'ThreadId:');
    if (tid) {
      try {
        const th = GmailApp.getThreadById(tid);
        const labels = th.getLabels();
        const isDone = labels.some(l => l.getName() === doneLabel.getName());
        if (isDone) {
          ev.setDescription(markEventDone_(desc));
          return;
        }
      } catch (e) {
        // 线程可能被删除/不可访问：忽略，继续按事件自身状态走
      }
    }

    // 3) 正常提醒逻辑：到了提醒点 → 发邮件 → 写入已发送记录，避免重复
    const sent = parseSentReminders_(desc);

    REMINDER_OFFSETS_MIN.forEach(min => {
      if (sent.has(min)) return;

      const remindAt = new Date(start.getTime() - min * 60000);
      const windowEnd = new Date(remindAt.getTime() + REMINDER_WINDOW_MIN * 60000);

      if (now.getTime() >= remindAt.getTime() && now.getTime() < windowEnd.getTime()) {
        const human = humanOffset_(min);
        const gmailThread = extractLineValue_(desc, 'GmailThread:');
        const parsedDeadline = extractLineValue_(desc, 'ParsedDeadline:');

        const mailSubj = `【期限提醒】${human}｜${ev.getTitle()}`;
        const mailBody = [
          `提醒时间点：${human}`,
          `截止事件时间：${start.toLocaleString('ja-JP')}`,
          parsedDeadline ? `解析到截止：${parsedDeadline}` : '',
          gmailThread ? `原邮件：${gmailThread}` : '',
          '',
          `如果你已经处理完：给该邮件线程加标签「${LABEL_DONE}」即可停止后续提醒（事件保留在日历里）。`
        ].filter(Boolean).join('\n');

        GmailApp.sendEmail(me, mailSubj, mailBody);

        sent.add(min);
        ev.setDescription(updateSentReminders_(desc, sent));
      }
    });
  });
}


/**
 * 每天晨报（可选）：只汇总 “ACTIVE 未完成” 的截止事件 + 未入日历的疑似截止邮件
 */
function sendMorningBriefing() {
  const cal = getCalendar_();
  const now = new Date();
  const end = new Date(now.getTime() + 14 * 24 * 60 * 60000);

  const events = cal.getEvents(now, end)
    .filter(ev => {
      const d = ev.getDescription() || '';
      return d.includes(MARKER) && !d.includes('Status: DONE');
    })
    .sort((a, b) => a.getStartTime() - b.getStartTime());

  const primaryOnly =
    'in:inbox ' +
    'label:^smartlabel_personal ' +
    '-label:^smartlabel_promo ' +
    '-label:^smartlabel_social ' +
    '-label:^smartlabel_updates ' +
    '-label:^smartlabel_forums ' +
    '-in:spam -in:trash -in:chats';

  const q = `${primaryOnly} ${buildKeywordQuery_()} -label:${LABEL_ADDED} -label:${LABEL_DONE}`;
  const threads = GmailApp.search(q, 0, 30);

  const me = Session.getEffectiveUser().getEmail();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const lines = [];
  lines.push(`【晨报】期限 & 待办（${dateStr}）`);
  lines.push(`生成时间：${now.toLocaleString('ja-JP')}`);
  lines.push('');

  // A) 未来14天截止（已进日历）
  lines.push(`A) 未来14天内未完成的截止（已进日历）：${events.length} 条`);
  if (!events.length) {
    lines.push('  - 无');
  } else {
    events.forEach((ev, i) => {
      const d = ev.getDescription() || '';
      const gmailLink = extractLineValue_(d, 'GmailThread:');
      const parsed = extractLineValue_(d, 'ParsedDeadline:');
      const fromLine = extractLineValue_(d, 'From:');

      const rawTitle = (ev.getTitle() || '').replace(/^【截止】/, '').trim();
      const snippet = sanitizeOneLine_(clip_(extractSnippetFromEventDesc_(d), 220));
      const searchHint = buildGmailSearchHint_(fromLine, rawTitle);

      lines.push(`  ${i + 1}. ${ev.getStartTime().toLocaleString('ja-JP')} | ${ev.getTitle()}`);
      if (parsed) lines.push(`     解析截止：${parsed}`);
      if (fromLine) lines.push(`     来自：${fromLine}`);
      if (snippet) lines.push(`     摘要：${snippet}`);
      lines.push(`     手机检索：${searchHint}`);
      if (gmailLink) lines.push(`     链接：${gmailLink}`); // 仍保留（iPad/PC可用）
      lines.push('');
    });
  }

  // B) Primary 里疑似截止但尚未进日历
  lines.push(`B) Primary 收件箱里疑似截止但尚未进日历：${threads.length} 条`);
  if (!threads.length) {
    lines.push('  - 无');
  } else {
    threads.forEach((th, i) => {
      const msg = th.getMessages().slice(-1)[0];
      const subject = msg.getSubject() || '(no subject)';
      const from = msg.getFrom() || '';
      const body = msg.getPlainBody() || '';

      const parsed = parseDeadline_(subject, body);
      const snippet = sanitizeOneLine_(clip_(body, 220));
      const searchHint = buildGmailSearchHint_(from, subject);

      lines.push(`  ${i + 1}. ${subject}`);
      lines.push(`     From: ${from}`);
      lines.push(`     ${parsed ? ('解析到：' + parsed.deadlineText) : '解析失败：建议正文出现明确日期时间（如 2026-01-05 18:00）'}`);
      if (snippet) lines.push(`     摘要：${snippet}`);
      lines.push(`     手机检索：${searchHint}`);
      lines.push('');
    });
  }

  GmailApp.sendEmail(me, `【晨报】期限&待办 ${dateStr}`, lines.join('\n'));
}

function extractSnippetFromEventDesc_(desc) {
  const marker = '--- Mail Snippet ---';
  const idx = (desc || '').indexOf(marker);
  if (idx === -1) return '';
  return (desc || '').slice(idx + marker.length).trim();
}

function sanitizeOneLine_(s) {
  return (s || '').replace(/\s+/g, ' ').trim();
}

// 生成一条“手机上可复制粘贴到 Gmail App 搜索框”的检索语句
function buildGmailSearchHint_(fromField, subject) {
  const emailMatch = (fromField || '').match(/<([^>]+)>/);
  const email = emailMatch ? emailMatch[1] : (fromField || '').replace(/.*\s/, '').trim();

  const key = (subject || '')
    .replace(/^【截止】/, '')
    .replace(/["']/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 18);

  // 例：from:xxx@yyy "締切のご案内"
  return email ? `from:${email} "${key}"` : `"${key}"`;
}


/* =========================
 * 解析截止日期（支持多格式）
 * ========================= */
function parseDeadline_(subject, body) {
  const text = toHalfWidth_(`${subject}\n${body}`);

  const zones = keywordNeighborhoods_(text);
  const candidates = [];
  zones.forEach(z => candidates.push(...extractDateTimes_(z)));
  if (!candidates.length) candidates.push(...extractDateTimes_(text));
  if (!candidates.length) return null;

  const now = new Date();
  candidates.sort((a, b) => a.date - b.date);

  let chosen = candidates.find(c => c.date.getTime() >= now.getTime());
  if (!chosen) chosen = candidates[candidates.length - 1];

  return { start: chosen.date, deadlineText: chosen.text };
}

function extractDateTimes_(text) {
  const out = [];
  const now = new Date();
  const curY = now.getFullYear();

  const monMap = {
    jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4,
    may:5, jun:6, june:6, jul:7, july:7, aug:8, august:8,
    sep:9, sept:9, september:9, oct:10, october:10, nov:11, november:11,
    dec:12, december:12
  };

  function normalizeYear_(yStr) {
    let y = parseInt(yStr, 10);
    if (yStr.length === 2) y = (y <= 69) ? (2000 + y) : (1900 + y);
    return y;
  }

  function timeNear_(idx) {
    return parseTime_(text.slice(idx, idx + 80));
  }

  function add_(y, m, d, t, raw) {
    const dt = new Date(y, m - 1, d, t.h, t.mi, 0);
    if (!isNaN(dt.getTime())) out.push({ date: dt, text: raw });
  }

  let m;

  const reYMD = /(\d{4})\s*(?:[\/\-\.年])\s*(\d{1,2})\s*(?:[\/\-\.月])\s*(\d{1,2})\s*(?:日)?/g;
  while ((m = reYMD.exec(text)) !== null) {
    const y = +m[1], mo = +m[2], d = +m[3];
    const t = timeNear_(m.index);
    add_(y, mo, d, t, `${y}-${pad2_(mo)}-${pad2_(d)} ${pad2_(t.h)}:${pad2_(t.mi)}`);
  }

  const reYYYYMMDD = /(?:^|[^\d])((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?:[^\d]|$)/g;
  while ((m = reYYYYMMDD.exec(text)) !== null) {
    const y = +m[1], mo = +m[2], d = +m[3];
    const t = timeNear_(m.index);
    add_(y, mo, d, t, `${y}-${pad2_(mo)}-${pad2_(d)} ${pad2_(t.h)}:${pad2_(t.mi)}`);
  }

  const reMonName1 = /\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\b\s+(\d{1,2})(?:st|nd|rd|th)?[,]?\s+(\d{2,4})/gi;
  while ((m = reMonName1.exec(text)) !== null) {
    const mo = monMap[m[1].toLowerCase()], d = +m[2], y = normalizeYear_(m[3]);
    const t = timeNear_(m.index);
    add_(y, mo, d, t, `${y}-${pad2_(mo)}-${pad2_(d)} ${pad2_(t.h)}:${pad2_(t.mi)}`);
  }

  const reMonName2 = /(\d{1,2})(?:st|nd|rd|th)?\s+\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\b[,]?\s+(\d{2,4})/gi;
  while ((m = reMonName2.exec(text)) !== null) {
    const d = +m[1], mo = monMap[m[2].toLowerCase()], y = normalizeYear_(m[3]);
    const t = timeNear_(m.index);
    add_(y, mo, d, t, `${y}-${pad2_(mo)}-${pad2_(d)} ${pad2_(t.h)}:${pad2_(t.mi)}`);
  }

  const reMD = /(\d{1,2})\s*(?:[\/\-\.月])\s*(\d{1,2})\s*(?:日)?/g;
  while ((m = reMD.exec(text)) !== null) {
    const mo = +m[1], d = +m[2];
    if (mo > 12 || d > 31) continue;

    const t = timeNear_(m.index);
    let y = curY;
    let dt = new Date(y, mo - 1, d, t.h, t.mi, 0);
    if (dt.getTime() < now.getTime() - 7 * 24 * 60 * 60000) {
      y = curY + 1;
      dt = new Date(y, mo - 1, d, t.h, t.mi, 0);
    }
    out.push({ date: dt, text: `${mo}/${d} ${pad2_(t.h)}:${pad2_(t.mi)} (assume ${y})` });
  }

  // 去重：同一分钟只留一个
  const seen = new Set();
  const dedup = [];
  out.forEach(c => {
    const key = Math.floor(c.date.getTime() / 60000);
    if (!seen.has(key)) { seen.add(key); dedup.push(c); }
  });
  return dedup;
}

function parseTime_(s) {
  s = toHalfWidth_(s || '');

  let m = s.match(/(上午|下午|晚上|中午|午前|午後)\s*(\d{1,2})(?:\s*[:：]\s*(\d{2}))?/);
  if (m) {
    let h = parseInt(m[2], 10);
    const mi = m[3] ? parseInt(m[3], 10) : 0;
    const p = m[1];
    const isPM = (p === '下午' || p === '晚上' || p === '中午' || p === '午後');
    const isAM = (p === '上午' || p === '午前');
    if (isPM && h < 12) h += 12;
    if (isAM && h === 12) h = 0;
    return { h, mi };
  }

  m = s.match(/(\d{1,2})\s*[:：]\s*(\d{2})\s*([AaPp][Mm])/);
  if (m) {
    let h = parseInt(m[1], 10);
    const mi = parseInt(m[2], 10);
    const ap = m[3].toLowerCase();
    if (ap === 'pm' && h < 12) h += 12;
    if (ap === 'am' && h === 12) h = 0;
    return { h, mi };
  }

  m = s.match(/(\d{1,2})\s*[:：]\s*(\d{2})/);
  if (m) return { h: parseInt(m[1], 10), mi: parseInt(m[2], 10) };

  m = s.match(/(\d{1,2})\s*(?:時|点)\s*(\d{1,2})?\s*(?:分)?/);
  if (m) return { h: parseInt(m[1], 10), mi: m[2] ? parseInt(m[2], 10) : 0 };

  return { h: 9, mi: 0 };
}


/* =========================
 * 查询/关键词
 * ========================= */
function buildKeywordQuery_() {
  const terms = KEYWORDS.map(k => `"${k}"`).join(' OR ');
  return `(${terms})`;
}

function keywordNeighborhoods_(text) {
  const zones = [];
  const lower = text.toLowerCase();
  KEYWORDS.forEach(k => {
    const kk = k.toLowerCase();
    let idx = 0;
    while ((idx = lower.indexOf(kk, idx)) !== -1) {
      const start = Math.max(0, idx - 80);
      const end = Math.min(text.length, idx + 260);
      zones.push(text.slice(start, end));
      idx += kk.length;
    }
  });
  return zones.length ? zones : [text];
}


/* =========================
 * 已发送提醒记录（写在事件描述里）
 * ========================= */
function parseSentReminders_(desc) {
  const set = new Set();
  const m = desc.match(/SentRemindersMin:\s*([0-9,\s]*)/);
  if (!m) return set;
  const raw = (m[1] || '').trim();
  if (!raw) return set;
  raw.split(',').map(x => x.trim()).filter(Boolean).forEach(x => set.add(parseInt(x, 10)));
  return set;
}

function updateSentReminders_(desc, sentSet) {
  const arr = Array.from(sentSet).sort((a, b) => b - a);
  const line = `SentRemindersMin: ${arr.join(',')}`;

  if (desc.match(/SentRemindersMin:\s*[0-9,\s]*/)) {
    return desc.replace(/SentRemindersMin:\s*[0-9,\s]*/, line);
  }
  return `${desc}\n${line}`;
}


/* =========================
 * DONE 标记（写在事件描述里，留痕）
 * ========================= */
function markEventDone_(desc) {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  let out = desc;

  if (out.includes('Status: ACTIVE')) {
    out = out.replace('Status: ACTIVE', `Status: DONE (${ts})`);
  } else if (!out.includes('Status: DONE')) {
    out = `Status: DONE (${ts})\n` + out;
  }

  return out;
}


/* =========================
 * 工具
 * ========================= */
function getCalendar_() {
  if (CALENDAR_ID && CALENDAR_ID.trim()) {
    return CalendarApp.getCalendarById(CALENDAR_ID.trim());
  }
  return CalendarApp.getDefaultCalendar();
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function extractLineValue_(desc, prefix) {
  const lines = (desc || '').split('\n');
  const line = lines.find(l => l.startsWith(prefix));
  if (!line) return '';
  return line.slice(prefix.length).trim();
}

function humanOffset_(min) {
  if (min === 12 * 60) return '12小时前';
  return `${Math.round(min / (24 * 60))}天前`;
}

function clip_(s, n) {
  if (!s) return '';
  s = s.replace(/\r/g, '');
  return s.length > n ? s.slice(0, n) + '\n... [clipped]' : s;
}

function pad2_(x) {
  return (x < 10 ? '0' : '') + x;
}

function toHalfWidth_(s) {
  return (s || '')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[：]/g, ':')
    .replace(/[／]/g, '/')
    .replace(/[－]/g, '-');
}
