# Gmail → Calendar Deadline Assistant (Google Apps Script)

A personal productivity Google Apps Script that:

- Scans **Gmail Primary (Inbox)** for emails containing deadline keywords
- Parses a deadline date/time from subject/body
- Creates **ONE** Google Calendar event (deadline log / record)
- Sends **email reminders** at multiple offsets (no extra calendar events)
- Sends a **daily morning briefing** (digest)
- Lets you stop future reminders by applying a **Done label** in Gmail while keeping the calendar event for record

> This runs entirely in your Google account via Apps Script (no external server required).

---

## What it does

### 1) Primary Inbox → Single Calendar Deadline Event
- Only searches Gmail **Primary tab** (`^smartlabel_personal`)
- Requires keyword match (e.g. `締切`, `締め切り`, `deadline`, etc.)
- If a date/time is parsed, creates **one** event:
  - Title: `【截止】<email subject>`
  - Description includes a marker, Gmail thread info, and optional snippet

### 2) Email Reminders (No extra calendar events)
Reminders are sent by the script via Gmail (NOT Calendar notifications), at:
- 14 days, 10 days, 7 days, 5 days, 3 days, 1 day, 12 hours before the deadline

✅ The script creates **only one calendar event** per deadline email.  
It does **not** create multiple events for reminders.

### 3) Done label stops reminders (calendar record stays)
Apply Gmail label:
- `GTC_处理完了`

Then the script:
- Stops future reminder emails for that deadline
- Keeps the Calendar event
- Writes `Status: DONE (timestamp)` into the event description

### 4) Daily morning briefing (1 email/day)
Every morning (default 08:00), sends one digest email:
- **A)** Upcoming deadlines (next 14 days) still ACTIVE
- **B)** “Deadline-like” emails in Primary not yet added to calendar (parse failed / missing date)

Includes short excerpt/snippet and a “mobile search hint” so you can find the email even if links don’t open well on phone.

### 5) Old mail protection (Cutoff)
To avoid importing old Primary emails, a script property `INGEST_CUTOFF_MS` is used.
You can set it to “now” so only new mail is processed.

---

## Gmail labels used

The script creates/uses these labels:

- `GTC_Added`  
  Added automatically after a calendar event is created (prevents duplicates)

- `GTC_处理完了`  
  You add this when finished → reminders stop, calendar stays

- `GTC_旧邮件忽略`  
  Used for ignoring old emails when cutoff logic is enabled

---

## Privacy & data handling (important)

This script runs inside your Google account and uses Gmail + Calendar APIs.

It may store the following **inside your own Calendar event description**:
- `From:` sender info
- Gmail thread ID + a Gmail link
- A snippet of email body (up to ~800 chars by default)
- Sent reminder offsets (to avoid duplicates)

The daily briefing email may include short excerpts/snippets of emails.

> No API keys are required and no external server is used by default.

If you want maximum privacy, disable snippets in the code (see Customization).

---

## Installation (copy & paste .gs)

1) Create a new Apps Script project  
   https://script.new

2) Create a file `GmailToCalendar.gs` and paste the code

3) Save, then run the function:
- `installTriggers()`

4) Grant permissions when prompted (Gmail, Calendar, Triggers)

After that, it runs automatically via triggers.

---

## First-time recommended step (avoid importing old emails)

If your Primary inbox contains many old emails with `締切/期限/deadline` etc., run:
- `setIngestCutoffToNow()`

This sets the cutoff to the current time so only new emails are processed.

Tip: If you enable/adjust cutoff logic later, label old threads with `GTC_旧邮件忽略` to prevent re-scans.

---

## How to use (daily workflow)

1) Receive email in Gmail Primary tab
2) If it contains a parseable deadline, a Calendar event is created automatically
3) You receive reminder emails at configured offsets
4) When finished, apply label:
- `GTC_处理完了`

Reminders stop for that deadline, and the Calendar event remains as a record.

---

## Triggers created

Running `installTriggers()` creates 3 time-based triggers:

- `ingestPrimaryInboxToCalendar` → every 10 minutes  
- `sendCustomReminderEmails` → every 1 hour  
- `sendMorningBriefing` → every day at 08:00

Note: `installTriggers()` deletes existing triggers in the same Apps Script project before recreating them.

---

## Customization

### Change reminder offsets
Edit:
```js
const REMINDER_OFFSETS_MIN = [ ... ];
