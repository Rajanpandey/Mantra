# ✦ Mantra

> Words that live with you. Delivered daily, at random times, to catch you off guard.

A personal system that emails you your own mantras throughout the day — at precise random times, equally distributed, grouped by themes you control.

No app to install. No server. Runs entirely on Google Sheets + Gmail, free forever.

---

## Setup

**1. Copy the sheet**

[Click here to copy the sheet to your Google Drive](https://docs.google.com/spreadsheets/d/10D7XjH7RXL48YH2OfRjrnxfuccVwRIymLCtGR3TKxxk/edit?usp=sharing)

Go to **File → Make a copy**. Save it to your Drive. The file name can be anything.

---

**2. Add the script**

In your copied sheet → **Extensions → Apps Script**

- Delete any existing code in the editor
- Paste the entire contents of `Code.gs` from this repo
- Save (`Cmd+S` or `Ctrl+S`)

---

**3. Add the manifest file**

In Apps Script → click the **gear icon** (Project Settings) → check **"Show appsscript.json manifest file in editor"**

Go back to the editor — you'll see an `appsscript.json` tab. Click it, replace everything inside with this:

```json
{
  "timeZone": "America/Los_Angeles",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/script.triggers",
    "https://www.googleapis.com/auth/script.scriptapp"
  ]
}
```

Save. This tells Google exactly what permissions the script needs so authorization works cleanly.

---

**4. Set your email**

Go to the **Config** sheet tab → find **Your Email Address** → replace `your@gmail.com` with your actual Gmail.

---

**5. Authorize**

Back in Apps Script → click **Run** → select `scheduleDailyMantras` → click **Run** again.

Google will ask for permissions. Click **Advanced → Go to Untitled project → Allow**.

This is your own script running on your own account — it's safe to authorize.

---

**6. Set the daily trigger**

In Apps Script → click the **clock icon** (Triggers) in the left sidebar → **+ Add Trigger**

| Setting | Value |
|---|---|
| Function | `scheduleDailyMantras` |
| Event source | Time-driven |
| Type | Day timer |
| Time | Midnight to 1am |

Save. That's it. Every night at midnight the script picks your active mantras, assigns each one a random exact minute within its delivery window, and schedules the emails. They arrive throughout the day.

---

## Adding your mantras

Open the **Mantras** sheet and fill in the rows. Delete the sample rows first.

| Column | What to put |
|---|---|
| **Mantra Text** | Your mantra, any length |
| **Theme** | A label — e.g. `General`, `Dates`, `Morning`, `Work` |
| **Active** | `YES` or `NO` |
| **Window Start / End** | Random delivery within this range — e.g. `09:00` to `20:00` |
| **Fixed Time** | Exact delivery time — e.g. `07:00`. Leave window blank if using this. |
| **Days of Week** | `Mon,Tue,Wed,Thu,Fri,Sat,Sun` — any order, any subset |
| **Pause Until** | `YYYY-MM-DD` — mantra is skipped until this date passes |
| **Notes** | Just for you — ignored by the script |

Delivery Count and Last Sent are updated automatically. Don't edit them.

---

## Themes

Themes are just labels you assign to mantras. Their power comes from the **Config sheet**.

Open **Config → Per-Theme Daily Limits** and set how many mantras from each theme you want per day.

```
General   →  1
Morning   →  1
Work      →  1
Dates     →  2
```

To pause an entire theme — change `Active` to `NO` on all mantras in that group. To enable it (say, before a date) — change them back to `YES`. Use **Ctrl+H (Find & Replace)** to do this in one step.

---

## Delivery windows vs fixed times

**Window** — the script picks a random minute within the range. Good for mantras you want to arrive unexpectedly.

```
Window Start: 09:00
Window End:   20:00
→ Delivers at e.g. 11:47am, or 3:23pm, or 6:08pm — different every day
```

**Fixed time** — delivers at exactly this time, every active day. Good for pre-event or morning mantras.

```
Fixed Time: 07:00
→ Always delivers at 7:00am
```

Leave the other fields blank depending on which you use.

---

## Pause Until

Set a date in `YYYY-MM-DD` format to temporarily pause a mantra until that date.

```
Pause Until: 2025-06-01
→ Mantra is skipped until June 1st, then automatically reactivates
```

Useful for seasonal mantras or ones you've overused.

---

## Testing

**Test email sending** (instant)
In Apps Script → select `testSendNow` → Run.
An email arrives in your inbox within seconds. Confirms your email config works.

**Test scheduled delivery** (2 minutes)
In Apps Script → select `testScheduledDelivery` → Run.
Picks one active mantra, schedules it 2 minutes from now, sends it via the same trigger system the real scheduler uses. If this works, the whole system works.

**Check what's pending today**
In Apps Script → select `debugPending` → Run.
Shows all mantras scheduled for delivery today, their trigger keys, and confirms your email is being read correctly.

---

## Year in review

Every December 1st, the script emails you an annual review covering:

- Total mantras delivered that year
- Most delivered mantra
- Longest streak
- Peak delivery time of day
- Theme breakdown
- Month-by-month chart

**To set this up** — run `setupYearlyReviewTrigger` once in Apps Script. That's it, it runs every year automatically.

**To get your current year's review right now** — run `sendYearReviewNow` in Apps Script. Good for checking what the email looks like, or reviewing mid-year.

---

## Delivery log

Every sent mantra is recorded in the **Delivery Log** sheet automatically — date, time, mantra text, theme. The Analytics sheet pulls from this to show you trends over time.

---

## Troubleshooting

**Script ran but no emails arrived**
Check the **Triggers** page (clock icon in Apps Script). You should see `sendPendingMantras` triggers listed — one per scheduled mantra. If they're there, the emails will come. If not, your mantra windows may have already passed for the day (windows are checked at midnight).

**Duplicate scheduling guard**
If you run `scheduleDailyMantras` manually more than once in the same day, the second run will skip silently. To reset — go to Apps Script → Project Settings → Script Properties → delete the `scheduled_date` property.

**Mantras after blank rows not being delivered**
The script scans the entire sheet and skips blank rows — all mantras will be picked up regardless of gaps.

---

## How it works

```
Midnight trigger fires
  → reads Mantras sheet
  → shuffles active mantras
  → applies per-theme daily limits
  → resolves delivery time (random minute within window, or fixed)
  → creates one Apps Script trigger per mantra at exact delivery time
  → stores mantra data in Script Properties keyed by trigger ID

Trigger fires at delivery time
  → looks up mantra by trigger ID
  → sends email via GmailApp
  → logs to Delivery Log sheet
  → updates Delivery Count + Last Sent in Mantras sheet
  → cleans up trigger
```

---

*Built with Google Apps Script. No servers. No subscriptions. Just your words, finding you.*
