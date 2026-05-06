# Excel Watchdog — Setup Guide

## 1 — Install Python dependencies

Open a Command Prompt as Administrator on the Windows Server:

```
pip install psutil pywin32
```

Verify it works:

```
python excel_watchdog.py --once
```

You should see log output. If you see a warning about pywin32, the install
didn't complete — run `pip install pywin32` again and restart the terminal.

---

## 2 — Set up the Microsoft Teams Incoming Webhook

There are two ways depending on your Teams version.

---

### Option A — Incoming Webhook connector (classic, most common)

1. Open **Microsoft Teams** and go to the **channel** where you want alerts
   (create a dedicated one like `#excel-alerts` if you like).

2. Click the **"..."** (More options) next to the channel name → **Connectors**.

3. Search for **"Incoming Webhook"** → click **Configure**.

4. Give it a name (e.g. `Excel Watchdog`) and optionally upload an icon.

5. Click **Create** → Teams generates a URL. **Copy it.**

6. Paste the URL into `excel_watchdog.py`:

   ```python
   "teams_webhook_url": "https://yourorg.webhook.office.com/webhookb2/...",
   ```

7. Test it immediately:

   ```
   python excel_watchdog.py --once
   ```

   If Excel is running and no dialog is stuck, nothing will post — that's correct.
   To force a test post, temporarily lower `max_runtime_minutes` to 1.

> ⚠️ **Note:** Microsoft is gradually phasing out classic connectors in favour
> of Power Automate Workflows. If you don't see the Connectors option, use
> Option B below.

---

### Option B — Power Automate Workflow webhook (newer Teams)

Microsoft is replacing connectors with Power Automate flows. If your tenant has
disabled connectors, follow these steps:

1. In the Teams channel, click **"..."** → **Workflows**.

2. Search for **"Post to a channel when a webhook request is received"**
   → click **Add workflow**.

3. Give it a name → click **Next** → select the team and channel → **Add workflow**.

4. Click **Copy** on the generated webhook URL.

5. Paste the URL into `excel_watchdog.py` exactly as in Option A.

> The script sends an Adaptive Card payload — both connector and workflow URLs
> accept the same format, so no code change is needed.

---

## 3 — Schedule in Windows Task Scheduler

### Step-by-step (GUI)

1. Open **Task Scheduler** (`taskschd.msc`).

2. In the right panel click **"Create Task"** (not "Create Basic Task" —
   you need the full options).

3. **General tab**
   - Name: `Excel Watchdog`
   - Description: `Monitors Excel for stuck processes and notifies Teams`
   - Select **"Run whether user is logged on or not"**
   - Check **"Run with highest privileges"**
   - Configure for: `Windows Server 2016` (or your actual version)

4. **Triggers tab** → New
   - Begin the task: `At startup`
   - Delay task for: `2 minutes` (gives Windows time to start services)
   - ✅ Enabled

5. **Actions tab** → New
   - Action: `Start a program`
   - Program/script:
     ```
     C:\Python312\python.exe
     ```
     *(adjust to your actual Python path — find it with `where python` in CMD)*
   - Add arguments:
     ```
     C:\ExcelWatchdog\excel_watchdog.py
     ```
   - Start in (optional but recommended):
     ```
     C:\ExcelWatchdog\
     ```

6. **Conditions tab**
   - Uncheck **"Start the task only if the computer is on AC power"**
     (important for servers)

7. **Settings tab**
   - ✅ "If the task fails, restart every" → `1 minute`, up to `3 times`
   - ✅ "If the running task does not end when requested, force it to stop"
   - Run task as soon as possible after a scheduled start is missed: ✅

8. Click **OK** — enter the service account password when prompted.

---

### Alternative: create the task from the command line

Save this as `register_task.bat` and run as Administrator:

```bat
schtasks /create ^
  /tn "Excel Watchdog" ^
  /tr "\"C:\Python312\python.exe\" \"C:\ExcelWatchdog\excel_watchdog.py\"" ^
  /sc ONSTART ^
  /delay 0002:00 ^
  /ru SYSTEM ^
  /rl HIGHEST ^
  /f
```

Using `SYSTEM` avoids password prompts but means the script runs as the local
system account — make sure `C:\ExcelWatchdog\` is writable by SYSTEM.

---

## 4 — Folder layout

After setup you should have:

```
C:\ExcelWatchdog\
    excel_watchdog.py    ← the script
    watchdog.log         ← created automatically on first run
    watchdog.db          ← SQLite database, created automatically
```

---

## 5 — Verify everything is working

```
# Check the log
type C:\ExcelWatchdog\watchdog.log

# Check the notification history database
python C:\ExcelWatchdog\excel_watchdog.py --history

# Run a single manual scan
python C:\ExcelWatchdog\excel_watchdog.py --once
```

To test the Teams webhook without waiting for a real stuck Excel, temporarily
set `max_runtime_minutes` to `0` and open any Excel file — the script will
flag it as timed out and post to Teams. Remember to revert the value after
testing.

---

## 6 — Tuning tips

| Setting | Default | When to change |
|---|---|---|
| `max_runtime_minutes` | 30 | Raise if your macros legitimately take a long time |
| `check_interval_seconds` | 20 | Lower (e.g. 10) if you want faster detection |
| `ERROR_DIALOG_PATTERNS` | see script | Add your specific VBA error message text |

To add a custom pattern — for example if your macro pops a dialog titled
"Data refresh failed" — add `"data refresh failed"` (lowercase) to the
`ERROR_DIALOG_PATTERNS` list in the script.
