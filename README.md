# NSE stock OHLC reports and CPR

Source backup of the daily and weekly NSE F&O-stock reporting system, including
the CPR columns and classifications. The running installation is separate from
this checkout; publishing this repository does not modify its scheduled tasks.

## Reports

- **Weekly:** cash-market OHLC for the union of stock underlyings in the selected
  week's NSE stock futures/options files. Runs Saturdays at 09:00 IST.
- **Daily:** each session's cash-market OHLC for that day's F&O stock universe.
  Checks every 15 minutes from 16:30 through 22:00 IST, with 09:00 and sign-in
  catch-up. One workbook per week, with a sheet per trading day.
- **Monthly:** completed calendar-month OHLC, produced at 09:00 IST on the first
  of each month, including when that date is a weekend or holiday. A separate
  monthly workbook includes Monthly OHLC, Daily Prices, Sources and Method, and
  CPR Guide. Uses the weekly six-category CPR bands and TradingView symbols.
- Both include Pivot, normalized TC/BC, nonnegative CPR width, width percentage,
  colour-coded classification, filters and a CPR Guide sheet.
- The final table column is **TradingView Symbol**, using `=A2&","` to produce
  text such as `ASTRAL,` without an exchange prefix. It is included in the table,
  so filtering and sorting keep it aligned with the stock symbol. Copy visible
  cells from this column after filtering to prepare a comma-separated list.

Prices are unadjusted NSE cash EQ/F1 prices in INR, not futures-contract prices.
NSE's official final CM and FO UDiFF archives are validated before a date is
processed. Holidays use NSE's official CM/FO calendar; a failed price download is
never silently treated as a holiday. Pending dates/weeks are retained for retry.

| CPR classification | Daily width % | Weekly width % |
| --- | --- | --- |
| Very Narrow | 0.000–0.050 | 0.000–0.050 |
| Narrow | 0.051–0.100 | 0.051–0.100 |
| Medium | 0.101–0.200 | 0.101–0.200 |
| Medium 1 | — | 0.201–0.500 |
| Medium 2 | — | 0.501–0.800 |
| Above Medium | 0.201 and above | 0.801 and above |

Column J stores percentage points: `0.05` means `0.05%`. Classification uses the
three-decimal rounded width, matching the three-decimal display in column J;
the underlying CPR calculations retain full precision. All table endpoints are
inclusive. TC is always the higher boundary. These categories describe width;
they do not establish breakout probabilities.

## Requirements and restore

This is a backup of a Windows installation that uses **Python 3.10+**, **Node.js**
and the **Codex bundled `@oai/artifact-tool` package** for Excel authoring. The
Python programs use the standard library; the report builders additionally need
that bundled Node package. Python integration tests use `openpyxl` for reading
and verification. Dependencies and market-data downloads are not committed.

1. Clone or copy the source into your chosen local report folder, outside
   OneDrive if local-only storage is desired. For the existing PC, the running
   folder is the local Desktop's `NSE Weekly OHLC` directory.
2. Copy `daily_config.example.json` to `daily_config.json`, and
   `schedule_config.example.json` to `schedule_config.json`. Existing local
   configuration files from the running installation can be reused instead.
3. Set `node_executable` to the installed bundled Node executable. Ensure
   `daily_runtime/node_modules` resolves to the matching bundled Node packages.
   On Windows, create a junction with PowerShell, using your actual package path:

   ```powershell
   New-Item -ItemType Junction -Path .\daily_runtime\node_modules -Target 'C:\path\to\bundled\node_modules'
   ```

4. Review the starting dates. `daily_config.json`'s `first_date` controls daily
   backfill. `schedule_config.json`'s `first_run` must be a Saturday at 09:00 IST
   and controls the first weekly report. Retain the example dates to recover the
   original November 2026 start, or choose the intended new start dates.
5. Test the programs manually with the full path to your Python executable:

   ```powershell
   python update_weekly_ohlc.py --auto-holidays
   python daily_update.py
   python scheduled_update.py --check-only
   ```

   On this PC, `python` may not be on PATH; use its full executable path instead.
6. Only when intentionally installing tasks, run the installers below. They
   refuse to overwrite an existing task of the same name:

   ```powershell
   .\install_schedule.ps1 -ReportFolder $PWD.Path -PythonExecutable 'C:\path\to\pythonw.exe'
   .\install_daily_schedule.ps1 -ReportFolder $PWD.Path -PythonExecutable 'C:\path\to\pythonw.exe'
   ```

   Both require Windows timezone **India Standard Time**. The weekly installer
   reads the first scheduled date from `schedule_config.json`. Daily trigger times
   are defined in `install_daily_schedule.ps1`; the config's polling fields are
   descriptive and do not automatically re-register Windows task triggers.

Once tasks are installed, do not move the installation folder without updating
their actions. The PC must be on and the configured user signed in. Wake timers
are requested, subject to hardware/Windows support. Missed starts and sign-in
checks catch up; failures retry three times at 15-minute intervals. Codex and
Excel do not need to remain open, but the bundled runtime must remain installed.

Use the stable `%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies`
runtime paths for unattended jobs when available. Version-numbered paths under
`Program Files\WindowsApps` can disappear when the Codex runtime updates.
If the daily task reports `0x80070002` (file not found), `repair_runtime.ps1`
updates its Python action, the shared Node config and the workbook-library
junction to the stable cache. It preserves daily scheduling settings and leaves
the existing weekly task unchanged. Run it from the installation directory, or
pass `-ReportFolder` explicitly; existing config/task definitions are backed up.

## Files

| File | Purpose |
| --- | --- |
| `update_weekly_ohlc.py` | Official downloads, holidays, weekly aggregation/export |
| `scheduled_update.py` | Weekly pending-run tracking and catch-up |
| `daily_update.py` | Daily collection, validation, per-week workbook updates |
| `daily_runtime/build_weekly.mjs` | Weekly workbook builder |
| `daily_runtime/build_daily.mjs` | Daily workbook builder |
| `daily_runtime/cpr.mjs` | CPR formulas, conditional formatting, guide and checks |
| `install_schedule.ps1` | Weekly Windows Task Scheduler installation |
| `install_daily_schedule.ps1` | Independent daily task installation |
| `Update Excel.cmd` | Manual weekly updater launcher |

The TXT guides describe the original installed schedules, report layout and
local file locations. Installers in this repository accept explicit paths for
restoration; the running Desktop installation has not been reconfigured.

## Tests

```powershell
node daily_runtime/test_cpr.mjs
node daily_runtime/test_tradingview.mjs
python test_scheduler.py
python test_daily.py
python test_daily_export.py
```

CPR checks cover reversed boundaries, zero/flat prices and exact classification
cutoffs. Python tests also cover catch-up, holidays, week rollover, duplicate
prevention, locked workbooks and full source-price reconciliation.

The historical integration tests expect the validated 2–6 and 9 November 2026
archives, the 2026 holiday calendar, and the original weekly JSON payload. These
fixtures are intentionally not committed. Generate the 2–6 November cache/JSON
using `update_weekly_ohlc.py --as-of 2026-11-10 --auto-holidays --json-only`;
running the daily updater with the original start date also retrieves 9 November
data when available. Offline tests need those files already cached.

## Local outputs and state

Reports are written to `outputs/YYYY-MM/`. Monthly workbooks use names such as
`Monthly Oct 2026.xlsx`, and weekly workbooks use `Weekly Nov 9-13 2026.xlsx`.
Daily workbook names remain unchanged. Weekly reports and their daily-workbook
counterparts belong to the month of their Monday: for example, `Weekly Nov
30-Dec 4 2026.xlsx` is stored in `outputs/2026-11/`. Cross-year weeks include both
years, e.g. `Weekly Dec 28 2026-Jan 1 2027.xlsx`. Monthly files belong to their
report month, and the year in names/folders prevents annual collisions.

Run `migrate_output_folders.py <installation-folder> --apply` once to move legacy
root-level files without rewriting their contents. It verifies file hashes,
updates completion receipts and unsent email locations, and leaves sent-email
records unchanged. Existing destinations are never overwritten. A move manifest
and previous state files are saved under `backups/`.

Original archives, calendars, logs, completion
receipts and machine-specific configs stay local and are excluded by `.gitignore`.
Generated daily workbooks are regenerated from their validated daily records;
keep personal edits in a separate analysis copy. Close workbooks during updates;
an Excel lock preserves the existing file and leaves the update pending.

## Monthly reports

Monthly reporting is independent: copy `monthly_config.example.json` to
`monthly_config.json`, review its `first_month` and special-session overrides,
then run `monthly_update.py`. Install with `install_monthly_schedule.ps1`
using `-ReportFolder` if needed. The first-of-month 09:00 task and sign-in
catch-up track each pending calendar month separately. `--month YYYY-MM`
generates a selected completed month; `--fetch-only` only validates/caches data.
Monthly notifications use the same secure sender and outbox as the other jobs.
Monthly aggregation consolidates symbol renames only when the official cash
files confirm the same ISIN. It uses the latest cash symbol for the stock row
and retains the original daily cash symbols and mapping notes for traceability.
The known 8 November 2026 Muhurat session is explicitly configured; add future
announced weekend sessions to `additional_sessions`. See `MONTHLY SCHEDULE.txt`.

## Email notifications

`email_notifications.py` sends short plain-text messages after a weekly report is
created or a new trading-day sheet is saved in a daily workbook. No attachment or
report contents are sent. Duplicate checks use the report week/trading date.
Email delivery failures keep the successful report intact and leave a durable
outbox event for retry on a later scheduled run. SMTP timeouts or a crash between
server acceptance and receipt recording can rarely cause a repeated message.

Copy `email_config.example.json` to `email_config.json`, set sender and recipient,
then run `setup_email.py` using the same signed-in Windows account as the tasks.
It presents a masked local entry for a Gmail App Password, checks Gmail, sends a
setup test message, and stores the password in Windows Credential Manager.
Never paste credentials into chat or put them in source/config files. Gmail App
Passwords require 2-Step Verification and may be unavailable for some account
policies: https://support.google.com/mail/answer/185833

Local `email_config.json`, setup status and `email_outbox/` are ignored by Git.
Notification code does not change task schedules or prevent a successful report
from being completed when email is unavailable. Run `python test_email_notifications.py`
for notification retry, duplicate suppression, secure SMTP and no-attachment checks.
