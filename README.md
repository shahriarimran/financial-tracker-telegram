# Financial Tracker Telegram Alerts

A Windows desktop-Excel tracker that refreshes market data, calculates Gold, Silver, and USD signals, records the Daily Log, and sends Telegram notifications from workbook VBA.

> Educational and informational use only. This project is not investment advice, a solicitation, or a recommendation to buy or sell any asset. Verify all data independently and make decisions appropriate to your circumstances.

## Requirements

- Windows 10/11
- Microsoft Excel desktop with macros and Power Query enabled
- Internet access for workbook data sources and Telegram
- A Telegram bot token and destination chat ID

## Setup

1. Copy `config/config.example.ini` to `config/config.ini`.
2. Add your own Telegram token and chat ID. Never commit `config/config.ini`.
3. Open `Financial_Trackers_AutoRun.xlsm`, enable macros only after reviewing `src/vba/`, and confirm that `DailyAutoRun` can refresh and notify.
4. Run `powershell -ExecutionPolicy Bypass -File .\scripts\Install-Task.ps1` to schedule weekday automation.

The scheduled flow is Windows Task Scheduler → `scripts/RunMacro.vbs` → Excel/VBA → Telegram. It runs Monday–Friday at 16:15 local Windows time, ignores overlapping launches, and has a 15-minute time limit.

## Manual testing

Open the workbook and run `DailyAutoRun` from Excel. Confirm Power Query refreshes, the Daily Log is updated as expected, and a test is directed only to your own Telegram destination. Do not test with credentials from another environment.

## Signals and Daily Log

The workbook owns the signal logic. `BUY`, `SELL`, and `HOLD` are outputs of its formulas; this repository does not alter the investment methodology. Daily Log keeps the workbook’s history structure. The public release contains a blank template rather than private history.

## Security and removal

Use a trusted local workbook location and review macro source before enabling macros. If your Office policy blocks VBA project access, use your organization’s approved process. To remove the scheduled task, run `powershell -ExecutionPolicy Bypass -File .\scripts\Uninstall-Task.ps1`.

## Limitations

Excel automation requires an interactive desktop Office installation. Data availability, Power Query sources, Telegram availability, and macro security policies can affect runs. No cloud service, Python runtime, or automatic trading is included. 
