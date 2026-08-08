# Installation

1. Copy `config/config.example.ini` to `config/config.ini` and enter your own Telegram values.
2. Keep the workbook and `scripts` folder together in the project root.
3. Open the workbook once in desktop Excel, review the VBA export, and enable macros only if trusted.
4. Run `powershell -ExecutionPolicy Bypass -File .\scripts\Install-Task.ps1` from the project root.

The installer registers only `FinancialTrackerDailyUpdate`. It starts `wscript.exe` with the portable `scripts\RunMacro.vbs` path at 16:15 Monday–Friday, ignores concurrent launches, and stops after 15 minutes.
