# Architecture

```text
Windows Task Scheduler
  -> scripts/RunMacro.vbs
    -> Financial_Trackers_AutoRun.xlsm / DailyAutoRun
      -> Power Query refresh + formulas + Daily Log
      -> Telegram API
```

`RunMacro.vbs` finds the workbook relative to the project directory. Local secrets belong only in `config/config.ini`, which is ignored by Git. The legacy Python bot is not production runtime.
