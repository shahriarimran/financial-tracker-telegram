# Troubleshooting

- **Workbook not found:** keep `scripts/RunMacro.vbs` two levels below the workbook, as distributed.
- **Task does not run:** verify the task with `Get-ScheduledTask -TaskName FinancialTrackerDailyUpdate` and check Task Scheduler history.
- **Macros blocked:** review the exported VBA and use a trusted location or your organization’s approved macro policy.
- **No Telegram message:** verify `config/config.ini`, network access, Power Query refresh, and the workbook’s signal output.
- **Excel remains open:** close the workbook, end only the Excel instance started for this task after confirming it is not a user session, then retry manually.
