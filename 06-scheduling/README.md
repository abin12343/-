# 定时调度

用途：定时任务封装、调度配置管理。

适合放在这里的脚本：
- 定时任务的入口封装（Python `schedule`、APScheduler 等）
- Windows 计划任务 / Linux Cron 的配置说明与注册脚本
- 任务依赖、重试、超时处理

命名建议：任务名 + 调度方式，例如 `run-daily-job.py`、`register-task.bat`。
