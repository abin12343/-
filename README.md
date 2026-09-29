# MyScripts 自动化脚本仓库

按功能分类管理自动化任务脚本，方便查找、维护和复用。

## 目录结构

| 目录 | 功能 | 说明 |
|---|---|---|
| `01-file-management` | 文件管理 | 文件整理、批量重命名、备份、同步、清理 |
| `02-data-processing` | 数据处理 | Excel/CSV 处理、数据库、报表生成 |
| `03-network-scraping` | 网络与爬虫 | 网页抓取、API 调用、数据采集 |
| `04-system-ops` | 系统与运维 | 系统监控、日志分析、磁盘清理、服务管理 |
| `05-notification` | 消息通知 | 邮件、钉钉/企业微信、Telegram、Webhook 推送 |
| `06-scheduling` | 定时调度 | 定时任务封装、计划任务/Cron 配置 |
| `07-testing` | 测试与验证 | 自动化测试、接口验证、回归检查 |
| `08-common-utils` | 通用工具 | 公共函数、复用模块、常量 |
| `config` | 配置 | 全局配置（私密配置不提交） |
| `logs` | 日志 | 脚本运行日志（不提交） |
| `99-archive` | 归档 | 停用或不再维护的脚本 |

## 使用约定

- 新增脚本先放进对应的功能分类目录，不要堆在仓库根目录。
- 需要多脚本复用的代码放入 `08-common-utils`。
- 私密配置（API Key、密码、Token）只放本地，不提交到仓库。
- 运行日志统一输出到 `logs` 目录。

## 添加新脚本的步骤

1. 在对应功能目录下新建脚本文件（参考该目录的 README）。
2. 脚本头部写清：用途、依赖、运行方式。
3. 有配置需求时，参照 `config/config.example.json` 创建本地配置。
4. 涉及删除、覆盖等危险操作的脚本，先进入试运行（dry-run）模式或由用户确认。
5. **新脚本直接 `from autolib import ...`**（08-common-utils/autolib），不要复制
   yongda-bill-check 那批私有实现；通用模式就在那里。

## 现有业务模块

| 目录 | 说明 |
|---|---|
| `02-data-processing/yongda-bill-check/` | 永达 UPS 账单对账（首套业务，3415 行私有实现） |
| `02-data-processing/skye-bill-check/`   | SKYE 账单对账（按 SOP 自动化，复用 autolib + 永达的天图引擎） |
