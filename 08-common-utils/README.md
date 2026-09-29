# 通用工具

用途：公共函数、复用模块、常量，供各分类脚本调用。

## 目录

```
08-common-utils/
├── autolib/                     公共库，新脚本直接 import
│   ├── config.py    配置加载（config.json + config.local.json 覆盖 + 环境变量）
│   ├── log.py       日志（统一写 logs/）+ 计时器
│   ├── paths.py     输出目录、带时间戳文件名、取最新文件、备份路径
│   ├── dialog.py    弹窗选文件/目录（独立 Tk 窗口，用完即销毁）
│   ├── notify.py    企业微信机器人 + 金山文档运行记录（标准库实现）
│   ├── browser.py   Playwright 辅助（复用本机 Edge，延迟导入）
│   └── excel.py     openpyxl 辅助（表头定位、列映射、去重追加、安全保存）
└── templates/                   可直接复制改写的脚本骨架
    ├── config.example.json
    ├── template_excel_batch.py      Excel / 表格批处理
    ├── template_web_automation.py   网页自动化（Playwright）
    └── template_file_batch.py       文件批量处理（整理/重命名/移动）
```

## 快速开始

**第一步，复制模板：**

```bash
cp templates/template_excel_batch.py ../02-data-processing/my-report.py
cp templates/config.example.json     ../02-data-processing/config.json
```

**第二步，改业务逻辑**：模板里只有一两个函数需要你改，其余骨架不用动。

**第三步，先 dry-run 再实跑：**

```bash
python my-report.py --input 数据.xlsx --dry-run
python my-report.py --input 数据.xlsx
```

## 在已有脚本里引入 autolib

```python
import sys
sys.path.insert(0, r"D:/MyScripts/08-common-utils")

from autolib import config, log, paths, dialog, notify

conf = config.load_config()                 # 读脚本同目录 config.json（+ .local.json 覆盖）
logg = log.setup_logger("my-script", conf)  # 日志写 logs/my-script.log
out  = paths.output_dir(conf)               # 输出目录
src  = dialog.ask_open_excel("请选择 Excel") # 弹窗选文件
```

## 各模块要点

### config.py

- `config.json` 提交到仓库，`config.local.json` 只放本地（已在 `.gitignore`）
- 值写成 `"${ENV_VAR}"` 会从环境变量取，密码类敏感信息推荐这么做
- `get(cfg, "paths.output_dir")` 按点号取值；`require()` 取不到直接报错，早失败

```python
from autolib import config
conf = config.load_config()
user, pwd = config.load_credentials(conf, "demo", "DEMO_USER", "DEMO_PASS")
```

### log.py

- 统一写到 `logs/<脚本名>.log`，UTF-8，控制台同步输出
- `Timer` 用于统计阶段耗时，直接塞进通知摘要

### paths.py

- `stamped_name(前缀, 后缀, 扩展名)` → `TTTX_UPS_核价结果_20260909_184000.xlsx`
- `latest(目录, 通配符)` 按修改时间取最新产物，用于脚本串联
- `backup_path(文件)` 生成不冲突的备份名

### dialog.py

- **每个弹窗独立创建 Tk 根窗口并立即销毁**。复用根窗口会导致第二次弹窗卡住不显示（仓库里踩过）
- `pick_until_valid()`：选完立刻校验，不合法弹错误提示并重选，最多 10 次
- 无 GUI 环境（计划任务后台）返回 `None`，调用方要能降级到命令行参数

### notify.py

- 只用标准库 `http.client`，零依赖
- **通知失败只返回消息、绝不抛异常**，不影响主流程结果

### browser.py

- 复用本机 Edge（`channel="msedge"`），不用下载浏览器内核
- `wait_stable()`：轮询等页面条数稳定再返回。后台系统常"先显示旧数据再刷新"，直接 sleep 会误判
- `collect_lines()`：边滚动边收集。虚拟表格按滚动分批渲染，不滚会漏数据
- `dump_debug()`：失败自动存 HTML + 截图到 `outputs/debug/`

### excel.py

- `read_rows(ws, aliases=...)`：按别名把列映射成逻辑名，**表格加列也不会错位**
- `append_rows()`：追加到首个空白行，按 (单号, 费用名称) 去重；重复行若原备注为空则补写
- `safe_save()`：文件被 Excel 占用时自动改名保存，不让整个流程崩掉
- `validate_bill_like()`：弹窗选完立刻校验表头，选错马上重选

### tiantu.py

仓库现状：永达项目下的 `check_tiantu.py`（1027 行 Playwright 引擎）已经能跑通天图逐单核验。
本模块是对它的**薄适配器**——新项目只把任务写成符合格式的 CSV（字段
`客户单号(去后缀)` + `期望标记` + `费用名称` + `金额USD` + `说明`），再调用
`tiantu.run(checklist_csv, out_csv=None, limit=0, config=...)` 即可拿到结果 CSV。
引擎路径从 `config.tiantu.engine` 取，默认指向
`02-data-processing/yongda-bill-check/check_tiantu.py`，需要时改这一个字段。

## 依赖

按需引入，不用全装：

| 模块 | 依赖 |
|---|---|
| config / log / paths / notify | 无（标准库） |
| dialog | tkinter（Python 自带） |
| excel | `pip install openpyxl` |
| browser | `pip install playwright`（复用本机 Edge，无需 `playwright install`） |

## 与 yongda-bill-check 的关系

`02-data-processing/yongda-bill-check/` 里有一批功能相同的私有实现
（`default_output_dir` 3 份、取最新文件 3 份、`dump_debug`/`first_visible`/
`try_auto_login`/`maximize_window` 各 2 份）。

**新脚本直接用 autolib，不要复制那些实现。** 老脚本在线上正常跑，暂不改动；
等下次有修改需求时再顺手迁移过来。

2026-09 起 SKYE 账单对账模块（`02-data-processing/skye-bill-check/`）按 autolib
风格从零搭建，并新增 `autolib/tiantu.py` 复用永达的天图核验引擎。
