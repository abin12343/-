# SKYE 打单数据整理（核价 + 差异写表 + 天图核验）

按 `影刀SKYE数据整理SOP.docx` 自动化的账单对账工具。
参考 `02-data-processing/yongda-bill-check/`（永达）的实现，抽出 `08-common-utils/autolib` 通用模块，
天图核验引擎直接复用永达那边的 `check_tiantu.py`，避免再维护一份 1000+ 行的 Playwright 代码。

## 目录

```
skye-bill-check/
├── config.json               业务规则外置（路径 / 报价表 / 费用映射 / 容差 / 天图标记）
├── credentials.example.json  账号模板（zhangfang@tiantutongxun.com / ***）
├── notify.example.json       通知模板
├── requirements.txt          openpyxl + playwright
├── README.md
├── quote_engine.py           报价表解析 + 查价（基础运费/百磅/附加费）
├── process_master.py         Master 加工：K/L 报价表/差异、附加费核对、报价差异
├── process_details.py        FedEx-Details / UPS-Details 加工 + 百磅透视
├── run_skye_check.py         核价引擎 CLI（一条命令跑完核价 → 加工副本 + 报告 + 天图清单）
├── main_flow.py              总流程：选账单 → 核价 → 天图核验（复用永达）→ 写统计表 → 通知
├── launcher.pyw              图形启动器（打包成 SkyeTool.exe，见「打包（EXE 版）」）
├── fetch_bill.py             SKYE 打单系统导出（Playwright + 本机 Edge）
└── outputs/                  产物
```

## 快速开始

```bash
# 一次性安装
pip install -r requirements.txt

# 复制本地凭证（不入库）
cp credentials.example.json credentials.local.json
cp notify.example.json     notify.local.json        # 可选：企业微信 key / 金山文档

# 跑一条账单（弹窗：依次问 账单 → 尾程渠道报价表 → 打单问题统计表）
python main_flow.py
python main_flow.py --input C:/path/to/M8123xxxxx.xlsx --quote C:/path/to/报价表.xlsx
# 全程不弹窗（定时任务/无人值守），一律用参数或 config 里的路径
python main_flow.py --input C:/path/to/M8123xxxxx.xlsx --quote C:/path/to/报价表.xlsx --no-dialog
# 或：先单独跑核价
python run_skye_check.py --input C:/path/to/M8123xxxxx.xlsx --quote C:/path/to/报价表.xlsx
# 末尾 --dry-run 只看不写盘
python run_skye_check.py --input C:/path/to/M8123xxxxx.xlsx --dry-run
```

**报价表时常更换**，所以 `main_flow.py` 默认每次弹窗让用户选一次（标题写明「尾程渠道报价表」，
与选账单、选统计表的弹窗区分开），并通过 `--quote` 传给核价引擎。

## 设计要点

### 1. 复用 08-common-utils/autolib

新代码不再"重造轮子"，配置 / 日志 / 路径 / 弹窗 / 通知 / 浏览器 / openpyxl 全部来自 autolib。
脚本顶部只需要 `sys.path.insert(0, .../08-common-utils); from autolib import config, log, ...`。

### 2. 天图核验引擎直接复用永达

仓库现状：`02-data-processing/yongda-bill-check/check_tiantu.py`（1027 行）是稳定可用的
Playwright 引擎，能直接消费我们生成的《天图核验清单》CSV。
本包不复制这份代码，**通过 `08-common-utils/autolib/tiantu.py`（约 100 行的薄适配器）** +
`main_flow.py` 里的 `run_tiantu_check()` 把它当外部脚本调用即可。

`config.tiantu.engine` 指向永达脚本路径，默认 `D:/MyScripts/02-data-processing/yongda-bill-check/check_tiantu.py`，
如果挪位置改这一个字段就行。

⚠️ 这份引擎是**两个项目共用**的：2026-09-16 修「住宅私人」假「否」时改的就是它（见 §6c），
改动对永达的住宅/地址核验同样生效。

### 3. SOP 步骤的代码映射

| SOP 步骤 | 代码入口 |
|---|---|
| 打开 SKYE 登录页 | `fetch_bill.py::main`（autolib.browser.edge_page） |
| 天图系统登录 / 逐单查标记 | 复用 永达 `check_tiantu.py`（`main_flow.py::run_tiantu_check`） |
| Master K/L 报价表/差异列 | `process_master._ensure_helper_columns` + `_price_master_freight` |
| Master 附加费列核对 | `process_master._price_master_surcharge`（带 fallback_sheets 兜底） |
| FedEx-Details 报价表/差异 + 费用类型 | `process_details.process_fedex_details` |
| UPS-Details 费用/类型/报价表/差异 + 百磅 | `process_details.process_ups_details` |
| 写《打单问题统计表》 | `main_flow.write_to_stats`（复用 autolib.excel.append_rows） |
| 金山文档运行记录 + 企微推送 | `main_flow.notify_all`（复用 autolib.notify） |

### 4. 报价表引擎（quote_engine.py）

单文件、零业务耦合、懒解析。三个核心数据结构：

- `BaseTable` — 基础运费：weight(向上取整) × Zone 2..8
- `HwtTable` — 百磅：200LB+/500LB+ 分档 + Min 价
- `SurchargeTable` — 附加费：支持 `by_zone`（按分区）/ `by_kind`（按商业/住宅）/ `flats`（固定价）

`surcharge_price()` 主 sheet 找不到费用项时按 `service_map.fallback_sheets` 兜底——这是 HWT
/Multiweigh 行的偏远/住宅等附加费能正确核对的关键。

### 5. 加工副本的列位（idempotent）

- Master：J 列右侧插入 K=报价表, L=差异（按 SOP）；翻译列**追加到表尾**（不破坏 SOP 既有列号）。
- FedEx-Details：I 列右侧插入 J=报价表, K=差异；AX 列右侧插入 AY=费用类型。
- UPS-Details：N 列右侧插入 O=费用, P=费用类型, Q=报价表, R=差异；S=原 O=Net Amount。

已经加工过的账单再次运行：表头已存在就不重复插入，直接更新数值。

**「报价表」列写公式，不写死数值**（2026-09-15 起，Master `K` / FedEx-Details `J` / UPS-Details `Q`）：

```excel
=VLOOKUP(ROUNDUP(V2,0),'[尾程渠道报价.xlsx]FedEx Home Delivery  '!$A:$H,8,0)
```

- **目录写相对路径**（2026-09-16 改）：报价表与账单同目录就只写 `[报价表.xlsx]<sheet>`，
  在上一级写 `..\SKYE\[报价表.xlsx]<sheet>`；Excel 原生就这么存同目录链接，**换台电脑照样匹配**。
  之前写死 `'C:\Users\TT1\Desktop\SKYE\[…]…'!`，在别的机器上打不开报价表，公式全 `#REF!`。
  只有跨盘符（如账单在 D:、报价表在 C:）才退化成绝对路径，那种情况本身就没法相对——把两个文件
  放一起即可。由 `QuoteBook.external_ref()` 按 `book_dir`（=账单所在目录）算，见
  `tmp/cmp_formulas.py`（两处不同目录跑出的公式逐格相同）、`tmp/probe_relpath.py`（无绝对路径）。
- 形状对着人工模板来：区间首列 = 重量列、末列 = 最后一个分区列、第 3 参数 = 分区号所在列
  （报价表的分区是 2..9 且首列是重量，所以"分区 N"正好落在第 N 列）。与人工模板里 698 条
  `=VLOOKUP(I4,'[1]UPS-Ground商业'!$B:$I,8,0)` 的几何完全一致（人工用 `[1]` 索引 + externalLinks 部件）。
- `ROUNDUP(...,0)` 是必需的：报价表阶梯只按整数磅排，小数重量精确匹配会 `#N/A`；
  取整口径与 `BaseTable.lookup(round_up=True)` 相同（`config.run.round_up_lookup`）。
- **表结构不合形状 / 重量超出阶梯上限时退回数值**（`vlookup_formula()` 返回 `None`）：
  超上限的票 `base_price` 会夹到最后一档，公式给不出这个值，写死才不会算错。
- **只有 `weight_zone` 服务（如 FedEx Home Delivery）逐行写**。百磅服务的明细行是多件**合并
  计费**的单行（本账单 20.25 vs Ground 阶梯 26.16），逐行套阶梯价会凭空多出 383 行假差异；
  人工模板也只给了凑巧对上的那几行写了公式。百磅的报价走两张透视页。
- `wb.calculation.fullCalcOnLoad = True`：openpyxl 建不了 Excel 的 `xl/externalLinks` 部件
  （人工模板的 `[1]...` 那种），本机算不出缓存值。打开账单时 Excel 会提示更新链接，
  确认后全文重算即得数（**首次打开需人工确认一次**）。

**幂等**：本轮不给报价的行（非 `weight_zone`、重量/分区缺失、低于百磅门槛…）会**擦掉上一轮
留在报价列/差异列的公式或数值**（`priced` 集合 + 收尾清扫，日志打印「擦掉 N 行」）。
否则改一次规则重跑，旧值会残留在不该有值的行上——自测时正是这样留下过 382 个 Multiweigh 公式。

### 5b. 两张百磅透视页（严格按 SOP 018-035 搭）

SOP 只让 FedEx-Details / UPS-Details 做百磅，其余费用一律在 Master 核对（SOP L36）。
两张页的列**只保留 SOP 要的**（早期版本多出来的 `公布价/明细行数/备注` 已去掉）：

| 页 | 列 | 取数 |
| --- | --- | --- |
| `FedEx-百磅透视` | 运单号 / 费用类型 / 分区 / 总重量 / 账单金额 / 报价 / 差异 | SOP L24-26：粘贴 `I`(金额) `V`(重量) `AX`(运单号) `AY`(费用类型) `BN`(分区) |
| `UPS-百磅透视` | 运单号 / 分区 / 总重量 / 账单金额 / 报价 / 差异 | SOP L33：粘贴 `G`(运单号) `I`(重量) `L`(分区) `S`(金额) |

- 列字母都是**按表头名现算**（`_col_by_name`），不写死；模型里 `BB` 就是 SOP 说的 `AX`
  那一列值（`AX` 带 `-2` 后缀、`BB` 已去后缀，透视要的是去后缀的）。
- 总重量/账单金额 = `SUMIF(S)` 回明细页按运单号聚合（等价于 SOP 的"做透视表"；
  **openpyxl 建不了真正的 Excel PivotTable 对象**，所以用公式得到同一组数，改重量/单价会自动跟着变）。
- 报价 = `MAX(ROUND(ROUND(档位单价×总重量,2)×折扣,2), 保底价)`，档位单价由"费用类型 + 分区 + 重量档"
  匹配报价表得到（已逐行与报价表独立核对一致）。
- 差异 = 报价 − 账单金额（SOP 说的"加减法"）。用户 2026-09-14 拍板：百磅两种计费并存
  （0.9625 折扣价 / 原价），差异取离账单更近的那个候选。
- **只有 Multiweigh / MWT 用百磅方式算**（SOP L22）：FedEx Home Delivery 等非百磅服务行照留，
  但报价/差异留空；这些行连同"取不到分区/报价表匹配不到"的行，原因写进**报告 txt** 的
  「百磅透视页无法定价的行」，不占用透视页的列。
- 改完用 `tmp/verify_pivot.py`（列名/公式引用/运单号集合）和 `tmp/verify_pivot_rate.py`
  （逐行回报价表查单价比对公式常量）自检；重跑账单时用 `tmp/sheet_digest.py` 比对逐页指纹，
  确认只有这两页变了。

### 6. 天图核验结果 → 《打单问题统计表》的数据流

```
核价引擎 → <账单>_天图核验清单_<时间戳>.csv   （本次要查的问题，列见下）
         ↓  永达 check_tiantu.py（浏览器）
        天图核验结果_<时间戳>.csv            （写在永达自己的 outputs/ 下）
         ↓  main_flow.collect_tiantu_missing()
       《打单问题统计表》的 skye sheet        （独立文件，不在账单里）
```

- **清单口径 = 所有被收费的附加费行**（SOP 5-13 项：超重附加费 / 超尺寸附加费 / 地址更正 /
  商业偏远附加费 / 住宅偏远附加费 / 商业超偏远附加费 / 私人超偏远 / 私人住宅附加费 / 超大尺寸费用），
  **不只**核价有差异的那些；按 (费用类别, 单号去后缀) 去重，一票一个费只需验一次标记。
  `异形包装`（Additional Handling）不在 SOP 的 5-13 项里，不列入。
- **清单列**：`Excel行号, 客户单号(去后缀), 费用类别, 费用名称, 期望标记, 金额USD, 说明`
- **结果列**：`单号, 费用名称, 期望标记, 是否出现, 命中条数, 明细, 缺失关键词`
  ⚠️ 两套列名**不一样**，合并时按结果 CSV 的列名读、按 `(单号, 费用名称)` 回清单补 `金额USD`。
  按清单列名去读结果（`客户单号(去后缀)`/`依据`）会全部读空并塌成 1 行空单号。
- **只登记「天图未查到标记」的**（`是否出现 != 是`）。核价差异走「报价差异」sheet，不进统计表。
- **只认本次清单里出现过的 `(单号, 费用名称)`**：否则 `--skip-tiantu` 或天图失败时会拿上一次的结果
  登记早已处理完的单号。天图阶段失败时**不退回旧结果**，只在状态里标「部分成功」。
- 统计表的各客户 sheet **列序/列数都不一样**（环洋 `运单号|金额USD|费用名称`、中盟 `费用名称|金额USD|运单号`、
  永达多一个`备注`），一律按表头名映射（`STATS_ALIASES`），按 `(运单号, 费用名称)` 去重后追加（`autolib.excel.append_rows`）。
- **备注要去登记**（2026-09-16 改）：备注内容是 `天图未查到标记：偏远（缺 XXX）`。原先的代码把备注
  拼好了却**丢掉**——`skye` sheet 只有 3 列、没有「备注」列，`append_rows` 就跳过不写。
  现在 `write_to_stats` 发现表里没有「备注」列就**在表头行末尾自动补一列**（与永达 sheet 的第 4 列同位置），
  再把备注写进去；已存在的重复行若备注为空也会补写。**在老的 3 列表上跑一次就会变成 4 列**，
  这是用户要求的（否则人工看不出为什么登记这行）。
- 旧流程往**账单里**新建的「打单问题统计」sheet 已被 `main_flow._drop_legacy_stats_sheet()` 清除。
- **统计表和报价表一样每次都要用户选**：不传 `--stats` 且不是 `--no-dialog` 时弹窗（标题写明
  「打单问题统计表」，与账单/报价表弹窗区分），并且**校验收到的文件里必须有 `paths.stats_sheet`
  那张 sheet**，选错文件当场提示重选。弹窗被取消才回落到 `config.paths.stats_file`。
  （早期版本的 bug：`stats_used` 先取了配置值，`if not stats_used` 永远为假 → 弹窗从来没弹过。）
  另外本次没有要登记的问题时不弹窗、不写入。

### 6c. 「住宅私人」核验的读取方式（2026-09-16 修过一次假「否」）

SOP L110（私人住宅附加费）要求「去天图-运单查看**应收栏是否收取住宅私人地址费**，没有则登记统计表」。
这一项在永达引擎里是**逐单查询**（应收列在表格最右侧，要横向滚动）。原先的实现有两个坑，
2026-09-16 用户报「私人住宅天图检验发现全部有标签」——自检后确认**4 个单号随机从「是」翻成「否」**：

- **坑 1：应收格子的内容是异步填的，读一次读空就当「否」。** 旧代码查完只等 2.5 秒就读一次。
  调试图证明：失败那次 `td[colid=应收]` 存在、**一个字都没有**，而同一行其它列（客户类型/服务/件数…）
  都有值 —— 不是找不到行，是内容还没到。现在改成**轮询到读出内容或超时（25s）**，
  仍为空就**重查一次**再轮询。
- **坑 2：`scroll_to_receivable()` 一次都没真的滚动过。** 旧实现是
  `if "应收" in body.inner_text(): return True` —— 表头里一直都有"应收"二字，于是第一轮就返回 True。
  现在改成按**表头元素的实际 bounding box** 判断它是否真进了可视区，另加
  `scroll_into_view_if_needed` + 逐次横向推动。
- **坑 3（假「是」方向）：弹层残留 + 整页文本扫描。** 旧的兜底会把「最后一个『应收』之后的整页文本」
  交给关键词匹配；而上一次核验若走「应收费用」兜底没关干净，残留弹层里的**内部备注**
  （写着"住宅地址费头程已收"）就会被下一单命中 —— 能把没有该费用的单判成「是」。
  现在每单开头先 `ensure_no_dialog()` 关掉残留弹层，关键词扫描**排除弹层内文字**，
  且必须先取到该单**自己那一行**的应收格子。

**认行方式（关键）**：vxe-table 把运单号列放在**左侧固定表（另一张 `<table>`）**里，
所以「行文本里有没有这个单号」永远不成立；三张表靠 `rowid="row_494"` 对应同一逻辑行。
`check_tiantu._RECEIVABLE_ROW_JS` 因此按 **rowid** 认行；单查只返回一行时退化为直接取那一行。

验收：24 个「住宅私人」单号**连跑两遍，24/24 完全一致（23 是 + 1 否，翻转 0 条）**
（改前：19 是 / 5 否，两遍间 4 个 是→否 翻转）。脚本 `tmp/resi_double_run.py`、
单号级探针 `tmp/probe_receivable_src.py`。

**唯一稳定的那个「否」是 `USGZ202605142168`：这是真实的**，不是漏读 ——
它的应收栏只有头程费用（海外仓入仓 279.24 / 快递出仓费 186.16 / 分拣费 335.09 / 贴快递标签 279.24），
确实没有住宅私人地址费（它的**内部备注**写着"快递派，一箱两张快递标签，住宅地址费头程已收"，
要不要算「有标签」需人工定夺，代码按 SOP 只认应收栏）。

### 6b. 天图耗时自测

天图慢只慢在**查询次数**上（每次查询 5-20 秒），跟 `--timeout`、`--batch` 的配置关系不大
（引擎里 `run_query` 的等待上限是 `max(wait_ms, 90)`，且结果稳定就提前返回，所以调 wait 基本没用）。

- 每次跑天图都会把引擎输出留一份日志：`outputs/<账单名>_天图运行_<时间戳>.log`（同时照常打在屏幕上）。
- 跑完自动统计并打印一行，例如：
  `[天图] 自测：共 50 次查询（批量查询 9、拆半补查 3、单查 12、应收核验 24），用时 640s，平均 12.8s/次`
- 从日志能看出时间花在哪：`住宅私人`（私人住宅附加费）和 `address`（地址更正）在永达引擎里是
  **逐单查询**（应收列要横向滚动），批量查询读不全时会 `拆半重试` 退化成逐单——这两类是最费时的。
- 缩短时间的可选做法（都需要先跑一次拿上面的数字）：
  1. 提高 `tiantu.batch`（默认 13）：清单偏"偏远"时批次数 9→按 40 算只要 3 次；但超过站点一次能
     返回的条数会触发拆半补查，反而更慢，要实测。
  2. 按标记分片并行跑多个引擎进程（各自一份 csv）：查询次数按进程数摊开，最省时间；但同一账号
     并发登录是否互相踢掉需要实测。
  3. 减少清单条数：目前 149 条里有 11 个单号命中了 2 个标记，引擎按标记分组会把它们查两次。

### 6d. 天图超时上限算漏了「逐单核验」（2026-09-17 用户报「天图核验步骤失败」）

现场（另一台机器，`M8123260904N0035.xlsx`）日志：`[主流程] 天图核验 用时 1270.68s（无结果）`，
引擎日志最后几行停在 11:38:24 且**仍在正常查询**（在为 `USSC2606150158` 读应收栏），
没有任何报错、也没有结果 CSV。1270.68s ≈ 当时算出的上限 `batches × 60 + 180` ——
**是我们自己的总时长上限把它掐掉的，不是引擎挂了。**

根因：原来的上限只按**批次数**估，但 `check_tiantu.py` 里 `if mark == "住宅私人":` 的分支
**不走批量查询**，是一单一单去读应收栏（实测 20-30 秒/单），这一项在上百单时会比所有批量查询
加起来还慢。修法三处（都在 `main_flow.py`）：

1. **`_count_per_waybill_rows()`** —— 数清单里「期望标记 = 住宅私人」的条数，加进估算：
   `总时长上限 = max(timeout, 批次数 × per_batch_seconds + 逐单核验条数 × per_waybill_seconds + 180)`。
   对着用户那份真实清单验证：`清单 149 条，其中逐单核验 24 条` / `批次数 12 → 旧上限 900s，新上限 1620s`，
   其中的 24 与 09:44 那次实跑的 24 次「应收核验」完全一致。
   新增配置键 `tiantu.per_waybill_seconds`（默认 30）。
2. **`_run_teeing()` 改成「静默看门狗 + 总时长兜底」** —— 原来是一次性 `threading.Timer`，
   到点就杀。现在另起一个守护线程，两条判据：连续 `tiantu.idle_seconds`（默认 300）没有任何输出
   → 判**卡死**；超过总时长上限 → 判总时长超限。返回 `(退出码, 中止原因)`，日志写明是哪一条触发的。
   三个分支各有单测：一直有输出但超总上限 → `总时长超过上限 3 秒`；只输出一行就卡住 → `卡死：3 秒没有任何输出`；
   正常跑完 → 原因为空、退出码 0。
3. **结果 CSV 只认本次产出** —— 结果由引擎写到它自己的输出目录，发布包/上一轮的 `outputs/` 里
   可能还躺着旧的 `天图核验结果_*.csv`。现场那条 `天图核验结果: …\outputs\天图核验结果_20260917_094413.csv`
   就是**九天前 09:44 那轮的残留**（随发布包一起被拷到了那台机器），不是本次结果。
   `find_tiantu_result()` 加 `newer_than` 参数，两处查找都传本次开始时间。
   （`--skip-tiantu` 是用户明确要求复用上一次结果，那一处**故意**不带 `newer_than`。）

**发布包已经清空 `outputs/`**（原先里面躺着我自己冒烟跑出来的 4 个文件，包括那个 094413 的旧结果）。
换机器前顺手清一下 `outputs\` 更不容易看混。

冻结版冒烟（发布包副本，通知置空、引擎换成 `cmd.exe` 假失败）：清单 5 条 → `约 1 批 + 5 条逐单核验，
总时长上限 390 秒` 正确；旧结果放在 `outputs/` 里仍报 `天图核验结果: 未生成`，
换成新时间戳的结果文件则正常认出。

## 打包（EXE 版）

交付给现场用 `release/SKYE账单整理工具_EXE版/`（布局照 永达 EXE 版）：

```
SKYE账单整理工具_EXE版/
├── 启动SKYE工具.bat          双击入口：cd /d "%~dp0" 后拉起启动器
├── SkyeTool/SkyeTool.exe     图形启动器（launcher.pyw，--windowed）
├── main_flow/main_flow.exe   主流程
├── run_skye_check/…exe       核价引擎
├── check_tiantu/…exe         天图引擎（135MB，含 playwright driver/node.exe）
├── config.json               paths 相对化 + tiantu.engine 指向包内 exe
├── credentials.local.json    明文账号（随包，不入库）
├── notify.local.json         明文企微 key + 金山 token（随包，不入库）
├── 打单费用名称中英文翻译.xlsx  翻译表（报价表不随包，每次弹窗选）
├── outputs/ logs/            空目录占位
└── 使用说明.md
```

构建（**不要用 venv**，里面没 PyInstaller；用装了 PyInstaller/openpyxl/playwright 的
`C:/Users/TT1/AppData/Local/Python/pythoncore-3.14-64/python.exe`）：

```bash
PY="C:/Users/TT1/AppData/Local/Python/pythoncore-3.14-64/python.exe"
cd 02-data-processing/skye-bill-check
C=(--noconfirm --onedir --distpath release/dist --workpath release/build --specpath release/build)

# 1) 主流程：autolib 靠 --paths 让 PyInstaller 自己发现；显式排除会连带 playwright 的 autolib.browser/tiantu
"$PY" -m PyInstaller "${C[@]}" --name main_flow --paths D:/MyScripts/08-common-utils \
  --hidden-import autolib.dialog --hidden-import autolib.excel --hidden-import autolib.notify \
  --exclude-module autolib.browser --exclude-module autolib.tiantu --exclude-module playwright main_flow.py

# 2) 核价引擎（纯 openpyxl）
"$PY" -m PyInstaller "${C[@]}" --name run_skye_check run_skye_check.py

# 3) 天图引擎：跨目录 import run_yongda_check，必须 --paths 指到永达目录；
#    PyInstaller 没有 playwright 的 hook，必须 --collect-all（+约 106MB）
"$PY" -m PyInstaller "${C[@]}" --name check_tiantu \
  --paths D:/MyScripts/02-data-processing/yongda-bill-check --collect-all playwright \
  D:/MyScripts/02-data-processing/yongda-bill-check/check_tiantu.py

# 4) 图形启动器：只有它用 --windowed
"$PY" -m PyInstaller "${C[@]}" --name SkyeTool --windowed --paths D:/MyScripts/08-common-utils launcher.pyw
```

产物落 `release/dist/<名>/`，再挪进发布根、改成上表的布局（`tmp/assemble_release.py` 做这件事）。
`release/`、`build/`、`*.spec` 都在 `.gitignore` 里。

### 冻结（frozen）适配：6 处

打 exe 后 `sys.executable` 是发布根下的 `main_flow.exe`、`__file__` 在 `_internal/` 里，
所以下面这些必须走 `getattr(sys, "frozen", False)` 分支，否则打包必坏（dev 行为不变）：

1. **`HERE = exe.parent.parent`** —— `main_flow.py` / `run_skye_check.py` / `fetch_bill.py` /
   `launcher.pyw` 都用它定位**发布根**（`config.json`、`outputs/`、各 exe 都在发布根下）。
2. **引擎启动命令** —— `main_flow._engine_cmd()`：冻结时用 `<发布根>/check_tiantu/check_tiantu.exe`，
   否则 `[PY, <脚本>]`。不这么做会变成"用 main_flow.exe 去跑 check_tiantu.py"，
   argparse 收到未知参数退出码 2，天图静默无结果。`tiantu.engine` 也支持相对发布根的路径。
3. **`run_stage()` 的存在性检查** —— 冻结时检查 `<HERE>/<脚本名>/<脚本名>.exe`，
   非冻结时才检查 `.py`（无条件查 `.py` 会让核价阶段一上来就 127）。
4. **autolib 路径** —— 冻结时 autolib 已打进 PYZ，`_autolib()` / `fetch_bill.py` 跳过
   `sys.path.insert(HERE.parents[1]/"08-common-utils")`（冻结后那个路径指向发布目录的祖父目录）。
5. **子进程强制 UTF-8** —— 见下。
6. **启动器 `sys.stdout` 为 None** —— `--windowed` 下 `print()` 会炸，`launcher.pyw` 给它一个黑洞 sink。

### 编码坑（打包后才暴露）

**PyInstaller 冻结后不认 `PYTHONUTF8` / `PYTHONIOENCODING`**：即使父进程设了这两个变量，
冻结的 exe 往管道里仍按 **cp936** 写中文。而 `main_flow._run_teeing()` 按 utf-8 解码子进程输出，
`summarize_tiantu_log()` 又靠中文措辞（`[批] 偏远 x13 -> `）数查询次数 —— 不修就会整段乱码、
查询次数全数成 0，启动器窗口里也是乱码。

修法是让每个冻结程序自己把 stdout/stderr 收到 utf-8（`sys.stdout.reconfigure(encoding="utf-8")`，
写在 `getattr(sys, "frozen", False)` 分支里）：`main_flow.py` / `run_skye_check.py` 的头部 +
永达 `check_tiantu.py` 的头部。`_child_env()` 里那两个环境变量保留，当双保险。

### 弹窗被子进程的前台锁定埋掉（2026-09-17）

用户报「最后的统计表弹窗没有弹出」。实测：`main_flow.exe` 是**子进程**、不是前台进程，
Windows 的前台锁定不允许它把新窗口激活到最前；如果这时用户正看着 Edge（天图刚跑完），
选择框就生成在 Edge **底下**，用户完全看不到 —— 窗口确实存在（EnumWindows 能枚举到
class `#32770`），只是不在最上层、也不是前台窗口。

`autolib/dialog.py::_root()` 原来「先置顶、`update()` 一下、再取消置顶」。改成**根窗口一直保持
`-topmost`**：filedialog 是根窗口的 owned window，owner 置顶弹窗才跟着置顶。同一次对照里，
旧包（26924）的对话框没有 `WS_EX_TOPMOST`，新包（29324）是 `FG TOPMOST`。

顺带修掉同一文件里的两处：`pick_until_valid` 的错误提示框原先 `_root()` 拿到根窗口就丢弃
（`_default_root` 还指向上一个已销毁的 root，`showerror` 会抛异常只剩 print），
`confirm()`/`alert()` 同样只建不销毁。现在都显式 `parent=root` 并在 `finally` 里 `destroy()`。

### 相对路径与工作目录

`run_skye_check.py` 的 `output_dir` / `translation_file` 是**按 CWD 解析**的，
天图引擎则按"自己 exe 的祖父目录"解析 —— 所以启动器用 `cwd=发布根` 拉起主流程，
`启动SKYE工具.bat` 也必须先 `cd /d "%~dp0"`。`outputs/` 与 `logs/` 在发布根下统一。

### 顺手记的构建事实

- `--collect-all playwright` 收进的是 `playwright/driver/node.exe`（约 92MB）与 `package`（约 14MB）；
  浏览器二进制不需要 —— 引擎复用本机 Edge（`channel="msedge"`）。
- 纯 Python 包（如 `autolib`）在 `_internal/` 里看不到，它打进 PYZ 归档了；
  想知道有没有收进去，翻 `release/build/<名>/PYZ-00.toc`（里面能看到 `autolib.notify` 等）。
- 天图引擎不自包含：`check_tiantu.py` 里有 `from run_yongda_check import load_credentials`，
  所以构建路径必须含 `yongda-bill-check/`（dev 与打包都靠这个）。

## 配置字段说明

- `paths.quote_file` — 报价表 xlsx（脚本会自动识别每个 sheet 内的基础运费/百磅/附加费块）
- `paths.translation_file` — 翻译表 xlsx（Sheet1 优先，Sheet2 兜底）
- `paths.stats_file` — 打单问题统计表
- `paths.output_dir` — 产物输出目录
- `paths.stats_sheet` — 统计表里 SKYE 用的 sheet 名（默认 "skye"）
- `paths.input_dir` — SKYE 账单初始目录（弹窗默认路径）
- `quote.service_map` — Service Code → 报价表 / 块 / surcharge_kind / 兜底 sheet
- `quote.hundredweight_blocks` — 百磅分档规则（<500 走 200LB+，>=500 走 500LB+，最低 81）
- `master.fee_checks` — 附加费列 → 报价表费用项的映射
  （`per_piece: false` = 该费用按票收、报价直接用单价；省略 = 按箱计费、乘本票件数）
- `details.FedEx-Details` / `details.UPS-Details` — 各 Details sheet 的列定义
- `tiantu_checks` — 费用类别 → 期望标记 / 天图查询口径
- `tiantu.engine` — 复用永达 check_tiantu.py 的路径
- `tiantu.batch` — 每批查询的单号数（默认 13，永达实测一次能返回的上限）
- `tiantu.timeout` / `tiantu.per_batch_seconds` / `tiantu.per_waybill_seconds` — 子进程总时长上限 =
  `max(timeout, 批次数 × per_batch_seconds + 逐单核验条数 × per_waybill_seconds + 180)`，
  批次数 = `ceil(清单条数 / batch)`，逐单核验条数 = 清单里期望标记为「住宅私人」的条数（宁松勿紧）
- `tiantu.idle_seconds` — 连续多少秒没有任何输出就判卡死并中止（**真正的中止依据**，
  总时长上限只是兜底）
- `notify` — 复用 autolib.notify 的字段

改价改列名改路径都改 `config.json`，**不动代码**。

## 已知校准点（运行后需要人工校准）

> 本节是 2026-09-04 首次冒烟（`M8123260904N0035.xlsx`）的校准记录，历史留档；
> 已修复的条目带 `✅` 与证据。以最新实跑的报告/「报价差异」sheet 为准。

冒烟跑 `M8123260904N0035.xlsx` 这份真实账单（2026-09 导出，689 行 Master）的输出：
- **Master 运费差异 216 行**：约 80% 是 FedEx / Multiweigh 行的运费——这部分运费其实记在 FedEx-Details / UPS-Details，Master 上记 0 是正常的；可在 config 里把 `master.columns.amount` 对 FedEx 服务行跳过（建议：service_code 包含 FedEx / Multiweigh / MWT 的行不核 Master 运费）。
  ✅ 2026-09-14：`M8123260821N0033.xlsx` 实跑只剩 7 行运费差异。
- **T01 行的差异（bill 23.06 vs R 15.05）**：这批差异集中在 26-30LB / zone 8 范围。可能是 SKYE 用了"商业+另一种变体"（block 2：K..S 段）报价，而脚本只取 block 0（A..I 段）。建议在 `service_map` 加 `block: 1` 的别名规则或确认 26-30LB 真的该用 block 1。
- **HWT 行百磅存在约 3.75% 折扣**：HWT 表只有 200LB+ / 500LB+ 档，没有 1000LB+ 档；>=1000 LB 也走 500LB+ 档。部分行 bill 实际是 `总重 × 0.9625 × 档位费率`，脚本按全额计算会差约 3.75%。在 `quote.hundredweight_blocks` 加 `discount_factor: 0.9625` 之类的字段即可支持。
  ✅ 2026-09-14：两种计费并存（约 60% 打 0.9625 折扣、约 40% 原价），`_hwt_accept()` 取最接近账单的候选；
  配置键 `quote.hundredweight_discount`。
- **附加费按"件"计 vs 按"票"计**：HWT 偏远费 22.5 = 10 × 2.25 或 5 × 4.5，需用 pieces 数乘回去；目前脚本只算单价，差异属正常，落到 报价差异 sheet 人工核对。
  ✅ 2026-09-14：已按 `单价 × 本票件数` 计，并叠加报价表的**封顶值**（`min(单价×件数, 封顶)`，封顶是整票上限）；
  本票 15 行触发封顶，与账单分毫不差。
  ✅ 2026-09-17：`超重附加费`（AHS - Weight）与 `超尺寸附加费`（AHS - Dimensions）改为**按票**收 ——
  `fee_checks[].per_piece: false` 时直接用单价、不乘件数（用户指出这两项不该乘片数）。
  行133（FedEx Multiweigh，14 件，账单只收 1 个单价 11.75）的假差异由此归零。
  **残留（待用户确认）**：行95 `USXM202607240025`（UPS GROUND HWT，8 件）账单收 194.4 = 24.3 × 8，
  改后变成 +170.10 的差异 —— 要么这行账单按件多收了，要么 UPS 侧这一项确实按件计费。
- **Master E → Details G/H 的对应关系**：M8123 的 Master E（Reference No = `USSZ202608050115`）与 UPS-Details G（Shipment Reference Number = `USSZAS2607250246`）前缀都是 `USSZ` 但位数不同，**不能直接 VLOOKUP**。需要确认 SKYE 真实用哪个字段关联，或在配置里加一个 `tracking_match.regex` 抽取主键。这导致当前 Details sheet 的核价 diff 几乎为 0。
- **翻译表的 Service Code 映射**：当前只对 Master C 列做了 en→zh，但 119 条翻译里很多 UPS 渠道名不在内；翻译列会留空，不影响核价。
- **危险品、超不可发、超偏远的"远 Alaska / Hawaii"等特殊区段**：报价表里已收录（`超偏远费` 的 `Remote/Alaska/Hawaii`），但 Service Code 决定 kind=商业/住宅，特殊区段未在 fee_checks 中显式区分。
- **Details 明细页的「报价表」列**（FedEx `J` / UPS `Q`）原先写死数值，人工模板是跨工作簿 VLOOKUP
  公式；SOP L21/L30-31 的口径也是"根据 `V` 列查找值、`BN` 列为行数匹配报价"。
  ✅ 2026-09-15：已改成公式（形状/取整/退回数值的规则见 §5）。用 `tmp/verify_detail_quote.py`
  逐行模拟 Excel 复算与改前快照比对：**776/776 一致、0 不一致**（FedEx 78 + UPS 698）；
  Master `K` 195 个公式亦逐行独立重算等于 `base_price`。逐页指纹证明只有
  Master / FedEx-Details / UPS-Details 三页变化，13 页不动。
  ✅ 2026-09-16：路径已改成**相对账单目录**（同目录只写 `[报价表.xlsx]<sheet>`），
  换台电脑只要报价表与账单放一起就能匹配，不再写死 `C:\Users\TT1\…`；跨盘符才退回绝对路径。
  **残留**：没有 `externalLinks` 部件（人工模板那种 `[1]` + `xl/externalLinks`），
  第一次用 Excel 打开需人工确认"更新链接"（本机没装 Excel，无法自测这一步）。
  非百磅行（FedEx Home Delivery 54 行）继续留在透视页，未动。
- **天图「住宅私人」核验会随机把有标签的单判成「否」**（用户 2026-09-16 报「全部有标签」）：
  根因是应收格子内容异步加载、旧代码只读一次；外加 `scroll_to_receivable` 从未真的滚动、
  弹层残留会污染整页关键词扫描。
  ✅ 2026-09-16：引擎已改（`check_tiantu.py`，按 rowid 认行 + 轮询 + 关残留弹层 + 关键词只在列表区扫），
  24 个单号连跑两遍 24/24 一致、翻转 0 条。详见 §6c。
  **残留（未改，仅记录）**：SOP L95/L105 对「住宅偏远附加费 / 私人超偏远」还要求同时确认
  **应收栏**收了偏远费和住宅私人地址费，而这两类走的是批量查询、只看了运单号下的标记文字。
  当前只松不紧（不会报假「否」）；且同一单的「私人住宅附加费」条目已在逐单查应收，
  只有"收了住宅偏远、没收私人住宅附加费"的单会漏掉这一半。要不要补需用户定夺（补法是逐单查，112 单会很慢）。

## 与永达的差异

| 维度 | 永达 | SKYE |
|---|---|---|
| 报价表 | 简单 weight×Zone 两张表 + 1 张百磅 | 9 张 sheet，每张含 weight×Zone + 附加费块 |
| 账单格式 | 单 Master sheet | Master + FedEx/UPS/USPS/Amazon 等多 Details sheet |
| Service Code | "Ground Commercial" / "Ground Residential" | "UPS-Ground-C-LA" / "FedEx®Ground Multiweigh" / ... |
| 附加费列 | 一栏 24 列 | Master 50+ 列、Details 200+ 列 |
| 天图 | 复用现有引擎 | 同上（autolib/tiantu.py 适配） |
| 代码库 | 自有实现 | 直接 `from autolib import ...` |

## 依赖

- Python 3.10+
- `openpyxl`（加工副本、报价表、统计表读写）
- `playwright`（`fetch_bill.py` + 复用的永达 `check_tiantu.py`，**复用本机 Edge，不需要 `playwright install`**）
- 网络仅在天图核验 + 通知阶段需要，弹窗核价本身完全离线
