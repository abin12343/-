# 探迹 Hot 联系方式补全

根据《探迹数据抓取》SOP，从指定 Excel 读取客户名称和不完整的 Hot 联系方式，在探迹 CRM 逐个精确搜索客户，点击联系方式旁的查看图标，读取变化后的 base64 图片并通过 OCR 还原号码。一个客户存在多个 Hot 联系方式时，结果按每个联系方式一行输出。默认使用 Chrome，也可用 `--browser edge` 切回 Edge。

## 安全说明

账号和密码不写入代码、配置、结果或日志。运行时优先读取 `TANZHI_USERNAME` 和 `TANZHI_PASSWORD` 环境变量；未设置时会在终端询问，密码不会回显。不要把真实凭据提交到 Git。

自动化只读取输入工作簿，不覆盖源文件。结果、断点和失败截图均写入独立的运行目录。

## 环境初始化

在本目录打开 PowerShell：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

需要本机已安装：

- Microsoft Edge
- Tesseract OCR，默认位置为 `C:\Program Files\Tesseract-OCR\tesseract.exe`

如 Tesseract 安装在其他位置，请修改 `config.json` 的 `tesseract_path`。

## 运行

推荐先检查输入和环境，不打开浏览器：

```powershell
.\.venv\Scripts\python.exe .\tanji_hot_contacts.py --input "C:\Users\TT1\Desktop\RPA抓取hot不全数据.xlsx" --validate-only
```

先试跑 5 条：

```powershell
$env:TANZHI_USERNAME = "你的账号"
$env:TANZHI_PASSWORD = "你的密码"
.\.venv\Scripts\python.exe .\tanji_hot_contacts.py --input "C:\Users\TT1\Desktop\RPA抓取hot不全数据.xlsx" --limit 5
```

使用 Chrome 续跑：

```powershell
.\.venv\Scripts\python.exe .\tanji_hot_contacts.py `
  --input "C:\Users\TT1\Desktop\RPA抓取hot不全数据.xlsx" `
  --output-dir "D:\MyScripts\output\tanji_hot_contacts_live" `
  --credential-docx "C:\Users\TT1\Desktop\探迹数据抓取.docx" `
  --browser chrome
```

确认无误后全量运行：

```powershell
.\run.ps1
```

常用参数：

- `--start-row 1002`：从 Excel 第 1002 行开始。
- `--limit 100`：最多处理 100 条，适合试跑。
- `--manual-login`：不自动填写账号密码，浏览器打开后手工登录。
- `--save-images all`：保存每个 OCR 原图；默认只保存失败或需复核的图片。
- `--no-resume`：忽略已有断点，创建新的运行目录。

自动登录会填写并提交账号密码。若随后提示“请先选择登录企业”，请在浏览器点击“我知道了”，选择正确企业并进入系统，再让操作端继续流程。企业选择属于账号的交互状态，脚本不会猜测应进入哪家企业。

## 输出

默认输出到 `output/run_年月日_时分秒/`：

- `探迹Hot联系方式结果.xlsx`：最终结果；多个联系方式拆为多行。
- `checkpoint.jsonl`：逐客户追加的断点文件；异常退出后用同一运行目录可继续。
- `run.log`：不含密码的运行日志。
- `evidence/`：OCR 失败或需复核时保存原始号码图片和页面截图。

结果状态：

- `成功`：OCR 得到有效号码，且与输入掩码一致，或多策略结果高度一致。
- `需复核`：得到了号码，但掩码不一致或 OCR 证据不足。
- `未找到客户`：搜索结果中没有客户名称完全一致的行。
- `未找到查看按钮`：找到客户，但对应 Hot 联系方式行没有可点击的查看图标。
- `OCR失败`：图片已取得，但无法可靠识别号码。
- `页面错误`：页面结构、网络或浏览器操作失败，已保留证据供定位。

## 断点续跑

程序每处理一个输入行就追加断点，并定期刷新 Excel。若中断，使用原运行目录继续：

```powershell
.\.venv\Scripts\python.exe .\tanji_hot_contacts.py `
  --input "C:\Users\TT1\Desktop\RPA抓取hot不全数据.xlsx" `
  --output-dir ".\output\run_20260922_170000"
```

已经完成的源行会自动跳过。失败行也会保留在结果中，不会被伪装成空号码或成功状态。
