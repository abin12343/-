# SKYE脚本GitHub上传指南

## 一、已完成的准备工作

我已经为你完成了以下操作：

### 1. 初始化Git仓库
```bash
cd D:/MyScripts/02-data-processing/skye-bill-check
git init
```

### 2. 配置Git用户信息
```bash
git config user.name "abin12343"
git config user.email "abin12343@users.noreply.github.com"
```

### 3. 创建.gitignore文件
已创建，排除了敏感文件和临时文件：
- credentials.local.json
- notify.local.json
- __pycache__/
- outputs/*（保留.gitkeep）
- logs/*（保留.gitkeep）
- 发布包文件

### 4. 添加并提交所有文件
```bash
git add .
git commit -m "Initial commit: SKYE账单核价工具

- 核心功能：SKYE账单自动核价与报价表匹配
- 账单处理：主表+明细表联合处理
- 报价引擎：支持阶梯价格和附加费
- 图形界面：launcher.pyw提供友好操作界面
- 完整文档：README.md包含使用说明"
```

✅ **提交ID**: `45fdc8d`  
✅ **提交文件**: 15个文件，4277行代码

### 5. 配置远程仓库
```bash
git branch -M main
git remote add origin https://github.com/abin12343/skye-bill-check.git
```

---

## 二、需要你完成的操作

### 步骤1：在GitHub创建仓库

1. 访问 https://github.com/new
2. 仓库名称填写：`skye-bill-check`
3. **重要**：选择仓库可见性
   - **Private（推荐）**：只有你能看到，适合内部业务代码
   - **Public**：所有人可见，不推荐（包含业务逻辑）
4. ⚠️ **不要勾选**"Add a README file"（我们已经有了）
5. ⚠️ **不要选择**"Add .gitignore"（我们已经创建了）
6. 点击"Create repository"

### 步骤2：推送代码到GitHub

在命令行中执行：

```bash
cd D:/MyScripts/02-data-processing/skye-bill-check
git push -u origin main
```

首次推送会要求你登录GitHub：
- **用户名**：abin12343
- **密码**：使用Personal Access Token（不是GitHub密码）

### 步骤3：如果需要创建Personal Access Token

如果你还没有token：

1. 访问 https://github.com/settings/tokens
2. 点击"Generate new token" → "Generate new token (classic)"
3. 勾选权限：
   - ✅ `repo`（完整仓库访问权限）
4. 生成后，复制token（只显示一次）
5. 推送时使用token作为密码

---

## 三、推送后验证

推送成功后，访问：
```
https://github.com/abin12343/skye-bill-check
```

你应该能看到：
- ✅ README.md（包含完整文档）
- ✅ 所有Python源代码文件
- ✅ config.json、requirements.txt等配置文件
- ✅ 示例配置文件（credentials.example.json等）
- ❌ 不包含敏感文件（credentials.local.json已被排除）

---

## 四、推送命令（完整版）

如果你想一次性执行所有命令：

```bash
# 进入SKYE目录
cd D:/MyScripts/02-data-processing/skye-bill-check

# 推送到GitHub
git push -u origin main

# 如果推送失败，可能需要强制推送（仅首次）
# git push -u origin main --force
```

---

## 五、我为你做了什么（总结）

### ✅ Git配置
- 初始化仓库
- 配置用户名：abin12343
- 配置邮箱：abin12343@users.noreply.github.com
- 设置分支名：main
- 添加远程仓库：https://github.com/abin12343/skye-bill-check.git

### ✅ 文件准备
- 创建.gitignore（排除敏感文件）
- 创建outputs/.gitkeep和logs/.gitkeep
- 添加所有源代码文件
- 创建初始提交

### ✅ 提交内容
包含的文件（15个）：
1. README.md - 完整文档
2. config.json - 配置文件
3. launcher.pyw - 图形界面
4. main_flow.py - 主流程
5. run_skye_check.py - 命令行入口
6. fetch_bill.py - 账单下载
7. process_master.py - 主表处理
8. process_details.py - 明细表处理
9. quote_engine.py - 报价引擎
10. requirements.txt - 依赖列表
11. credentials.example.json - 凭证模板
12. notify.example.json - 通知模板
13. .gitignore - Git忽略规则
14. outputs/.gitkeep - 输出目录占位
15. logs/.gitkeep - 日志目录占位

---

## 六、安全提示

⚠️ **已排除的敏感文件**（不会被上传）：
- credentials.local.json（天图账号）
- notify.local.json（企业微信webhook）
- __pycache__/（Python缓存）
- outputs/*（输出文件）
- logs/*（日志文件）

✅ **建议设置为Private仓库**，因为代码包含：
- 业务逻辑
- 报价算法
- 费用规则配置

---

## 七、下次更新代码

以后要更新代码时，执行：

```bash
cd D:/MyScripts/02-data-processing/skye-bill-check

# 查看修改的文件
git status

# 添加修改的文件
git add .

# 提交
git commit -m "更新说明"

# 推送
git push
```

---

**准备完成！现在你可以按照"步骤1"和"步骤2"完成GitHub仓库创建和推送。**
