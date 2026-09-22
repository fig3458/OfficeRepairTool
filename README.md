# OfficeRepairTool

一键修复 Microsoft Word「上次启动失败 / 是否在安全模式中启动」问题的 Windows 小工具。

带中文 GUI，启动时自动申请管理员权限，按官方常见排障思路**重命名（备份）**相关注册表与缓存，而不是直接删除，方便出问题后还原。

---

## 它能解决什么问题？

Word 反复弹出类似提示：

> Word 上次启动失败。安全模式可以帮助您解决问题……是否要在安全模式中启动？

常见原因是：用户配置损坏、加载项异常、`Normal.dotm` / 模板缓存异常、Office 漫游与本地缓存不一致等。

本工具把手动排障步骤自动化：

1. 清理（重命名）`HKCU\Software\Microsoft\Office`
2. 清理（重命名）`%AppData%\Microsoft` 下的 `Office` / `Excel` / `Word` / `PowerPoint` / `Templates`
3. 清理（重命名）`%LocalAppData%\Microsoft` 下的 `Office` / `Excel`

全部改为 `*.old` 备份名；若已存在同名备份，会自动追加时间戳（如 `Office.old.20260922_143000`）。

---

## 功能一览

| 功能 | 说明 |
|------|------|
| 报错参考图 | 界面内嵌常见「安全模式」弹窗截图，方便对照 |
| 进程检测 | 显示仍在运行的 Word / Excel 等前台进程，以及 Click-to-Run 等后台进程 |
| 一键结束进程 | 结束 Office 前台 + 常见后台占用进程 |
| 暂停 Click-to-Run | 修复前尝试停止 `ClickToRunSvc`，降低「文件正在使用」概率；结束后再拉起 |
| 一键修复 | 按上面三步重命名注册表与缓存目录 |
| 占用提示 | 若重命名失败（正在使用 / 拒绝访问），弹窗列出疑似占用进程，可选择结束后重试 |
| 结果统计 | 成功 / 失败分项统计；有失败时不会假装「已修好」 |
| 还原备份 | 把 `*.old` 还原成原名（仅当原名不存在时），适合修坏后回退 |

---

## 使用方法

### 直接使用（推荐）

1. 运行 `dist\OfficeRepairTool.exe`（或自行用 `build.bat` 打包后的 exe）
2. 在 UAC 提示中允许管理员权限
3. **建议顺序：**
   1. 手动保存并关闭所有 Office 窗口  
   2. 点击「一键结束 Office 进程」  
   3. 点击「一键执行修复」  
   4. 看日志是否全部成功，再重新打开 Word  

### 如果修完后 Word 打不开

1. 打开本工具  
2. 点击「还原备份 (*.old)」  
3. 再按上面流程重新修复一次，并确认日志中无失败项  

---

## 修复时具体改了什么？

### 1. 注册表（当前用户）

- 路径：`HKEY_CURRENT_USER\Software\Microsoft\Office`
- 操作：重命名为 `Office.old`（或带时间戳的备份名）

### 2. Roaming 缓存

目录：`%AppData%\Microsoft\`（通常是 `C:\Users\<用户名>\AppData\Roaming\Microsoft\`）

重命名（若存在）：

- `Office`
- `Excel`
- `Word`
- `PowerPoint`
- `Templates`

### 3. Local 缓存

目录：`%LocalAppData%\Microsoft\`

重命名（若存在）：

- `Office`
- `Excel`

### 为什么要停 Click-to-Run？

`%LocalAppData%\Microsoft\Office` 经常被 `OfficeClickToRun` / `ClickToRunSvc` 占用。  
若在占用时强行半改成功、半改失败，可能导致 Word「修完反而打不开」。  
因此工具会：

- 结束更多关联进程（含后台）
- 尝试 `sc stop ClickToRunSvc`
- 修复结束后再 `sc start ClickToRunSvc`
- 单步失败时给出占用提示，而不是静默跳过并宣称成功

---

## 从源码运行 / 打包

### 环境

- Windows 10 / 11
- Python 3.10+（开发与打包）
- 依赖见 `requirements.txt`：`psutil`、`pyinstaller`

### 开发运行

```bat
pip install -r requirements.txt
python main.py
```

（非管理员时会自动弹 UAC 提权后重启。）

### 打包成 exe

双击或在项目根目录执行：

```bat
build.bat
```

生成文件：

```text
dist\OfficeRepairTool.exe
```

打包参数要点：

- `--onefile --windowed`：单文件 GUI
- `--uac-admin`：清单要求管理员
- 内嵌 `assets\word_safe_mode_error.png`
- 排除 `PIL` / `numpy` / `PySide6`，控制体积（约 10MB 量级）

---

## 项目结构

```text
OfficeRepairTool/
├── main.py                 # 主程序（tkinter GUI + 修复逻辑）
├── assets/
│   └── word_safe_mode_error.png   # 界面内报错参考图
├── requirements.txt
├── build.bat               # PyInstaller 一键打包
├── MainForm.cs             # 早期 WinForms 试验版（不推荐，见下方说明）
├── app.manifest            # C# 版 UAC 清单（遗留）
└── README.md
```

### 关于 `MainForm.cs`（遗留）

仓库里保留过一版极轻量的 C# WinForms 实现。实测在未处理好 Click-to-Run 占用时，容易把 Local 缓存改到「半残」状态，出现修完打不开的问题。  
**请以 `main.py` / 打包出的 Python 版 exe 为准。** C# 源码仅作参考，默认不要用它作为日常修复工具。

---

## 注意事项与风险

- 请先保存所有 Office 文档再结束进程。  
- 本工具只动**当前用户**的 HKCU 与用户目录下的缓存，不会卸载 Office。  
- 重命名后 Word 首次启动可能稍慢（会重建部分配置）。  
- 加载项、登录状态、部分个性化设置可能需要重新配置。  
- 若企业环境有策略锁定注册表/目录，部分步骤可能失败，请以日志为准。  
- 杀毒软件可能拦截对 Office 目录的重命名，必要时临时允许。  

---

## 故障排查

| 现象 | 建议 |
|------|------|
| 提示「正在使用」 | 点「一键结束 Office 进程」，关闭资源管理器中对该文件夹的窗口/预览，再修 |
| 修复有失败项 | 不要强行开 Word；处理占用后重试，或先「还原备份」 |
| 修完打不开 | 立刻用「还原备份」，确认还原成功后再修 |
| UAC 被取消 | 需允许管理员权限；或右键 exe → 以管理员身份运行 |
| 参考图不显示 | 确认 `assets\word_safe_mode_error.png` 存在；重新执行 `build.bat` |

---

## 许可证

按个人 / 内部运维工具使用即可。若二次分发，请自行评估对用户数据（配置重命名）的影响，并保留风险说明。

---

## 更新摘要

- GUI + 自动 UAC
- 注册表 / Roaming / Local 一键备份式清理
- 占用检测与关联进程提示
- Click-to-Run 服务停启，降低 Local 缓存损坏概率
- 一键还原 `*.old` 备份
