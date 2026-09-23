<div align="center">

# WinPrintDiag

**Windows 打印子系统便携体检工具 · Portable diagnostic tool for the Windows printing subsystem**

双击即用 · 免安装 · 默认只读不碰系统

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011-lightgrey.svg)](#)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](#)

[**中文文档**](#中文文档) · [**English**](#english) · [**🌐 宣传页 Landing Page**](https://dc1024.github.io/winprintdiag/)

</div>

---

## 中文文档

### 它是干什么的

针对一类很烦人的故障：**打印后台处理程序（Print Spooler / `spoolsv.exe`）反复自动停止** —— 重启服务后过一会儿又挂；打印任务卡在队列里删不掉；打印机时好时坏，事件日志翻半天也看不出原因。

WinPrintDiag 把散落在**事件日志、注册表、队列目录**里的线索一次性收齐并给出结论。

关键在于：**它自身不依赖打印服务**。打印机、端口、驱动信息直接读注册表，所以服务已经挂掉、队列已经卡死的时候，它照样能跑。

### 它能查什么

| 检测项 | 说明 |
|---|---|
| **崩溃历史** | 聚合应用程序日志里的 `spoolsv` 崩溃（事件 1000），按日统计 + **出错模块统计** —— 一眼看出是不是 `usbmon.dll`、`localspl.dll` 这类组件在崩 |
| **崩溃循环** | 近 30 分钟崩溃次数，≥ 3 次判定为崩溃循环 |
| **服务状态** | Spooler 服务状态 / 启动类型，以及事件日志能回溯多久（决定证据窗口） |
| **组件完好性** | 8 个打印核心文件（`spoolsv.exe` / `usbmon.dll` / `localspl.dll` / `win32spl.dll` …）的体积、版本、修改时间、**数字签名**。签名异常直接判定为被第三方替换 |
| **组件存储比对** | 签名异常时列出 `WinSxS` 里的正版候选文件，供还原使用 |
| **打印机配对** | 每台打印机的 **连接方式**（USB 直连 / 网络 WSD / 网络 IPP / 网络 IP / 虚拟）、端口、驱动；检测端口与驱动错配 |
| **重复条目** | 检测**同一台物理打印机被注册成多个条目** —— 最常见的原因就是 USB 和无线同时连着。`USB + 网络` 组合判定为高危 |
| **打印队列** | 队列文件数、合计体积、最新 / 最早任务时间；解码 `.SHD` 描述文件里的任务内容 |
| **变更相关性** | 近 45 天补丁与软件安装、近 30 天意外关机、近 7 天开机次数、蓝屏转储数量 |
| **审计日志** | 打印操作日志（`PrintService/Operational`）是否开启 |

### 快速开始

三种用法任选，全部免安装：

**① 图形界面（推荐）** —— 双击 `WinPrintDiagUI.exe`

**② 命令行** —— 双击 `WinPrintDiag.cmd`，体检完自动用记事本打开报告

**③ 直接跑脚本**（连 exe 都不需要）

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\WinPrintDiag.ps1
```

报告默认输出到工具所在目录，文件名带时间戳。

### 报告长什么样

```text
===== [2] 打印服务状态 =====
服务状态   : 运行中     启动类型: 自动
进程状态   : 运行中，实例数 1

===== [3] 崩溃历史 =====
spoolsv 崩溃事件（应用程序日志 1000）: 27
最早       : 2026-08-25 10:10:15
最近       : 2026-09-23 01:29:03
-- 出错模块统计 --
  usbmon.dll  x20
  localspl.dll  x1

===== [5] 打印组件文件完好性 =====
文件              大小      版本               修改时间          签名
spoolsv.exe       991232    10.0.26100.8875    2026-09-22 00:11  有效
usbmon.dll        1351680   10.0.26100.8875    2026-09-22 00:11  有效

===== [6] 打印机 / 端口 / 驱动配对 =====
打印机                                  端口                         连接方式    驱动
Adobe PDF                               Documents\*.pdf              虚拟/本地   Adobe PDF Converter
Pantum M6200NW Series 0001              IPP_Pantum-13CEDD_1          网络 IPP    Pantum M6200NW Series
Pantum-13CEDD (M6200NW series)          WSD-c0c4001a-e93a-4cd5-8f4f~ 网络 WSD    Microsoft IPP Class Driver

===== [10] 结论与建议 =====
发现以下问题:
  [警告] 同一台打印机注册了多个条目: Pantum M6200NW Series 0001 [网络 IPP] 与
         Pantum-13CEDD (M6200NW series) [网络 WSD]  (共有标识: m6200nw/13cedd)
```

发现项按 `[严重]` / `[警告]` / `[提示]` 三档分级，图形界面顶部结论条会按最严重的一档变色。

### 修复能力（可选，需管理员）

**默认全程只读，不改动系统任何东西。** 需要工具动手时才加开关：

| 命令 | 作用 |
|---|---|
| `WinPrintDiag.exe -Repair` | 从组件存储还原签名异常的打印组件；**原文件先隔离备份、不直接删除**；顺带开启打印操作日志 |
| `WinPrintDiag.exe -ClearQueue` | 清理卡死的打印队列；**整个队列目录先备份到带时间戳的文件夹**，不做删除 |
| `WinPrintDiag.exe -OutDir <目录>` | 指定报告输出目录 |
| `WinPrintDiag.exe -OpenReport` | 结束后用记事本打开报告 |

图形界面里对应「修复（需管理员）」和「清理队列」两个按钮，会自动请求提权。

### 设计取舍：宁可漏报，不可误报

判定「同一台打印机是否注册了多个条目」时，工具从**打印机名 + 端口名**提词，只认**长度 ≥ 5 且同时含字母和数字**的共有词（如 `13cedd`、`m6200nw`、`l2350dw`）。

这样 `pdf`、`series`、`0001` 这类通用词或纯数字，不会把 Adobe PDF、PDF-XChange、PDF24、导出为 WPS PDF 这些虚拟打印机凑成一对。

代价是 `HP LaserJet 1020 (Copy 1)` 与 `(Copy 2)` 这类同型号重复副本会漏报 —— 这是**刻意的**：诊断工具误报比漏报更伤信任。规则有 19 条回归测试兜底。

### 自测

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\token_rules_test.ps1
```

19 条用例，覆盖连接方式识别、应告警场景（USB + 网络双条目）、以及 6 种不应误报的场景。测试直接从主脚本里**抽取函数原文**来跑，所以不会出现「改了主脚本、忘了改测试」的情况。

### 常见问题

**Q：会不会误删我的打印机或文件？**
不会。诊断阶段纯只读。`-ClearQueue` 是**移动**队列文件到备份目录，`-Repair` 是把可疑组件**隔离**而非删除，两者都保留原件。

**Q：需要联网吗？需要装什么吗？**
都不需要。单文件、零依赖，只用 Windows 自带的 PowerShell 5.1。

**Q：图形界面打不开怎么办？**
运行 `WinPrintDiag-GUI.cmd`，它会在 exe 启动失败时自动回落到脚本方式。

**Q：我想改脚本，改了以后中文变乱码了？**
`.ps1` 必须保存为 **UTF-8 带 BOM**。PowerShell 5.1 读没有 BOM 的脚本会按 ANSI 解码，中文字面量会损坏，而且报错会伪装成语法错误（`意外的标记` / `字符串缺少终止符`）—— 遇到这种"莫名语法错误"，先检查 BOM。

### 已知限制

- 仅支持 Windows 10 / 11 + PowerShell 5.1（系统自带）
- 部分机器不记录服务启停事件（7036），「最近状态」会显示不可用，此时崩溃循环判定改由崩溃次数承担
- 图形界面版是打包产物，无法在本项目作者的自动化环境里做运行时验证，需要手动双击确认

### 许可

[MIT](LICENSE)

---

## English

### What it is

A portable, no-install diagnostic tool for a very annoying class of failures: **the Windows Print Spooler (`spoolsv.exe`) keeps stopping on its own**. The service restarts fine, then dies again a while later. Print jobs get stuck in the queue and refuse to delete. The printer works intermittently, and the Event Log gives you nothing obvious.

WinPrintDiag gathers the clues scattered across the **Event Log, the registry and the spool queue directory**, and turns them into a single verdict.

The important part: **it does not depend on the print service.** Printer, port and driver information is read straight from the registry, so it still works when the Spooler is already down and the queue is jammed.

### What it checks

| Check | Details |
|---|---|
| **Crash history** | Aggregates `spoolsv` crashes from Application log (Event ID 1000), per-day counts, and **faulting-module statistics** — so you can see at a glance whether `usbmon.dll` or `localspl.dll` is the culprit |
| **Crash loop** | Crash count within the last 30 minutes; ≥ 3 is flagged as a crash loop |
| **Service state** | Spooler status / startup type, plus how far back the Event Log actually reaches (the evidence window) |
| **Binary integrity** | Size, version, modified time and **Authenticode signature** of 8 print-stack binaries (`spoolsv.exe`, `usbmon.dll`, `localspl.dll`, `win32spl.dll`, …). An invalid signature is flagged as a third-party replacement |
| **Component store diff** | When a signature is invalid, lists the genuine candidates in `WinSxS` for restoration |
| **Printer pairing** | Per-printer **connection type** (USB / network WSD / network IPP / network IP / virtual), port and driver; detects port-driver mismatches |
| **Duplicate entries** | Detects the **same physical printer registered as multiple entries** — most often because USB and Wi-Fi are connected at the same time. A `USB + network` combination is flagged critical |
| **Print queue** | Queue file count, total size, newest / oldest job timestamps; decodes job descriptions from `.SHD` files |
| **Change correlation** | Updates and software installed in the last 45 days, unexpected shutdowns in 30 days, boot count in 7 days, crash dump count |
| **Audit log** | Whether the `PrintService/Operational` log is enabled |

### Quick start

Three ways, all portable and install-free:

**1. GUI (recommended)** — double-click `WinPrintDiagUI.exe`

**2. Command line** — double-click `WinPrintDiag.cmd`; the report opens in Notepad when done

**3. Run the script directly** (no exe needed)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\WinPrintDiag.ps1
```

The report is written next to the tool with a timestamp in the filename.

### Sample output

```text
===== [3] 崩溃历史 / crash history =====
spoolsv 崩溃事件（应用程序日志 1000）: 27
-- 出错模块统计 / faulting modules --
  usbmon.dll  x20
  localspl.dll  x1

===== [6] 打印机 / 端口 / 驱动配对 =====
打印机 / printer       端口 / port                  连接方式 / link   驱动 / driver
Pantum M6200NW 0001    IPP_Pantum-13CEDD_1          网络 IPP          Pantum M6200NW Series
Pantum-13CEDD          WSD-c0c4001a-e93a-...        网络 WSD          Microsoft IPP Class Driver

===== [10] 结论与建议 / verdict =====
发现以下问题 / issues found:
  [警告] 同一台打印机注册了多个条目: ...  (共有标识: m6200nw/13cedd)
```

Findings are graded `[严重]` critical / `[警告]` warning / `[提示]` info, and the GUI summary bar takes the most severe color.

### Optional repair (admin required)

**By default the tool is entirely read-only.** Switches are opt-in:

| Command | Effect |
|---|---|
| `WinPrintDiag.exe -Repair` | Restores print-stack binaries with invalid signatures from the component store. **The original file is quarantined, never deleted.** Also enables the print audit log |
| `WinPrintDiag.exe -ClearQueue` | Clears a jammed print queue. **The whole queue directory is backed up to a timestamped folder first** |
| `WinPrintDiag.exe -OutDir <dir>` | Choose the report output directory |
| `WinPrintDiag.exe -OpenReport` | Open the report in Notepad when finished |

### Design trade-off: prefer a miss over a false alarm

To decide whether two printer records are the same physical device, the tool extracts tokens from the **printer name + port name** and only accepts shared tokens that are **at least 5 chars long and contain both letters and digits** (`13cedd`, `m6200nw`, `l2350dw`).

This keeps generic words and pure numbers (`pdf`, `series`, `0001`) from pairing up unrelated virtual printers such as Adobe PDF, PDF-XChange and PDF24.

The cost: same-model duplicates like `HP LaserJet 1020 (Copy 1)` and `(Copy 2)` are missed. That is deliberate — a diagnostic tool loses credibility faster from false alarms than from misses. The rules are covered by 19 regression tests.

### Tests

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\token_rules_test.ps1
```

19 cases covering connection-type detection, cases that *must* warn (USB + network duplicates), and 6 cases that must *not* warn. The tests extract the real function bodies from the main script, so they can never drift into testing a stale copy.

### FAQ

**Will it delete my printers or files?**
No. Diagnosis is strictly read-only. `-ClearQueue` *moves* queue files to a backup folder, and `-Repair` *quarantines* the suspect binary instead of deleting it.

**Does it need internet access or any installation?**
Neither. Single file, zero dependencies, uses only the built-in PowerShell 5.1.

**The GUI won't open — what now?**
Run `WinPrintDiag-GUI.cmd`, which falls back to launching the PowerShell script when the exe fails.

**I edited the script and the Chinese text turned into garbage.**
`.ps1` files must be saved as **UTF-8 with BOM**. PowerShell 5.1 decodes BOM-less scripts as ANSI, which corrupts non-ASCII literals — and the resulting error *looks* like a syntax error (`unexpected token` / `string is missing the terminator`). If you hit a bizarre syntax error in a Chinese script, check the BOM first.

### Known limitations

- Windows 10 / 11 + PowerShell 5.1 only
- Some machines don't record service start/stop events (7036); "recent state" then shows as unavailable and crash-loop detection falls back to crash counts
- The packaged GUI binary could not be runtime-verified in the author's automated environment and should be confirmed by double-clicking

### License

[MIT](LICENSE)
