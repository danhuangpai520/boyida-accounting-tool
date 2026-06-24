# 保谊达做账执行工具

这是运输磅单做账工具项目。仓库历史中包含源码、公开执行规范、版本记录和发布包；后续自动同步以公开安全文档和公开索引为白名单，不使用 `git add .` 提交私有批次、密钥、安装包输出或本机完整手册。

## 本机执行边界

- 原项目 `C:\Users\Administrator\Desktop\zhipu` 只读，仅作历史参考；不得在那里打包、写配置、改源码、覆盖 EXE 或移动批次数据。
- 当前实际工作副本是 `C:\Users\Administrator\Desktop\token\zhipu_workcopy`；后续复刻、文档更新、验证命令和 GitHub 同步都先回到这个目录执行。
- GitHub 只同步公开安全文档：`README.md`、`做账执行规范_公开版.md`、`项目复刻公开版.md`、`项目源码索引_公开版.md`、`版本历史.md`、`项目资料\说明与发布\使用说明.md`。
- 本机完整复刻手册 `项目资料\项目复刻手册\完整复刻手册.md` 只在本机保留；可更新，但不得提交到公开 GitHub。
- 每次功能、UI、OCR/API、规则、Excel、打包、安装、扫描监听或发布流程变化，都必须同步更新上面的公开文档；没有同步文档和验证状态的功能变更，不得标记为可交付。

## 文档入口

- [做账执行规范（GitHub 公开版）](做账执行规范_公开版.md)
- [项目复刻公开版](项目复刻公开版.md)
- [项目源码索引公开版](项目源码索引_公开版.md)
- [版本历史](版本历史.md)

做账路线规则必须在软件“规则管理”中明文可见；一键出表先执行内置做账规则，再应用人工纠错/路线覆盖规则，人工覆盖会在备注中留痕。

当前主界面以左侧“今日批次状态”和“出表设置”为入口：一键出表默认使用已保存的出表设置，不再每次弹列确认；出表完成后进入交付工作台，集中查看待复核、缺重量、路线待核、重复图片和最终文件。

疑问复核支持连续处理、仍待核标记、原图点击打开和键盘操作；规则管理会显示规则来源、数量、冲突处理方式，以及每行实际命中的内置/人工规则。

当前稳定版发布形态是根目录 `做账执行工具.exe` 加同级 `_internal` 运行库目录，避免 PyInstaller onefile 首次自解包时出现 `CreateProcessW: 拒绝访问`。发给别人时应压缩整个稳定包，不要只发单独 EXE。

当前工作副本安装包方案使用 Inno Setup，把稳定版 onedir 安装到 `%LocalAppData%\Programs\BoyidaAccountingTool`，并创建 `扫描导入` 文件夹。扫描仪或人工保存图片到 `扫描导入\日期批次` 或 `扫描导入\日期\批次` 后，软件会检测新批次并询问是否立即一键出表。

2026-06-24 工作副本状态：安装器已改为默认创建当前用户桌面快捷方式 `{userdesktop}\保谊达车队做账工具`，目标为 `{app}\做账执行工具.exe`；主界面已做小屏/VM 自适应，侧栏、左右栏、顶栏和流程 pipeline canvas 会按可用高度压缩，避免按钮或流程被裁切；OCR 快速控制已从窄 `ttk.Combobox` 改为点击按钮弹出选择面板，并保留旧 adapter 数据流兼容。包体大小本轮按要求暂不处理。

本轮公开验证状态：Python 3.11 `py_compile`、`--self-test`、`--startup-smoke-test`、`--tray-menu-probe`、`--scanner-watch-probe` 已通过；`--qwen-live-probe` 已接入，当前机器因未配置本机安全凭据返回 `QWEN_LIVE_PROBE_MISSING_KEY`；安装器已在本机临时目录完成静默安装/卸载 smoke，确认桌面快捷方式创建、目标正确，卸载后快捷方式、安装目录、配置目录和卸载注册表清理干净。尚未在用户真实 VM 中亲手运行。

当前工作副本产物哈希：

- `做账执行工具.exe` SHA256：`2610906320A30BF951911E457297CF586C8D3277A89AD2FAB897C297AF6D112D`
- `项目资料\安装包\输出\BoyidaAccountingTool_Setup_v2.3.exe` SHA256：`C10FC1109A1456C7C73A823313155874D15AE80112BAB639FE533D907DDCD750`

## 公开安全复刻范围

公开 GitHub 文档能复刻：

- 仓库目录结构、公开源码入口、关键模块职责和验证命令。
- 公开可追踪 UI 图片资产：`项目资料\开发源码\assets\boyida_truck.png`、`boyida_truck.ico`、`jingzhe_header_line.png`。
- `做账执行工具.exe + _internal` onedir 运行形态、Inno 安装器构建/验证流程、默认当前用户桌面快捷方式和 `扫描导入` 文件夹。
- UI/交互合同：左侧导航、顶部主动作、中间流程与日志、右侧 GPS/OCR 控制；OCR 引擎和接口/模式已从窄下拉框改为按钮弹出选择面板。
- OCR/API 安全边界：千问视觉 OCR（百炼 OpenAI 兼容接口）已作为独立在线 OCR 引擎接入源码，默认 / 优先模型为官方 OCR 专用 `qwen-vl-ocr`，同时可选 `qwen3-vl-plus` 和兼容旧 `qwen3.7-plus`；模型名、`workspace_id`、`region` 和 endpoint/base URL 可通过本机安全配置覆盖，配置 `workspace_id` 时会按百炼官方工作空间地址自动生成 endpoint；软件内 `OCR 设置 -> 千问配置` 可保存本机 DPAPI 配置并直接测试，`OCR 设置 -> 测试千问` 可立刻跑真实探针；`--qwen-preflight` 可先做无联网、无泄密的本机凭据/模型/endpoint 预检，`--qwen-key-template` 只生成本机私有模板，`--qwen-configure-key` 弹窗保存本机 DPAPI 配置后立即探针，`--qwen-import-key-from-clipboard` 只从当前剪贴板提取 Key 形态并隐藏保存后立即探针，`--qwen-live-probe` 只做本机安全凭据探针；公开文档只记录安全配置口径，不记录任何真实密钥。
- 扫描监听行为：批次目录发现、文件稳定等待、立即/稍后/忽略、成功后标记、失败后可重试。
- 卸载清理范围：安装目录、快捷方式、安装目录下扫描/批次目录、用户配置目录和卸载注册表项。
- 当前公开工作副本产物哈希。

公开 GitHub 文档不能复刻：

- 真实 API Key、Cookie、GPS Authorization、JSESSIONID、账号密码或客户真实批次。
- 真实千问 / 百炼 / DashScope 密钥；公开文档只允许列出 `DASHSCOPE_API_KEY`、`BAILIAN_API_KEY`、`QWEN_API_KEY` 这些环境变量名，以及 `qwen.txt` / `qwen.csv` 这类本机私有导入文件名，不允许填值。
- 客户原图、私有 OCR JSON、私有出表结果、私有完整执行稿和本机绝对路径细节。
- 真实 GPS 后端授权结果；公开版只复刻 GPS 配置入口、保存/清理方式和无授权时的安全行为。
- 当前本机完整打包所用的本地 vendor SDK 目录和 PyInstaller hook 目录未作为公开文档同步对象；公开克隆若缺这些材料，需要按源码依赖重新安装/生成等价构建材料后再打包。

新电脑复刻时只应在本机软件配置、环境变量或私下授权文件中填入凭据，不把凭据写入 Markdown、命令行、JSON 样例、Excel 或 Git 历史。

## 冷启动复刻核对表

| 步骤 | 所需文件 | 验证命令/动作 | 通过标志 | 公开边界 |
|---|---|---|---|---|
| 1. 读公开入口 | `README.md`、`做账执行规范_公开版.md`、`项目复刻公开版.md`、`项目源码索引_公开版.md`、`版本历史.md` | 人工确认版本、边界和源码入口一致 | 能定位 `项目资料\开发源码\zhipu_accounting_app.py`、安装包脚本和验证命令 | 不需要私有规范或客户批次 |
| 2. 检查源码入口 | `项目资料\开发源码` | `Test-Path .\项目资料\开发源码\zhipu_accounting_app.py` | 返回 `True` | 只验证公开源码结构 |
| 3. 跑基础验证 | 主程序源码 | `python -m py_compile .\项目资料\开发源码\zhipu_accounting_app.py`；`python .\项目资料\开发源码\zhipu_accounting_app.py --self-test`；`python .\项目资料\开发源码\zhipu_accounting_app.py --startup-smoke-test` | 退出码 0，且启动冒烟输出 `STARTUP_SMOKE_OK` | 不验证真实在线 OCR 准确率 |
| 4. 跑交互探针 | 主程序源码 | `python .\项目资料\开发源码\zhipu_accounting_app.py --tray-menu-probe`；`python .\项目资料\开发源码\zhipu_accounting_app.py --scanner-watch-probe` | `TRAY_MENU_PROBE_OK`、`SCANNER_WATCH_PROBE_OK` | 不导入客户原图 |
| 4a. 生成千问 Key 模板 | 主程序源码 | `python .\项目资料\开发源码\zhipu_accounting_app.py --qwen-key-template` | 输出 `QWEN_KEY_TEMPLATE_READY`，只在本机私有 `API KEY` 文件夹生成 `qwen.example.txt` | 复制为 `qwen.txt` 或改用 `qwen.csv` 时由用户本机填写；不提交、不打印真实 Key |
| 4b. 本机配置千问 Key | 主程序源码 + 用户本机授权 | 软件内点 `OCR 设置 -> 千问配置`，或运行 `--qwen-configure-key`；也可先复制 Key/workspace 配置后运行 `--qwen-import-key-from-clipboard` | 保存后立即执行真实探针；成功输出 `QWEN_LIVE_PROBE_OK` | 弹窗/剪贴板入口不打印 Key，不把 Key 写入源码、文档、命令行或 Git |
| 4c. 千问调用预检 | 主程序源码 + 本机安全凭据 | `python .\项目资料\开发源码\zhipu_accounting_app.py --qwen-preflight` | 有 Key 时输出 `QWEN_PREFLIGHT_READY`；无 Key 时输出 `QWEN_PREFLIGHT_MISSING_KEY` | 不联网、不打印 Key，用于先确认模型和 endpoint 候选 |
| 4d. 跑千问在线探针 | 主程序源码 + 本机安全凭据 | `python .\项目资料\开发源码\zhipu_accounting_app.py --qwen-live-probe` | 有 Key 时输出 `QWEN_LIVE_PROBE_OK`；无 Key 时输出 `QWEN_LIVE_PROBE_MISSING_KEY` | 不在命令行、源码、文档或 Git 中写真实 Key |
| 5. 验证安装器 | `做账执行工具.exe`、`_internal`、安装包脚本 | `powershell -NoProfile -ExecutionPolicy Bypass -File .\项目资料\安装包\构建安装包.ps1 -SkipAppBuild`，再做临时目录安装/卸载 smoke | 桌面快捷方式目标正确；卸载后快捷方式、安装目录、配置目录和卸载注册表清理干净 | 不改变安装脚本里的安全清理边界 |
| 6. 校验产物 | 根目录 EXE、安装器输出 | `Get-FileHash .\做账执行工具.exe -Algorithm SHA256`；`Get-FileHash .\项目资料\安装包\输出\BoyidaAccountingTool_Setup_v2.3.exe -Algorithm SHA256` | 分别等于本文档列出的两个 SHA256 | 哈希只能证明当前公开产物一致 |
| 7. 安全扫文档 | 公开 Markdown | 运行下方敏感扫描命令 | 无命中 | 边界说明中的普通词 `Cookie`、`Authorization` 不等于泄密 |

敏感扫描命令：

```powershell
$secretPattern = 'AKIA[0-9A-Z]{16}|sk-' + '[A-Za-z0-9_-]{20,}|JSESSIONID' + '=[^;\s]+|Authorization' + ':\s*Bearer\s+\S+|Cookie' + ':\s*\S+='
Select-String -Path .\README.md,.\做账执行规范_公开版.md,.\项目复刻公开版.md,.\项目源码索引_公开版.md,.\版本历史.md -Pattern $secretPattern
```

## 隐私规则

- 不在 GitHub 文档、源码或配置样例里提交真实智谱、阿里云、百度、腾讯云 API Key。
- 不在 GitHub 文档、源码或配置样例里提交真实千问 / 百炼 / DashScope API Key；只允许记录环境变量名、本机安全配置规则和 `qwen.txt` / `qwen.csv` 私有导入文件名。
- 不提交 GPS Authorization、Cookie、JSESSIONID、账号密码或客户原始批次数据。
- 本地私有执行稿 `做账执行规范.md` 只放在当前电脑使用；GitHub 使用 `做账执行规范_公开版.md`。
- 发布给别人使用时，通过 GitHub Release 下载对应版本的 EXE。

## 当前版本

当前正式版本：`v2.3`

下载入口：[GitHub Releases](https://github.com/danhuangpai520/boyida-accounting-tool/releases)
