# 保谊达做账执行工具

这是运输磅单做账工具项目。仓库历史中包含源码、公开执行规范、版本记录和发布包；后续自动同步以公开安全文档和公开索引为白名单，不使用 `git add .` 提交私有批次、密钥、安装包输出或本机完整手册。

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

本轮公开验证状态：Python 3.11 `py_compile`、`--self-test`、`--startup-smoke-test`、`--tray-menu-probe`、`--scanner-watch-probe` 已通过；安装器已在本机临时目录完成静默安装/卸载 smoke，确认桌面快捷方式创建、目标正确，卸载后快捷方式、安装目录、配置目录和卸载注册表清理干净。尚未在用户真实 VM 中亲手运行。

当前工作副本产物哈希：

- `做账执行工具.exe` SHA256：`2610906320A30BF951911E457297CF586C8D3277A89AD2FAB897C297AF6D112D`
- `项目资料\安装包\输出\BoyidaAccountingTool_Setup_v2.3.exe` SHA256：`C10FC1109A1456C7C73A823313155874D15AE80112BAB639FE533D907DDCD750`

## 公开安全复刻范围

公开 GitHub 文档能复刻：

- 仓库目录结构、公开源码入口、关键模块职责和验证命令。
- 公开可追踪 UI 图片资产：`项目资料\开发源码\assets\boyida_truck.png`、`boyida_truck.ico`、`jingzhe_header_line.png`。
- `做账执行工具.exe + _internal` onedir 运行形态、Inno 安装器构建/验证流程、默认当前用户桌面快捷方式和 `扫描导入` 文件夹。
- UI/交互合同：左侧导航、顶部主动作、中间流程与日志、右侧 GPS/OCR 控制；OCR 引擎和接口/模式已从窄下拉框改为按钮弹出选择面板。
- 扫描监听行为：批次目录发现、文件稳定等待、立即/稍后/忽略、成功后标记、失败后可重试。
- 卸载清理范围：安装目录、快捷方式、安装目录下扫描/批次目录、用户配置目录和卸载注册表项。
- 当前公开工作副本产物哈希。

公开 GitHub 文档不能复刻：

- 真实 API Key、Cookie、GPS Authorization、JSESSIONID、账号密码或客户真实批次。
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
- 不提交 GPS Authorization、Cookie、JSESSIONID、账号密码或客户原始批次数据。
- 本地私有执行稿 `做账执行规范.md` 只放在当前电脑使用；GitHub 使用 `做账执行规范_公开版.md`。
- 发布给别人使用时，通过 GitHub Release 下载对应版本的 EXE。

## 当前版本

当前正式版本：`v2.3`

下载入口：[GitHub Releases](https://github.com/danhuangpai520/boyida-accounting-tool/releases)
