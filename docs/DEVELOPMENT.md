# 开发笔记：Excel 加载宏的技术坑

本文面向贡献者，记录开发本项目时踩到并已解决的问题，避免后来者重蹈覆辙。

## 1. Ribbon 关系必须挂在包级 `_rels/.rels`

自定义 Ribbon 的三个要素必须成套：部件放 `customui/customui.xml`、XML 命名空间用 `http://schemas.microsoft.com/office/2006/01/customui`、关系挂在**包级** `_rels/.rels` 且类型为 `http://schemas.microsoft.com/office/2006/relationships/ui/extensibility`。

把关系写进 `xl/_rels/workbook.xml.rels`、或把部件命名成 `customUI14.xml` 配 2009 命名空间，都可能被 Excel **静默忽略**（不报错，选项卡就是不出现）。最稳妥的做法是照抄 Office 自带加载宏（SOLVER.XLAM / EUROTOOL.XLAM / FUNCRES.XLAM）的接法。

## 2. zip 条目名必须是正斜杠

OPC 规范要求包内条目名用 `/` 分隔。.NET 的 `ZipFile.CreateFromDirectory` 在 .NET Framework 下会写入反斜杠（`customui\customui.xml`），Excel 同样静默读不到 Ribbon。Python 的 `zipfile` 天然用正斜杠。构建脚本末尾必须加断言校验条目名。

## 3. `%APPDATA%\Microsoft\AddIns` 目录里的文件会被 Excel 自动识别

把旧版备份成 `公式助手.xlam.bak` 放在同一目录，Excel 照样会把它当加载项列出来。旧版备份请放到目录之外（如 `AddIns-backup\`）。

## 4. pywin32 后期绑定下不要用 `AddIns.Add` / `AddIns.Remove`

实测 `Add` 报「类 AddIns 的 Add 方法无效」，`Remove` 直接 `AttributeError`。加载项注册的本质是 `HKCU\Software\Microsoft\Office\<版本>\Excel\Options` 下的 `OPEN` / `OPEN1` / … 值（内容是带双引号的完整路径），直接读写注册表更可控。

另外两条相关经验：

- **Excel 运行时不要直写这些注册表键**：Excel 退出时会用内存中的加载项列表覆盖 Options，运行中的写入会被冲掉。
- **Excel 加载了加载项后经常无视 `Quit()`**，留下无窗口的僵尸进程，还会阻塞 Windows 关机。自动化脚本必须"Quit → 轮询本次启动的 PID → 超时按 PID 精确强杀"，绝不误伤用户自己的实例（见 `excelutil.quit_excel` / `excel-common.ps1` 的 `Stop-ExcelGracefully`）。

## 5. `Application.InputBox(Type:=8)` 框选会把活动单元格带走

两处叠加的坑：

- `Range.Address(False, False)` **永远不含工作表名**（要 `External:=True` 才有）。跨表框选后拿到 `A1:B10`，写公式时按目标表解析，就指到了目标表的同名区域。修法是统一走 `QualifiedAddress`。
- 用户切到别的表框选后，`ActiveSheet` / `ActiveCell` 已经漂走。凡是「界面打开时确定、之后还要用」的位置（公式写进哪个单元格），必须在界面打开那一刻存进模块级变量，之后全程用它。

回归测试见 `build/test_vlookup.py`（专门模拟框选后活动单元格漂移）。

## 6. 模块级声明必须写在所有过程之前

写在过程之间整个工程编译不过，症状极其隐蔽：Excel 只弹一个编译错误对话框，通过 COM 调用时脚本直接卡死（看不到对话框、没有输出）。`build.py` 的 `check_module_layout` 在注入源码前拦截此类错误。

同理，用 `Application.Run` 从外部驱动加载项时，**参数尽量只用字符串/数字**：传 Range 对象容易失败并卡死。需要区域的探针入口一律收地址字符串，在 VBA 侧自己 `Application.Range(...)` 转换。

## 7. 多格未合并区域的 `.MergeArea` 会直接抛 1004

单格与"已合并区域"的 `MergeArea` 都安全，唯独"多格但未合并"的范围读它必炸（`Range("A2:A3").MergeArea` → 运行时错误 1004，描述只有一句"应用程序定义或对象定义错误"）。合并功能的幂等判断最初写在多格范围的 `MergeArea` 上，导致每次合并都在这条语句上崩掉。

修法：读合并状态一律经由**单个单元格**（`Cells(r, c).MergeArea`），拿它的地址与目标范围比对。

## 8. COM 新建工作簿会带入本机 Excel 用户名

`Workbooks.Add` 出来的工作簿，`docProps/core.xml` 的 `dc:creator` / `cp:lastModifiedBy` 会自动填上本机 Office 的用户名。公开发布的二进制必须清洗：通过 `BuiltinDocumentProperties` 置空会被 Excel 静默忽略，可靠的做法是在打包阶段直接改写 zip 内的 `core.xml` / `app.xml`（见 `build.py` 的 `_scrub_doc_props`）。

## 9. 构建期需要临时开启 VBOM 访问

访问 `Workbook.VBProject` 要求信任中心开启「信任对 VBA 工程对象模型的访问」（`HKCU\...\Office\<版本>\Excel\Security\AccessVBOM`），否则报晦涩的 COM 错误。构建脚本应在构建前按本机所有 Office 版本临时写入、构建后还原原值（见 `build.py` 的 `VbomGuard`）。

## 10. 事件过程必须自带错误处理

VBA 错误如果在向导窗体的事件链（Initialize / Change / Click）里未被捕获，会以裸的运行时错误对话框形式弹出并阻塞自动化；通过 COM 调用时则表现为调用方卡死或收到晦涩的 com_error。每个事件过程都应有本地 `On Error` 处理，把错误转成界面提示或返回值。
