# 安装说明

## 普通安装

1. 保存文档并关闭所有 Microsoft Word 窗口。
2. 从项目 Releases 下载 `WordLatexVSTO_Setup.exe`。
3. 双击安装并同意 Windows 管理员确认；安装器会调用 MSI，默认安装到 `%ProgramFiles%\WordLatexVSTO`。
4. 打开 Word，在顶部确认出现“论文工具”。

社区构建使用临时自签名证书保护 VSTO 清单完整性，并安装到 Microsoft VSTO 支持的 `Program Files` 受信任位置。安装器不会把社区自签名证书写入系统根证书库。请只安装从本项目 Release 下载且哈希符合发布页的文件；机构环境建议由管理员使用组织证书重新签名。

如果 Word 启动时显示“隐藏的模块中的编译错误”，错误来自 `.dotm`/`.dot` VBA 模板，而不是 C# VSTO 清单。请先关闭 Word，将 `%APPDATA%\Microsoft\Word\STARTUP` 中对应的旧模板移出该目录后再启动 Word；不要直接删除，便于需要时恢复。

## 系统要求

- Windows 10 或 Windows 11
- Microsoft 365 或 Office 2021 桌面版 Word（32 位和 64 位均可）
- .NET Framework 4.8
- Microsoft Visual Studio 2010 Tools for Office Runtime 4.0

受支持的 Microsoft 365 / Office 2021 安装通常包含 VSTO Runtime。如果 Word 提示无法加载 VSTO，请安装微软当前提供的 VSTO Runtime 后重试。

## 检查加载状态

在 Word 中打开：

`文件 → 选项 → 加载项 → 管理：COM 加载项 → 转到`

应看到并勾选 `WordLatexVSTO`。如果它位于“禁用的应用程序加载项”，请先在“管理：禁用项目”中启用，再返回 COM 加载项勾选。

## 升级

关闭 Word，直接运行新版 `WordLatexVSTO_Setup.exe`。MSI 会自动替换旧版本并保留当前用户设置。

## 卸载

在 Windows“设置 → 应用 → 已安装的应用”中卸载 `WordLatexVSTO Word 插件`。卸载会移除插件文件和 Word 加载项注册项；用户设置文件保留在 `%LocalAppData%\WordLatexVSTO\settings.xml`，可手动删除。

## 故障排查

### 没有“论文工具”选项卡

1. 确认 Word 已完全退出后再启动。
2. 按上面的“检查加载状态”启用加载项。
3. 在“信任中心 → 加载项”确认没有禁止所有应用程序加载项。
4. 确认 VSTO Runtime 与 .NET Framework 4.8 已安装。

### Word 将加载项设为禁用

Word 在加载项启动异常或启动过慢时可能自动禁用。先在“禁用项目”中恢复；如果再次发生，请到项目 Issues 提交 Office 版本、Word 位数和错误截图。

### 安装器提示 Word 正在运行

关闭所有 Word 窗口；若任务管理器中仍有 `WINWORD.EXE`，先保存工作并结束该进程，再重新安装。
