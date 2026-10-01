# 安装说明

## 普通安装

1. 保存文档并关闭所有 Microsoft Word 窗口。
2. 从项目 Releases 下载 `WordLatexVSTO_Setup.exe`。
3. 双击安装；默认安装到当前用户的 `%LocalAppData%\Programs\WordLatexVSTO`，不需要管理员权限。
4. 打开 Word，在顶部确认出现“论文工具”。

首次安装的社区构建使用自签名 VSTO 清单，Windows 或 Office 可能显示“未知发布者”。请只安装从本项目 Release 下载且哈希符合发布页的文件。机构环境建议由管理员使用组织证书重新签名。

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

关闭 Word，直接运行新版 `WordLatexVSTO_Setup.exe`。安装器使用固定产品 ID，会覆盖程序文件并保留当前用户设置。

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
