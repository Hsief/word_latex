# WordLatexVSTO

WordLatexVSTO 是面向 Windows Microsoft Word 的 VSTO 插件。它把文档中的 LaTeX 公式转换为 Word 内置的可编辑公式（OMML / OMath），不会生成图片，也不依赖 MathType。

## 功能

- 在“论文工具”Ribbon 中转换当前公式、当前段落或全文。
- 识别 `$...$`、`\(...\)`、`$$...$$` 和 `\[...\]`。
- 生成真正的 Word `OMath`，可继续在 Word 公式编辑器中修改。
- 自动模式：键入结束符（空格或常见标点）后转换刚输入的公式。
- 公式编号：使用 Word `SEQ` 域自动编号；使用 `REF` 域插入可更新的引用。
- 科研排版：粗体向量、粗体希腊字母、矩阵、上下标、算子，以及 `SO(3)`、`SE(3)`、`FMCW`、`LiDAR`、`IMU` 的正体处理。
- 批量转换从文档后部向前执行；单个公式失败时保留 LaTeX 原文并继续。

## 支持环境

- Windows 10/11
- Microsoft 365 或 Office 2021 桌面版 Word
- .NET Framework 4.8
- VSTO Runtime（受支持的 Office 安装通常已包含）
- 开发：Visual Studio 2022，“Office/SharePoint 开发”工作负载

## 安装

从 [Releases](https://github.com/Hsief/word_latex/releases) 下载 `WordLatexVSTO_Setup.exe`，关闭 Word 后双击安装。重新打开 Word，顶部应出现“论文工具”。详见 [安装说明](docs/INSTALL.md)。

## 快速使用

在 Word 输入：

```latex
$x_k=f(x_{k-1})$
```

选中公式并单击“转换当前公式”，或开启“自动模式”后在公式末尾输入空格。结果是 Word 原生公式对象，可双击继续编辑。

更多示例、编号和引用方法见 [用户指南](docs/USER_GUIDE.md)。

## 支持的 LaTeX 子集

当前重点支持科研论文常用结构：

- `\frac{a}{b}`、`x_i`、`x^2`、`\sqrt{x}`、`\sqrt[3]{x}`
- 常用希腊字母、关系符号、箭头、集合符号
- `\sum`、`\int`、`\prod` 及其上下限
- `\hat`、`\bar`、`\tilde`、`\vec`、`\dot`、`\ddot`
- `matrix`、`pmatrix`、`bmatrix`、`cases`、`aligned`
- `\mathbf`、`\boldsymbol`、`\mathrm`、`\mathit`、`\operatorname`、`\mathbb`
- `\left` / `\right` 与常见定界符

这不是完整 TeX 引擎。自定义宏、宏包、复杂对齐和部分高级排版命令会保留为原文，并在转换摘要中报告。

## 项目结构

```text
WordLatexVSTO/
├── src/
│   ├── WordLatexAddin/       # VSTO、Ribbon、Word 事件、编号与设置
│   └── AfterMathCore/        # C# LaTeX 解析器与 OMML 生成器
├── tests/AfterMathCore.Tests/
├── installer/
│   ├── setup/                # WiX MSI 与 Inno Setup 启动器
│   └── publish/              # CI 生成的 MSI/EXE 发布文件
├── docs/
├── scripts/
└── WordLatexVSTO.sln
```

## 构建

1. 安装 Visual Studio 2022，选择“.NET 桌面开发”和“Office/SharePoint 开发”。
2. 安装 WiX Toolset v3 和 Inno Setup 6。
3. 打开 `WordLatexVSTO.sln`，选择 `Release | Any CPU` 并生成。
4. 执行 `scripts\Build-Release.ps1 -Version 1.0.6`，安装包输出到 `installer\publish\WordLatexVSTO_Setup.exe`。

核心测试可单独执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Test-Core.ps1
```

## 自动发布

`.github/workflows/build.yml` 在 `windows-2022` 上执行以下流程：

1. 编译完整 Visual Studio Solution。
2. 运行转换器和扫描器测试。
3. 生成已签名的 VSTO 部署清单。
4. 使用 WiX 生成真正的 Windows Installer（MSI），再由 Inno Setup 生成 `WordLatexVSTO_Setup.exe`。
5. 上传 Actions artifact；主分支构建同时创建或更新与 `PRODUCT_VERSION` 对应的 Release。

CI 使用临时自签名证书保护 VSTO 清单完整性。安装器以管理员权限安装到 `Program Files` 的 VSTO 受信任位置，不会向系统根证书库写入社区自签名证书。正式企业分发时，建议通过仓库机密提供组织的代码签名证书。

## AfterMath 集成

`AfterMathCore` 是对开源 [AfterMath](https://github.com/axobase001/aftermath) LaTeX→OMML 核心思路与解析树结构的 C# 移植和扩展。插件不启动 Python、不修改 `.docx` 压缩包，而是把生成的 OMML 直接插入当前 Word 文档。许可和来源见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

## 许可

MIT，见 [LICENSE](LICENSE)。
