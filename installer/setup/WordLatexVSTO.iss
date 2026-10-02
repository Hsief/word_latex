#ifndef MyAppVersion
  #define MyAppVersion "1.1.0"
#endif

#define MyAppName "WordLatexVSTO"
#define MyPublisher "WordLatexVSTO contributors"
#define LegacyAppId "{BFEF1444-9C43-4FF2-B458-F45AFD58D4DC}_is1"

[Setup]
AppId={{0C2A0E62-EB9D-4CDD-9686-065D370F948A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyPublisher}
AppPublisherURL=https://github.com/Hsief/word_latex
AppSupportURL=https://github.com/Hsief/word_latex/issues
CreateAppDir=no
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=..\publish
OutputBaseFilename=WordLatexVSTO_Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
SetupLogging=yes
Uninstallable=no
ArchitecturesAllowed=x64compatible

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\publish\WordLatexVSTO.msi"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Code]
function GetLegacyUninstaller(Param: String): String;
var
  UninstallKey: String;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#LegacyAppId}';
  Result := '';
  if IsWin64 then
    RegQueryStringValue(HKLM64, UninstallKey, 'UninstallString', Result);
  if Result = '' then
    RegQueryStringValue(HKLM32, UninstallKey, 'UninstallString', Result);
  Result := RemoveQuotes(Result);
end;

function LegacyUninstallerExists(): Boolean;
begin
  Result := FileExists(GetLegacyUninstaller(''));
end;

function InitializeSetup(): Boolean;
var
  Release: Cardinal;
begin
  Result := True;
  if not RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) or (Release < 528040) then
  begin
    MsgBox('WordLatexVSTO 需要 .NET Framework 4.8。请先通过 Windows Update 安装后再继续。', mbError, MB_OK);
    Result := False;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if FindWindowByClassName('OpusApp') <> 0 then
    Result := 'Microsoft Word 正在运行。请保存文档并关闭所有 Word 窗口，然后重新单击“安装”。';
end;

function MsiEnumRelatedProducts(UpgradeCode: String; Reserved: Cardinal; Index: Cardinal; ProductCode: String): Cardinal;
  external 'MsiEnumRelatedProductsW@msi.dll stdcall';

// Windows Installer costs files before removing the old product, so a lower version is silently skipped.
// Uninstall every installed WordLatexVSTO product first and install fresh.
procedure UninstallRelatedProducts();
var
  ProductCode: String;
  ExitCode, Guard: Integer;
begin
  for Guard := 1 to 10 do
  begin
    ProductCode := '';
    SetLength(ProductCode, 39);
    if MsiEnumRelatedProducts('{6F1A5592-9AC8-48CE-93F3-BB17E878E7FD}', 0, 0, ProductCode) <> 0 then Exit;
    ProductCode := Copy(ProductCode, 1, 38);
    if not Exec(ExpandConstant('{sys}\msiexec.exe'), '/x ' + ProductCode + ' /qn /norestart', '',
      SW_HIDE, ewWaitUntilTerminated, ExitCode) or ((ExitCode <> 0) and (ExitCode <> 3010)) then
      RaiseException('旧版 WordLatexVSTO 卸载失败，退出码：' + IntToStr(ExitCode));
  end;
end;

// MSI failure must not be reported as a successful setup.
procedure CurStepChanged(CurStep: TSetupStep);
var
  ExitCode: Integer;
  Manifest: String;
begin
  if CurStep <> ssPostInstall then Exit;
  if LegacyUninstallerExists then
    if not Exec(GetLegacyUninstaller(''), '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '',
      SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
      RaiseException('旧版卸载失败，退出码：' + IntToStr(ExitCode));

  UninstallRelatedProducts();

  WizardForm.StatusLabel.Caption := '正在安装并验证 WordLatexVSTO...';
  if not Exec(ExpandConstant('{sys}\msiexec.exe'),
    '/i "' + ExpandConstant('{tmp}\WordLatexVSTO.msi') + '" /qn /norestart /L*v "' +
    ExpandConstant('{tmp}\WordLatexVSTO-msi.log') + '"', '',
    SW_HIDE, ewWaitUntilTerminated, ExitCode) then
    RaiseException('无法启动 MSI 安装，错误码：' + IntToStr(ExitCode));
  Log('MSI exit code: ' + IntToStr(ExitCode));
  if (ExitCode <> 0) and (ExitCode <> 3010) then
    RaiseException('MSI 安装失败，退出码：' + IntToStr(ExitCode) + '。日志：' +
      ExpandConstant('{tmp}\WordLatexVSTO-msi.log'));

  if not RegQueryStringValue(HKLM64,
    'Software\Microsoft\Office\Word\Addins\WordLatexVSTOAddin', 'Manifest', Manifest) then
    RaiseException('MSI 完成但没有创建 Word 加载项注册项。');
  // A per-user Addins key overrides the machine-wide MSI registration and stops the VSTO Ribbon from loading.
  ExecAsOriginalUser(ExpandConstant('{sys}\reg.exe'), 'delete "HKCU\Software\Microsoft\Office\Word\Addins\WordLatexVSTOAddin" /f /reg:64', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  ExecAsOriginalUser(ExpandConstant('{sys}\reg.exe'), 'delete "HKCU\Software\Microsoft\Office\Word\Addins\WordLatexAddin" /f /reg:64', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
end;
