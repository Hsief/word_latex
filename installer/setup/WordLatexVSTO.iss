#ifndef MyAppVersion
  #define MyAppVersion "1.2.1"
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

// MSI failure must not be reported as a successful setup. Repair HKCU under the
// original desktop user, not whichever account supplied administrator credentials.
procedure CurStepChanged(CurStep: TSetupStep);
var
  ExitCode: Integer;
  Manifest: String;
  RepairPath: String;
begin
  if CurStep <> ssPostInstall then Exit;
  if LegacyUninstallerExists then
    if not Exec(GetLegacyUninstaller(''), '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '',
      SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
      RaiseException('旧版卸载失败，退出码：' + IntToStr(ExitCode));

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
  RepairPath := ExpandConstant('{pf64}\WordLatexVSTO\WordLatexVSTO_Repair.exe');
  if not FileExists(RepairPath) then
    RaiseException('安装缺少当前用户注册修复工具：' + RepairPath);
  if not ExecAsOriginalUser(RepairPath, '/quiet', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then
    RaiseException('无法为原始 Windows 用户设置 Word 自动加载，错误码：' + IntToStr(ExitCode));
  Log('Original-user registration repair exit code: ' + IntToStr(ExitCode));
  if ExitCode <> 0 then
    RaiseException('当前用户注册失败。请关闭 Word，运行 ' + RepairPath + ' 查看具体原因。');
end;
