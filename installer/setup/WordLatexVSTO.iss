#ifndef MyAppVersion
  #define MyAppVersion "1.0.6"
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
Uninstallable=no
ArchitecturesAllowed=x64compatible

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\publish\WordLatexVSTO.msi"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
Filename: "{code:GetLegacyUninstaller}"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated; Check: LegacyUninstallerExists; StatusMsg: "正在移除旧版 WordLatexVSTO..."
Filename: "{sys}\msiexec.exe"; Parameters: "/i ""{tmp}\WordLatexVSTO.msi"" /qn /norestart"; Flags: runhidden waituntilterminated; StatusMsg: "正在安装 WordLatexVSTO..."

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
