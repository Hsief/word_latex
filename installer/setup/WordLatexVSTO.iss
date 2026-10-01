#ifndef MyAppVersion
  #define MyAppVersion "1.0.3"
#endif

#define MyAppName "WordLatexVSTO"
#define MyPublisher "WordLatexVSTO contributors"
#define MyAppId "{{BFEF1444-9C43-4FF2-B458-F45AFD58D4DC}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyPublisher}
AppPublisherURL=https://github.com/Hsief/word_latex
AppSupportURL=https://github.com/Hsief/word_latex/issues
DefaultDirName={autopf}\WordLatexVSTO
DefaultGroupName=WordLatexVSTO
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=..\publish
OutputBaseFilename=WordLatexVSTO_Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=WordLatexVSTO Word 插件
CloseApplications=yes
RestartApplications=no
ArchitecturesAllowed=x86compatible x64compatible

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\publish\WordLatexAddin.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\publish\Application Files\*"; DestDir: "{app}\Application Files"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\docs\INSTALL.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\..\docs\USER_GUIDE.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\THIRD_PARTY_NOTICES.md"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordLatexAddin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "WordLatexVSTO"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordLatexAddin"; ValueType: string; ValueName: "Description"; ValueData: "LaTeX 自动转换为 Word 原生公式"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordLatexAddin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordLatexAddin"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri}"; Flags: uninsdeletekey

[Icons]
Name: "{group}\使用说明"; Filename: "{app}\docs\USER_GUIDE.md"
Name: "{group}\安装说明"; Filename: "{app}\docs\INSTALL.md"
Name: "{group}\卸载 WordLatexVSTO"; Filename: "{uninstallexe}"

[Run]
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Install ""{app}\WordLatexAddin.vsto"" /Silent"; Flags: runhidden waituntilterminated runasoriginaluser skipifdoesntexist; StatusMsg: "正在注册 WordLatexVSTO..."

[UninstallRun]
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Uninstall ""{app}\WordLatexAddin.vsto"" /Silent"; Flags: runhidden waituntilterminated skipifdoesntexist

[Code]
function GetManifestUri(Param: String): String;
var
  ManifestPath: String;
begin
  ManifestPath := ExpandConstant('{app}\WordLatexAddin.vsto');
  StringChangeEx(ManifestPath, '\', '/', True);
  StringChangeEx(ManifestPath, ' ', '%20', True);
  Result := 'file:///' + ManifestPath + '|vstolocal';
end;

function GetVstoInstallerPath(Param: String): String;
var
  Candidate: String;
begin
  if IsWin64 then
  begin
    Candidate := ExpandConstant('{commoncf64}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');
    if FileExists(Candidate) then
    begin
      Result := Candidate;
      exit;
    end;
  end;

  Result := ExpandConstant('{commoncf32}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');
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

