; Axon Outlook add-in installer (per-user, no admin). Installs the DLL + icons, registers the
; managed COM add-in so the Move/Download buttons appear in Outlook, asks for the on-site Ollama
; server URL, and writes the client config.
;
; Build (after build.ps1 has produced ..\AxonAddin.dll):
;   "C:\...\Inno Setup 6\ISCC.exe" installer\AxonOutlook.iss
; Output: installer\Output\AxonOutlook-Setup.exe

[Setup]
AppId={{C3D8F1A2-4B5E-4C6D-9E7F-1A2B3C4D5E6F}}
AppName=Axon Outlook add-in
AppVersion=1.0.0
AppPublisher=Axon Group
DefaultDirName={localappdata}\AxonOutlook
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=AxonOutlook-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\AxonAddin.dll";          DestDir: "{app}"; Flags: ignoreversion
Source: "..\icons\axon-move.png";     DestDir: "{app}"; Flags: ignoreversion
Source: "..\icons\axon-download.png"; DestDir: "{app}"; Flags: ignoreversion

[Code]
const
  CLSID  = '{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}';
  PROGID = 'Axon.OutlookAddin';
  ASM    = 'AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null';
var
  UrlPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  UrlPage := CreateInputQueryPage(wpWelcome, 'Model server',
    'Where is your Ollama server?',
    'Folder suggestions are produced by your on-site Ollama server. Email content is sent only to ' +
    'this server and never leaves your network. Enter its address (ask IT if unsure).');
  UrlPage.Add('Ollama server URL:', False);
  UrlPage.Values[0] := 'http://localhost:11434';
end;

function OllamaUrl(): String;
begin
  Result := Trim(UrlPage.Values[0]);
  if Result = '' then Result := 'http://localhost:11434';
end;

procedure WriteConfig();
var
  dir, path, json: String;
begin
  dir := ExpandConstant('{userappdata}\AxonOutlook');
  ForceDirectories(dir);
  path := dir + '\config.json';
  json := '{"ollama_url": "' + OllamaUrl() + '", "model": "llama3.2:3b"}';
  SaveStringToFile(path, json, False);
end;

procedure RegisterAddin();
var
  code, clsKey, inproc, addins: String;
begin
  code := ExpandConstant('{app}\AxonAddin.dll');
  StringChangeEx(code, '\', '/', True);
  code := 'file:///' + code;
  RegWriteStringValue(HKCU, 'Software\Classes\' + PROGID + '\CLSID', '', CLSID);
  clsKey := 'Software\Classes\CLSID\' + CLSID;
  inproc := clsKey + '\InprocServer32';
  RegWriteStringValue(HKCU, inproc, '', ExpandConstant('{sys}\mscoree.dll'));
  RegWriteStringValue(HKCU, inproc, 'ThreadingModel', 'Both');
  RegWriteStringValue(HKCU, inproc, 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(HKCU, inproc, 'Assembly', ASM);
  RegWriteStringValue(HKCU, inproc, 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(HKCU, inproc, 'CodeBase', code);
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'Assembly', ASM);
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'CodeBase', code);
  RegWriteStringValue(HKCU, clsKey + '\ProgId', '', PROGID);
  addins := 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID;
  RegWriteStringValue(HKCU, addins, 'FriendlyName', 'Axon intelligence');
  RegWriteStringValue(HKCU, addins, 'Description', 'File and download emails with Axon intelligence');
  RegWriteDWordValue(HKCU, addins, 'LoadBehavior', 3);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    WriteConfig();
    RegisterAddin();
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\CLSID\' + CLSID);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\' + PROGID);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
  end;
end;
