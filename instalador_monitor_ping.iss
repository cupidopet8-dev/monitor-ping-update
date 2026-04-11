#define MyAppName "Monitor de Ping"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "Anderson Ribeiro"
#define MyAppExeName "MonitorPing.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-47A8-91B2-C3D4E5F6A7B8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Monitor de Ping
DefaultGroupName=Monitor de Ping
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=instalador_monitor_ping
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
SetupIconFile=icone.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "dist\MonitorPing.exe"; DestDir: "{app}"; DestName: "{#MyAppExeName}"; Flags: ignoreversion

[Icons]
Name: "{group}\Monitor de Ping"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Monitor de Ping"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Monitor de Ping"; Flags: nowait postinstall skipifsilent