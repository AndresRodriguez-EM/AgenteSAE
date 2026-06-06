; AgenteSAE - script de Inno Setup (genera AgenteSAE-Setup.exe)
; Empaqueta la carpeta dist_app\AgenteSAE generada por PyInstaller (--windowed, sin --onefile).
[Setup]
AppName=AgenteSAE
AppVersion=1.0
AppPublisher=SAE
DefaultDirName={autopf}\AgenteSAE
DefaultGroupName=AgenteSAE
DisableProgramGroupPage=yes
OutputBaseFilename=AgenteSAE-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\AgenteSAE.exe

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "dist_app\AgenteSAE\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\AgenteSAE"; Filename: "{app}\AgenteSAE.exe"
Name: "{group}\Desinstalar AgenteSAE"; Filename: "{uninstallexe}"
Name: "{autodesktop}\AgenteSAE"; Filename: "{app}\AgenteSAE.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\AgenteSAE.exe"; Description: "Ejecutar AgenteSAE ahora"; Flags: nowait postinstall skipifsilent
