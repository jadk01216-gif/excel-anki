[Setup]
AppName=劍橋字典Excel轉Anki
AppVersion=0.0.3
DefaultDirName={pf}\CambridgeDictAnki
DefaultGroupName=CambridgeDictAnki
UninstallDisplayIcon={app}\ExcelToAnki.exe
Compression=lzma2
SolidCompression=yes
OutputDir=output
OutputBaseFilename=EnglishLearningSetup_v0.0.3

[Files]
Source: "dist\ExcelToAnki\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\劍橋字典Excel轉Anki"; Filename: "{app}\ExcelToAnki.exe"
Name: "{userdesktop}\劍橋字典Excel轉Anki"; Filename: "{app}\ExcelToAnki.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "建立桌面快捷方式"; GroupDescription: "其他任務:"

[Run]
Filename: "{app}\ExcelToAnki.exe"; Description: "啟動程式"; Flags: nowait postinstall skipifsilent
