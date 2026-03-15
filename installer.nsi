!include "MUI2.nsh"

Name "劍橋字典Excel轉Anki v0.0.2"
OutFile "installer\EnglishLearningSetup_v0.0.2.exe"
InstallDir "$PROGRAMFILES\CambridgeDictAnki"
RequestExecutionLevel admin

!define MUI_ABORTWARNING

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "TradChinese"

Section "Install"
  SetOutPath "$INSTDIR"
  File "dist\劍橋字典Excel轉Anki_v0.0.2.exe"
  File "converter.py"
  
  WriteUninstaller "$INSTDIR\uninstall.exe"
  
  CreateDirectory "$SMPROGRAMS\CambridgeDictAnki"
  CreateShortcut "$SMPROGRAMS\CambridgeDictAnki\劍橋字典Excel轉Anki.lnk" "$INSTDIR\劍橋字典Excel轉Anki_v0.0.2.exe"
  CreateShortcut "$DESKTOP\劍橋字典Excel轉Anki.lnk" "$INSTDIR\劍橋字典Excel轉Anki_v0.0.2.exe"
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\uninstall.exe"
  Delete "$INSTDIR\劍橋字典Excel轉Anki_v0.0.2.exe"
  Delete "$INSTDIR\converter.py"
  RMDir "$INSTDIR"
  
  Delete "$SMPROGRAMS\CambridgeDictAnki\劍橋字典Excel轉Anki.lnk"
  RMDir "$SMPROGRAMS\CambridgeDictAnki"
  Delete "$DESKTOP\劍橋字典Excel轉Anki.lnk"
SectionEnd
