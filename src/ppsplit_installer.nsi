; NullSoft installer script for PPspliT
; Written by Massimo Rimondini

;--------------------------------

; Use modern user interface
!include "MUI2.nsh"
; Support easier-to-write conditional expressions
!include LogicLib.nsh


; Define some variables
Var /global OFFICE_RELEASE
Var /global OFFICE_REGKEY
Var /global OFFICE_ARCH
Var /global PPSPLIT_RELEASE

Function .onInit
	StrCpy $PPSPLIT_RELEASE "1.1"
FunctionEnd

;--------------------------------

; Installer package attributes
Name "PPspliT"
!define MUI_ICON common_resources\ppsplit.ico
!define MUI_UNICON common_resources\ppsplit.ico
OutFile "..\PPspliT-setup.exe"

; User interface options
BrandingText "PPspliT $PPSPLIT_RELEASE installer"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "common_resources\ppsplit-wide.bmp"
!define MUI_WELCOMEFINISHPAGE_BITMAP "common_resources\ppsplit-large.bmp"
!define MUI_WELCOMEFINISHPAGE_BITMAP_NOSTRETCH
!define MUI_PAGE_HEADER_TEXT "PPspliT version $PPSPLIT_RELEASE"
!define MUI_PAGE_HEADER_SUBTEXT "Setup procedure"
; Ask for confirmation on abort request from the user
!define MUI_ABORTWARNING
!define MUI_ABORTWARNING_CANCEL_DEFAULT

; Give the user some time to read the installation log
!define MUI_FINISHPAGE_NOAUTOCLOSE

; Request application privileges for Windows Vista
RequestExecutionLevel user

;--------------------------------

!define MUI_WELCOMEPAGE_TITLE "PPspliT setup"
!define MUI_WELCOMEPAGE_TEXT "Welcome to the PPspliT installer!$\n$\nThis tool will guide you to the easy process of setting up PPspliT on your computer.$\n$\nPlease make sure that PowerPoint is not running before proceeding."

!insertmacro MUI_PAGE_WELCOME

;--------------------------------

!define MUI_LICENSEPAGE_TEXT_TOP "Please review the following licensing and usage information:"
!define MUI_LICENSEPAGE_TEXT_BOTTOM "After you have read the above conditions, check the box below to continue."
!define MUI_LICENSEPAGE_CHECKBOX
!define MUI_LICENSEPAGE_CHECKBOX_TEXT "I have read the licensing and usage information."

!insertmacro MUI_PAGE_LICENSE license.txt

;--------------------------------

InstallDir $APPDATA\Microsoft\AddIns\PPspliT

!insertmacro MUI_PAGE_INSTFILES

;--------------------------------

!define MUI_FINISHPAGE_TEXT "Setup of PPspliT is now complete!$\n$\nTo start using the add-in, simply start PowerPoint and look for the PPspliT toolbar.$\n$\n$\nIf you want to remove the add-in, use the $\"Add/Remove Programs$\" tool in the Control Panel."

!insertmacro MUI_PAGE_FINISH

;--------------------------------

; Uninstaller attributes
!define MUI_UNABORTWARNING

!define MUI_UNWELCOMEFINISHPAGE_BITMAP "common_resources\ppsplit-uninst-large.bmp"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP_NOSTRETCH
!insertmacro MUI_UNPAGE_WELCOME

!insertmacro MUI_UNPAGE_CONFIRM

!insertmacro MUI_UNPAGE_INSTFILES

!define MUI_UNFINISHPAGE_NOAUTOCLOSE
!insertmacro MUI_UNPAGE_FINISH

;--------------------------------

Function GetOfficeRelease
	EnumRegKey $OFFICE_REGKEY HKCU Software\Microsoft\Office 0
	StrCpy $0 $OFFICE_REGKEY -2
	${If} $0 <= 11
		; Prior to Office 2007
		StrCpy $OFFICE_RELEASE "11-"
		DetailPrint "Detected Office release $0 (prior to Office 2007)."
	${ElseIf} $0 >= 12
		; Office 2007 or newer
		StrCpy $OFFICE_RELEASE "12+"
		DetailPrint "Detected Office release $0 (Office 2007 or newer)."
	${Else}
		MessageBox MB_OK|MB_ICONEXCLAMATION "Could not recognize the installed Office release. Aborting :-("
		Abort "Installation failed: could not recognize the installed Office version."
	${EndIf}
FunctionEnd

;--------------------------------

Function GetOfficeArchitecture
	StrCmp "$PROGRAMFILES64" "$PROGRAMFILES32" 0 +4
	DetailPrint "32-bit OS detected. Assuming 32-bit architecture for Office."
	StrCpy $OFFICE_ARCH "x86"
	Return
	
	StrCpy $0 $OFFICE_REGKEY -2
	IfFileExists "$PROGRAMFILES64\Microsoft Office\Office$0\powerpnt.exe" +3
	StrCpy $OFFICE_ARCH "x86"
	Goto +2
	StrCpy $OFFICE_ARCH "amd64"
	DetailPrint "Detected Office architecture: $OFFICE_ARCH"
FunctionEnd

;--------------------------------

Function un.UnregisterAddin
	StrCpy $0 0
	EnumRegKey $1 HKCU Software\Microsoft\Office $0
	${DoWhile} $1 != ""
		EnumRegKey $2 HKCU "Software\Microsoft\Office\$1\PowerPoint\AddIns\PPspliT" 0
		IfErrors +3
		DeleteRegKey HKCU "Software\Microsoft\Office\$1\PowerPoint\AddIns\PPspliT"
		DetailPrint "Unregistered for Office version $1."
		IntOp $0 $0 + 1
		EnumRegKey $1 HKCU Software\Microsoft\Office $0
	${Loop}
FunctionEnd

;--------------------------------

Section ""

	SetDetailsView show
	
	SetOutPath $INSTDIR

	Call GetOfficeRelease
	Call GetOfficeArchitecture
  
  	IfFileExists $INSTDIR\mouse-button.gif 0 +2
  	DetailPrint "Upgrading existing installation."
  	
  	File changelog.txt
	File common_resources\mouse-button.gif
	File common_resources\ppsplit-button.gif
	File common_resources\ppsplit.ico
	${If} $OFFICE_RELEASE == "11-"
		File PPT11-\*.*
	${ElseIf} $OFFICE_RELEASE == "12+"
		${If} $OFFICE_ARCH == "amd64"
			File PPT12+\amd64\*.*
		${Else}
			File PPT12+\x86\*.*
		${EndIf}
	${EndIf}
	
	DetailPrint "Registering add-in"
	${If} $OFFICE_RELEASE == "11-"
		WriteRegStr HKCU "Software\Microsoft\Office\$OFFICE_REGKEY\PowerPoint\AddIns\PPspliT" "Path" "$INSTDIR\PPspliT.ppa"
	${ElseIf} $OFFICE_RELEASE == "12+"
		WriteRegStr HKCU "Software\Microsoft\Office\$OFFICE_REGKEY\PowerPoint\AddIns\PPspliT" "Path" "$INSTDIR\PPspliT-$OFFICE_ARCH.ppam"
	${EndIf}
	WriteRegDWORD HKCU "Software\Microsoft\Office\$OFFICE_REGKEY\PowerPoint\AddIns\PPspliT" "AutoLoad" 1

	WriteUninstaller "$INSTDIR\ppsplit-uninstall.exe"

	; Create an entry under "Add/Remove Programs"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "DisplayName" "PPspliT"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "UninstallString" "$INSTDIR\ppsplit-uninstall.exe"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "DisplayIcon" "$INSTDIR\ppsplit.ico"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "DisplayVersion" "$PPSPLIT_RELEASE"
	WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "NoModify" 1
	WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "NoRepair" 1

SectionEnd


Section "Uninstall"

	SetDetailsView show

	DetailPrint "Unregistering add-in"

	Call un.UnregisterAddin
	
	; WARNING: The following command should only be used if the InstallDir
	; cannot be changed by the user (like it is the case here). Otherwise, you
	; risk to wipe out important folders!!
	RMDir /r "$INSTDIR"

	DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT"

SectionEnd

!insertmacro MUI_LANGUAGE "English"
LangString MUI_UNTEXT_WELCOME_INFO_TEXT ${LANG_ENGLISH} "This wizard will guide you through the uninstallation of $(^NameDA).$\r$\n$\r$\nBefore starting the uninstallation, make sure PowerPoint is not running.$\r$\n$\r$\n$_CLICK"
