; NullSoft installer script for PPspliT
; Written by Massimo Rimondini

;--------------------------------

; Use modern user interface
!include "MUI2.nsh"
; Support easier-to-write conditional expressions
!include LogicLib.nsh


; Define some variables
Var HOST_ARCH
Var PPSPLIT_RELEASE
Var REGISTRATION_HANDLER
Var CURRENT_OFFICE_RELEASE
Var SHORT_OFFICE_RELEASE
Var RELEASE_ARCH
Var CURRENT_OFFICE_REGKEY
Var ADDIN_FILE

Var ERRORS		; if "yes" at the end of the install, then errors have occurred in setting up the add-in
Var CONFIGURED		; if "" at the end of the install, then the add-in has not been set up for any Office releases

;--------------------------------

; This function must be shared between installer and uninstaller
!macro define_init_callback un
Function ${un}.onInit
	StrCpy $PPSPLIT_RELEASE "1.7"
	StrCpy $ERRORS ""
	StrCpy $CONFIGURED ""
	ReadRegStr $HOST_ARCH HKLM "System\CurrentControlSet\Control\Session Manager\Environment" "PROCESSOR_ARCHITECTURE"
FunctionEnd
!macroend
!insertmacro define_init_callback ""
!insertmacro define_init_callback "un"

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

; This function must be shared between installer and uninstaller
!macro define_addin_registration_function un
Function ${un}Handle_Addin_Registration
	Push $0
	Push $1
	Push $2
	Push $3
	StrCpy $0 0	; Iterates over registry keys representing candidate Office releases
	SetRegView 32
${un}Loop:
	; Enumerate all Office releases found in the registry
	EnumRegKey $CURRENT_OFFICE_RELEASE HKLM "Software\Microsoft\Office" $0
	StrCmp "" $CURRENT_OFFICE_RELEASE ${un}Done

	; We now check that the currently found Office release is really installed
	StrCpy $RELEASE_ARCH ""
	StrCpy $CURRENT_OFFICE_REGKEY "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE"
	ReadRegStr $2 HKLM "$CURRENT_OFFICE_REGKEY\PowerPoint\InstallRoot" "Path"
	IfErrors ${un}Next
	
			
	; We are now sure that the current Office release is really installed
	
	; Only the first 2 characters of the Office release are relevant
	StrCpy $SHORT_OFFICE_RELEASE $CURRENT_OFFICE_RELEASE 2
	StrCpy $1 "unknown release"
	; Assume we could not recognize the Office release until we actually do
	StrCpy $ERRORS "yes"
	StrCmp "10" $SHORT_OFFICE_RELEASE 0 +3
	StrCpy $1 "Office XP"	
	StrCpy $ERRORS ""
	StrCmp "11" $SHORT_OFFICE_RELEASE 0 +3
	StrCpy $1 "Office 2003"
	StrCpy $ERRORS ""
	StrCmp "12" $SHORT_OFFICE_RELEASE 0 +3
	StrCpy $1 "Office 2007"
	StrCpy $ERRORS ""
	StrCmp "14" $SHORT_OFFICE_RELEASE 0 +3
	StrCpy $1 "Office 2010"
	StrCpy $ERRORS ""
	
	${If} $3 = 0
		; First iteration: 32-bit
		StrCpy $RELEASE_ARCH "x86"
	${Else}
		; Second iteration: 64-bit
		StrCpy $RELEASE_ARCH "amd64"
	${EndIf}

	DetailPrint "      Configuring PowerPoint $CURRENT_OFFICE_RELEASE ($1), architecture $RELEASE_ARCH"
	StrCpy $CONFIGURED "yes"

	; Determine the correct version of the add-in to install	
	${If} $SHORT_OFFICE_RELEASE <= 11
		; Prior to Office 2007
		StrCpy $ADDIN_FILE "$INSTDIR\PPspliT.ppa"
	${Else}
		; Office 2007 or newer
		${If} $RELEASE_ARCH == "x86"
			StrCpy $ADDIN_FILE "$INSTDIR\PPspliT.ppam"
		${Else}
			StrCpy $ADDIN_FILE "$INSTDIR\PPspliT.ppam"
		${EndIf}
	${EndIf}
	Call $REGISTRATION_HANDLER
${un}Next:
	IntOp $0 $0 + 1
	Goto ${un}Loop
${un}Done:
	${If} $3 = 0
	${AndIf} $HOST_ARCH != "x86"
		; Reiterate on 64-bit releases
		StrCpy $3 1
		StrCpy $0 0
		SetRegView 64
		Goto ${un}Loop
	${EndIf}
	Pop $3
	Pop $2
	Pop $1
	Pop $0
FunctionEnd
!macroend
!insertmacro define_addin_registration_function ""
!insertmacro define_addin_registration_function "un."

;--------------------------------

Function RegisterAddin
	WriteRegStr HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\PPspliT" "Path" $ADDIN_FILE
	WriteRegDWORD HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\PPspliT" "AutoLoad" 1
	DetailPrint "            Add-in registered"
FunctionEnd

;--------------------------------

Function un.UnregisterAddin
	EnumRegKey $9 HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\PPspliT" 0
	IfErrors +3
	DeleteRegKey HKCU "Software\Microsoft\Office\$CURRENT_OFFICE_RELEASE\PowerPoint\AddIns\PPspliT"
	DetailPrint "            Add-in unregistered"
FunctionEnd

;--------------------------------

Section ""

	SetDetailsView show
	
	SetOutPath $INSTDIR

  	IfFileExists $INSTDIR\mouse-button.gif 0 +2
  	DetailPrint "Upgrading existing installation."
  	
  	File changelog.txt
	File common_resources\about-button.gif
	File common_resources\mouse-button.gif
	File common_resources\slide-numbers.gif
	File common_resources\ppsplit-button.gif
	File common_resources\ppsplit.ico
	File PPT11-\*.*
	File PPT12+\*.*
	
	DetailPrint "Registering add-in for all installed PowerPoint releases..."
	
	GetFunctionAddress $REGISTRATION_HANDLER RegisterAddin
	Call Handle_Addin_Registration

	WriteUninstaller "$INSTDIR\ppsplit-uninstall.exe"

	; Create an entry under "Add/Remove Programs"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "DisplayName" "PPspliT"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "UninstallString" "$INSTDIR\ppsplit-uninstall.exe"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "DisplayIcon" "$INSTDIR\ppsplit.ico"
	WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "DisplayVersion" "$PPSPLIT_RELEASE"
	WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "NoModify" 1
	WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT" "NoRepair" 1

	StrCmp $ERRORS "" +2
	MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to recognize at least one of the installed Office releases: the add-in may have been left unconfigured."
	
	StrCmp $CONFIGURED "" 0 +2
	MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to automatically detect any Office releases. The add-in has been left unconfigured."
SectionEnd


Section "Uninstall"

	SetDetailsView show

	DetailPrint "Unregistering add-in for all installed PowerPoint releases..."

	GetFunctionAddress $REGISTRATION_HANDLER un.UnregisterAddin
	Call un.Handle_Addin_Registration
	
	; WARNING: The following command should only be used if the InstallDir
	; cannot be changed by the user (like it is the case here). Otherwise, you
	; risk to wipe out important folders!!
	RMDir /r "$INSTDIR"

	DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PPspliT"

SectionEnd

!insertmacro MUI_LANGUAGE "English"
LangString MUI_UNTEXT_WELCOME_INFO_TEXT ${LANG_ENGLISH} "This wizard will guide you through the uninstallation of $(^NameDA).$\r$\n$\r$\nBefore starting the uninstallation, make sure PowerPoint is not running.$\r$\n$\r$\n$_CLICK"
