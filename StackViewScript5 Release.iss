; InnoScript Version 9.2  Build 5
; Randem Systems, Inc.
; Copyright (c) 2002 - 2007, Randem Systems, Inc.
; Website:  http://www.randem.com
; Support:  http://www.randem.com/cgi-bin/discus/discus.cgi
; OS: Windows XP 5.1 build 2600 (Service Pack 3)

; Derived from VB VBP Project File

; Local Machine Settings. Use these settings as a template for your installation folders

; {app}           : C:\Program Files\Randem Systems\innoscript
; {appdata}       : C:\Documents and Settings\Parents\Application Data\Randem Systems\innoscript\
; {localappdata}  : C:\Documents and Settings\Parents\Local Settings\Application Data\Randem Systems\innoscript\
; {cf}            : C:\Program Files\Common Files\Randem Systems
; {tmp}           : C:\Documents and Settings\Parents\Local Settings\Temp\
; {commonappdata} : C:\Documents and Settings\All Users\Application Data\Randem Systems\innoscript
; {pf}            : C:\Program Files\

; Date: July 27, 2008

;              VB Runtime Files Folder:   C:\Program Files\Randem Systems\InnoScript\InnoScript 9\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Documents and Settings\Parents\My Documents\Data\Home Stuff\Magic\Visual Basic Programs\StackView V5\StackView.vbp
; Inno Setup Script Output File (.iss):   C:\Documents and Settings\Parents\My Documents\Data\Home Stuff\Magic\Visual Basic Programs\StackView V5\StackViewScript5 Release.iss
;         Script Template Files (.tpl):   C:\Documents and Settings\Parents\Local Settings\Application Data\Randem Systems\innoscript\Templates\Release.tpl

; ------------------------
;        References
; ------------------------

; Microsoft Scripting Runtime - (scrrun.dll)


; --------------------------
;        Components
; --------------------------



[Setup]
SetupLogging=Yes
AppId=StackView 

;------------------------------------------------------------------------------------------------------------------------
; Taken from VBP/VBG Project File Parameters AppName, AppName AppVersion and Company
;------------------------------------------------------------------------------------------------------------------------

AppName=StackView
AppVerName=StackView 5.0.3
AppPublisher=www.stackview.com
AppUpdatesURL=www.stackview.com
AppVersion=5.0.3
VersionInfoVersion=5.0.3
WizardImageFile=C:\Documents and Settings\Parents\My Documents\Data\Home Stuff\Magic\Magic\Cards\StackView Installer Logo Large.bmp
WizardSmallImageFile=C:\Documents and Settings\Parents\My Documents\Data\Home Stuff\Magic\Magic\Cards\StackView Installer Logo Small.bmp
AllowNoIcons=no
DefaultGroupName=StackView
DefaultDirName={pf}\StackView
AppCopyright=Copyright © 2006 by Nick Pudar
PrivilegesRequired=None
MinVersion=4.0,4.0
Compression=lzma
OutputBaseFilename=StackViewSetup503

[Tasks]
Name: ScriptingRuntime; Description: Install Microsoft's Scripting Runtime; GroupDescription: Install Scripting Runtime:

[Files]
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\msvbvm60.dll; DestDir: {sys}; Flags:  uninsneveruninstall regserver restartreplace sharedfile; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\oleaut32.dll; DestDir: {sys}; Flags:  uninsneveruninstall restartreplace sharedfile; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\olepro32.dll; DestDir: {sys}; Flags:  uninsneveruninstall regserver restartreplace sharedfile; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\asycfilt.dll; DestDir: {sys}; Flags:  uninsneveruninstall restartreplace sharedfile; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\stdole2.tlb; DestDir: {sys}; Flags:  uninsneveruninstall restartreplace sharedfile regtypelib; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\comcat.dll; DestDir: {sys}; Flags:  uninsneveruninstall restartreplace sharedfile; OnlyBelowVersion: 0,6
Source: c:\program files\randem systems\innoscript\innoscript 9\vb 6 redist files\vb5db.dll; DestDir: {sys}; Flags:  uninsneveruninstall restartreplace sharedfile; OnlyBelowVersion: 0,6
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\scripten.exe; DestDir: {app}; Flags:  restartreplace sharedfile ignoreversion nocompression; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02; Tasks: ScriptingRuntime
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\scr56en.exe; DestDir: {app}; Flags:  restartreplace sharedfile ignoreversion nocompression; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01; Tasks: ScriptingRuntime
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\stackview user guide.pdf; DestDir: {app}; Flags:  ignoreversion; 
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\comdlg32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile; 
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\mscomctl.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile; 
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\richtx32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile; 
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\tabctl32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile; 
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\stackview.exe; DestDir: {app}; Flags:  restartreplace ignoreversion; 

[INI]
Filename: {app}\StackView.url; Section: InternetShortcut; Key: URL; String: 

[Icons]
Name: {group}\StackView ; Filename : {app}\StackView.exe; WorkingDir: {app}
Name: {group}\{cm:ProgramOnTheWeb, StackView }; Filename: {app}\StackView.url; IconFilename: {app}\StackView.ico
Name: {group}\{cm:UninstallProgram, StackView }; Filename: {uninstallexe}

[Run]
Filename: {tmp}\scr56en.exe; Parameters: /r:n /q:a; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: ScriptingRuntime
Filename: {tmp}\scripten.exe; Parameters: /r:n /q:a; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: ScriptingRuntime
Filename: {app}\StackView.exe; Description: {cm:LaunchProgram, StackView }; Flags: nowait postinstall skipifsilent; WorkingDir: {app}

[UninstallDelete]
Type: files; Name: {app}\StackView.url
Type: dirifempty; Name: {app}

[Comments]

 Template Processing first character indicators usage.

	   	No Indicator		Attempt replacement if cannot replace then add the line
  + 		Plus Sign		Force addition of template line into script (no attempted replacement).
  ;		Semi-colon		Add line as a comment only. (No attempted replacement).
  -		Minus Sign		Delete the line. (No attempted replacement).
  &		Ampersand		Comment the line. (No attempted replacement).
