; InnoScript Version 5.3  Build 9
; Randem Systems, Inc.
; Copyright 2003-2006
; website:  http://www.randem.com
; support:  http://www.innoscript.com/cgi-bin/discus/discus.cgi

; Date: September 27, 2006

;              VB Runtime Files Folder:   C:\Program Files\Randem Systems\InnoScript\InnoScript 5.3\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Documents and Settings\Pudar\My Documents\Data\Home Stuff\Magic\Visual Basic Programs\StackView V5\StackView.vbp
; Inno Setup Script Output File (.iss):   C:\Documents and Settings\Pudar\My Documents\Data\Home Stuff\Magic\Visual Basic Programs\StackView V5\StackViewScript1.iss

; ------------------------
;        References
; ------------------------

; Visual Basic runtime objects and procedures - (MSVBVM60.DLL)
; Standard OLE Types - (OLEPRO32.DLL)
; OLE Automation - (STDOLE2.TLB)
; Microsoft Scripting Runtime - (scrrun.dll)
; Microsoft Shell Controls And Automation - (SHELL32.dll)


; --------------------------
;        Components
; --------------------------

; Microsoft Common Dialog Control 6.0 (SP6) - (comdlg32.ocx)
; Microsoft Windows Common Controls 6.0 (SP6) - (mscomctl.ocx)
; Microsoft Rich Textbox Control 6.0 (SP6) - (richtx32.ocx)
; Microsoft Tabbed Dialog Control 6.0 (SP6) - (tabctl32.ocx)


[Setup]
AppName=StackView
AppVerName=StackView 5.0.1
AppPublisher=
AppUpdatesURL=www.stackview.com
AppVersion=5.0.1
VersionInfoVersion=5.0.1
AllowNoIcons=no
WizardImageFile=C:\Documents and Settings\Pudar\My Documents\Data\Home Stuff\Magic\Magic\Cards\StackView Installer Logo Large.bmp
WizardSmallImageFile=C:\Documents and Settings\Pudar\My Documents\Data\Home Stuff\Magic\Magic\Cards\StackView Installer Logo Small.bmp
DefaultGroupName=StackView
DefaultDirName=StackView
AppCopyright=
PrivilegesRequired=Admin
MinVersion=4.0,4.0
Compression=lzma
OutputBaseFilename=StackView501Release

[Tasks]
Name: desktopicon; Description: Create a &desktop icon; GroupDescription: Additional Icons:

[Files]
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\msvbvm60.dll; DestDir: {sys}; Flags:  sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\oleaut32.dll; DestDir: {sys}; Flags:  sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\olepro32.dll; DestDir: {sys}; Flags:  sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\asycfilt.dll; DestDir: {sys}; Flags:  sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\stdole2.tlb; DestDir: {sys}; Flags:  regtypelib
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\comcat.dll; DestDir: {sys}; Flags:  sharedfile
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\stackview user guide.pdf; DestDir: {app}; Flags:  ignoreversion
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\scrrun.dll; DestDir: {sys}; Flags:  regserver restartreplace sharedfile
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\comdlg32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\mscomctl.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\richtx32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\tabctl32.ocx; DestDir: {sys}; Flags:  regserver restartreplace sharedfile
Source: c:\documents and settings\pudar\my documents\data\home stuff\magic\visual basic programs\stackview v5\stackview.exe; DestDir: {app}; Flags:  ignoreversion

[INI]
Filename: {app}\StackView.url; Section: InternetShortcut; Key: URL; String: 

[Icons]
Name: {group}\StackView; Filename: {app}\StackView.exe; WorkingDir: {app}
Name: {group}\StackView on the Web; Filename: {app}\StackView.url
Name: {group}\Uninstall StackView; Filename: {uninstallexe}
Name: {userdesktop}\StackView; Filename: {app}\StackView.exe; Tasks: desktopicon; WorkingDir: {app}

[Run]
Filename: {app}\StackView.exe; Description: Launch StackView; Flags: nowait postinstall skipifsilent; WorkingDir: {app}

[UninstallDelete]
Type: files; Name: {app}\StackView.url
