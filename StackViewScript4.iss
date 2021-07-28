; InnoScript Version 5.3  Build 9
; Randem Systems, Inc.
; Copyright 2003-2006
; website:  http://www.randem.com
; support:  http://www.innoscript.com/cgi-bin/discus/discus.cgi

; Date: September 26, 2006

;              VB Runtime Files Folder:   C:\Program Files\Randem Systems\InnoScript\InnoScript 5.3\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Documents and Settings\Pudar\My Documents\Data\Home Stuff\Magic\Visual Basic Programs\StackView V5\StackView.vbp
; Inno Setup Script Output File (.iss):   C:\Documents and Settings\Pudar\My Documents\Data\Home Stuff\Magic\Visual Basic Programs\StackView V5\StackViewScript3.iss

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
AppVerName=StackView 5.0.2
AppPublisher=www.stackview.com
AppUpdatesURL=www.stackview.com
AppVersion=5.0.2
VersionInfoVersion=5.0.2
WizardImageFile=C:\Documents and Settings\Parents\My Documents\Data\Home Stuff\Magic\Magic\Cards\StackView Installer Logo Large.bmp
WizardSmallImageFile=C:\Documents and Settings\Parents\My Documents\Data\Home Stuff\Magic\Magic\Cards\StackView Installer Logo Small.bmp
AllowNoIcons=no
DefaultGroupName=StackView
DefaultDirName={pf}\StackView
AppCopyright=Copyright © 2006 by Nick Pudar
PrivilegesRequired=None
MinVersion=4.0,4.0
Compression=lzma
OutputBaseFilename=StackViewSetup502

[Tasks]
Name: desktopicon; Description: Create a &desktop icon; GroupDescription: Additional Icons:

[Files]
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\msvbvm60.dll; DestDir: {sys}; Flags: sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\oleaut32.dll; DestDir: {sys}; Flags: sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\olepro32.dll; DestDir: {sys}; Flags: sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\asycfilt.dll; DestDir: {sys}; Flags: sharedfile
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\stdole2.tlb; DestDir: {sys}; Flags: regtypelib
Source: c:\program files\randem systems\innoscript\innoscript 5.3\vb 6 redist files\comcat.dll; DestDir: {sys}; Flags: sharedfile
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\stackview user guide.pdf; DestDir: {app}; Flags: ignoreversion
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\scr56en.exe; DestDir: {tmp}; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\scripten.exe; DestDir: {tmp}; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\comdlg32.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\mscomctl.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\richtx32.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\tabctl32.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\support\msstdfmt.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile
Source: c:\documents and settings\parents\my documents\data\home stuff\magic\visual basic programs\stackview v5\stackview.exe; DestDir: {app}; Flags: ignoreversion

[INI]

[Icons]
Name: {group}\StackView; Filename: {app}\StackView.exe; WorkingDir: {app}
Name: {group}\Uninstall StackView; Filename: {uninstallexe}
Name: {userdesktop}\StackView; Filename: {app}\StackView.exe; Tasks: desktopicon; WorkingDir: {app}

[Run]
; Install Windows 98, Me, and NT 4.0 version
Filename: {tmp}\scr56en.exe; Parameters: /r:n /q:1; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01
; Install Windows 2000 and XP version
Filename: {tmp}\scripten.exe; Parameters: /r:n /q:1; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02
Filename: {app}\StackView.exe; Description: Launch StackView; Flags: nowait postinstall skipifsilent; WorkingDir: {app}
