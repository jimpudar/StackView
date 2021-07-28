 How to deploy scrrun.dll and the rest of Windows Script Runtime

This shows how you can deploy the Windows Script Runtime. Note that on Windows NT 4.0, users must have Internet Explorer 3.02 or later installed, according to the Microsoft web site.

1. Download scr56en.exe (Windows Script 5.6 for Windows 98, Windows Me, and Windows NT 4.0).
2. Download scripten.exe (Windows Script 5.6 for Windows 2000 and XP).
3. Add these lines to your script:

[Files]
Source: "scr56en.exe"; DestDir: "{tmp}"; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01
Source: "scripten.exe"; DestDir: "{tmp}"; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02

[Run]
; Install Windows 98, Me, and NT 4.0 version
Filename: "{tmp}\scr56en.exe"; Parameters: "/r:n /q:1"; MinVersion: 4.1,4.0; OnlyBelowVersion: 0,4.01
; Install Windows 2000 and XP version
Filename: "{tmp}\scripten.exe"; Parameters: "/r:n /q:1"; MinVersion: 0,5.0; OnlyBelowVersion: 0,5.02