; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=Gescom
AppVerName=Gescom 1.0
AppPublisher=NetConf-Simmetric
AppPublisherURL=http://www.coding-web.com
AppSupportURL=http://www.coding-web.com
AppUpdatesURL=http://www.coding-web.com
DefaultDirName={pf}\Gescom
DefaultGroupName=Gescom
;AlwaysCreateUninstallIcon=yes
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableStartupPrompt=yes
DisableReadyMemo=yes
DisableReadyPage=yes
; uncomment the following line if you want your installation to run on NT 3.51 too.
; MinVersion=4,3.51

[Files]
Source: "..\GescomUI.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\GescomPrint.dll"; DestDir: "{app}"; Flags: regserver ignoreversion
Source: "..\Logo.bmp"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\GCServerMTS.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "..\CoreServer.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "..\GescomObjects.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "..\CoreObjects.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "..\EntityProxy.ocx"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: ".\Componentes\COMCT332.OCX"; DestDir: "{sys}"; Flags: regserver
Source: ".\Componentes\Comdlg32.ocx"; DestDir: "{sys}"; Flags: regserver
Source: ".\Componentes\mscomctl.ocx"; DestDir: "{sys}"; Flags: regserver
Source: ".\Componentes\MSCOMCT2.OCX"; DestDir: "{sys}"; Flags: regserver
Source: ".\Componentes\MSMAPI32.OCX"; DestDir: "{sys}"; Flags: regserver
Source: ".\Componentes\GRIDEX20.OCX"; DestDir: "{sys}"; Flags: regserver
Source: ".\Componentes\TEXT2___.PFB"; DestDir: "{fonts}"; FontInstall: "TextileLH PiTwo"; Flags: onlyifdoesntexist uninsneveruninstall
;fontisnttruetype
Source: ".\Componentes\TEXT2___.PFM"; DestDir: "{fonts}"; FontInstall: "TextileLH PiTwo"; Flags: onlyifdoesntexist uninsneveruninstall
;fontisnttruetype


[INI]
Filename: "WIN.INI"; Section: "GesCom"; Key: "PERSIST_SERVER"; String: "HONGOPDC"; Flags: createkeyifdoesntexist uninsdeleteentry
Filename: "WIN.INI"; Section: "GesCom"; Key: "TerminalID"; String: "1"; Flags: createkeyifdoesntexist uninsdeleteentry


