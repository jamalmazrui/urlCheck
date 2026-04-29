; =====================================================================
; urlCheck installer script for Inno Setup 6.x
;
; Compile with the Inno Setup IDE (ISCC.exe) to produce urlCheck_setup.exe.
; The resulting installer:
;   - Requires administrator privileges.
;   - Installs urlCheck.exe and supporting documentation files to
;     C:\Program Files\urlCheck (standard GUI-program install path).
;   - Registers the product for "Apps & Features" uninstall.
;   - Creates a desktop shortcut with hotkey Alt+Ctrl+U that
;     launches urlCheck in GUI mode with saved-configuration loading
;     enabled (equivalent to urlCheck -g -u).
;   - Does NOT register a File Explorer right-click verb. urlCheck
;     accepts a URL, a single local HTML file, or a single plain text
;     file listing URLs one per line; that range of inputs is best
;     entered through the GUI dialog or the command line, not through
;     a per-extension shell verb.
;   - On the final wizard page, offers two PostInstall checkboxes
;     (both checked by default): launch urlCheck, and read the HTML
;     documentation.
; =====================================================================

#define cAppName       "urlCheck"
#define cAppVersion    "1.10.0"
#define cAppPublisher  "Jamal Mazrui"
#define cAppUrl        "https://github.com/JamalMazrui/urlCheck"
#define cAppExeName    "urlCheck.exe"
#define cAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."

[Setup]
AppId={{B2C4F1A8-3D9E-4F7B-8C5D-9E1A2B3C4D5E}

AppName={#cAppName}
AppVersion={#cAppVersion}
AppVerName={#cAppName} {#cAppVersion}
AppPublisher={#cAppPublisher}
AppPublisherURL={#cAppUrl}
AppSupportURL={#cAppUrl}
AppUpdatesURL={#cAppUrl}/releases
AppCopyright={#cAppCopyright}
VersionInfoVersion={#cAppVersion}

; Install under Program Files (standard GUI-program location).
DefaultDirName={pf}\{#cAppName}
DefaultGroupName={#cAppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
UsePreviousGroup=yes

OutputDir=.
OutputBaseFilename=urlCheck_setup
Compression=lzma2
SolidCompression=yes
SetupIconFile=urlCheck.ico

; Installer requires admin to write to Program Files.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=

; 64-bit Windows only.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\{#cAppExeName}
UninstallDisplayName={#cAppName} {#cAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Note: urlCheck.ico is NOT copied to {app} because PyInstaller embeds it
; directly into urlCheck.exe at build time (via --icon=urlCheck.ico in
; buildUrlCheck.cmd). The shortcut icons in [Icons] inherit from the exe's
; embedded icon by default. The .ico file IS still needed at COMPILE time
; for the SetupIconFile= directive above, which gives urlCheck_setup.exe
; itself an icon -- but that's a compile-time dependency only and does not
; need to ship with the installed program.
Source: "urlCheck.exe";       DestDir: "{app}"; Flags: ignoreversion
Source: "urlCheck.cmd";       DestDir: "{app}"; Flags: ignoreversion
Source: "README.htm";         DestDir: "{app}"; Flags: ignoreversion
Source: "README.md";          DestDir: "{app}"; Flags: ignoreversion
Source: "license.htm";        DestDir: "{app}"; Flags: ignoreversion
Source: "announce.md";        DestDir: "{app}"; Flags: ignoreversion
Source: "announce.htm";       DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "urlCheck.py";        DestDir: "{app}"; Flags: ignoreversion
Source: "buildUrlCheck.cmd";  DestDir: "{app}"; Flags: ignoreversion
Source: "urlCheck.iss";       DestDir: "{app}"; Flags: ignoreversion
Source: "requirements.txt";   DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group. WorkingDir is set to the user's Documents folder so
; output folders and the optional urlCheck.log land somewhere writable
; (the install dir under Program Files is not writable for non-admins).
Name: "{group}\{#cAppName}"; \
  Filename: "{app}\{#cAppExeName}"; \
  Parameters: "-g -u"; \
  WorkingDir: "{userdocs}"; \
  Comment: "Check web pages and HTML files for accessibility problems"

Name: "{group}\{#cAppName} README"; \
  Filename: "{app}\README.htm"; \
  WorkingDir: "{app}"; \
  Comment: "Documentation for {#cAppName}"

Name: "{group}\Uninstall {#cAppName}"; \
  Filename: "{uninstallexe}"; \
  Comment: "Remove {#cAppName} from this computer"

; Desktop shortcut with the Alt+Ctrl+U hotkey. Launches urlCheck in
; GUI mode (-g) with saved-configuration loading (-u). The hotkey is
; not used by Windows or major office applications by default, but
; individual applications may intercept it when they have focus.
; WorkingDir is the user's Documents folder for the same writability
; reason as the Start Menu shortcut above.
Name: "{userdesktop}\{#cAppName}"; \
  Filename: "{app}\{#cAppExeName}"; \
  WorkingDir: "{userdocs}"; \
  Parameters: "-g -u"; \
  HotKey: Alt+Ctrl+U; \
  Comment: "Check accessibility (Alt+Ctrl+U)"

[Run]
; Post-install checkboxes shown on the final wizard page. Both
; default to checked; the user can uncheck either to skip.

; Launch urlCheck (GUI mode). WorkingDir is the user's Documents folder
; so any output folders or log file land somewhere writable.
FileName: "{app}\{#cAppExeName}"; \
  Parameters: "-g"; \
  WorkingDir: "{userdocs}"; \
  Description: "Launch {#cAppName} now"; \
  Flags: nowait postinstall skipifsilent

; Open the HTML documentation.
FileName: "{app}\README.htm"; \
  Description: "Read documentation for {#cAppName}"; \
  Flags: postinstall shellexec skipifsilent
