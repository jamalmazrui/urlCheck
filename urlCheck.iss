; =====================================================================
; urlCheck installer script for Inno Setup 6.x
;
; Compile with the Inno Setup IDE (ISCC.exe) to produce urlCheck_setup.exe.
; The resulting installer:
;   - Requires administrator privileges.
;   - Prompts the user for the installation directory; default is
;     C:\Program Files\urlCheck.
;   - Shows a brief MIT license summary on the welcome page (no extra
;     wizard screen). The full license text is installed alongside
;     the program as License.htm.
;   - Registers the product for "Apps & Features" uninstall.
;   - Creates a desktop shortcut with hotkey Alt+Ctrl+U that
;     launches urlCheck in GUI mode with saved-configuration loading
;     enabled (equivalent to urlCheck -g -u).
;   - Does NOT register a File Explorer right-click verb.
;   - On the final wizard page, offers two PostInstall checkboxes
;     (both checked by default): launch urlCheck (with a hotkey
;     reminder), and read the HTML documentation.
;
; This installer ships only the runtime distribution (the .exe, the
; documentation in HTML form, and the license). The Markdown sources,
; the Python source, the build script, and this .iss script live in
; the GitHub repository.
; =====================================================================

#define sAppName       "urlCheck"
#define sAppVersion    "1.10.0"
#define sAppPublisher  "Jamal Mazrui"
#define sAppUrl        "https://github.com/JamalMazrui/urlCheck"
#define sAppExeName    "urlCheck.exe"
#define sAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."
#define sHotKey        "Alt+Ctrl+U"

[Setup]
AppId={{B2C4F1A8-3D9E-4F7B-8C5D-9E1A2B3C4D5E}

AppName={#sAppName}
AppVersion={#sAppVersion}
AppVerName={#sAppName} {#sAppVersion}
AppPublisher={#sAppPublisher}
AppPublisherURL={#sAppUrl}
AppSupportURL={#sAppUrl}
AppUpdatesURL={#sAppUrl}/releases
AppCopyright={#sAppCopyright}
VersionInfoVersion={#sAppVersion}

; Install under Program Files. {autopf} resolves to "Program Files"
; on 64-bit Windows when the installer runs in 64-bit mode (see
; ArchitecturesInstallIn64BitMode below). The user can override this
; default on the wizard's directory page.
DefaultDirName={autopf}\{#sAppName}
DefaultGroupName={#sAppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
UsePreviousGroup=yes

OutputDir=.
OutputBaseFilename={#sAppName}_setup
Compression=lzma2
SolidCompression=yes
SetupIconFile={#sAppName}.ico
WizardStyle=modern

; Installer requires admin to write to Program Files.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=

; 64-bit Windows only.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\{#sAppExeName}
UninstallDisplayName={#sAppName} {#sAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
; Replace the default welcome-page body text with one that includes a
; brief MIT license notice. This satisfies the requirement that the
; license summary appear on an existing wizard screen rather than on
; an additional dedicated page (which is what LicenseFile= would
; produce). The full license text is installed alongside the program.
WelcomeLabel2=This will install [name/ver] on your computer.%n%n[name] is licensed under the MIT License: free to use, copy, modify, and distribute; provided "as is" with no warranty. The full license text will be installed as License.htm in the program folder.%n%nIt is recommended that you close all other applications before continuing.

[Files]
; The runtime distribution: just the executable, the cosmetic-warning
; suppression wrapper, the HTML docs, and the license. The icon is
; embedded in urlCheck.exe at build time (PyInstaller --icon flag),
; so the .ico does not need to ship in the install directory.
Source: "{#sAppName}.exe";    DestDir: "{app}"; Flags: ignoreversion
Source: "{#sAppName}.cmd";    DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.htm";         DestDir: "{app}"; Flags: ignoreversion
Source: "Announce.htm";       DestDir: "{app}"; Flags: ignoreversion
Source: "License.htm";        DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group. WorkingDir is set to the user's Documents folder so
; output folders and the optional urlCheck.log land somewhere writable
; (the install dir under Program Files is not writable for non-admins).
Name: "{group}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  Parameters: "-g -u"; \
  WorkingDir: "{userdocs}"; \
  Comment: "Check web pages and HTML files for accessibility problems"

Name: "{group}\{#sAppName} ReadMe"; \
  Filename: "{app}\ReadMe.htm"; \
  WorkingDir: "{app}"; \
  Comment: "Documentation for {#sAppName}"

Name: "{group}\Uninstall {#sAppName}"; \
  Filename: "{uninstallexe}"; \
  Comment: "Remove {#sAppName} from this computer"

; Desktop shortcut with the Alt+Ctrl+U hotkey. Launches urlCheck in
; GUI mode (-g) with saved-configuration loading (-u). The hotkey is
; not used by Windows or major office applications by default, but
; individual applications may intercept it when they have focus.
; WorkingDir is the user's Documents folder for the same writability
; reason as the Start Menu shortcut above.
Name: "{userdesktop}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  WorkingDir: "{userdocs}"; \
  Parameters: "-g -u"; \
  HotKey: {#sHotKey}; \
  Comment: "Check accessibility ({#sHotKey})"

[Run]
; Post-install checkboxes shown on the final wizard page. Both
; default to checked; the user can uncheck either to skip. The launch
; checkbox label includes a reminder of the desktop hotkey so the
; user notices and remembers it.

FileName: "{app}\{#sAppExeName}"; \
  Parameters: "-g"; \
  WorkingDir: "{userdocs}"; \
  Description: "Launch {#sAppName} now (desktop hotkey: {#sHotKey})"; \
  Flags: nowait postinstall skipifsilent

FileName: "{app}\ReadMe.htm"; \
  Description: "Read documentation for {#sAppName}"; \
  Flags: postinstall shellexec skipifsilent
