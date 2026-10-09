; urlCheck_setup.iss -- installer for urlCheck, from the HomerDev template
; Templates\_APP__setup.iss (kit 1.43.2).
;
; What the template settles, kept here:
;   MACHINE WIDE, ADMINISTRATOR. Program Files, no per-user fallback. Somebody
;   who wants a portable copy takes the zip.
;   THE VERSION COMES FROM version.txt, read at compile time; no literal here.
;   THE HOMER LAYOUT: the program in exec\, the full guide and the other
;   documents in help\, ReadMe and License at the top, where a person
;   looking for them expects them.
;   THE RESULTS BOX COMES BEFORE THE LAUNCH. The Launch box only leaves a
;   marker; the program starts after the Results box has been read.
;
; urlCheck installs no components of its own: it drives the Microsoft Edge
; that Windows already has, and everything else is inside urlCheck.exe. The
; finish page therefore offers only Launch (ticked) and Open the user guide
; (unticked), which is what the installer before the kit offered, with the
; guide box now unticked as the kit's finish-page rule asks.

#define AppName       "urlCheck"

#define VerFile FileOpen(AddBackslash(SourcePath) + "version.txt")
#define AppVersion Trim(FileRead(VerFile))
#expr FileClose(VerFile)
#undef VerFile

#define AppPublisher  "Jamal Mazrui"
#define AppUrl        "https://github.com/JamalMazrui/urlCheck"
#define AppExeName    "urlCheck.exe"
#define AppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."

; The desktop shortcut's hotkey. HotKey is Inno's own syntax; HotKeyDisplay is
; the same key as a person reads it. Alt+Control+letter belongs to desktop
; shortcuts, and a shortcut's own hotkey is the sanctioned use of it.
; ALT+CONTROL+SHIFT+U SINCE 1.12.2. urlFido, the companion tool used more
; often, keeps Alt+Control+U, as urlFido's installer long said urlCheck
; would move; two desktop shortcuts with one hotkey leave Windows to give
; it to whichever was made last.
#define HotKey        "Alt+Ctrl+Shift+U"
#define HotKeyDisplay "Alt+Control+Shift+U"

; What the program is started with after the Results box: the dialog, with
; saved settings, exactly as the desktop shortcut starts it.
#define AppLaunchParams "-g -u"

; The build passes /DHomerDev=<kit>; the kit's component table is included
; from there in [Code].
#ifndef HomerDev
#define HomerDev "C:\HomerDev"
#endif

[Setup]
; Unchanged from every earlier urlCheck installer, so an upgrade finds the
; previous install.
AppId={{B2C4F1A8-3D9E-4F7B-8C5D-9E1A2B3C4D5E}

AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}
AppUpdatesURL={#AppUrl}/releases
AppCopyright={#AppCopyright}

; The version resource of the built setup. release reads the FileVersion
; STRING from it and tags v<that>, so the text form is set explicitly.
VersionInfoVersion={#AppVersion}
VersionInfoTextVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoProductTextVersion={#AppVersion}
VersionInfoCompany={#AppPublisher}
VersionInfoCopyright={#AppCopyright}
VersionInfoDescription={#AppName} Setup

DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
; A reinstall asks nothing and goes where the last one went; a first install
; still chooses the folder.
DisableDirPage=auto
UsePreviousGroup=yes

; Inno's own detailed log, copied at the end to
; %LOCALAPPDATA%\urlCheck\logs\urlCheck-setup-<yyyymmdd-hhmmss>.log.
SetupLogging=yes

OutputDir=.
OutputBaseFilename={#AppName}_setup
SolidCompression=yes
WizardStyle=modern
Compression=lzma2/max
MinVersion=10.0
AppComments=A Homer Tools program for keyboard and screen-reader users.
SetupIconFile={#AppName}.ico

PrivilegesRequired=admin
; THE PER-USER AREAS ARE USED ON PURPOSE -- the setup log and the launch marker
; -- so Inno's warning about them is turned off rather than read past on every
; build.
UsedUserAreasWarning=no
PrivilegesRequiredOverridesAllowed=

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\exec\{#AppExeName}
UninstallDisplayName={#AppName} {#AppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%n[name] is licensed under the MIT License: free to use, copy, modify, and distribute; provided "as is" with no warranty. The full license is installed as License.htm in the program folder.%n%nIt is recommended that you close all other applications before continuing.

[Dirs]
Name: "{app}\exec"
Name: "{app}\help"

[Files]
; THE PROGRAM goes in exec. Everything urlCheck needs is inside it.
Source: "exec\{#AppExeName}"; DestDir: "{app}\exec"; Flags: ignoreversion
; The command-line wrapper stays at the top, where typing urlCheck in the
; program folder finds it; it runs exec\urlCheck.exe.
Source: "{#AppName}.cmd"; DestDir: "{app}"; Flags: ignoreversion
; ReadMe and License at the top, both forms: Markdown for an editor or a
; braille display, HTML for a browser.
Source: "ReadMe.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.htm"; DestDir: "{app}"; Flags: ignoreversion
Source: "License.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "License.htm"; DestDir: "{app}"; Flags: ignoreversion
; Every other document is in help, and every one of them ships: the build
; checks that each file in help\ is matched by a line here.
Source: "help\*.md"; DestDir: "{app}\help"; Flags: ignoreversion
Source: "help\*.htm"; DestDir: "{app}\help"; Flags: ignoreversion
; The spoken tutorials, in the Homer pattern of ten, and their audio, made by the
; build before this is compiled (9 October 2026).
Source: "help\Tutorial_*.inix"; DestDir: "{app}\help"; Flags: ignoreversion skipifsourcedoesntexist
Source: "help\TutorialFeed.xml"; DestDir: "{app}\help"; Flags: ignoreversion skipifsourcedoesntexist
Source: "help\tutorials\*.mp3"; DestDir: "{app}\help\tutorials"; Flags: ignoreversion skipifsourcedoesntexist
Source: "help\tutorials\*.m3u"; DestDir: "{app}\help\tutorials"; Flags: ignoreversion skipifsourcedoesntexist
Source: "help\tutorials\*.sha256"; DestDir: "{app}\help\tutorials"; Flags: ignoreversion skipifsourcedoesntexist

[InstallDelete]
; What the installers before the kit put at the top of the program folder.
; The program is in exec now, and the documents in help; a stale copy left at
; the top would be the one a person found first.
Type: files; Name: "{app}\{#AppExeName}"
Type: files; Name: "{app}\Announce.htm"

[Icons]
; WorkingDir is Documents, so output folders land somewhere the person can
; write to; Program Files is not.
Name: "{group}\{#AppName}"; Filename: "{app}\exec\{#AppExeName}"; Parameters: "-g -u"; WorkingDir: "{userdocs}"; Comment: "Check web pages and HTML files for accessibility problems"
Name: "{group}\{#AppName} guide"; Filename: "{app}\help\{#AppName}.htm"; Comment: "The full guide to {#AppName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"; Comment: "Remove {#AppName} from this computer"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\exec\{#AppExeName}"; Parameters: "-g -u"; WorkingDir: "{userdocs}"; HotKey: "{#HotKey}"; Comment: "Check accessibility ({#HotKeyDisplay})"

[Run]
; FINISH-PAGE ORDER, the HomerDev rule: Install, Update, Reinstall groups
; (none here: urlCheck installs no components), then Launch ticked, then Open
; the user guide unticked.
;
; Launch only leaves a marker. The program starts from CurStepChanged(ssDone),
; after the Results box has been read and closed. TWO pairs of quotes: this
; Parameters value starts with a quote, the one case where cmd /s strips the
; outer pair correctly.
FileName: "{cmd}"; \
  Parameters: "/c echo launch > ""{localappdata}\{#AppName}\logs\{#AppName}_launch.flag"""; \
  Description: "Launch {#AppName} (desktop hotkey {#HotKeyDisplay})"; \
  Flags: postinstall skipifsilent runhidden runasoriginaluser

FileName: "{app}\help\{#AppName}.htm"; \
  Description: "Open the user guide (F1 in {#AppName})"; \
  Flags: postinstall shellexec nowait skipifsilent skipifdoesntexist runasoriginaluser unchecked

[UninstallDelete]
; ONLY WHAT urlCheck WROTE: its logs and its saved settings, including the
; urlCheck.ini of the versions before the kit. Never the whole folder.
Type: filesandordirs; Name: "{localappdata}\{#AppName}\logs"
Type: files; Name: "{localappdata}\{#AppName}\configs\{#AppName}.inix"
Type: dirifempty; Name: "{localappdata}\{#AppName}\configs"
Type: files; Name: "{localappdata}\{#AppName}\urlCheck.ini"
Type: dirifempty; Name: "{localappdata}\{#AppName}"

[Code]
//  THE HOMER INSTALLER PATTERN, from the kit's Templates\_APP__setup.iss:
//    - the kit's component table, HomerComponents.iss, included INSIDE [Code]
//      and from the kit the build names. urlCheck registers no component --
//      it drives the Edge that Windows keeps current itself, and the
//      installer before the kit offered none -- so the table is empty; a
//      component added later is one homerAdd and three [Run] entries;
//    - the ticked boxes recorded when Finish is pressed (homerNoteTicked), so
//      the Results box reports only what was asked for this session;
//    - the Results box first, the launch after it.
//
//  WHAT THE CODE SECTION DOES:
//    - reads the version of any previous install, so the welcome page says
//      Install, Update or Reinstall truthfully;
//    - keeps the setup log with the program's own logs;
//    - shows one Results box of what was done, then starts urlCheck if the
//      Launch box was ticked -- in that order, so the box is not hidden behind
//      the program's window.
//  Comments inside [Code] use // or (* *), never ;.
#include HomerDev + "\Templates\HomerComponents.iss"

var
  sActions: String;
  sPriorVersion: String;

function priorVersion(): String;
var
  sKey, sFound: String;
begin
  Result := '';
  sKey := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1';
  if RegQueryStringValue(HKLM, sKey, 'DisplayVersion', sFound) then Result := sFound
  else if RegQueryStringValue(HKCU, sKey, 'DisplayVersion', sFound) then Result := sFound;
end;

//  The previous install is read ONCE, before anything is installed: read
//  again at the end it would find this very install and call it a reinstall.
function InitializeSetup(): Boolean;
begin
  sPriorVersion := priorVersion();
  Result := True;
end;

function isFreshInstall(): Boolean;
begin
  Result := (sPriorVersion = '');
end;

//  versionPart: the Nth dotted number of a version, or 0 where there is none.
//  Written out rather than using PackVersionString, which not every Inno Setup
//  6 has.
function versionPart(sVersion: String; iWanted: Integer): Integer;
var
  iAt, iPart: Integer;
  sNumber: String;
begin
  Result := 0;
  iPart := 1;
  sNumber := '';
  for iAt := 1 to Length(sVersion) do
  begin
    if sVersion[iAt] = '.' then
    begin
      if iPart = iWanted then begin Result := StrToIntDef(sNumber, 0); exit; end;
      iPart := iPart + 1;
      sNumber := '';
    end
    else if (sVersion[iAt] >= '0') and (sVersion[iAt] <= '9') then
      sNumber := sNumber + sVersion[iAt];
  end;
  if iPart = iWanted then Result := StrToIntDef(sNumber, 0);
end;

function versionIsOlder(sHave, sWant: String): Boolean;
var
  iPart, iHave, iWant: Integer;
begin
  Result := False;
  for iPart := 1 to 4 do
  begin
    iHave := versionPart(sHave, iPart);
    iWant := versionPart(sWant, iPart);
    if iHave < iWant then begin Result := True; exit; end;
    if iHave > iWant then exit;
  end;
end;

function isOlderInstalled(): Boolean;
begin
  Result := False;
  if isFreshInstall() then exit;
  Result := versionIsOlder(sPriorVersion, '{#AppVersion}');
end;

procedure addAction(sText: String);
begin
  if sText = '' then exit;
  if sActions <> '' then sActions := sActions + #13#10;
  sActions := sActions + '  ' + sText;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
//  Finish pressed: the boxes are settled, and nothing on the finish page has
//  run yet.
begin
  Result := True;
  if CurPageID = wpFinished then homerNoteTicked();
end;

procedure startIfAsked();
//  Starts urlCheck's dialog if the Launch box left its marker, and removes the
//  marker. Started through cmd so it runs as the person, not as the elevated
//  installer, and with Documents as its folder like the shortcuts.
var
  sFlag: String;
  iResult: Integer;
begin
  sFlag := ExpandConstant('{localappdata}\{#AppName}\logs\{#AppName}_launch.flag');
  if not FileExists(sFlag) then exit;
  DeleteFile(sFlag);
  Exec(ExpandConstant('{cmd}'),
       '/s /c ""' + ExpandConstant('{app}\exec\{#AppExeName}') + '" {#AppLaunchParams}"',
       ExpandConstant('{userdocs}'), SW_SHOW, ewNoWait, iResult);
end;

procedure reportWhatHappened();
var
  sBody: String;
begin
  if isFreshInstall() then
    sBody := 'Installed {#AppName} {#AppVersion}.'
  else if isOlderInstalled() then
    sBody := 'Updated {#AppName} from ' + sPriorVersion + ' to {#AppVersion}.'
  else
    sBody := 'Reinstalled {#AppName} {#AppVersion}.';
  //  One addAction per component, in the order of the [Run] section; none yet.
  if sActions <> '' then sBody := sBody + #13#10 + #13#10 + sActions;
  sBody := sBody + #13#10 + #13#10
         + 'Logs are kept in ' + ExpandConstant('{localappdata}\{#AppName}\logs') + '.';
  homerResultsBox(sBody);
  startIfAsked();
end;

procedure keepSetupLog();
//  Inno writes its log to the temporary folder, where nobody finds it. The
//  installer runs elevated, so {localappdata} is the profile of whoever
//  answered the elevation prompt.
var
  sFolder, sTarget: String;
begin
  sFolder := ExpandConstant('{localappdata}\{#AppName}\logs');
  if not DirExists(sFolder) then
    if not ForceDirectories(sFolder) then exit;
  sTarget := sFolder + '\{#AppName}-setup-' + GetDateTimeString('yyyymmdd-hhnnss', #0, #0) + '.log';
  CopyFile(ExpandConstant('{log}'), sTarget, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssDone then
  begin
    keepSetupLog();
    reportWhatHappened();
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
//  Say on the welcome page what is about to happen, because a user reading by
//  ear should not have to work it out from a version number in a caption.
begin
  if CurPageID = wpWelcome then
  begin
    if isFreshInstall() then
      WizardForm.WelcomeLabel1.Caption := 'Install {#AppName} {#AppVersion}'
    else if isOlderInstalled() then
      WizardForm.WelcomeLabel1.Caption := 'Update {#AppName} from ' + sPriorVersion + ' to {#AppVersion}'
    else
      WizardForm.WelcomeLabel1.Caption := 'Reinstall {#AppName} {#AppVersion}';
  end;
end;
