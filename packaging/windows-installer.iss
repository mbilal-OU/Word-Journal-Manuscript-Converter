#define MyAppName "Word Journal Manuscript Converter"
#define MyAppVersion "0.5.0-beta.1"
#define MyAppPublisher "Muhammad Bilal"
#define MyAppURL "https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/"
#define MyAppExeName "Word-Journal-Manuscript-Converter.exe"

[Setup]
AppId={{E6D3B6B7-8D3E-4D3E-A1C1-59D7E16D6D13}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} - Pre-Launch Beta 1
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\Word Journal Manuscript Converter
DefaultGroupName=Word Journal Manuscript Converter
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\release
OutputBaseFilename=Word-Journal-Manuscript-Converter-Setup
SetupIconFile=..\src\word_journal_manuscript_converter\assets\app-icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\dist\Word-Journal-Manuscript-Converter.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\word-journal-converter.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\PRIVACY.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\integrations\word-addin\manifest.xml"; DestDir: "{app}\word-addin"; Flags: ignoreversion
Source: "..\docs\WORD_ADDIN.md"; DestDir: "{app}\word-addin"; Flags: ignoreversion

[Icons]
Name: "{group}\Word Journal Manuscript Converter"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Word Add-in Guide"; Filename: "https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/word-addin/"
Name: "{group}\Uninstall Word Journal Manuscript Converter"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Word Journal Manuscript Converter"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Word Journal Manuscript Converter"; Flags: nowait postinstall skipifsilent
