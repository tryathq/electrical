; Inno Setup script for Electrical Report app
; Prerequisite: Run build_for_customer.bat first so ElectricalReport_Ready folder exists.
; Then open this .iss in Inno Setup Compiler (or run: iscc installer.iss) to create setup.exe

#define MyAppName "Back Down Calculator"
#define MyAppVersion "1.0"
#define MyAppPublisher "Your Company"
#define MyAppURL "https://example.com"
#define MyAppExeName "START_APP.bat"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Installer size can be large (includes Python venv)
Compression=lzma2/ultra64
SolidCompression=yes
OutputDir=installer_output
OutputBaseFilename=BackDownCalculator_Setup
; Optional: add app.ico in this folder and uncomment next line for custom installer/icon
; SetupIconFile=app.ico
WizardStyle=modern
PrivilegesRequired=lowest
; Optional: use an .ico file for the installer/shortcut (put app.ico in this folder and uncomment next line)
; WizardImageFile=app.bmp
; WizardSmallImageFile=app_small.bmp

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Pack everything from the pre-built customer folder (create it first with build_for_customer.bat)
Source: "BackDownCalculator_Ready\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu shortcut: runs START_APP.bat from the install folder
Name: "{group}\{#MyAppName}"; Filename: "{cmd}"; Parameters: "/c ""{app}\START_APP.bat"""; WorkingDir: "{app}"; Comment: "Generate calculation sheet for BD and non compliance"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; Desktop shortcut (optional, off by default)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{cmd}"; Parameters: "/c ""{app}\START_APP.bat"""; WorkingDir: "{app}"; Tasks: desktopicon; Comment: "Generate calculation sheet for BD and non compliance"

[Run]
; Optional: launch app after install
Filename: "{cmd}"; Parameters: "/c ""{app}\START_APP.bat"""; WorkingDir: "{app}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}"
