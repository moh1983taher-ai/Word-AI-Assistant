#define AppName "ResearchTools Setup"
#define AppVersion "1.1.0"

[Setup]
AppId={{6F31C2B4-8D31-4C81-97A9-4B9E7932C8C1}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=ResearchTools
DefaultDirName={localappdata}\ResearchTools\Setup
DisableDirPage=yes
DisableProgramGroupPage=yes
CreateAppDir=no
Uninstallable=no
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64
OutputDir=..\dist
OutputBaseFilename=ResearchTools-Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
CloseApplications=no
UsePreviousAppDir=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\dist\DuckAI-Runtime-release.zip"; DestDir: "{tmp}"; Flags: dontcopy
Source: "..\dist\DuckAI-Runtime-release.zip.sha256"; DestDir: "{tmp}"; Flags: dontcopy
Source: "..\dist\ResearchTools-MCP-Host-release.zip"; DestDir: "{tmp}"; Flags: dontcopy
Source: "..\dist\ResearchTools-MCP-Host-release.zip.sha256"; DestDir: "{tmp}"; Flags: dontcopy
Source: "ResearchTools-Launcher.vbs"; DestDir: "{tmp}"; Flags: dontcopy
Source: "Install-ResearchTools.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "Word-AI-Assistant.xml"; DestDir: "{tmp}"; Flags: dontcopy

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  PowerShellPath: String;
  Parameters: String;
begin
  if CurStep = ssPostInstall then
  begin
    ExtractTemporaryFile('DuckAI-Runtime-release.zip');
    ExtractTemporaryFile('DuckAI-Runtime-release.zip.sha256');
    ExtractTemporaryFile('ResearchTools-MCP-Host-release.zip');
    ExtractTemporaryFile('ResearchTools-MCP-Host-release.zip.sha256');
    ExtractTemporaryFile('ResearchTools-Launcher.vbs');
    ExtractTemporaryFile('Install-ResearchTools.ps1');
    ExtractTemporaryFile('Word-AI-Assistant.xml');

    PowerShellPath := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
    Parameters := '-NoLogo -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass' +
      ' -File "' + ExpandConstant('{tmp}\Install-ResearchTools.ps1') + '"' +
      ' -RuntimeArchive "' + ExpandConstant('{tmp}\DuckAI-Runtime-release.zip') + '"' +
      ' -RuntimeShaFile "' + ExpandConstant('{tmp}\DuckAI-Runtime-release.zip.sha256') + '"' +
      ' -McpHostArchive "' + ExpandConstant('{tmp}\ResearchTools-MCP-Host-release.zip') + '"' +
      ' -McpHostShaFile "' + ExpandConstant('{tmp}\ResearchTools-MCP-Host-release.zip.sha256') + '"' +
      ' -LauncherFile "' + ExpandConstant('{tmp}\ResearchTools-Launcher.vbs') + '"' +
      ' -ManifestFile "' + ExpandConstant('{tmp}\Word-AI-Assistant.xml') + '"';

    WizardForm.StatusLabel.Caption := 'Installing ResearchTools, DuckAI, MCP host and Word add-in...';
    WizardForm.Refresh;

    if not Exec(PowerShellPath, Parameters, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      MsgBox('The ResearchTools configuration process could not be started.', mbError, MB_OK);
      RaiseException('ResearchTools configuration process failed to start.');
    end;

    if ResultCode <> 0 then
    begin
      MsgBox('Setup did not complete. Review the installation log at:' + #13#10 +
        ExpandConstant('{localappdata}\ResearchTools\Logs\ResearchTools-Setup.log'),
        mbError, MB_OK);
      RaiseException('ResearchTools setup failed with exit code ' + IntToStr(ResultCode) + '.');
    end;
  end;
end;
