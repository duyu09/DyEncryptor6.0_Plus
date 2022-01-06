; �ű��� Inno Setup �ű��� ���ɣ�
; �йش��� Inno Setup �ű��ļ�����ϸ��������İ����ĵ���

#define MyAppName "DyEncryptor5.0�ļ�����ϵͳ"
#define MyAppVersion "5.0.0.47"
#define MyAppPublisher "��³��ҵ��ѧ ������̿���1�� ����"
#define MyAppExeName "DyEncryptor5.0_Plus.exe"

[Setup]
; ע: AppId��ֵΪ������ʶ��Ӧ�ó���
; ��ҪΪ������װ����ʹ����ͬ��AppIdֵ��
; (��Ҫ�����µ� GUID�����ڲ˵��е�� "����|���� GUID"��)
AppId={{0AF87FD2-DF38-417B-8249-1B3B1C364B41}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
VersionInfoVersion={#MyAppVersion}
VersionInfoTextVersion={#MyAppVersion}
AppCopyright={#MyAppPublisher}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; [Icons] �ġ�quicklaunchicon����Ŀʹ�� {userappdata}������ [Tasks] ��Ŀ�����ʺ� IsAdminInstallMode �ļ�顣
UsedUserAreasWarning=no
InfoBeforeFile=C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEnc5.0��װ���.txt
; ������ȡ��ע�ͣ����ڷǹ���װģʽ�����У���Ϊ��ǰ�û���װ����
;PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=C:\Users\Administrator\Desktop
OutputBaseFilename=DyEncryptor5.0�ļ�����ϵͳ
SetupIconFile=C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncIcon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"
Name: "armenian"; MessagesFile: "compiler:Languages\Armenian.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "catalan"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "corsican"; MessagesFile: "compiler:Languages\Corsican.isl"
Name: "czech"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "danish"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "english"; MessagesFile: "compiler:Languages\English.isl"
Name: "finnish"; MessagesFile: "compiler:Languages\Finnish.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"
Name: "icelandic"; MessagesFile: "compiler:Languages\Icelandic.isl"
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "norwegian"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "slovenian"; MessagesFile: "compiler:Languages\Slovenian.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\Dy_EncCore.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEnc_BulidEXE.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEnc_FileDestroyModule.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEnc5.0.HISTORY"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEnc5.0��װ���.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEnc5_BuildEXE_Core.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncGUI5.0.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncGUI5.0.OtherSettings.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncIcon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncryptor5.0_installini.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncryptor5.0_Plus.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\GUI_Color.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\MSVBVM60.DLL"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\msvcrt.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\Rar.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\RarExt.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\rarreg.key"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\WinRAR.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncGUI_FontsLib\*"; DestDir: "{app}\DyEncGUI_FontsLib"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\Administrator\Desktop\DyEnc5.0_Plus\DyEncGUI_IconLib\*"; DestDir: "{app}\DyEncGUI_IconLib"; Flags: ignoreversion recursesubdirs createallsubdirs
; ע��: ��Ҫ���κι���ϵͳ�ļ���ʹ�á�Flags: ignoreversion��

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\DyEncryptor5.0_installini.bat";WorkingDir:"{app}";StatusMsg:"���������ļ�����"
