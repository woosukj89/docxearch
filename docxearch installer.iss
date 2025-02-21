; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "docxearch"
#define MyAppVersion "1.5"
#define MyAppPublisher "Oh Jung Il"
#define MyAppExeName "docxearch.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{C73BFFF7-9DB9-42DE-B040-36C6C4C8D877}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=C:\Users\woosu\Projects\Apps\doc search app\docxearch license.txt
; Remove the following line to run in administrative install mode (install for all users.)
PrivilegesRequired=lowest
OutputBaseFilename=mysetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_asyncio.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_bz2.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_ctypes.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_decimal.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_hashlib.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_lzma.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_multiprocessing.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_overlapped.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_queue.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_socket.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_ssl.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\_win32sysloader.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-console-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-datetime-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-debug-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-errorhandling-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-file-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-file-l1-2-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-file-l2-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-handle-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-heap-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-interlocked-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-libraryloader-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-localization-l1-2-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-memory-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-namedpipe-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-processenvironment-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-processthreads-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-processthreads-l1-1-1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-profile-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-rtlsupport-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-string-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-synch-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-synch-l1-2-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-sysinfo-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-timezone-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-core-util-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-conio-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-convert-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-environment-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-filesystem-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-heap-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-locale-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-math-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-multibyte-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-process-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-runtime-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-stdio-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-string-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-time-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\api-ms-win-crt-utility-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\base_library.zip"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\ffi.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\LIBBZ2.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\liblzma.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\mfc140u.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\MSVCP140.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\pyexpat.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\python3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\python310.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\select.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\ucrtbase.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\unicodedata.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\VCRUNTIME140.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\VCRUNTIME140_1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\win32api.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\win32trace.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\win32ui.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\woosu\Projects\Apps\doc search app\dist\docxearch\zlib.dll"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

