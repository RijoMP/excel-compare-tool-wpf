[Setup]
AppName=Excel Compare Tool
AppVersion=2.1
DefaultDirName={pf}\ExcelCompare
DefaultGroupName=Excel Compare Tool
OutputDir=.
OutputBaseFilename=ExcelCompareInstaller_v2_1
Compression=lzma
SolidCompression=yes

[Files]
Source: "M:\WPF\excelcompare\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Excel Compare Tool"; Filename: "{app}\excelcompare.exe"
Name: "{commondesktop}\Excel Compare Tool"; Filename: "{app}\excelcompare.exe"

[Run]
Filename: "{app}\excelcompare.exe"; Description: "Launch Excel Compare Tool"; Flags: nowait postinstall skipifsilent