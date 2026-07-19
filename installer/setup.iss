; =============================================================================
; CSV.Convertor — Inno Setup скрипт
; Генерирует установщик Setup.exe с русским интерфейсом
; =============================================================================

#define MyAppName "CSV.Convertor"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "CSV.Convertor"
#define MyAppURL "https://github.com/LucasVakhos/CSV.Convertor"
#define MyAppExeName "CSV.Convertor.exe"

[Setup]
; Уникальный идентификатор приложения (генерируется один раз)
AppId={{B8F4A3D2-7C1E-4E5A-9D6F-2B8C0E7A5F3D}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Папка установки по умолчанию — Program Files
DefaultDirName={autopf}\{#MyAppName}

; Файл лицензии (показывается на странице с соглашением)
LicenseFile=..\LICENSE.txt

; Иконка установщика и приложения в списке программ
SetupIconFile=..\EXCEL_257.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Сжатие и выходной файл
Compression=lzma2/ultra64
SolidCompression=yes
OutputDir=Output
OutputBaseFilename=CSV.Convertor.Setup

; Интерфейс — только русский
; Требуется файл Languages\Russian.isl в папке Inno Setup
WizardStyle=modern

; Запрет запуска нескольких копий установщика
DisableWelcomePage=no

; Требования к ОС
MinVersion=10.0.17763

; Архитектура
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
; Ярлык на рабочем столе (отмечен по умолчанию)
Name: "desktopicon"; Description: "Создать ярлык на &рабочем столе"; GroupDescription: "Дополнительные ярлыки:"; Flags: checkedonce

[Files]
; Основной исполняемый файл (single-file self-contained .exe)
; Путь относительно папки installer\ — ведёт к выходной папке dotnet publish
Source: "..\bin\Release\net10.0-windows\win-x64\publish\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Ярлык в меню «Пуск»
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

; Ярлык на рабочем столе (если выбрана соответствующая задача)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

; Ярлык для удаления в меню «Пуск»
Name: "{autoprograms}\Удалить {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
; Запустить приложение после установки (с галочкой по умолчанию)
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Удалить папку приложения, если она пуста (пользовательские файлы в Документах не затрагиваются)
Type: dirifempty; Name: "{app}"

[Code]
// Проверка, что .exe существует перед установкой
function InitializeSetup: Boolean;
begin
  if not FileExists(ExpandConstant('{src}\..\bin\Release\net10.0-windows\win-x64\publish\{#MyAppExeName}')) then
  begin
    MsgBox('Не найден файл {#MyAppExeName} в папке публикации.' + #13#10 +
           'Сначала выполните dotnet publish -c Release.', mbError, MB_OK);
    Result := False;
  end
  else
    Result := True;
end;
