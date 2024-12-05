unit uToSiteMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian,
  pFIBQuery,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green, Data.DB,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black, FIB,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, pFIBDataSet,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxLabel, dxGDIPlusClasses, Vcl.ExtCtrls, cxProgressBar,
  cxTextEdit, cxMemo, FIBDatabase, pFIBDatabase, inifiles, pFIBErrorHandler,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdFTPList,
  IdExplicitTLSClientServerBase, IdFTP, IdFTPCommon, IdHTTP, Vcl.StdCtrls,
  dxSkinMetropolis, dxSkinMetropolisDark, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinOffice2016Colorful,
  dxSkinOffice2016Dark, dxSkinVisualStudio2013Blue, dxSkinVisualStudio2013Dark,
  dxSkinVisualStudio2013Light, pFIBProps, System.UITypes, System.Types,
  frxClass, frxPreview,
  frxDBSet, frxChart, frxCrypt, frxGZip, frxDMPExport,
  frxGradient, frxChBox, frxCross, frxRich, frxOLE, frxBarcode, frxDCtrl,
  frxExportODF, frxExportMail, frxExportCSV, frxExportText, frxExportImage,
  frxExportRTF, frxExportXML, frxExportXLS, frxExportHTML, frxExportPDF,
  frxExportDBF, frxExportBIFF, frxExportDOCX, frxExportSVG,
  frxExportHTMLDiv, frxExportXLSX, frxExportPPTX,
  fs_iilparser, IB_Services, FIBQuery, pFIBStoredProc, Vcl.Menus, Vcl.AppEvnts,
  dxSkinTheBezier, shellapi, DateUtils, FIBDataSet, SynEdit, HTTPApp, JSon,
  ComObj;

type
  TAppParams = record
    bServerLocal: boolean;
    sServerName: string;
    sCharset: string;
    sProtocol: string;
    sFileName: string;
    bDebug: boolean;
    bBackupShow: boolean;
    bBackupCopy2OnlyOne: boolean;
    sBackupFilesMask: string;
  end;

  TfmToSiteMain = class(TForm)
    mmLog: TcxMemo;
    trUpdate: TpFIBTransaction;
    trMain: TpFIBTransaction;
    dbMain: TpFIBDatabase;
    trPingServer: TTimer;
    FibErrorHandler: TpFibErrorHandler;
    IdFTP: TIdFTP;
    IdHTTP: TIdHTTP;
    Button1: TButton;
    mmWork22: TcxMemo;
    BackupService1: TpFIBBackupService;
    pFIBStoredProc1: TpFIBStoredProc;
    ti: TTrayIcon;
    ae: TApplicationEvents;
    pm: TPopupMenu;
    miShowWin: TMenuItem;
    N3: TMenuItem;
    miExit: TMenuItem;
    Panel1: TPanel;
    Image2: TImage;
    lbStatus: TcxLabel;
    cxProgressBar: TcxProgressBar;
    tbScreens: TpFIBDataSet;
    cxProgressBar2: TcxProgressBar;
    mmWork2: TSynEdit;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure trPingServerTimer(Sender: TObject);
    procedure dbMainAfterRestoreConnect(Database: TFIBDatabase);
    procedure dbMainLostConnect(Database: TFIBDatabase; E: EFIBError;
      var Actions: TOnLostConnectActions; var DoRaise: boolean);
    procedure dbMainErrorRestoreConnect(Database: TFIBDatabase; E: EFIBError;
      var Actions: TOnLostConnectActions; var DoRaise: boolean);
    procedure FibErrorHandlerFIBErrorEvent(Sender: TObject;
      ErrorValue: EFIBError; KindIBError: TKindIBError; var DoRaise: boolean);
    procedure tiDblClick(Sender: TObject);
    procedure aeMinimize(Sender: TObject);
    procedure miShowWinClick(Sender: TObject);
    procedure miExitClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    bExecBackup, bClose: boolean;
    iToLogCnt: integer;
    sBatPath: string;
    rWorkArea: TRect;
    bCanExit: boolean;
    slFilesFTP: TStringList;
    tbData: TpFIBDataSet;
    M: TStringStream;
    slRepParams: TStringList;
    iAttemptRestore: integer;
    AppParams: TAppParams;
    sHidePrintRepName: string;
    procedure SaveLog;
    procedure PingDB;
    function ExecAndWait(const FileName, Params: ShortString; const WinState: Word): boolean; export;
    procedure SaveParams;
    procedure LoadParams;
    procedure Connect2DB;
    procedure CommitTrasaction(tr: TpFIBTransaction);
    procedure SetStatus(s: string; bToLog: boolean = true);
    procedure DoAfterReconnect;
    procedure ExecSQLCode(sSQL: string; bCommit: boolean = true);
    procedure ProcessTasks;
    function Backup(sFile: string): boolean;
    function Archive(sBACKUP_ARCHIVATOR, sBACKUP_PARAMS, sBackupFile,
      sArchiveFile, sCopyToPath, sCopyToPath2: string; iDelAftDays: integer;
      bOnlyOneFile4Copy2: boolean): boolean;
    procedure DelOldFiles(sDir: string; iDelAftDays: integer);
    procedure DelFilesExceptOne(sDir, sExceptFName: string);
    procedure ListFiles(Path: string; FileList: TStrings; sFilter: string = '*.*');
    procedure SetProps_DataSet(DataSet: TpFIBDataSet; sName, sCpt, sTableName, sSQL: string; bActive: boolean);
    function GetParValueSys(sParamName: string): variant;
    function GetLastOperMonthDate(dDate: TDate): TDate;
    function GetCurLogFileName: string;
    function GetArchiveFileFullName(sDir, sFName: string): string;
    function GetFieldValue(sSQL: string): string;
    function AppDirFileName(sFName: string): string;
  public
    function GetUserOptionValue(sName: string): string;
  end;

const
  sFDBFName = 'USU.FDB';
  sLogFolder = 'logs';
  sIniFNameParams = 'srv_params.ini';
  sIniFNameLang = 'srv_lang.txt';
  sIniDBSection = 'db';
  sIniDBServerLocal = 'local';
  sIniDBServerName = 'ServerName';
  sIniDBCharset = 'Charset';
  sIniDBProtocol = 'Protocol';
  sIniDBFileName = 'FileName';
  sIniDBDebug = 'Debug';
  sIniDBBackupShow = 'BackupShow';
  sIniDBBackupCopy2OnlyOne = 'BackupCopy2OnlyOne';
  sIniDBBackupFilesMask = 'BackupFilesMask';
  sLangSect_FormMain = 'FORM_MAIN';
  sPrefixExportFile = 'EXP_';
  sFieldNAME = 'NAME';
  sFieldFILEDOC = 'FILEDOC';
  sFieldNEW_FILE_NAME = 'FILE_NAME';
  sFieldMODULE = 'MODULE';
  sFieldACT = 'ACT';
  sFieldID = 'ID';
  sFieldDEPID = 'DEPID';
  sFormatTimeProgress = 'hh:nn:ss:zzz';
  sFormatFloat = '#,##0.00';
  iNoClient = -1;
  sNULL = 'NULL';
  sFKFlagLast = '__';
  sRepAliasSep = sFKFlagLast;
  sNameFieldName = 'NAME';
  sPrefixDataSource = 'ds';
  sPrefixGenerator = 'GEN';
  sKeyFieldName = 'ID';
  sTblKeyFields = sKeyFieldName;
  sFormatDateTime = 'dd.mm.yyyy hh:mm:ss';
  sFormatTimeFIB = 'hh:mm:ss';
  sLangRepPrefix = 'REP_';
  sRepParSys = 'SYS_';
  sRepParUser = 'USER_';
  sRepParParam = 'PAR_';
  sLeftParented = '(';
  sRightParented = ')';
  sORG_NAME = 'ORG_NAME';
  sOperday = 'OPERDAY';
  sDep = 'DEP';
  sDepIDField = sDep + sKeyFieldName;
  sIniAppSection = 'app';
  sIniFormLeft = 'FormLeft';
  sIniFormTop = 'FormTop';
  sIniFormWidth = 'FormWidth';
  sIniFormHeight = 'FormHeight';
  sIniBatPath = 'BatPath';
  sIniExecBackup = 'ExecBackup';
  iDelPriorPhotosEach = 30;
  iDefFTPPort = 21;
  sSepPalka = '|';

  sWait = 'Ожидание';
  sErrorExec = 'Ошибка выполнения: %s';
  sSoftwareClose = 'Программа закрыта';
  sSoftwareOpen = 'Программа открыта';
  sSiteExchange = 'Универсальная Система Учета';
  sCantSoftwareOpen = 'Программа не смогла открыться';
  sMsgCantDBConnect = 'Не удалось подключиться к базе данных: %s';
  sMsgConnectionRestored = 'Соединение с базой восстановлено';
  sFormTryConnectLabel = 'Попытка соединения с базой №';
  sMsgConnectionLost = 'Разрыв соединения';

var
  fmToSiteMain: TfmToSiteMain;

implementation

{$R *.dfm}

procedure TfmToSiteMain.SetStatus(s: string; bToLog: boolean = true);
const
  iMaxToLog = 200;
begin
  lbStatus.Caption := s;
  lbStatus.Refresh;

  if bToLog then
  begin
    if s <> '' then
      s := DateTimeToStr(Now) + ' ' + s;

    mmLog.Lines.Add(s);
    mmLog.InnerControl.Perform(EM_LINESCROLL, 0, mmLog.Lines.Count);

    inc(iToLogCnt);
    if iToLogCnt >= iMaxToLog then
    begin
      iToLogCnt := 0;
      try
        SaveLog;
      except
        // --
      end;
    end;
  end;
end;

function TfmToSiteMain.GetCurLogFileName: string;
var
  sFName, sDirName: string;
begin
  sFName := FormatDateTime('YYYYMMDD', Date) + '.txt';
  sDirName := ExtractFilePath(Application.ExeName) + sLogFolder;
  if not DirectoryExists(sDirName) then
    CreateDir(sDirName);
  Result := sDirName + '\' + sFName;
end;

function TfmToSiteMain.GetArchiveFileFullName(sDir, sFName: string): string;
begin
  Result := ExtractFilePath(Application.ExeName) + sDir;
  if not DirectoryExists(Result) then
    CreateDir(Result);

  Result := Result + '\' + FormatDateTime('YYYYMMDD', Date);
  if not DirectoryExists(Result) then
    CreateDir(Result);

  Result := Result + '\' + sFName;
end;

procedure TfmToSiteMain.ProcessTasks;
var
  tbTasks: TpFIBDataSet;
  sFile, sCopy2: string;
  iDays: integer;
begin
  tbTasks := TpFIBDataSet.Create(nil);
  slRepParams := TStringList.Create;
  try
    try
      with tbTasks do
      begin
        Database := dbMain;
        SQLs.SelectSQL.Text := 'select * from PL_LIST';
        Active := true;

        First;
        while not Eof do
        begin
          if bClose then
            Exit;

          if FieldByName('IS_BACKUP').AsInteger = 1 then
          begin
            if AppParams.bBackupShow then
              miShowWin.Click;

            ExecSQLCode('execute procedure PL_SET_STATUS(' + FieldByName('ID')
              .AsString + ', 1)');
            SetStatus(FieldByName('TASK_NAME').AsString);

            SetStatus('BACKUP. Начало');
            sFile := FieldByName('BACKUP_PATH').AsString;
            if Backup(sFile) then
            begin
              SetStatus('BACKUP. Файл создан: ' + sFile);

              Application.ProcessMessages;
              SetStatus('BACKUP. Архивирование. Начало');

              sCopy2 := '';
              if FindField('BACKUP_COPY_TO2') <> nil then
                sCopy2 := FindField('BACKUP_COPY_TO2').AsString;

              iDays := 0;
              if FindField('BACKUP_DEL_AFT_DAYS') <> nil then
                iDays := StrToIntDef(FindField('BACKUP_DEL_AFT_DAYS')
                  .AsString, 0);

              if Archive(FieldByName('BACKUP_ARCHIVATOR').AsString,
                FieldByName('BACKUP_PARAMS').AsString, sFile,
                FieldByName('BACKUP_ARCHIVE').AsString,
                FieldByName('BACKUP_COPY_TO').AsString, sCopy2, iDays,
                AppParams.bBackupCopy2OnlyOne) then
              begin
                SetStatus('BACKUP. Конец');
                ExecSQLCode('execute procedure PL_SET_STATUS(' +
                  FieldByName('ID').AsString + ', 0)');
              end;
            end;
          end;

          Next;
        end
      end;
    except
      on E: Exception do
        SetStatus(Format(sErrorExec,[E.Message]));
    end;
  finally
    tbTasks.Active := false;
    tbTasks.Free;
    slRepParams.Free;
  end;
end;

function TfmToSiteMain.Archive(sBACKUP_ARCHIVATOR, sBACKUP_PARAMS, sBackupFile,
  sArchiveFile, sCopyToPath, sCopyToPath2: string; iDelAftDays: integer;
  bOnlyOneFile4Copy2: boolean): boolean;
var
  sCopyArchiveFileName, sCopyArchiveFileName2: string;
  bHasCopy: boolean;
begin
  if sBACKUP_ARCHIVATOR = '' then
  begin
    Result := true;
    Exit;
  end;

  sCopyArchiveFileName := '';
  if sCopyToPath <> '' then
  begin
    if not(sCopyToPath[Length(sCopyToPath)] in ['\', '/']) then
      sCopyToPath := sCopyToPath + '\';
    sCopyArchiveFileName := sCopyToPath + ExtractFileName(sArchiveFile);
  end;

  sCopyArchiveFileName2 := '';
  if sCopyToPath2 <> '' then
  begin
    if not(sCopyToPath2[Length(sCopyToPath2)] in ['\', '/']) then
      sCopyToPath2 := sCopyToPath2 + '\';
    sCopyArchiveFileName2 := sCopyToPath2 + ExtractFileName(sArchiveFile);
  end;

  Result := false;
  try
    if ExecAndWait(sBACKUP_ARCHIVATOR, sBACKUP_PARAMS, SW_HIDE) then
    begin
      SetStatus('BACKUP. Файл заархивирован: ' + sArchiveFile);
      DeleteFile(sBackupFile);
      SetStatus('BACKUP. Файл до архивации удален: ' + sBackupFile);
      bHasCopy := false;

      SetStatus('BACKUP. sCopyArchiveFileName = ' + sCopyArchiveFileName);
      if sCopyArchiveFileName <> '' then
        if CopyFile(PWideChar(sArchiveFile), PWideChar(sCopyArchiveFileName),
          false) then
        begin
          SetStatus('BACKUP. Архив скопирован: ' + sCopyArchiveFileName);
          bHasCopy := true;

          DelOldFiles(ExtractFileDir(sCopyArchiveFileName), iDelAftDays);
        end;

      SetStatus('BACKUP. sCopyArchiveFileName2 = ' + sCopyArchiveFileName2);
      if sCopyArchiveFileName2 <> '' then
        if CopyFile(PWideChar(sArchiveFile), PWideChar(sCopyArchiveFileName2),
          false) then
        begin
          SetStatus('BACKUP. Архив2 скопирован: ' + sCopyArchiveFileName2);
          bHasCopy := true;

          if bOnlyOneFile4Copy2 then
            DelFilesExceptOne(ExtractFileDir(sCopyArchiveFileName2),
              ExtractFileName(sCopyArchiveFileName2))
          else
            DelOldFiles(ExtractFileDir(sCopyArchiveFileName2), iDelAftDays);
        end;

      if bHasCopy then
        if FileExists(sArchiveFile) then
        begin
          DeleteFile(sArchiveFile);
          SetStatus('Есть копия. Файл после архивации удален: ' + sArchiveFile);
        end;

      Result := true;
    end;
  except
    on E: Exception do
      SetStatus(Format(sErrorExec,[E.Message]));
  end;
end;

procedure TfmToSiteMain.ListFiles(Path: string; FileList: TStrings;
  sFilter: string = '*.*');
var
  SR: TSearchRec;
begin
  FileList.Clear;
  if FindFirst(Path + sFilter, faAnyFile, SR) = 0 then
  begin
    repeat
      if (SR.Attr <> faDirectory) then
        if (SR.Name <> '.') and (SR.Name <> '..') then
          FileList.Add(SR.Name);
    until FindNext(SR) <> 0;

    FindClose(SR);
  end;
end;

procedure TfmToSiteMain.DelOldFiles(sDir: string; iDelAftDays: integer);
var
  SR: TSearchRec;
  slFiles: TStringList;
  iFilesCnt: integer;
begin
  if iDelAftDays > 0 then
  begin
    slFiles := TStringList.Create;
    try
      ListFiles(sDir + '\', slFiles);
      iFilesCnt := slFiles.Count;
    finally
      slFiles.Free;
    end;

    if iFilesCnt > iDelAftDays then
      if FindFirst(sDir + '\' + AppParams.sBackupFilesMask, faAnyFile, SR) = 0
      then
      begin
        repeat
          if (SR.Attr <> faDirectory) then
            if DaysBetween(Now, SR.TimeStamp) > iDelAftDays then
            begin
              DeleteFile(sDir + '\' + SR.Name);
              SetStatus('Удалили файл старше ' + IntToStr(iDelAftDays) +
                ' дней: ' + ' ' + SR.Name + ' ' + DateTimeToStr(SR.TimeStamp));
            end;
        until FindNext(SR) <> 0;
        FindClose(SR);
      end;
  end;
end;

procedure TfmToSiteMain.DelFilesExceptOne(sDir, sExceptFName: string);
var
  SR: TSearchRec;
  slFiles: TStringList;
begin
  if sExceptFName <> '' then
    if FindFirst(sDir + '\' + AppParams.sBackupFilesMask, faAnyFile, SR) = 0
    then
    begin
      repeat
        if (SR.Attr <> faDirectory) then
          if SR.Name <> sExceptFName then
          begin
            DeleteFile(sDir + '\' + SR.Name);
            SetStatus('Удалили файл "' + SR.Name +
              '", чтобы оставить в папке только один (' + sExceptFName + ')');
          end;
      until FindNext(SR) <> 0;
      FindClose(SR);
    end;
end;

function TfmToSiteMain.ExecAndWait(const FileName, Params: ShortString;
  const WinState: Word): boolean; export;
var
  StartInfo: TStartupInfo;
  ProcInfo: TProcessInformation;
  CmdLine: ShortString;
begin
  CmdLine := '"' + FileName + '" ' + Params;
  FillChar(StartInfo, SizeOf(StartInfo), #0);
  with StartInfo do
  begin
    cb := SizeOf(StartInfo);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := WinState;
  end;
  Result := CreateProcess(nil, PChar(String(CmdLine)), nil, nil, false,
    CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil,
    PChar(ExtractFilePath(FileName)), StartInfo, ProcInfo);
  if Result then
  begin
    WaitForSingleObject(ProcInfo.hProcess, INFINITE);
    CloseHandle(ProcInfo.hProcess);
    CloseHandle(ProcInfo.hThread);
  end;
end;

function TfmToSiteMain.Backup(sFile: string): boolean;
const
  iPingFTPEvery = 500;
var
  sDir: string;
  bBACKUP_DETAILS: boolean;
  iLineNo: integer;
begin
  iLineNo := 0;
  Result := false;
  Screen.Cursor := crHourGlass;
  try
    try
      ExecSQLCode
        ('update usu_options_user ou set ou.par_value = 1 where ou.par_name = ''WT_NOT_SHOW_ERROR''');

      bBACKUP_DETAILS := true;

      sDir := ExtractFilePath(sFile);
      if not DirectoryExists(sDir) then
        ForceDirectories(sDir);
      BackupService1.BackupFile.Text := sFile;

      SetStatus('BACKUP. ServerName' + AppParams.sServerName);
      SetStatus('BACKUP. DatabaseName' + AppParams.sFileName);

      BackupService1.ServerName := AppParams.sServerName;
      BackupService1.DatabaseName := AppParams.sFileName;
      BackupService1.LoginPrompt := false;

      BackupService1.Params.Clear;
      BackupService1.Params.Add('user_name=SYSDBA');
      BackupService1.Params.Add('password=masterkey');

      BackupService1.Active := true;
      BackupService1.Verbose := bBACKUP_DETAILS;
      BackupService1.ServiceStart;
      if bBACKUP_DETAILS then
        while not BackupService1.Eof do
        begin
          Application.ProcessMessages;
          SetStatus(BackupService1.GetNextLine);

          inc(iLineNo);
          if iLineNo >= iPingFTPEvery then
          begin
            iLineNo := 0;
          end;

        end;
      BackupService1.Active := false;
      Result := true;
    except
      on E: Exception do
        SetStatus(Format(sErrorExec,[E.Message]));
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmToSiteMain.tiDblClick(Sender: TObject);
begin
  miShowWin.Click;
end;

procedure TfmToSiteMain.PingDB;
var
  sVal, sSQL: string;
begin
  sSQL := 'select 1 from usu_options';
  sVal := GetFieldValue(sSQL);
  SetStatus('- PingDB');
end;

procedure TfmToSiteMain.trPingServerTimer(Sender: TObject);
begin
  (Sender as TTimer).Enabled := false;
  try
    try
      if not bClose then
      begin
        if AppParams.bDebug then
          SetStatus('1..1');

        if dbMain.Connected then
        begin
          if AppParams.bDebug then
            SetStatus('1..2');

          dbMain.Connected := false;
        end;

        if AppParams.bDebug then
          SetStatus('1..3');

        dbMain.Connected := true;

        if 1 = 1 then
        begin
          SetStatus('-');

          (Sender as TTimer).Interval := StrToIntDef(GetUserOptionValue('QIWI_TIMER_SEC'), 5) * 1000;

          if not bClose then
          begin
            ProcessTasks;
          end;
        end
        else
        begin
          dbMain.Connected := true;
          DoAfterReconnect;
        end;
      end;
    except
      on E: Exception do
        SetStatus(Format(sErrorExec,[E.Message]));
    end;
  finally
    (Sender as TTimer).Enabled := true;
    SetStatus(sWait, false);
    CommitTrasaction(trMain);
    CommitTrasaction(trUpdate);
  end;
end;

procedure TfmToSiteMain.SetProps_DataSet(DataSet: TpFIBDataSet;
  sName, sCpt, sTableName, sSQL: string; bActive: boolean);
begin
  with DataSet do
  begin
    Name := sName;
    Description := sCpt;
    if sTableName <> '' then
    begin
      AutoCommit := true;
      AutoUpdateOptions.GeneratorName := sPrefixGenerator + '_' + sTableName +
        '_' + sKeyFieldName;
      AutoUpdateOptions.KeyFields := sTblKeyFields;
      AutoUpdateOptions.UpdateTableName := sTableName;
      AutoUpdateOptions.WhenGetGenID := wgBeforePost;
    end;
    DefaultFormats.DateTimeDisplayFormat := sFormatDateTime;
    DefaultFormats.DisplayFormatTime := sFormatTimeFIB;
    Database := dbMain;
    Transaction := trMain;
    UpdateTransaction := trUpdate;
    SQLs.SelectSQL.Text := sSQL;
    if bActive then
      Active := true;
  end;
end;

function TfmToSiteMain.GetLastOperMonthDate(dDate: TDate): TDate;
  function DaysOfMonth(mm, yyyy: integer): integer;
  begin
    if mm = 2 then
    begin
      Result := 28;
      if IsLeapYear(yyyy) then
        Result := 29;
    end
    else
    begin
      if mm < 8 then
      begin
        if (mm mod 2) = 0 then
          Result := 30
        else
          Result := 31;
      end
      else
      begin
        if (mm mod 2) = 0 then
          Result := 31
        else
          Result := 30;
      end;
    end;
  end;

begin
  Result := EncodeDate(StrToInt(FormatDateTime('YYYY', dDate)),
    StrToInt(FormatDateTime('MM', dDate)),

    DaysOfMonth(StrToInt(FormatDateTime('MM', dDate)),
    StrToInt(FormatDateTime('YYYY', dDate))));
end;

function TfmToSiteMain.GetParValueSys(sParamName: string): variant;
begin
  if sParamName = sORG_NAME then
    Result := GetFieldValue('select o.org_name from usu_options o')
  else if sParamName = sOperday then
    Result := GetFieldValue('select current_date from usu_options')
  else if sParamName = 'OPER_MONTH_AND_YEAR' then
    Result := AnsiUpperCase(FormatDateTime('MMMM, YYYY',
      StrToDate(GetFieldValue('select current_date from usu_options'))))
  else if sParamName = 'LAST_MONTH_DATE' then
    Result := GetLastOperMonthDate
      (StrToDate(GetFieldValue('select current_date from usu_options')))
  else if sParamName = 'OPER_YEAR' then
    Result := FormatDateTime('YYYY',
      StrToDate(GetFieldValue('select current_date from usu_options')))
  else if sParamName = sDepIDField then
    Result := '1'
  else if sParamName = 'USER_NAME' then
    Result := GetFieldValue('select current_user from usu_options')
  else if sParamName = 'ROLE' then
    Result := GetFieldValue('select current_role from usu_options')
  else if sParamName = 'CUR_DATE' then
    Result := Date
  else if sParamName = 'FILTER' then
    Result := '';
end;

function TfmToSiteMain.GetFieldValue(sSQL: string): string;
var
  tbTemp: TpFIBDataSet;
begin
  tbTemp := TpFIBDataSet.Create(nil);
  try
    with tbTemp do
    begin
      Database := dbMain;
      SQLs.SelectSQL.Text := sSQL;
      Active := true;
      Result := Fields[0].AsString;
    end;
  finally
    tbTemp.Active := false;
    tbTemp.Free;
  end;
end;

procedure TfmToSiteMain.ExecSQLCode(sSQL: string; bCommit: boolean = true);
begin
  if not bClose then
    with TpFIBQuery.Create(nil) do
    begin
      try
        try
          Database := dbMain;
          SQL.Text := sSQL;
          ExecQuery;

          if bCommit then
          begin
            CommitTrasaction(trMain);
            CommitTrasaction(trUpdate);
          end;
        except
          on E: Exception do
            MessageDlg(Format(sErrorExec, [E.Message]), mtError, [mbOk], 0);
        end;
      finally
        Free;
      end;
    end;
end;

function TfmToSiteMain.GetUserOptionValue(sName: string): string;
begin
  Result := GetFieldValue
    ('select PAR_VALUE from USU_OPTIONS_USER where PAR_NAME = ''' +
    sName + '''');
end;

procedure TfmToSiteMain.aeMinimize(Sender: TObject);
begin
  ShowWindow(Handle, SW_HIDE);
  ShowWindow(Application.Handle, SW_HIDE);
  SetWindowLong(Application.Handle, GWL_EXSTYLE,
    GetWindowLong(Application.Handle, GWL_EXSTYLE) or (not WS_EX_APPWINDOW));

  ti.BalloonTitle := Self.Caption;
  ti.BalloonHint :=
    'Программа просто свернута и находится в области уведомлений!';
  ti.ShowBalloonHint;
end;

function TfmToSiteMain.AppDirFileName(sFName: string): string;
begin
  Result := ExtractFilePath(Application.ExeName) + sFName;
end;

procedure TfmToSiteMain.FibErrorHandlerFIBErrorEvent(Sender: TObject;
  ErrorValue: EFIBError; KindIBError: TKindIBError; var DoRaise: boolean);
begin
  if KindIBError = keLostConnect then
    DoRaise := false;
end;

procedure TfmToSiteMain.SaveLog;
begin
  mmLog.Lines.SaveToFile(GetCurLogFileName, TEncoding.UTF8);
end;

procedure TfmToSiteMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if not bCanExit then
  begin
    aeMinimize(nil);
    Action := caNone;
  end
  else
  begin
    bClose := true;

    SaveLog;

    SetStatus(sSoftwareClose);
    trPingServer.Enabled := false;

    if tbData <> nil then
      tbData.Free;

    if M <> nil then
      M.Free;

    CommitTrasaction(trMain);
    CommitTrasaction(trUpdate);

    with dbMain do
      if Connected then
        Connected := false;

    SaveParams;
  end;
end;

procedure TfmToSiteMain.SaveParams;
var
  sIniFileName: string;
begin
  sIniFileName := AppDirFileName(sIniFNameParams);
  with TMemIniFile.Create(sIniFileName, TEncoding.UTF8) do
  begin
    try
      WriteInteger(sIniAppSection, sIniFormLeft, Self.Left);
      WriteInteger(sIniAppSection, sIniFormTop, Self.Top);
      WriteInteger(sIniAppSection, sIniFormWidth, Self.Width);
      WriteInteger(sIniAppSection, sIniFormHeight, Self.Height);
    finally
      UpdateFile;
      Free;
    end;
  end;
end;

procedure TfmToSiteMain.LoadParams;
const
  iCheckVal = -1000;
var
  sIniFileName: string;
begin
  sIniFileName := AppDirFileName(sIniFNameParams);
  if FileExists(sIniFileName) then
    with TMemIniFile.Create(sIniFileName, TEncoding.UTF8) do
    begin
      try
        sBatPath := ReadString(sIniAppSection, sIniBatPath, '');
        bExecBackup := ReadInteger(sIniAppSection, sIniExecBackup, 0) = 1;

        if ReadInteger(sIniAppSection, sIniFormLeft, iCheckVal) = iCheckVal then
        begin
          SystemParametersInfo(SPI_GETWORKAREA, 0, @rWorkArea, 0);
          with Self do
          begin
            Left := rWorkArea.Right div 4;
            Top := rWorkArea.Top;
            Width := rWorkArea.Right div 2;
            Height := rWorkArea.Bottom;
          end;
        end
        else
        begin
          Self.Left := ReadInteger(sIniAppSection, sIniFormLeft, Self.Left);
          Self.Top := ReadInteger(sIniAppSection, sIniFormTop, Self.Top);
          Self.Width := ReadInteger(sIniAppSection, sIniFormWidth, Self.Width);
          Self.Height := ReadInteger(sIniAppSection, sIniFormHeight,
            Self.Height);
        end;

        with AppParams do
        begin
          bServerLocal := ReadBool(sIniDBSection, sIniDBServerLocal, true);
          sServerName := ReadString(sIniDBSection, sIniDBServerName,
            'localhost');
          sCharset := ReadString(sIniDBSection, sIniDBCharset, 'WIN1251');
          sProtocol := ReadString(sIniDBSection, sIniDBProtocol, 'TCP/IP');
          sFileName := ReadString(sIniDBSection, sIniDBFileName,
            AppDirFileName(sFDBFName));
          bDebug := ReadInteger(sIniDBSection, sIniDBDebug, 0) = 1;
          bBackupShow := ReadInteger(sIniDBSection, sIniDBBackupShow, 0) = 1;
          bBackupCopy2OnlyOne := ReadInteger(sIniDBSection,
            sIniDBBackupCopy2OnlyOne, 0) = 1;
          sBackupFilesMask := ReadString(sIniDBSection,
            sIniDBBackupFilesMask, '*.*');
        end;
      finally
        Free;
      end;
    end;
end;

procedure TfmToSiteMain.FormCreate(Sender: TObject);
var
  sLogFileName: string;
begin
  M := TStringStream.Create('');

  tbData := TpFIBDataSet.Create(nil);
  tbData.Database := dbMain;

  bClose := false;
  iAttemptRestore := 0;
  iToLogCnt := 0;

  sLogFileName := GetCurLogFileName;
  if FileExists(sLogFileName) then
    mmLog.Lines.LoadFromFile(sLogFileName);

  fmToSiteMain.Caption := 'USU.kz - ' + sSiteExchange;
  ti.Hint := fmToSiteMain.Caption;
  bCanExit := false;
end;

procedure TfmToSiteMain.FormResize(Sender: TObject);
begin
  cxProgressBar.Width := Self.Width - cxProgressBar.Left - 30;
  cxProgressBar2.Width := Self.Width - cxProgressBar2.Left - 30;
  lbStatus.Width := cxProgressBar.Width;
end;

procedure TfmToSiteMain.FormShow(Sender: TObject);
begin
  ti.Visible := true;
  LoadParams;
  Connect2DB;
end;

procedure TfmToSiteMain.Connect2DB;
begin
  with dbMain, AppParams do
  begin
    DBName := sFileName;
    if not bServerLocal then
    begin
      if AnsiUpperCase(sProtocol) = 'NetBEUI' then
        DBName := '\\' + sServerName + '\' + DBName
      else if AnsiUpperCase(sProtocol) = 'Novell SPX' then
        DBName := sServerName + '@' + DBName
      else
        DBName := sServerName + ':' + DBName;
    end;

    DBParams.Clear;
    DBParams.Add('user_name=SYSDBA');
    DBParams.Add('password=masterkey');
    DBParams.Add('lc_ctype=' + sCharset);
    DBParams.Add('sql_role_name=');

    try
      Connected := true;
      DoAfterReconnect;

      if Trim(mmLog.Text) <> '' then
        SetStatus('');
      SetStatus(sSoftwareOpen);
    except
      on E: Exception do
      begin
        SetStatus(sCantSoftwareOpen);
        MessageDlg(Format(sMsgCantDBConnect,[E.Message]), mtError, [mbOk], 0);
        Application.Terminate;
      end;
    end;
  end;
end;

procedure TfmToSiteMain.DoAfterReconnect;
begin
  if dbMain.Connected then
  begin
    trMain.Active := true;
    trUpdate.Active := true;



    trPingServer.Interval := 2 * 1000;
    trPingServer.Enabled := true;
  end;
end;

procedure TfmToSiteMain.dbMainAfterRestoreConnect(Database: TFIBDatabase);
begin
  iAttemptRestore := 0;
  SetStatus(sMsgConnectionRestored);
  DoAfterReconnect;
end;

procedure TfmToSiteMain.dbMainErrorRestoreConnect(Database: TFIBDatabase;
  E: EFIBError; var Actions: TOnLostConnectActions; var DoRaise: boolean);
begin
  inc(iAttemptRestore);

  SetStatus(sFormTryConnectLabel + ' ' + IntToStr(iAttemptRestore), false);
end;

procedure TfmToSiteMain.dbMainLostConnect(Database: TFIBDatabase; E: EFIBError;
  var Actions: TOnLostConnectActions; var DoRaise: boolean);
begin
  trPingServer.Enabled := false;
  SetStatus(sMsgConnectionLost);
  Actions := laWaitRestore;
end;

procedure TfmToSiteMain.CommitTrasaction(tr: TpFIBTransaction);
begin
  with tr do
    if Active then
      if InTransaction then
        CommitRetaining;
end;

procedure TfmToSiteMain.miShowWinClick(Sender: TObject);
begin
  ShowWindow(Handle, SW_RESTORE);
  SetForegroundWindow(Handle);
end;

procedure TfmToSiteMain.miExitClick(Sender: TObject);
begin
  bCanExit := true;
  Close;
end;

end.
