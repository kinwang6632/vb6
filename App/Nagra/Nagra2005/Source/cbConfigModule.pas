unit cbConfigModule;

interface

uses
  SysUtils, Classes, Windows, Variants, IniFiles, Forms, DB, ADODB, Controls,
  DBClient, cbClass, cxTL, LbCipher, LbClass, Provider;

type
  TConfigModule = class(TDataModule)
    AccessConnection: TADOConnection;
    DataReader: TADOQuery;
    HighCmdEnv: TClientDataSet;
    LowCmdEnv: TClientDataSet;
    CmdError: TClientDataSet;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FSoList: TThreadList;
    FAppStartupInfo: TAppStartupInfo;
    FCommonEnv: TCommonEnv;
    FCmdTransEnv: TCmdTransEnv;
    FCASEnv: TCASEnv;
    FMsgSubject: TMessageExSubject;
    FFeedbackSo: PSoInfo;
    FConfigFileManager: TConfigFileManager;
    FLbDES: TLbDES;
    function GetSoCount: Integer;
    function ConnectToAccess: Boolean;
    procedure DisconnectFromAccess;
    procedure InternalClearSoList(const aList: TList);
    procedure GetConfigFormValue;
    function GetStartupInfo: Boolean;
    function GetSoList: Boolean;
    function GetCommon: Boolean;
    function GetHighCmdEnv: Boolean;
    function GetLowCmdEnv: Boolean;
    function GetCmdTransEnv: Boolean;
    function GetCASEnv: Boolean;
    function GetCmdError: Boolean;
    function GetFeedbackSo: Boolean;
    {}
    function SetCASEnv: Boolean;
    function SetCommon: Boolean;
    function SetCmdTransEnv: Boolean;
    function SetSoList: Boolean;
    function SetHighCmdEnv: Boolean;
    function SetLowCmdEnv: Boolean;
    function SetCmdError: Boolean;
  public
    { Public declarations }
    constructor Create; reintroduce;
    function ShowConfigForm: TModalResult;
    procedure ClearSoList;
    procedure ConfigLoadFromFile;
    procedure ConfigSaveToFile;
    property AppStartupInfo: TAppStartupInfo read FAppStartupInfo;
    property SoList: TThreadList read FSoList;
    property SoCount: Integer read GetSoCount;
    property FeedbackSo: PSoInfo read FFeedbackSo;
    property CommEnv: TCommonEnv read FCommonEnv;
    property CASEnv: TCASEnv read FCASEnv;
    property CmdTransEnv: TCmdTransEnv read FCmdTransEnv;
    property ConfigFileManager: TConfigFileManager read FConfigFileManager;
  end;

var
  ConfigModule: TConfigModule;

implementation

uses cbResStr, cbUtilis, cbMain, cbConfig, cxTextEdit;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

constructor TConfigModule.Create;
begin
  inherited Create( nil );
  FSoList := TThreadList.Create;
  FMsgSubject := TMessageExSubject.Create;
  FConfigFileManager := TConfigFileManager.Create;
  New( FFeedbackSo );
  FLbDES := TLbDES.Create( nil );
  FLbDES.GenerateKey( #67#83 );
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleCreate(Sender: TObject);
begin
  FMsgSubject.AutoNotify := True;
  FMsgSubject.AddObServer( fmMain.StartupObServer );
  GetStartupInfo;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleDestroy(Sender: TObject);
var
  aFileName: String;
begin
  DisconnectFromAccess;
  ClearSoList;
  FMsgSubject.Free;
  Dispose( FFeedbackSo );
  aFileName :=  IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + DEFAULT_APPFILE;
  FConfigFileManager.SaveTo( aFileName );
  FConfigFileManager.Free;
  FLbDES.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ConfigLoadFromFile;
begin
  if ConfigFileManager.ActiveFile = EmptyStr then
  begin
    FMsgSubject.Error( Format( SConfigOpenError,
      [EmptyStr, SConfigFileNotAssign] ) );
    Exit;
  end;
  if ConnectToAccess then
  begin
    fmMain.dxEnableControlSend.ImageIndex := 18;
    fmMain.dxEnableFeedbackRecv.ImageIndex := 18;
    try
      GetCommon;
      GetCASEnv;
      fmMain.dxNagraServer.Caption := Format( SlblConfigCASAddr,
        [FCASEnv.IP, SControlSend, FCASEnv.ControlSourceId, SControlRecv,
         FCASEnv.FeedbackSourceId] );
      Delay( COMMON_DELAY * 6 );
      GetHighCmdEnv;
      Delay( COMMON_DELAY * 6 );
      GetLowCmdEnv;
      Delay( COMMON_DELAY * 6 );
      GetCmdError;
      Delay( COMMON_DELAY * 6 );
      GetCmdTransEnv;
      Delay( COMMON_DELAY * 6 );
      GetSoList;
      Delay( COMMON_DELAY * 6 );
      GetFeedbackSo;
    finally
      DisconnectFromAccess;
    end;
    {}
    if FCmdTransEnv.EnableControlSend then
      fmMain.dxEnableControlSend.ImageIndex := 17;
    if FCmdTransEnv.EnableFeedbackRecv then
      fmMain.dxEnableFeedbackRecv.ImageIndex := 17;
    {}  
    fmMain.dxDockPanel2.Visible := ( fmMain.dxEnableControlSend.ImageIndex = 17 );
    fmMain.dxDockPanel7.Visible := ( fmMain.dxEnableControlSend.ImageIndex = 17 );
    fmMain.dxDockPanel4.Visible := ( fmMain.dxEnableFeedbackRecv.ImageIndex = 17 );
    fmMain.dxDockPanel8.Visible := ( fmMain.dxEnableFeedbackRecv.ImageIndex = 17 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ConfigSaveToFile;
begin
  if ConnectToAccess then
  begin
    try
      SetCASEnv;
      Delay( 500 );
      SetCommon;
      Delay( 500 );
      SetCmdTransEnv;
      Delay( 500 );
      SetSoList;
      Delay( 500 );
      SetHighCmdEnv;
      Delay( 500 );
      SetLowCmdEnv;
      Delay( 500 );
      SetCmdError;
      Delay( 500 );
    finally
      DisconnectFromAccess;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetSoCount: Integer;
var
  aList: TList;
begin
  aList := FSoList.LockList;
  try
    Result := aList.Count;
  finally
    FSoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ConnectToAccess: Boolean;
const
  aDbPassword = 'cyc84177282';
var
  aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    FConfigFileManager.ActiveFile;
  AccessConnection.Connected := False;
  AccessConnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [aFileName, aDbPassword] );
  try
    AccessConnection.Connected := True;
    FMsgSubject.OK( Format( SConfigOpenSuccess, [aFileName] ) );
  except
    on E: Exception do
      FMsgSubject.Error( Format( SConfigOpenError, [aFileName,E.Message] ) );
  end;
  Result := AccessConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DisconnectFromAccess;
begin
  if AccessConnection.Connected then
    AccessConnection.Connected := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.InternalClearSoList(const aList: TList);
var
  aIndex: Integer;
begin
  for aIndex := 0 to aList.Count - 1  do
  begin
    if Assigned( aList.Items[aIndex] ) then
    begin
      Dispose( PSoInfo( aList.Items[aIndex] ) );
      aList.Items[aIndex] := nil;
    end;
  end;
  aList.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetSoList: Boolean;
var
  aLockList: TList;
  aSoInfo: PSOInfo;
  aSoCount: Integer;
begin
  aLockList := FSoList.LockList;
  try
    InternalClearSoList( aLockList );
    DataReader.Close;
    DataReader.SQL.Text :=
      ' SELECT A.SO_SELECTED,               ' +
      '        A.SO_POS,                    ' +
      '        B.SO_POS_NAME,               ' +
      '        A.SO_TYPE,                   ' +
      '        A.SO_COMP_CODE,              ' +
      '        A.SO_COMP_NAME,              ' +
      '        A.SO_NETWORK_ID,             ' +
      '        A.SO_LOGINUSER,              ' +
      '        A.SO_LOGINPASS,              ' +
      '        A.SO_DBALIASE                ' +
      '   FROM SO_COMP A, SOPOS_ENV B       ' +
      '  WHERE A.SO_POS = B.SO_POS          ' +
      '    AND A.SO_TYPE = ''OvC24yZLk0s='' '; // 'SEND'
    try
      DataReader.Open;
      DataReader.First;
      aSoCount := 0;
      while not DataReader.Eof do
      begin
        New( aSoInfo );
        aSoInfo.Selected := ( FLbDES.DecryptString( DataReader.FieldByName( 'SO_SELECTED' ).AsString ) = 'Y' );
        aSoInfo.Pos := StrToIntDef( FLbDES.DecryptString( DataReader.FieldByName( 'SO_POS' ).AsString ), 0 );
        aSoInfo.PosName := FLbDES.DecryptString( DataReader.FieldByName( 'SO_POS_NAME' ).AsString );
        aSoInfo.SoType := FLbDES.DecryptString( DataReader.FieldByName( 'SO_TYPE' ).AsString );
        aSoInfo.CompCode := FLbDES.DecryptString( DataReader.FieldByName( 'SO_COMP_CODE' ).AsString );
        aSoInfo.CompName := FLbDES.DecryptString( DataReader.FieldByName( 'SO_COMP_NAME' ).AsString );
        aSoInfo.NetworkId := FLbDES.DecryptString( DataReader.FieldByName( 'SO_NETWORK_ID' ).AsString );
        aSoInfo.LoginUser := FLbDES.DecryptString( DataReader.FieldByName( 'SO_LOGINUSER' ).AsString );
        aSoInfo.LoginPass := FLbDES.DecryptString( DataReader.FieldByName( 'SO_LOGINPASS' ).AsString );
        aSoInfo.DbAliase := FLbDES.DecryptString( DataReader.FieldByName( 'SO_DBALIASE' ).AsString );
        if aSoInfo.Selected then
          aSoInfo.DbConnectStatus := dbNone
        else
          aSoInfo.DbConnectStatus := dbNoSelect;
        aSoInfo.RecordCount := 0;
        aSoInfo.BufferCount := 0;
        aSoInfo.Data := nil;
        aSoInfo.NotifyMsg := EmptyStr;
        aSoInfo.CriticalError := 0;
        aSoInfo.LastCriticalErrorTickCount := 0;
        aLockList.Add( aSoInfo );
        Inc( aSoCount );
        FMsgSubject.Info( Format( SConfigSoRead, [aSoInfo.CompName] ) );
        Delay( COMMON_DELAY );
        DataReader.Next;
      end;
      FMsgSubject.OK( Format( SConfigSoReadSuccess, [aSoCount] ) );
      Result := True;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SConfigSoReadError, [E.Message] ) );
        InternalClearSoList( aLockList );
        Result := False;
      end;
    end;
    DataReader.Close;
  finally
    FSoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCommon: Boolean;
begin
  FCommonEnv.ProcessRecordCount := 20;
  FCommonEnv.ProcessIPPV := 'A';
  FCommonEnv.ProcessBatch := 'A';
  FCommonEnv.DbMultiThread := True;
  FCommonEnv.DbRetryFrequence := 10;
  FCommonEnv.DbResendCount := 3;
  FCommonEnv.DbWriteErrorWhenSocketFail := False;
  FCommonEnv.BusyTimeStart := EmptyStr;
  FCommonEnv.BusyTimeEnd := EmptyStr;
  FCommonEnv.BusyTimeDbReadFrequence := 3;
  FCommonEnv.NormalTimeDbReadFrequence := 30;
  FCommonEnv.CASRetryFrequence := 10;
  FCommonEnv.CASSendErrMax := 3;
  FCommonEnv.CASRecvErrMax := 3;
  FCommonEnv.CASCommCheck := 5;
  { 這個值不在於設定檔中, 程式內部用, 有須要再加進設定檔 }
  FCommonEnv.DbWaterMark := 100; 
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.PROCESSRECORDCOUNT,         ' +
    '        A.PROCESSIPPV,                ' +
    '        A.PROCESSBATCH,               ' +
    '        A.DBMULTITHREAD,              ' +
    '        A.DBRETRYFREQUENCE,           ' +
    '        A.DBRESENDCOUNT,              ' +
    '        A.DBWRITEERRORWHENSOCKETFAIL, ' +
    '        A.BUSY_TIME_START,            ' +
    '        A.BUSY_TIME_END,              ' +
    '        A.BUSY_TIME_READ_FREQUENCY,   ' +
    '        A.NORMAL_TIME_READ_FREQUENCY, ' +
    '        A.BATCHOPERATOR,              ' +
    '        A.CASRETRYFREQUENCE,          ' +
    '        A.CASSENDERRMAX,              ' +
    '        A.CASRECVERRMAX,              ' +
    '        A.CASCOMMCHECK                ' +
    '   FROM COMMON_ENV A                  ';
  try
    DataReader.Open;
    DataReader.First;
    { }
    FCommonEnv.ProcessRecordCount := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'PROCESSRECORDCOUNT' ).AsString ), 20 );
    FMsgSubject.Info(
      Format( SConfigCommonProcessRecordCount, [FCommonEnv.ProcessRecordCount] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.ProcessIPPV :=
      FLbDES.DecryptString( DataReader.FieldByName( 'PROCESSIPPV' ).AsString );
    FMsgSubject.Info(
      Format( SConfigCommonProcessIPPV, [FCommonEnv.ProcessIPPV] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.ProcessBatch :=
      FLbDES.DecryptString( DataReader.FieldByName( 'PROCESSBATCH' ).AsString );
    FMsgSubject.Info(
      Format( SConfigCommonProcessBatch, [FCommonEnv.ProcessBatch] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbMultiThread :=
      ( FLbDES.DecryptString( DataReader.FieldByName( 'DBMULTITHREAD' ).AsString ) = 'Y' );
    FMsgSubject.Info( Format( SConfigCommonDbMultiThread,
      [FLbDES.DecryptString( DataReader.FieldByName( 'DBMULTITHREAD' ).AsString )] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbRetryFrequence := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'DBRETRYFREQUENCE' ).AsString ), 60 );
    FMsgSubject.Info(
      Format( SConfigCommonDbRetryFrequence, [FCommonEnv.DbRetryFrequence] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbResendCount := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'DBRESENDCOUNT' ).AsString ), 0 );
    FMsgSubject.Info(
      Format( SConfigCommonDbRetryFrequence, [FCommonEnv.DbResendCount] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbWriteErrorWhenSocketFail :=
      ( FLbDES.DecryptString( DataReader.FieldByName( 'DBWRITEERRORWHENSOCKETFAIL' ).AsString ) = 'Y' );
    FMsgSubject.Info( Format( SConfigCommonDbWriteErrorWhenSocketFail,
      [FLbDES.DecryptString( DataReader.FieldByName( 'DBWRITEERRORWHENSOCKETFAIL' ).AsString )] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BusyTimeStart :=
      FLbDES.DecryptString( DataReader.FieldByName( 'BUSY_TIME_START' ).AsString );
    FMsgSubject.Info( Format( SConfigCommonBusyTimeStart,
      [Nvl( FCommonEnv.BusyTimeStart, SConfigNoneSet )] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BusyTimeEnd :=
      FLbDES.DecryptString( DataReader.FieldByName( 'BUSY_TIME_END' ).AsString );
    FMsgSubject.Info( Format( SConfigCommonBusyTimeEnd,
      [Nvl( FCommonEnv.BusyTimeEnd, SConfigNoneSet )] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BusyTimeDbReadFrequence := StrToFloatDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'BUSY_TIME_READ_FREQUENCY' ).AsString ), 3 );
    FMsgSubject.Info( Format( SConfigCommonBusyTimeDbReadFrequence,
      [FloatToStr( FCommonEnv.BusyTimeDbReadFrequence )] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.NormalTimeDbReadFrequence := StrToFloatDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'NORMAL_TIME_READ_FREQUENCY' ).AsString ), 3 );
    FMsgSubject.Info( Format( SConfigCommonNormalTimeDbReadFrequence,
      [FloatToStr( FCommonEnv.NormalTimeDbReadFrequence )] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BatchOperator :=
      FLbDES.DecryptString( DataReader.FieldByName( 'BATCHOPERATOR' ).AsString );
    FMsgSubject.Info(
      Format( SConfigCommonBatchOperator, [FCommonEnv.BatchOperator] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.CASRetryFrequence := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'CASRETRYFREQUENCE' ).AsString ), 10 );
    FMsgSubject.Info(
      Format( SConfigCommonCASRetryFrequence, [FCommonEnv.CASRetryFrequence] ) );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.CASSendErrMax := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'CASSENDERRMAX' ).AsString ), 3 );
    FMsgSubject.Info( Format( SConfigCommonCASSendErrMax, [FCommonEnv.CASSendErrMax] ) );
    { }
    FCommonEnv.CASRecvErrMax := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'CASRECVERRMAX' ).AsString ), 3 );
    FMsgSubject.Info( Format( SConfigCommonCASRecvErrMax, [FCommonEnv.CASRecvErrMax] ) );
    { }
    FCommonEnv.CASCommCheck := StrToIntDef(
      FLbDES.DecryptString( DataReader.FieldByName( 'CASCOMMCHECK' ).AsString ), 10 );
    FMsgSubject.Info( Format( SConfigCommonCASCommCheck, [FCommonEnv.CASCommCheck] ) );
    { }
    FMsgSubject.OK( SConfigCommonReadSuccess );
    Result := True;
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SConfigCommonReadError, [E.Message] ) );
      Result := False;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetHighCmdEnv: Boolean;
begin
  if not VarIsNull( HighCmdEnv.Data ) then
  begin
    HighCmdEnv.EmptyDataSet;
    if ( HighCmdEnv.IndexName <> EmptyStr ) then
      HighCmdEnv.DeleteIndex( HighCmdEnv.IndexName );
    HighCmdEnv.Data := Null;
  end;
  HighCmdEnv.CreateDataSet;
  HighCmdEnv.AddIndex( 'IDX1', 'HIGH_LEVEL_CMD', [] );
  HighCmdEnv.IndexName := 'IDX1';
  {}
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.CMD_TYPE,            ' +
    '        A.HIGH_LEVEL_CMD,      ' +
    '        A.DESCRIPTION,         ' +
    '        A.LOW_LEVEL_CMD,       ' +
    '        A.IRD_CMD,             ' +
    '        A.NOTES_USE,           ' +
    '        A.CMD_TYPE_PRIORITY,   ' +
    '        A.CMD_PRIORITY,        ' +
    '        A.NEW_CMD_COUNT,       ' +
    '        A.OLD_CMD_COUNT        ' +
    '   FROM HIGH_CMD_ENV A         ';
  DataReader.Open;
  if DataReader.IsEmpty then
  begin
    FMsgSubject.Error( SConfigHighCmdReadIsEmpty );
    Result := False;
  end else
  begin
    Result := True;
    DataReader.First;
    while not DataReader.Eof do
    begin
      try
        HighCmdEnv.AppendRecord( [
          FLbDES.DecryptString( DataReader.FieldByName( 'CMD_TYPE' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'HIGH_LEVEL_CMD' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'DESCRIPTION' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'LOW_LEVEL_CMD' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'IRD_CMD' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'NOTES_USE' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'CMD_TYPE_PRIORITY' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'CMD_PRIORITY' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'NEW_CMD_COUNT' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'OLD_CMD_COUNT' ).AsString )] );
        FMsgSubject.Info( Format( SConfigHighCmdReading,
          [HighCmdEnv.FieldByName( 'HIGH_LEVEL_CMD' ).AsString,
           HighCmdEnv.FieldByName( 'DESCRIPTION' ).AsString] ) );
      except
        on E: Exception do
        begin
          FMsgSubject.Error( Format( SConfigHighCmdReadError, [E.Message] ) );
          HighCmdEnv.EmptyDataSet;
          Result := False;
          Break;
        end;
      end;
      Delay( COMMON_DELAY );
      DataReader.Next;
    end;
    if Result then
    begin
      FMsgSubject.OK( SConfigHighCmdReadSuccess );
      Result := True;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetLowCmdEnv: Boolean;
begin
  if not VarIsNull( LowCmdEnv.Data ) then
  begin
    LowCmdEnv.EmptyDataSet;
    if ( LowCmdEnv.IndexName <> EmptyStr ) then
      LowCmdEnv.DeleteIndex( LowCmdEnv.IndexName );
    LowCmdEnv.Data := Null;
  end;
  LowCmdEnv.CreateDataSet;
  LowCmdEnv.AddIndex( 'IDX1', 'LOW_LEVEL_CMD', [] );
  LowCmdEnv.IndexName := 'IDX1';
  {}
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.LOW_LEVEL_CMD ,          ' +
    '        A.DESCRIPTION              ' +
    '   FROM LOW_CMD_ENV A              ';
  DataReader.Open;
  if DataReader.IsEmpty then
  begin
    FMsgSubject.Error( SConfigLowCmdReadIsEmpty );
    Result := False;
  end else
  begin
    Result := True;
    DataReader.First;
    while not DataReader.Eof do
    begin
      try
        LowCmdEnv.AppendRecord( [
          FLbDES.DecryptString( DataReader.FieldByName( 'LOW_LEVEL_CMD' ).AsString ),
          FLbDES.DecryptString( DataReader.FieldByName( 'DESCRIPTION' ).AsString )] );
        FMsgSubject.Info( Format( SConfigLowCmdReading,
          [LowCmdEnv.FieldByName( 'LOW_LEVEL_CMD' ).AsString,
           LowCmdEnv.FieldByName( 'DESCRIPTION' ).AsString] ) );
      except
        on E: Exception do
        begin
          FMsgSubject.Error( Format( SConfigLowCmdReadError, [E.Message] ) );
          LowCmdEnv.EmptyDataSet;
          Result := False;
          Break;
        end;
      end;
      Delay( COMMON_DELAY );
      DataReader.Next;
    end;
    if Result then
    begin
      FMsgSubject.OK( SConfigLowCmdReadSuccess );
      Result := True;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCmdTransEnv: Boolean;
begin
  FCmdTransEnv.UseSimulator := False;
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.USESIMULATOR,                  ' +
    '        A.ENABLECONTROLSEND,             ' +
    '        A.ENABLEFEEDBACKRECV,            ' +
    '        A.SENDCOMMANDDELAY               ' +
    '   FROM CMDTRANS_ENV A                   ';
  DataReader.Open;
  Result := True;
  DataReader.First;
  try
    FCmdTransEnv.UseSimulator :=
      ( FLbDES.DecryptString( DataReader.FieldByName( 'USESIMULATOR' ).AsString ) = 'Y' );
    FMsgSubject.Info( Format( SConfigCmdTransUseSimulator,
      [FLbDES.DecryptString( DataReader.FieldByName( 'USESIMULATOR' ).AsString )] ) );
    Delay( COMMON_DELAY );
    {}
    FCmdTransEnv.EnableControlSend :=
      ( FLbDES.DecryptString( DataReader.FieldByName( 'ENABLECONTROLSEND' ).AsString ) = 'Y' );
    FMsgSubject.Info( Format( SConfigCmdTransEnableControlSend,
      [FLbDES.DecryptString( DataReader.FieldByName( 'ENABLECONTROLSEND' ).AsString )] ) );
    Delay( COMMON_DELAY );
    {}
    FCmdTransEnv.EnableFeedbackRecv :=
      ( FLbDES.DecryptString( DataReader.FieldByName( 'ENABLEFEEDBACKRECV' ).AsString ) = 'Y' );
    FMsgSubject.Info( Format( SConfigCmdTransEnableFeedbackRecv,
      [FLbDES.DecryptString( DataReader.FieldByName( 'ENABLECONTROLSEND' ).AsString )] ) );
    Delay( COMMON_DELAY );
    {}
    FCmdTransEnv.SendCommandDelay :=
      StrToFloatDef( FLbDES.DecryptString( DataReader.FieldByName( 'SENDCOMMANDDELAY' ).AsString ), 0.5 );
    FMsgSubject.Info( Format( SConfigCmdTransSendCommandDelay,
      [FCmdTransEnv.SendCommandDelay] ) );
    Delay( COMMON_DELAY );
    { }
    FMsgSubject.OK( SConfigCmdTransReadSuccess );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SConfigCmdTransReadError, [E.Message] ) );
      Result := False;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCASEnv: Boolean;
begin
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.CAS_IP,                         ' +
    '        A.CAS_CONTROL_PORT,               ' +
    '        A.CAS_FEEDBACK_PORT,              ' +
    '        A.CAS_CONTROL_SOURCE_ID,          ' +
    '        A.CAS_CONTROL_DEST_ID,            ' +
    '        A.CAS_CONTROL_MOPP_ID,            ' +
    '        A.CAS_FEEDBACK_SOURCE_ID,         ' +
    '        A.CAS_FEEDBACK_DEST_ID,           ' +
    '        A.CAS_FEEDBACK_MOPP_ID,           ' +
    '        A.CAS_CONTROL_TRANS_ID,           ' +
    '        A.CAS_FEEDBACK_TRANS_ID,          ' +
    '        A.CAS_MAIL_ID,                    ' +
    '        A.CAS_CMD_MAX_COUNTER             ' +
    '   FROM CAS_ENV A               ';
  DataReader.Open;
  DataReader.First;
  FCASEnv.IP := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_IP' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingIP, [FCASEnv.IP] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.ControlPort := StrToInt( FLbDES.DecryptString( DataReader.FieldByName( 'CAS_CONTROL_PORT' ).AsString ) );
  FMsgSubject.Info( Format( SConfigCASReadingSPort, [FCASEnv.ControlPort] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.FeedbackPort := StrToInt( FLbDES.DecryptString( DataReader.FieldByName( 'CAS_FEEDBACK_PORT' ).AsString ) );
  FMsgSubject.Info( Format( SConfigCASReadingRPort, [FCASEnv.FeedbackPort] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.ControlSourceId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_CONTROL_SOURCE_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingControlSourceId, [FCASEnv.ControlSourceId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.ControlDestId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_CONTROL_DEST_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingControlDestId, [FCASEnv.ControlDestId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.ControlMopPPId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_CONTROL_MOPP_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingControlMopPPId, [FCASEnv.ControlMopPPId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.FeedbackSourceId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_FEEDBACK_SOURCE_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingFeedbackSourceId, [FCASEnv.FeedbackSourceId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.FeedbackDestId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_FEEDBACK_DEST_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingFeedbackDestId, [FCASEnv.FeedbackDestId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.FeedbackMopPPId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_FEEDBACK_MOPP_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadingFeedbackMopPPId, [FCASEnv.FeedbackMopPPId] ) );
  { }
  FCASEnv.ControlTransId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_CONTROL_TRANS_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadControlTransId, [FCASEnv.ControlTransId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.FeedbackTransId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_FEEDBACK_TRANS_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadFeedbackTransId, [FCASEnv.FeedbackTransId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.MailId := FLbDES.DecryptString( DataReader.FieldByName( 'CAS_MAIL_ID' ).AsString );
  FMsgSubject.Info( Format( SConfigCASReadMailId, [FCASEnv.MailId] ) );
  Delay( COMMON_DELAY );
  { }
  FCASEnv.CmdMaxCounter := StrToInt( FLbDES.DecryptString( DataReader.FieldByName( 'CAS_CMD_MAX_COUNTER' ).AsString ) );
  FMsgSubject.Info( Format( SConfigCASReadCmdMaxCounter, [FCASEnv.CmdMaxCounter] ) );
  Delay( COMMON_DELAY );
  { }
  FMsgSubject.OK( SConfigCASReadSuccess );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCmdError: Boolean;
begin
  if not VarIsNull( CmdError.Data ) then
  begin
    CmdError.EmptyDataSet;
    CmdError.Data := Null;
  end;
  CmdError.CreateDataSet;
  FMsgSubject.Info( SConfigCmdErrorReading );
  Delay( COMMON_DELAY );
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.ERROR_FLAG, A.ERROR_CODE, A.ERROR_DESC   ' +
    '   FROM CMD_ERROR A                                ' +
    '  ORDER BY A.ERROR_FLAG, A.ERROR_CODE              ';
  DataReader.Open;
  DataReader.First;
  while not DataReader.Eof do
  begin
    CmdError.AppendRecord( [
      FLbDES.DecryptString( DataReader.FieldByName( 'ERROR_FLAG' ).AsString ),
      FLbDES.DecryptString( DataReader.FieldByName( 'ERROR_CODE' ).AsString ),
      FLbDES.DecryptString( DataReader.FieldByName( 'ERROR_DESC' ).AsString )] );
    DataReader.Next;
  end;
  DataReader.Close;
  FMsgSubject.OK( SConfigCmdErrorReadSuccess );
  Delay( COMMON_DELAY );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetFeedbackSo: Boolean;
begin
  FMsgSubject.Info( SConfigFeedbackDbRead );
  Delay( COMMON_DELAY * 2 );
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.SO_SELECTED,                      ' +
    '        A.SO_POS,                           ' +
    '        B.SO_POS_NAME,                      ' +
    '        A.SO_TYPE,                          ' +
    '        A.SO_COMP_CODE,                     ' +
    '        A.SO_COMP_NAME,                     ' +
    '        A.SO_NETWORK_ID,                    ' +
    '        A.SO_LOGINUSER,                     ' +
    '        A.SO_LOGINPASS,                     ' +
    '        A.SO_DBALIASE                       ' +
    '   FROM SO_COMP A, SOPOS_ENV B              ' +
    '  WHERE A.SO_POS = B.SO_POS                 ' +
    '    AND A.SO_TYPE = ''UyGoR8/qpug=''        '; //RECV
  DataReader.Open;
  if ( DataReader.IsEmpty ) then
    FFeedbackSo.Pos := -1
  else begin
    DataReader.First;
    FFeedbackSo.Selected := ( FLbDES.DecryptString(
      DataReader.FieldByName( 'SO_SELECTED' ).AsString ) = 'Y' );
    FFeedbackSo.Pos := StrToInt( FLbDES.DecryptString( DataReader.FieldByName( 'SO_POS' ).AsString ) );
    FFeedbackSo.PosName := FLbDES.DecryptString( DataReader.FieldByName( 'SO_POS_NAME' ).AsString );
    FFeedbackSo.SoType := FLbDES.DecryptString( DataReader.FieldByName( 'SO_TYPE' ).AsString );
    FFeedbackSo.CompCode := FLbDES.DecryptString( DataReader.FieldByName( 'SO_COMP_CODE' ).AsString );
    FFeedbackSo.CompName := FLbDES.DecryptString( DataReader.FieldByName( 'SO_COMP_NAME' ).AsString );
    FFeedbackSo.NetworkId := FLbDES.DecryptString( DataReader.FieldByName( 'SO_NETWORK_ID' ).AsString );
    FFeedbackSo.LoginUser := FLbDES.DecryptString( DataReader.FieldByName( 'SO_LOGINUSER' ).AsString );
    FFeedbackSo.LoginPass := FLbDES.DecryptString( DataReader.FieldByName( 'SO_LOGINPASS' ).AsString );
    FFeedbackSo.DbAliase := FLbDES.DecryptString( DataReader.FieldByName( 'SO_DBALIASE' ).AsString );
    if FFeedbackSo.Selected then
      FFeedbackSo.DbConnectStatus := dbNone
    else
      FFeedbackSo.DbConnectStatus := dbNoSelect;
    FFeedbackSo.RecordCount := 0;
    //FFeedbackSo.ItemIndex := -1;
    FFeedbackSo.Data := nil;
    FFeedbackSo.NotifyMsg := EmptyStr;
    FFeedbackSo.CriticalError := 0;
    FFeedbackSo.LastCriticalErrorTickCount := 0;
  end;
  FMsgSubject.OK( SConfigFeedbackDbSuccess );
  Delay( COMMON_DELAY );
  DataReader.Close;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetStartupInfo: Boolean;
var
  aIndex: Integer;
  aFileName: String;
begin
  FAppStartupInfo.AutoExecute := False;
  FAppStartupInfo.StartupAssignedConfigFile := EmptyStr;
  for aIndex := 0 to ParamCount do
  begin
    if CompareText( ParamStr( aIndex ), 'AUTO_INVOKE' ) = 0 then
      FAppStartupInfo.AutoExecute := True
    else if AnsiPos( DEFAULT_CONFIGFILE_EXT, AnsiUpperCase( ParamStr( aIndex ) ) ) > 0 then
      FAppStartupInfo.StartupAssignedConfigFile := ParamStr( aIndex );
  end;
  FConfigFileManager.ActiveFile := FAppStartupInfo.StartupAssignedConfigFile;
  { 啟動參數沒給設定檔, 則從參數檔 AppInfo.ini 載入 }
  if FConfigFileManager.ActiveFile = EmptyStr then
  begin
    aFileName :=  IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + DEFAULT_APPFILE;
    FConfigFileManager.Restore( aFileName );
  end;  
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ClearSoList;
var
  aLockList: TList;
begin
  aLockList := FSoList.LockList;
  try
    InternalClearSoList( aLockList );
  finally
    FSoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ShowConfigForm: TModalResult;
begin
  fmConfig := TfmConfig.Create( nil );
  try
    Result := fmConfig.ShowModal;
    if Result = mrOk then
    begin
      Screen.Cursor := crHourGlass;
      try
        GetConfigFormValue;
        ConfigSaveToFile;
      finally
        Screen.Cursor := crDefault;
      end;  
      { 若是由選項設定按進來的, 必須把自動執行關掉 }
      { AppStartupInfo.AutoExecute 值, 只在 ConfigModule 啟動時讀取一次,
        再次讀取設定檔, 不須再次讀入此值 }
      FAppStartupInfo.AutoExecute := False;
    end;
  finally
    fmConfig.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.GetConfigFormValue;
var
  aList: TList;
  aIndex: Integer;
  aNode: TcxTreeListNode;
  aSo: PSoInfo;
  aFlag: Integer;
begin
  FCASEnv.IP := fmConfig.txtCASIP.Text;
  FCASEnv.ControlPort := StrToIntDef( fmConfig.txtControlPort.Text, 60002 );
  FCASEnv.FeedbackPort := StrToIntDef( fmConfig.txtFeedbackPort.Text, 60003 );
  {}
  FCASEnv.ControlSourceId := fmConfig.txtControlSourceId.Text;
  FCASEnv.ControlDestId := fmConfig.txtControlDestId.Text;
  FCASEnv.ControlMopPPId := fmConfig.txtControlMopPpid.Text;
  {}
  FCASEnv.FeedbackSourceId := fmConfig.txtFeedbackSourceId.Text;
  FCASEnv.FeedbackDestId := fmConfig.txtFeedbackDestId.Text;
  FCASEnv.FeedbackMopPPId := fmConfig.txtFeedbackMopPpid.Text;
  {}
  FCASEnv.ControlTransId := fmConfig.txtControlTransId.Text;
  FCASEnv.FeedbackTransId := fmConfig.txtFeedbackTransId.Text;
  FCASEnv.CmdMaxCounter := StrToIntDef( fmConfig.txtCasCmdMaxCounter.Text, 1000 );
  {}
  if fmConfig.chkProcessIPPV_A.Checked then
    FCommonEnv.ProcessIPPV := 'A'
  else if fmConfig.chkProcessIPPV_N.Checked then
    FCommonEnv.ProcessIPPV := 'N'
  else
    FCommonEnv.ProcessIPPV := 'O';
  {}
  if fmConfig.ChkProcessBatch_A.Checked then
    FCommonEnv.ProcessBatch := 'A'
  else if fmConfig.ChkProcessBatch_N.Checked then
    FCommonEnv.ProcessBatch := 'N'
  else
    FCommonEnv.ProcessBatch := 'O';
  {}
  FCommonEnv.ProcessRecordCount := StrToIntDef( fmConfig.txtProcessRecordCount.Text, 20 );
  FCommonEnv.DbMultiThread := fmConfig.chkDbMultiThread.Checked;
  FCommonEnv.DbRetryFrequence := StrToIntDef( fmConfig.txtDbRetryFrequence.Text, 60 );
  { 程式沒用到, 給預設值 }
  FCommonEnv.DbResendCount := 3;
  FCommonEnv.DbWriteErrorWhenSocketFail := fmConfig.chkWriteYes.Checked;
  FCommonEnv.BusyTimeStart := fmConfig.txtBusyTimeStart.Text;
  FCommonEnv.BusyTimeEnd := fmConfig.txtBusyTimeEnd.Text;
  FCommonEnv.BusyTimeDbReadFrequence := StrToFloatDef( fmConfig.txtlBusyTimeReadFrquence.Text, 3 );
  FCommonEnv.NormalTimeDbReadFrequence := StrToFloatDef( fmConfig.txtNormalTime.Text, 3 );
  FCommonEnv.BatchOperator := fmConfig.txtBatchOperator.Text;
  FCommonEnv.CASRetryFrequence := StrToIntDef( fmConfig.txtCasRetryFrquence.Text, 60 );
  FCommonEnv.CASSendErrMax := StrToIntDef( fmConfig.txtCasSendErrMax.Text, 3 );
  FCommonEnv.CASRecvErrMax := StrToIntDef( fmConfig.txtCasRecvErrMax.Text, 3 );
  FCommonEnv.CASCommCheck := StrToIntDef( fmConfig.txtCasCommCheck.Text, 10 );
  {}
  FCmdTransEnv.UseSimulator := fmConfig.chkUseSimulator.Checked;
  FCmdTransEnv.EnableControlSend := fmConfig.chkEnableControlSend.Checked;
  FCmdTransEnv.EnableFeedbackRecv := fmConfig.chkEnableFeedbackRecv.Checked;
  FCmdTransEnv.SendCommandDelay := StrToFloatDef( fmConfig.txtSendCommandDelay.Text, 0.5 );
  {}
  aList := FSoList.LockList;
  try
    InternalClearSoList( aList );
    for aIndex := 0 to fmConfig.SoTree.Nodes.Count - 1 do
    begin
      if fmConfig.SoTree.Nodes.Items[aIndex].Level = 0 then Continue;
      aNode := fmConfig.SoTree.Nodes.Items[aIndex];
      New( aSo );
      aSo.Selected := ( aNode.Values[fmConfig.SoTreeColumn1.ItemIndex] = 'Y' );
      aSo.Pos := StrToIntDef( aNode.Texts[fmConfig.SoTreeColumn9.ItemIndex], 1 );
      aSo.PosName := aNode.Parent.Texts[fmConfig.SoTreeColumn3.ItemIndex];
      aSo.SoType := aNode.Texts[fmConfig.SoTreeColumn8.ItemIndex];
      aSo.CompCode := aNode.Texts[fmConfig.SoTreeColumn2.ItemIndex];
      aSo.CompName := aNode.Texts[fmConfig.SoTreeColumn3.ItemIndex];
      aSo.NetworkId := aNode.Texts[fmConfig.SoTreeColumn4.ItemIndex];
      aSo.LoginUser := aNode.Texts[fmConfig.SoTreeColumn5.ItemIndex];
      aSo.LoginPass := aNode.Texts[fmConfig.SoTreeColumn6.ItemIndex];
      aSo.DbAliase := aNode.Texts[fmConfig.SoTreeColumn7.ItemIndex];
      if not aSo.Selected then
        aSo.DbConnectStatus := dbNoSelect
      else
        aSo.DbConnectStatus := dbNone;
      aSo.RecordCount := 0;
      aSo.Data := nil;
      aSo.NotifyMsg := EmptyStr;
      aSo.CriticalError := 0;
      aSo.LastCriticalErrorTickCount := 0;
      if aNode.Texts[fmConfig.SoTreeColumn8.ItemIndex] = 'SEND' then
        aList.Add( aSo )
      else begin
        Dispose( FFeedbackSo );
        FFeedbackSo := aSo;
      end;
    end;
  finally
    FSoList.UnlockList;
  end;
  {}
  HighCmdEnv.EmptyDataSet;
  for aIndex := 0 to fmConfig.HighCmdTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.HighCmdTree.Nodes.Items[aIndex];
    HighCmdEnv.AppendRecord( [
      //指令別 CMD_TYPE
      aNode.Texts[fmConfig.HighCmdTreeColumn3.ItemIndex],
      //高階指令 HIGH_LEVEL_CMD
      aNode.Texts[fmConfig.HighCmdTreeColumn1.ItemIndex],
      //名稱 DESCRIPTION
      aNode.Texts[fmConfig.HighCmdTreeColumn2.ItemIndex],
      //對應低階指令 LOW_LEVEL_CMD
      aNode.Texts[fmConfig.HighCmdTreeColumn4.ItemIndex],
      //對應IRD指令 IRD_CMD
      aNode.Texts[fmConfig.HighCmdTreeColumn8.ItemIndex],
      //多組拆解 NOTES_USE
      aNode.Texts[fmConfig.HighCmdTreeColumn5.ItemIndex],
      //指令別優先權 CMD_TYPE_PRIORITY
      aNode.Texts[fmConfig.HighCmdTreeColumn6.ItemIndex],
      //傳送優先權 CMD_PRIORITY
      aNode.Texts[fmConfig.HighCmdTreeColumn7.ItemIndex]] );
  end;
  {}
  LowCmdEnv.EmptyDataSet;
  for aIndex := 0 to fmConfig.LowCmdTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.LowCmdTree.Nodes.Items[aIndex];
    LowCmdEnv.AppendRecord( [
      aNode.Texts[fmConfig.LowCmdTreeColumn1.ItemIndex],
      aNode.Texts[fmConfig.LowCmdTreeColumn2.ItemIndex]] );
  end;
  {}
  CmdError.EmptyDataSet;
  for aIndex := 0 to fmConfig.ErrorTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.ErrorTree.Nodes.Items[aIndex];
    aFlag := 0;
    if aNode.Texts[fmConfig.ErrorTreeColumn1.ItemIndex] = 'EXT' then
      aFlag := 1;
    CmdError.AppendRecord( [
      aFlag,
      aNode.Texts[fmConfig.ErrorTreeColumn2.ItemIndex],
      aNode.Texts[fmConfig.ErrorTreeColumn3.ItemIndex] ] );
  end;
end;

{ ---------------------------------------------------------------------------- }


function TConfigModule.SetCASEnv: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM CAS_ENV ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO CAS_ENV (        ' +
    '     CAS_IP,                  ' +
    '     CAS_CONTROL_PORT,        ' +
    '     CAS_FEEDBACK_PORT,       ' +
    '     CAS_CONTROL_SOURCE_ID,   ' +
    '     CAS_CONTROL_DEST_ID,     ' +
    '     CAS_CONTROL_MOPP_ID,     ' +
    '     CAS_FEEDBACK_SOURCE_ID,  ' +
    '     CAS_FEEDBACK_DEST_ID,    ' +
    '     CAS_FEEDBACK_MOPP_ID,    ' +
    '     CAS_CONTROL_TRANS_ID,    ' +
    '     CAS_FEEDBACK_TRANS_ID,   ' +
    '     CAS_MAIL_ID,             ' +
    '     CAS_CMD_MAX_COUNTER  )   ' +
    ' VALUES (                     ' +
    '     :1,                      ' +
    '     :2,                      ' +
    '     :3,                      ' +
    '     :4,                      ' +
    '     :5,                      ' +
    '     :6,                      ' +
    '     :7,                      ' +
    '     :8,                      ' +
    '     :9,                      ' +
    '     :10,                     ' +
    '     :11,                     ' +
    '     :12,                     ' +
    '     :13       )              ';
  DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( FCASEnv.IP );
  DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( IntToStr( FCASEnv.ControlPort ) );
  DataReader.Parameters.ParamByName( '3' ).Value := FLbDES.EncryptString( IntToStr( FCASEnv.FeedbackPort ) );
  DataReader.Parameters.ParamByName( '4' ).Value := FLbDES.EncryptString( FCASEnv.ControlSourceId );
  DataReader.Parameters.ParamByName( '5' ).Value := FLbDES.EncryptString( FCASEnv.ControlDestId );
  DataReader.Parameters.ParamByName( '6' ).Value := FLbDES.EncryptString( FCASEnv.ControlMopPPId );
  DataReader.Parameters.ParamByName( '7' ).Value := FLbDES.EncryptString( FCASEnv.FeedbackSourceId );
  DataReader.Parameters.ParamByName( '8' ).Value := FLbDES.EncryptString( FCASEnv.FeedbackDestId );
  DataReader.Parameters.ParamByName( '9' ).Value := FLbDES.EncryptString( FCASEnv.FeedbackMopPPId );
  DataReader.Parameters.ParamByName( '10' ).Value := FLbDES.EncryptString( FCASEnv.ControlTransId );
  DataReader.Parameters.ParamByName( '11' ).Value := FLbDES.EncryptString( FCASEnv.FeedbackTransId );
  DataReader.Parameters.ParamByName( '12' ).Value := FLbDES.EncryptString( FCASEnv.MailId );
  DataReader.Parameters.ParamByName( '13' ).Value := FLbDES.EncryptString( IntToStr( FCASEnv.CmdMaxCounter ) );
  try
    DataReader.ExecSQL;
    FMsgSubject.OK( SConfigCASWritSuccess );
  except
    on E: Exception do
      FMsgSubject.Error( Format( SConfigCASWriteError, [E.Message] ) );
  end;
  Result := True;  
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetCommon: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM COMMON_ENV ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO COMMON_ENV (         ' +
    '   DBMULTITHREAD,                 ' +
    '   DBRETRYFREQUENCE,              ' +
    '   BUSY_TIME_START,               ' +
    '   BUSY_TIME_END,                 ' +
    '   BUSY_TIME_READ_FREQUENCY,      ' +
    '   NORMAL_TIME_READ_FREQUENCY,    ' +
    '   PROCESSRECORDCOUNT,            ' +
    '   DBRESENDCOUNT,                 ' +
    '   DBWRITEERRORWHENSOCKETFAIL,    ' +
    '   PROCESSIPPV,                   ' +
    '   PROCESSBATCH,                  ' +
    '   BATCHOPERATOR,                 ' +
    '   CASRETRYFREQUENCE,             ' +
    '   CASSENDERRMAX,                 ' +
    '   CASRECVERRMAX,                 ' +
    '   CASCOMMCHECK     )             ' +
    ' VALUES (                         ' +
    '   :1,                            ' +
    '   :2,                            ' +
    '   :3,                            ' +
    '   :4,                            ' +
    '   :5,                            ' +
    '   :6,                            ' +
    '   :7,                            ' +
    '   :8,                            ' +
    '   :9,                            ' +
    '   :10,                           ' +
    '   :11,                           ' +
    '   :12,                           ' +
    '   :13,                           ' +
    '   :14,                           ' +
    '   :15,                           ' +
    '   :16              )             ';
  if FCommonEnv.DbMultiThread then
    DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( 'Y' )
  else
    DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( 'N' );
  DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.DbRetryFrequence ) );
  DataReader.Parameters.ParamByName( '3' ).Value := FLbDES.EncryptString( FCommonEnv.BusyTimeStart );
  DataReader.Parameters.ParamByName( '4' ).Value := FLbDES.EncryptString( FCommonEnv.BusyTimeEnd );
  DataReader.Parameters.ParamByName( '5' ).Value := FLbDES.EncryptString( FloatToStr( FCommonEnv.BusyTimeDbReadFrequence ) );
  DataReader.Parameters.ParamByName( '6' ).Value := FLbDES.EncryptString( FloatToStr( FCommonEnv.NormalTimeDbReadFrequence ) );
  DataReader.Parameters.ParamByName( '7' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.ProcessRecordCount ) );
  DataReader.Parameters.ParamByName( '8' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.DbResendCount ) );
  if FCommonEnv.DbWriteErrorWhenSocketFail then
    DataReader.Parameters.ParamByName( '9' ).Value := FLbDES.EncryptString( 'Y' )
  else
    DataReader.Parameters.ParamByName( '9' ).Value := FLbDES.EncryptString( 'N' );
  DataReader.Parameters.ParamByName( '10' ).Value := FLbDES.EncryptString( FCommonEnv.ProcessIPPV );
  DataReader.Parameters.ParamByName( '11' ).Value := FLbDES.EncryptString( FCommonEnv.ProcessBatch );
  DataReader.Parameters.ParamByName( '12' ).Value := FLbDES.EncryptString( FCommonEnv.BatchOperator );
  DataReader.Parameters.ParamByName( '13' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.CASRetryFrequence ) );
  DataReader.Parameters.ParamByName( '14' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.CASSendErrMax ) );
  DataReader.Parameters.ParamByName( '15' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.CASRecvErrMax ) );
  DataReader.Parameters.ParamByName( '16' ).Value := FLbDES.EncryptString( IntToStr( FCommonEnv.CASCommCheck ) );
  try
    DataReader.ExecSQL;
    FMsgSubject.OK( SConfigCommonWritSuccess );
  except
    on E: Exception do
      FMsgSubject.Error( Format( SConfigCommonWriteError, [E.Message] ) );
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetCmdTransEnv: Boolean;
var
  aIsError: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM CMDTRANS_ENV ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO CMDTRANS_ENV (        ' +
    '    USESIMULATOR,                  ' +
    '    ENABLECONTROLSEND,             ' +
    '    ENABLEFEEDBACKRECV,            ' +
    '    SENDCOMMANDDELAY   )           ' +
    ' VALUES (                          ' +
    '    :1,                            ' +
    '    :2,                            ' +
    '    :3,                            ' +
    '    :4    )                        ';
  DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( 'N' );
  DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( 'N' );
  DataReader.Parameters.ParamByName( '3' ).Value := FLbDES.EncryptString( 'N' );
  {}
  if fmConfig.chkUseSimulator.Checked then
    DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( 'Y' );
  if fmConfig.chkEnableControlSend.Checked then
    DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( 'Y' );
  if fmConfig.chkEnableFeedbackRecv.Checked then
    DataReader.Parameters.ParamByName( '3' ).Value := FLbDES.EncryptString( 'Y' );
  DataReader.Parameters.ParamByName( '4' ).Value :=
    FLbDES.EncryptString( fmConfig.txtSendCommandDelay.Text );
  aIsError := False;  
  try
    DataReader.ExecSQL;
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( SCOnfigCmdTransWriteError, [E.Message] ) );
      aIsError := True;
    end;
  end;
  if not aIsError then FMsgSubject.OK( SCOnfigCmdTransWriteSuccess );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetSoList: Boolean;
var
  aIndex: Integer;
  aList: TList;
  aSo: PSoInfo;
  aIsError: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM SO_COMP ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO SO_COMP (         ' +
    '    SO_TYPE,                   ' +
    '    SO_SELECTED,               ' +
    '    SO_POS,                    ' +
    '    SO_COMP_CODE,              ' +
    '    SO_COMP_NAME,              ' +
    '    SO_NETWORK_ID,             ' +
    '    SO_LOGINUSER,              ' +
    '    SO_LOGINPASS,              ' +
    '    SO_DBALIASE )              ' +
    ' VALUES (                      ' +
    '    :1,                        ' +
    '    :2,                        ' +
    '    :3,                        ' +
    '    :4,                        ' +
    '    :5,                        ' +
    '    :6,                        ' +
    '    :7,                        ' +
    '    :8,                        ' +
    '    :9          )              ';
  aList := FSoList.LockList;
  try
    aIsError := False;
    for aIndex := 0 to aList.Count - 1 do
    begin
      aSo := PSoInfo( aList.Items[aIndex] );
      DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( aSo.SoType );
      if aSo.Selected then
        DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( 'Y' )
      else
        DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( 'N' );
      DataReader.Parameters.ParamByName( '3' ).Value := FLbDES.EncryptString( IntToStr( aSo.Pos ) );
      DataReader.Parameters.ParamByName( '4' ).Value := FLbDES.EncryptString( aSo.CompCode );
      DataReader.Parameters.ParamByName( '5' ).Value := FLbDES.EncryptString( aSo.CompName );
      DataReader.Parameters.ParamByName( '6' ).Value := FLbDES.EncryptString( aSo.NetworkId );
      DataReader.Parameters.ParamByName( '7' ).Value := FLbDES.EncryptString( aSo.LoginUser );
      DataReader.Parameters.ParamByName( '8' ).Value := FLbDES.EncryptString( aSo.LoginPass );
      DataReader.Parameters.ParamByName( '9' ).Value := FLbDES.EncryptString( aSo.DbAliase );
      try
        DataReader.ExecSQL;
      except
        on E: Exception do
        begin
          FMsgSubject.Error( Format( SConfigSoWriteError, [E.Message] ) );
          aIsError := True;
          Break;
        end;
      end;
    end;
    DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( FFeedbackSo.SoType );
    if FFeedbackSo.Selected then
      DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( 'Y' )
    else
      DataReader.Parameters.ParamByName( '2' ).Value := FLbDES.EncryptString( 'N' );
    DataReader.Parameters.ParamByName( '3' ).Value := FLbDES.EncryptString( IntToStr( FFeedbackSo.Pos ) );
    DataReader.Parameters.ParamByName( '4' ).Value := FLbDES.EncryptString( FFeedbackSo.CompCode );
    DataReader.Parameters.ParamByName( '5' ).Value := FLbDES.EncryptString( FFeedbackSo.CompName );
    DataReader.Parameters.ParamByName( '6' ).Value := FLbDES.EncryptString( FFeedbackSo.NetworkId );
    DataReader.Parameters.ParamByName( '7' ).Value := FLbDES.EncryptString( FFeedbackSo.LoginUser );
    DataReader.Parameters.ParamByName( '8' ).Value := FLbDES.EncryptString( FFeedbackSo.LoginPass );
    DataReader.Parameters.ParamByName( '9' ).Value := FLbDES.EncryptString( FFeedbackSo.DbAliase );
    try
      DataReader.ExecSQL;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SConfigSoWriteError, [E.Message] ) );
        aIsError := True;
      end;
    end;
    if not aIsError then FMsgSubject.OK( SConfigSoWriteSuccess );
  finally
    FSoList.UnlockList;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetHighCmdEnv: Boolean;
var
  aIndex: Integer;
  aNode: TcxTreeListNode;
  aIsError: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM HIGH_CMD_ENV ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO HIGH_CMD_ENV (         ' +
    '    CMD_TYPE,                       ' +
    '    HIGH_LEVEL_CMD,                 ' +
    '    DESCRIPTION,                    ' +
    '    LOW_LEVEL_CMD,                  ' +
    '    IRD_CMD,                        ' +
    '    NOTES_USE,                      ' +
    '    CMD_TYPE_PRIORITY,              ' +
    '    CMD_PRIORITY,                   ' +
    '    NEW_CMD_COUNT,                  ' +
    '    OLD_CMD_COUNT  )                ' +
    ' VALUES (                           ' +
    '    :1,                             ' +
    '    :2,                             ' +
    '    :3,                             ' +
    '    :4,                             ' +
    '    :5,                             ' +
    '    :6,                             ' +
    '    :7,                             ' +
    '    :8,                             ' +
    '    :9,                             ' +
    '    :10    )                        ';
  aIsError := False;
  for aIndex := 0 to fmConfig.HighCmdTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.HighCmdTree.Nodes.Items[aIndex];
    DataReader.Parameters.ParamByName( '1' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn3.ItemIndex] );
    DataReader.Parameters.ParamByName( '2' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn1.ItemIndex] );
    DataReader.Parameters.ParamByName( '3' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn2.ItemIndex] );
    DataReader.Parameters.ParamByName( '4' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn4.ItemIndex] );
    DataReader.Parameters.ParamByName( '5' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn8.ItemIndex] );
    DataReader.Parameters.ParamByName( '6' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn5.ItemIndex] );
    DataReader.Parameters.ParamByName( '7' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn6.ItemIndex] );
    DataReader.Parameters.ParamByName( '8' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn7.ItemIndex] );
    DataReader.Parameters.ParamByName( '9' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn9.ItemIndex] );
    DataReader.Parameters.ParamByName( '10' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.HighCmdTreeColumn10.ItemIndex] );
    try
      DataReader.ExecSQL;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SConfigHighCmdWriteError, [E.Message] ) );
        aIsError := True;
        Break;
      end;
    end;
  end;
  if not aIsError then FMsgSubject.OK( SConfigHighCmdWriteSuccess );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetLowCmdEnv: Boolean;
var
  aIndex: Integer;
  aNode: TcxTreeListNode;
  aIsError: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM LOW_CMD_ENV ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO LOW_CMD_ENV (          ' +
    '    LOW_LEVEL_CMD,                  ' +
    '    DESCRIPTION )                   ' +
    ' VALUES (                           ' +
    '    :1,                             ' +
    '    :2 )                            ';
  aIsError := False;  
  for aIndex := 0 to fmConfig.LowCmdTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.LowCmdTree.Nodes.Items[aIndex];
    DataReader.Parameters.ParamByName( '1' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.LowCmdTreeColumn1.ItemIndex] );
    DataReader.Parameters.ParamByName( '2' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.LowCmdTreeColumn2.ItemIndex] );
    try
      DataReader.ExecSQL;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SConfigLowCmdWriteError, [E.Message] ) );
        aIsError := True;
        Break;
      end;
    end;
  end;
  if not aIsError then FMsgSubject.OK( SConfigLowCmdWriteSuccess );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SetCmdError: Boolean;
var
  aIndex: Integer;
  aNode: TcxTreeListNode;
  aIsError: Boolean;
begin
  DataReader.SQL.Text := ' DELETE FROM CMD_ERROR ';
  DataReader.ExecSQL;
  DataReader.SQL.Text :=
    ' INSERT INTO CMD_ERROR (            ' +
    '    ERROR_FLAG,                     ' +
    '    ERROR_CODE,                     ' +
    '    ERROR_DESC  )                   ' +
    ' VALUES (                           ' +
    '    :1,                             ' +
    '    :2,                             ' +
    '    :3    )                         ';
  aIsError := False;  
  for aIndex := 0 to fmConfig.ErrorTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.ErrorTree.Nodes.Items[aIndex];
    if aNode.Texts[fmConfig.ErrorTreeColumn1.ItemIndex] = 'CODE' then
      DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( '0' )
    else
      DataReader.Parameters.ParamByName( '1' ).Value := FLbDES.EncryptString( '1' );
    DataReader.Parameters.ParamByName( '2' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.ErrorTreeColumn2.ItemIndex] );
    DataReader.Parameters.ParamByName( '3' ).Value :=
      FLbDES.EncryptString( aNode.Texts[fmConfig.ErrorTreeColumn3.ItemIndex] );
    try
      DataReader.ExecSQL;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( SConfigCmdErrorWriteError, [E.Message] ) );
        aIsError := True;
        Break;
      end;
    end;
  end;
  if not aIsError then FMsgSubject.OK( SConfigCmdErrorWriteSuccess );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

end.
