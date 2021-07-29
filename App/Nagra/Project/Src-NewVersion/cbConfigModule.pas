unit cbConfigModule;

interface

uses
  SysUtils, Classes, Windows, Variants, IniFiles, DB, ADODB, cbClass,
  DBClient;

type
  TConfigModule = class(TDataModule)
    AccessConnection: TADOConnection;
    DataReader: TADOQuery;
    CmdEnv: TClientDataSet;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FSoList: TThreadList;
    FConfigFileName: String;
    FConfigIsReady: Boolean;
    FAppStartupInfo: TAppStartupInfo;
    FCommonEnv: TCommonEnv;
    FCASEnv: TCASEnv;
    FMsgSubject: TMessageSubject;
    function GetSoCount: Integer;
    function ConnectToAccess: Boolean;
    procedure DisconnectFromAccess;
    procedure InternalClearSoList(const aList: TList);
    procedure ClearSoList;
    function GetSoList: Boolean;
    function GetStartupInfo: Boolean;
    function GetCommon: Boolean;
    function GetCmdEnv: Boolean;
    function GetCASEnv: Boolean;
  public
    { Public declarations }
    constructor Create; reintroduce;
    procedure ConfigSaveToFile;
    procedure ConfigLoadFromFile;
    property ConfigFileName: String read FConfigFileName;
    property AppStartupInfo: TAppStartupInfo read FAppStartupInfo;
    property SoList: TThreadList read FSoList;
    property SoCount: Integer read GetSoCount;
    property CommEnv: TCommonEnv read FCommonEnv;
    property CASEnv: TCASEnv read FCASEnv;
    property ConfigIsReady: Boolean read FConfigIsReady;
  end;

var
  ConfigModule: TConfigModule;

implementation

uses cbResStr, cbUtils, cbMain;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

constructor TConfigModule.Create;
begin
  inherited Create( nil );
  FSoList := TThreadList.Create;
  FMsgSubject := TMessageSubject.Create;
  FConfigIsReady := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleCreate(Sender: TObject);
begin
  FMsgSubject.AutoNotify := True;
  FMsgSubject.AddObServer( fmMain.StartupObServer );
  FMsgSubject.NotifyMessage := Format( SConfigNoneRead, [FConfigFileName] ) ;
  GetStartupInfo;
  ConfigLoadFromFile;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleDestroy(Sender: TObject);
begin
  DisconnectFromAccess;
  ClearSoList;
  FMsgSubject.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ConfigLoadFromFile;
var
  aLoadResult: Boolean;
begin
  FConfigIsReady := FileExists( FAppStartupInfo.ConfigFileName );
  if not FConfigIsReady then
  begin
    FMsgSubject.NotifyMessage := Format( SConfigFileNotExists, [FConfigFileName] ) ;
    Exit;
  end;
  aLoadResult := ConnectToAccess;
  try
    if aLoadResult then
    begin
      GetCommon;
      GetCASEnv;
      GetCmdEnv;
      GetSoList;
    end;
  finally
    DisconnectFromAccess;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ConfigSaveToFile;
begin
  { ... }
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
begin
  AccessConnection.Connected := False;
  AccessConnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [FConfigFileName, aDbPassword] );
  try
    AccessConnection.Connected := True;
    FMsgSubject.NotifyMessage := Format( SConfigOpenSuccess, [FConfigFileName] );
  except
    on E: Exception do
      FMsgSubject.NotifyMessage := Format( SConfigOpenError, [E.Message] );
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
      ' SELECT A.SO_SELECTED, A.SO_POS, B.SO_POS_NAME,     ' +
      '        A.SO_COMP_CODE, A.SO_COMP_NAME,             ' +
      '        A.SO_NETWORK_ID, A.SO_LOGINUSER,            ' +
      '        A.SO_LOGINPASS, A.SO_DBALIASE               ' +
      '   FROM SO_COMP A, SOPOS_ENV B                      ' +
      '  WHERE A.SO_POS = B.SO_POS                         ' +
      '  ORDER BY A.SO_COMP_CODE                           ';
    try
      DataReader.Open;
      DataReader.First;
      aSoCount := 0;
      while not DataReader.Eof do
      begin
        New( aSoInfo );
        aSoInfo.Selected := ( DataReader.FieldByName( 'SO_SELECTED' ).AsString = 'Y' );
        aSoInfo.Pos := DataReader.FieldByName( 'SO_POS' ).AsInteger;
        aSoInfo.PosName := DataReader.FieldByName( 'SO_POS_NAME' ).AsString;
        aSoInfo.CompCode := DataReader.FieldByName( 'SO_COMP_CODE' ).AsString;
        aSoInfo.CompName := DataReader.FieldByName( 'SO_COMP_NAME' ).AsString;
        aSoInfo.NetworkId := DataReader.FieldByName( 'SO_NETWORK_ID' ).AsString;
        aSoInfo.LoginUser := DataReader.FieldByName( 'SO_LOGINUSER' ).AsString;
        aSoInfo.LoginPass := DataReader.FieldByName( 'SO_LOGINPASS' ).AsString;
        aSoInfo.DbAliase := DataReader.FieldByName( 'SO_DBALIASE' ).AsString;
        if aSoInfo.Selected then
          aSoInfo.DbConnectStatus := dbNone
        else
          aSoInfo.DbConnectStatus := dbNoSelect;
        aSoInfo.RecordCount := 0;
        aSoInfo.ItemIndex := -1;
        aSoInfo.NotifyMsg := EmptyStr;
        aSoInfo.CriticalError := 0;
        aSoInfo.LastCriticalErrorTickCount := 0;
        aLockList.Add( aSoInfo );
        Inc( aSoCount );
        FMsgSubject.NotifyMessage := Format( SConfigSoRead, [aSoInfo.CompName] );
        Delay( COMMON_DELAY );
        DataReader.Next;
      end;
      FMsgSubject.NotifyMessage := Format( SConfigSoReadSuccess, [aSoCount] );
      Result := True;
    except
      on E: Exception do
      begin
        FMsgSubject.NotifyMessage := Format( SConfigSoReadError, [E.Message] );
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
  FCommonEnv.BusyTimeStart := EmptyStr;
  FCommonEnv.BusyTimeEnd := EmptyStr;
  FCommonEnv.BusyTimeDbReadFrequence := 3;
  FCommonEnv.NormalTimeDbReadFrequence := 30; 
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.PROCESSRECORDCOUNT,         ' +
    '        A.PROCESSIPPV,                ' +
    '        A.PROCESSBATCH,               ' +
    '        A.DBMULTITHREAD,              ' +
    '        A.DBRETRYFREQUENCE,           ' +
    '        A.DBRESENDCOUNT,              ' +
    '        A.BUSY_TIME_START,            ' +
    '        A.BUSY_TIME_END,              ' +
    '        A.BUSY_TIME_READ_FREQUENCY,   ' +
    '        A.NORMAL_TIME_READ_FREQUENCY, ' +
    '        A.BATCHOPERATOR               ' + 
    '   FROM COMMON_EVN A                  ';
  try
    DataReader.Open;
    DataReader.First;
    { }
    FCommonEnv.ProcessRecordCount :=
      DataReader.FieldByName( 'PROCESSRECORDCOUNT' ).AsInteger;
    FMsgSubject.NotifyMessage :=
      Format( SConfigCommonProcessRecordCount, [FCommonEnv.ProcessRecordCount] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.ProcessIPPV :=
      DataReader.FieldByName( 'PROCESSIPPV' ).AsString;
    FMsgSubject.NotifyMessage :=
      Format( SConfigCommonProcessIPPV, [FCommonEnv.ProcessIPPV] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.ProcessBatch :=
      DataReader.FieldByName( 'PROCESSBATCH' ).AsString;
    FMsgSubject.NotifyMessage :=
      Format( SConfigCommonProcessBatch, [FCommonEnv.ProcessBatch] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbMultiThread :=
      ( DataReader.FieldByName( 'DBMULTITHREAD' ).AsString = 'Y' );
    FMsgSubject.NotifyMessage := Format( SConfigCommonDbMultiThread,
      [DataReader.FieldByName( 'DBMULTITHREAD' ).AsString] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbRetryFrequence :=
      DataReader.FieldByName( 'DBRETRYFREQUENCE' ).AsInteger;
    FMsgSubject.NotifyMessage :=
      Format( SConfigCommonDbRetryFrequence, [FCommonEnv.DbRetryFrequence] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.DbResendCount :=
      DataReader.FieldByName( 'DBRESENDCOUNT' ).AsInteger;
    FMsgSubject.NotifyMessage :=
      Format( SConfigCommonDbRetryFrequence, [FCommonEnv.DbResendCount] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BusyTimeStart :=
      DataReader.FieldByName( 'BUSY_TIME_START' ).AsString;
    FMsgSubject.NotifyMessage := Format( SConfigCommonBusyTimeStart,
      [Nvl( FCommonEnv.BusyTimeStart, SConfigNoneSet )] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BusyTimeEnd :=
      DataReader.FieldByName( 'BUSY_TIME_END' ).AsString;
    FMsgSubject.NotifyMessage := Format( SConfigCommonBusyTimeEnd,
      [Nvl( FCommonEnv.BusyTimeEnd, SConfigNoneSet )] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BusyTimeDbReadFrequence :=
      DataReader.FieldByName( 'BUSY_TIME_READ_FREQUENCY' ).AsFloat;
    FMsgSubject.NotifyMessage := Format( SConfigCommonBusyTimeDbReadFrequence,
      [FloatToStr( FCommonEnv.BusyTimeDbReadFrequence )] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.NormalTimeDbReadFrequence :=
      DataReader.FieldByName( 'NORMAL_TIME_READ_FREQUENCY' ).AsFloat;
    FMsgSubject.NotifyMessage := Format( SConfigCommonNormalTimeDbReadFrequence,
      [FloatToStr( FCommonEnv.NormalTimeDbReadFrequence )] );
    Delay( COMMON_DELAY );
    { }
    FCommonEnv.BatchOperator :=
      DataReader.FieldByName( 'BATCHOPERATOR' ).AsString;
    FMsgSubject.NotifyMessage :=
      Format( SConfigCommonBatchOperator, [FCommonEnv.BatchOperator] );
    Delay( COMMON_DELAY );
    FMsgSubject.NotifyMessage := SConfigCommonReadSuccess;
    Result := True;
  except
    on E: Exception do
    begin
      FMsgSubject.NotifyMessage := Format( SConfigCommonReadError, [E.Message] );
      Result := False;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCmdEnv: Boolean;
begin
  if not VarIsNull( CmdEnv.Data ) then
  begin
    CmdEnv.EmptyDataSet;
    CmdEnv.Data := Null;
  end;
  CmdEnv.CreateDataSet;
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.HIGH_LEVEL_CMD,      ' +
    '        A.DESCRIPTION,         ' +
    '        A.LOW_LEVEL_CMD,       ' +
    '        A.CMD_TYPE,            ' +
    '        A.CMD_TYPE_PRIORITY,   ' +
    '        A.CMD_PRIORITY         ' +
    '   FROM CMD_ENV A              ' +
    '  WHERE IIF( ISNULL( A.CMD_STOP ), ''N'', A.CMD_STOP ) = ''N'' ';
  DataReader.Open;
  if DataReader.IsEmpty then
  begin
    FMsgSubject.NotifyMessage := SConfigCmdReadIsEmpty;
    Result := False;
  end else
  begin
    Result := True;
    DataReader.First;
    while not DataReader.Eof do
    begin
      try
        CmdEnv.AppendRecord( [DataReader.FieldByName( 'HIGH_LEVEL_CMD' ).Value,
          DataReader.FieldByName( 'DESCRIPTION' ).Value,
          DataReader.FieldByName( 'LOW_LEVEL_CMD' ).Value,
          DataReader.FieldByName( 'CMD_TYPE' ).Value,
          DataReader.FieldByName( 'CMD_TYPE_PRIORITY' ).Value,
          DataReader.FieldByName( 'CMD_PRIORITY' ).Value] );
        FMsgSubject.NotifyMessage := Format( SConfigCmdReading,
          [DataReader.FieldByName( 'HIGH_LEVEL_CMD' ).AsString,
           DataReader.FieldByName( 'DESCRIPTION' ).AsString] );
      except
        on E: Exception do
        begin
          FMsgSubject.NotifyMessage := Format( SConfigCmdReadError, [E.Message] );
          CmdEnv.EmptyDataSet;
          Result := False;
          Break;
        end;
      end;
      Delay( COMMON_DELAY );
      DataReader.Next;
    end;
    if Result then
    begin
      FMsgSubject.NotifyMessage := SConfigCmdReadSuccess;
      Result := True;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCASEnv: Boolean;
begin
  FCASEnv.IP := '127.0.0.1';
  FCASEnv.ControlPort := 20000;
  FCASEnv.FeedbackPort := 20001;
  FCASEnv.SourceId := '1';
  FCASEnv.DestId := '1';
  FCASEnv.MopPPId := '1';
  DataReader.Close;
  DataReader.SQL.Text :=
    ' SELECT A.CAS_IP, A.CAS_CONTROL_PORT, A.CAS_FEEDBACK_PORT, ' +
    '        A.CAS_SOURCE_ID, A.CAS_DEST_ID, A.CAS_MOPP_ID      ' +
    '   FROM CAS_ENV A                                          ';
  DataReader.Open;
  DataReader.First;
  FCASEnv.IP := DataReader.FieldByName( 'CAS_IP' ).AsString;
  FMsgSubject.NotifyMessage := Format( SConfigCASReadingIP, [FCASEnv.IP] );
  { }
  FCASEnv.ControlPort := DataReader.FieldByName( 'CAS_CONTROL_PORT' ).AsInteger;
  FMsgSubject.NotifyMessage := Format( SConfigCASReadingSPort, [FCASEnv.ControlPort] );
  { }
  FCASEnv.FeedbackPort := DataReader.FieldByName( 'CAS_FEEDBACK_PORT' ).AsInteger;
  FMsgSubject.NotifyMessage := Format( SConfigCASReadingRPort, [FCASEnv.FeedbackPort] );
  { }
  FCASEnv.SourceId := DataReader.FieldByName( 'CAS_SOURCE_ID' ).AsString;
  FMsgSubject.NotifyMessage := Format( SConfigCASReadingSourceId, [FCASEnv.SourceId] );
  { }
  FCASEnv.DestId := DataReader.FieldByName( 'CAS_DEST_ID' ).AsString;
  FMsgSubject.NotifyMessage := Format( SConfigCASReadingDestId, [FCASEnv.DestId] );
  { }
  FCASEnv.MopPPId := DataReader.FieldByName( 'CAS_DEST_ID' ).AsString;
  FMsgSubject.NotifyMessage := Format( SConfigCASReadingMopPPId, [FCASEnv.MopPPId] );
  FMsgSubject.NotifyMessage := SConfigCASReadSuccess;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetStartupInfo: Boolean;
var
  aIndex: Integer;
begin
  FAppStartupInfo.AutoExecute := False;
  FAppStartupInfo.ConfigFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + CONFIG_FILE_NAME;
  for aIndex := 1 to ParamCount - 1 do
  begin
    if CompareText( ParamStr( aIndex ), 'AUTO_INVOKE' ) = 0 then
      FAppStartupInfo.AutoExecute := True
    else if AnsiPos( '.INI', UpperCase( ParamStr( aIndex ) ) ) > 0 then
    begin
      FAppStartupInfo.ConfigFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
        ParamStr( 0 ) ) ) + ParamStr( aIndex );
    end;
  end;
  FConfigFileName := FAppStartupInfo.ConfigFileName;
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

end.
