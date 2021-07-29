unit cbSoDbThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, CommCtrl, DB, ADODB, DBClient,
  {$IFDEF DEBUG} CsIntf, {$ENDIF} cbClass;


type
  TSoDbThread = class(TCommonThread)
  private
    { Private declarations }
    FSoInfo: TList;
    FADOControlList: TList;
    FCurrentActiveSo: PSoInfo;
    FCurrentActiveADO: PSoADOControl;
    function GetSelectSQL: String;
    function GetDbRertyFrquence: Cardinal;
    function CanDbRertyConnect: Boolean;
    function CanAppsendRecord(aIndex: Integer; aSeqno: String): Boolean;
    function CheckDbStatus(aIndex: Integer): Boolean;
    procedure InternalCreate;
    procedure PrepareADOControl;
    procedure ConnectToDb; overload;
    procedure ConnectToDb(const aIndex: Integer); overload;
    procedure DisconnectFromDb; overload;
    procedure DisconnectFromDb(const aIndex: Integer); overload;
    procedure AssignSoInfo(const aSource: TThreadList);
    procedure ReleaseSoInfoResource;
    procedure ReleaseADOResource;
    { 從資料庫取出指令, 寫入傳送指令 List 裏 }
    procedure GetData(aIndex: Integer);
    procedure AppendRecord(aIndex: Integer);
    procedure CustomSort(aIndex: Integer);
    procedure AddToList(aIndex: Integer);
    { 將已傳送的指令取出, 回寫 Send_Nagra 狀態成 P --> 已處理 }
    procedure PhysicalWriteAlreadySend(aIndex: Integer); overload;
    procedure PhysicalWriteAlreadySend; overload;
    { 將 Nagra 已回應指令取出, 回寫 Send_Nagra 狀態成 C --> 已完成或 E ---> 錯誤  }
    procedure PhysicalWriteNagraAck(aIndex: Integer); overload;
    procedure PhysicalWriteNagraAck; overload;
    { 寫 Log }
    procedure PhysicalWriteLog(aIndex: Integer); overload;
    procedure PhysicalWriteLog; overload;
  protected
    function GetWaitWhileFrquence: Cardinal; override;
    function GetDelayFrquence: Cardinal; override;
    procedure Execute; override;
    procedure BeginActive; override;
    procedure EndActive; override;
    procedure Notify; override;
    procedure Update; override;
  public
    constructor Create(const aSoInfo: TThreadList); overload;
    constructor Create(const aSoInfo: PSoInfo); overload;
    destructor Destroy; override;
  end;

implementation

uses cbMain, cbResStr, cbUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSoDbThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }


{ ---------------------------------------------------------------------------- }

{ TSoDbThread }

constructor TSoDbThread.Create(const aSoInfo: TThreadList);
begin
  inherited Create;
  InternalCreate;
  AssignSoInfo( aSoInfo );
end;

{ ---------------------------------------------------------------------------- }

constructor TSoDbThread.Create(const aSoInfo: PSoInfo);
begin
  inherited Create;
  InternalCreate;
  FSoInfo.Add( aSoInfo );
end;

{ ---------------------------------------------------------------------------- }

destructor TSoDbThread.Destroy;
begin
  FSoInfo.Free;
  FADOControlList.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.InternalCreate;
begin
  FSoInfo := TList.Create;
  FADOControlList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.AssignSoInfo(const aSource: TThreadList);
var
  aIndex: Integer;
  aSourceList: TList;
begin
  aSourceList := aSource.LockList;
  try
    FSoInfo.Clear;
    for aIndex := 0 to aSourceList.Count - 1 do
    begin
      if PSoInfo( aSourceList.Items[aIndex] ).Selected then
      begin
        PSoInfo( aSourceList.Items[aIndex] ).CriticalError := 0;
        PSoInfo( aSourceList.Items[aIndex] ).LastCriticalErrorTickCount := 0;
        FSoInfo.Add( aSourceList.Items[aIndex] );
      end;  
    end;
  finally
    aSource.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ReleaseSoInfoResource;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoInfo.Count - 1 do
  begin
    if Assigned( FSoInfo.Items[aIndex] ) then
      FSoInfo.Items[aIndex] := nil;
  end;
  FSoInfo.Clear;
  FCurrentActiveSo := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ReleaseADOResource;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FADOControlList.Count - 1 do
  begin
    if Assigned( FADOControlList.Items[aIndex] ) then
    begin
      PSoADOControl( FADOControlList.Items[aIndex] ).Connection.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).DataReader.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).DataWriter.Free;
      {
      PSoADOControl( FADOControlList.Items[aIndex] ).StoredProc1.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).StoredProc2.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).StoredProc3.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).StoredProc4.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).StoredProc5.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).StoredProc6.Free;
      }
      FADOControlList.Items[aIndex] := nil;
    end;
  end;
  FADOControlList.Clear;
end;

{ ---------------------------------------------------------------------------- }

{$WRITEABLECONST ON}
function TSoDbThread.GetSelectSQL: String;

  { ----------------------------------------------------- }

  function GetIppvWhereCondition(const aType: String): String;
  const
    aConditionO: String = '';
    aConditionN: String = '';
  var
    aBuildConditon: Boolean;  
  begin
    aBuildConditon :=
      ( ( aConditionO = EmptyStr ) and ( aType = 'O' ) or
        ( aConditionN = EmptyStr ) and ( aType = 'N' ) );
    if aBuildConditon then
    begin
      CmdEnv.Filter := 'CMD_TYPE IN ( ''PPV'', ''IPPV'' )';
      CmdEnv.Filtered := True;
      try
        CmdEnv.First;
        while not CmdEnv.Eof do
        begin
          Result := ( Result +
            Format( '''%s''', [CmdEnv.FieldByName( 'HIGH_LEVEL_CMD' ).AsString] ) );
          CmdEnv.Next;
          if ( not CmdEnv.Eof ) and ( Result <> EmptyStr ) then
            Result := Result + ',';
        end;
        if aType = 'O' then
          aConditionO := Format( ' IN ( %s ) '#13, [Result] )
        else
          aConditionN := Format( ' NOT IN ( %s ) '#13, [Result] );
      finally
        CmdEnv.Filtered := False;
        CmdEnv.Filter := EmptyStr;
      end;
    end;  
    if aType = 'O' then
      Result := aConditionO
    else
      Result := aConditionN;
  end;
  
  { ----------------------------------------------------- }


var
  aSQL: String;
begin
  aSQL := Format(
    ' SELECT A.HIGH_LEVEL_CMD_ID,                                              '#13 +
    '        A.ICC_NO,                                                         '#13 +
    '        A.STB_NO,                                                         '#13 +
    '        TO_CHAR( A.SUBSCRIPTION_BEGIN_DATE, ''YYYYMMDD'' )                '#13 +
    '            AS SUBSCRIPTION_BEGIN_DATE,                                   '#13 +
    '        TO_CHAR( A.SUBSCRIPTION_END_DATE, ''YYYYMMDD'' )                  '#13 +
    '            AS SUBSCRIPTION_END_DATE,                                     '#13 +
    '        A.CMD_STATUS,                                                     '#13 +
    '        A.OPERATOR,                                                       '#13 +
    '        A.IMS_PRODUCT_ID,                                                 '#13 +
    '        A.ZIP_CODE,                                                       '#13 +
    '        A.CREDIT_MODE,                                                    '#13 +
    '        A.CREDIT,                                                         '#13 +
    '        A.CREDIT_LIMIT,                                                   '#13 +
    '        A.THRESHOLD_CREDIT,                                               '#13 +
    '        A.EVENT_NAME,                                                     '#13 +
    '        A.PRICE,                                                          '#13 +
    '        A.CC_NUMBER_1,                                                    '#13 +
    '        A.IP_ADDR,                                                        '#13 +
    '        A.CC_PORT,                                                        '#13 +
    '        A.CALLBACK_DATE,                                                  '#13 +
    '        A.CALLBACK_TIME,                                                  '#13 +
    '        A.CALLBACK_FREQUENCY,                                             '#13 +
    '        A.FIRST_CALLBACK_DATE,                                            '#13 +
    '        A.PHONE_NUM_1,                                                    '#13 +
    '        A.PHONE_NUM_2,                                                    '#13 +
    '        A.PHONE_NUM_3,                                                    '#13 +
    '        A.SEQNO,                                                          '#13 +
    '        A.MIS_IRD_CMD_ID,                                                 '#13 +
    '        A.MIS_IRD_CMD_DATA,                                               '#13 +
    '        A.COMPCODE,                                                       '#13 +
    '        A.NOTES,                                                          '#13 +
    '        TO_CHAR( A.CLEANUP_DATE, ''YYYYMMDD'' )                           '#13 +
    '            AS CLEANUP_DATE,                                              '#13 +
    '        TO_CHAR( A.CONDITION_DATE, ''YYYYMMDD'' )                         '#13 +
    '            AS CONDITION_DATE,                                            '#13 +
    '        TO_CHAR( A.COLLECT_DATE, ''YYYYMMDD'' )                           '#13 +
    '            AS COLLECT_DATE,                                              '#13 +
    '        NVL( A.RESENTTIMES, 0 ) AS RESENTTIMES                            '#13 +
    '   FROM SEND_NAGRA A                                                      '#13 +
    '   WHERE ( ( A.CMD_STATUS = ''W'' ) OR                                    '#13 +
    '           ( A.CMD_STATUS = ''P'' AND NVL( A.RESENTTIMES, 0 ) <= %d ) )   '#13 +
    '     AND ( A.PROCESSINGDATE < SYSDATE OR A.PROCESSINGDATE IS NULL )       '#13 +
    '     AND A.COMPCODE = ''%s''                                              ',
    [CommEnv.DbResendCount, FCurrentActiveSo.CompCode] );

  { 只處理 IPPV資料, 不處理 IPPV 資料 }
  if ( CommEnv.ProcessIPPV = 'O' ) or ( CommEnv.ProcessIPPV = 'N' ) then
  begin
    aSQL := ( aSQL + GetIppvWhereCondition( CommEnv.ProcessIPPV ) );
  end;
  { 只處理批次資料 }
  if ( CommEnv.ProcessBatch = 'O' ) then
  begin
    aSQL := ( aSQL +
      Format( ' AND NVL( A.OPERATOR, ''X'' ) = ''%s'' '#13, [CommEnv.BatchOperator] ) );
  end else
  { 不處理批次資料 }
  if ( CommEnv.ProcessBatch = 'N' ) then
  begin
    aSQL := ( aSQL +
      Format( ' AND NVL( A.OPERATOR, ''X'' ) <> ''%s'' '#13, [CommEnv.BatchOperator] ) );
  end;
  aSQL := ( aSQL + ' ORDER BY A.CMD_STATUS DESC, NVL( A.RESENTTIMES,0 ), A.SEQNO '#13 );
  Result := Format( ' SELECT * FROM ( %s ) WHERE ROWNUM <= %d ',
    [aSQL, CommEnv.ProcessRecordCount] );
end;
{$WRITEABLECONST OFF}

{ ---------------------------------------------------------------------------- }

{$WRITEABLECONST ON}
function TSoDbThread.GetDelayFrquence: Cardinal;
const
  aLastCheckTickCount: Cardinal = 0;
  aStoredResult: Cardinal = 0;
var
  aTimeStart, aTimeEnd: TDateTime;
  aIsSetBusyTime, aIsBusyTime, aCanCheck: Boolean;
begin
  Result := aStoredResult;
  { 不須每次都檢查時段, 每 1 分鐘檢查一次即可 }
  aCanCheck := ( ( GetTickCount - aLastCheckTickCount ) <= 0 ) or
    ( aLastCheckTickCount = 0 ) or ( ( GetTickCount - aLastCheckTickCount ) > 60000 );
  if aCanCheck then
  begin
    aIsBusyTime := False;
    aIsSetBusyTime :=
      ( CommEnv.BusyTimeStart <> EmptyStr ) and ( CommEnv.BusyTimeEnd <> EmptyStr );
    if ( aIsSetBusyTime ) then
    begin
      aTimeStart := StrToTime( CommEnv.BusyTimeStart );
      aTimeEnd := StrToTime( CommEnv.BusyTimeEnd );
      aIsBusyTime := ( GetTime >= aTimeStart ) and ( GetTime <= aTimeEnd );
    end;
    if ( aIsSetBusyTime ) and ( aIsBusyTime ) then
      Result := Trunc( CommEnv.BusyTimeDbReadFrequence * 1000 )
    else
      Result := Trunc( CommEnv.NormalTimeDbReadFrequence * 1000 );
    aLastCheckTickCount := GetTickCount;
  end;
  aStoredResult := Result;
end;
{$WRITEABLECONST OFF}

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetDbRertyFrquence: Cardinal;
var
  aTestResult: Integer;
begin
  aTestResult := ( CommEnv.DbRetryFrequence * 1000 );
  if FCurrentActiveSo.LastCriticalErrorTickCount > 0 then
  begin
    aTestResult := ( GetTickCount - FCurrentActiveSo.LastCriticalErrorTickCount );
    if aTestResult > 0 then
      aTestResult := ( GetTickCount - FCurrentActiveSo.LastCriticalErrorTickCount );
  end;
  Result := aTestResult;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetWaitWhileFrquence: Cardinal;
begin
  Result := inherited GetWaitWhileFrquence;
  if CommEnv.DbMultiThread then Result := ( Result * 5 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanDbRertyConnect: Boolean;
var
  aRetry: Integer;
begin
  aRetry := GetDbRertyFrquence;
  Result := aRetry >= ( CommEnv.DbRetryFrequence * 1000 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanAppsendRecord(aIndex: Integer; aSeqno: String): Boolean;
begin
  Result := ( aSeqno <> EmptyStr );
  if Result then
  begin
     ControlSendManager.Items[aIndex].BeginRead;
     try
       Result := ( ControlSendManager.Items[aIndex].IndexOf( aSeqno ) < 0 );
    finally
      ControlSendManager.Items[aIndex].EndRead;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckDbStatus(aIndex: Integer): Boolean;
var
  aADO: PSoADOControl;
begin
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  if ( not aADO.Connection.Connected ) and CanDbRertyConnect then
  begin
    FCurrentActiveSo.NotifyMsg := Format( SDbRertyConnect, [FCurrentActiveSo.CompName] );
    Synchronize( Notify );
    ConnectToDb( aIndex );
  end;
  Result := aADO.Connection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PrepareADOControl;
var
  aIndex: Integer;
  aADO: PSoADOControl;
begin
  for aIndex := 0 to FSoInfo.Count - 1 do
  begin
    if ( PSoInfo( FSoInfo.Items[aIndex] ).Selected ) then
    begin
      New( aADO );
      aADO.CompCode := PSoInfo( FSoInfo.Items[aIndex] ).CompCode;
      aADO.Connection := TADOConnection.Create( nil );
      aADO.DataReader := TADOQuery.Create( nil );
      aADO.DataWriter := TADOQuery.Create( nil );
      aADO.DataSet := TClientDataSet.Create( nil );
      {
      aADO.StoredProc1 := TADOStoredProc.Create( nil );
      aADO.StoredProc2 := TADOStoredProc.Create( nil );
      aADO.StoredProc3 := TADOStoredProc.Create( nil );
      aADO.StoredProc4 := TADOStoredProc.Create( nil );
      aADO.StoredProc5 := TADOStoredProc.Create( nil );
      aADO.StoredProc6 := TADOStoredProc.Create( nil );
      }
      aADO.Connection.LoginPrompt := False;
      aADO.DataReader.CacheSize := 1000;
      aADO.DataReader.Connection := aADO.Connection;
      aADO.DataReader.LockType := ltBatchOptimistic;
      aADO.DataWriter.CacheSize := 1000;
      aADO.DataWriter.Connection := aADO.Connection;
      {
      aADO.StoredProc1.Connection := aADO.Connection;
      aADO.StoredProc2.Connection := aADO.Connection;
      aADO.StoredProc3.Connection := aADO.Connection;
      aADO.StoredProc4.Connection := aADO.Connection;
      aADO.StoredProc5.Connection := aADO.Connection;
      aADO.StoredProc6.Connection := aADO.Connection;
      }
      aADo.DataSet.FieldDefs.Clear;
      aADo.DataSet.FieldDefs.Add( 'SENDSTATUS', ftString, 2 );
      aADo.DataSet.FieldDefs.Add( 'TRANSACTIONNUMBER', ftString, 9 );
      aADo.DataSet.FieldDefs.Add( 'CMD_TYPE_PRIORITY', ftInteger  );
      aADo.DataSet.FieldDefs.Add( 'CMD_PRIORITY', ftInteger );
      aADo.DataSet.FieldDefs.Add( 'HIGH_LEVEL_CMD_ID', ftString, 4 );
      aADo.DataSet.FieldDefs.Add( 'ICC_NO', ftString, 4 );
      aADo.DataSet.FieldDefs.Add( 'STB_NO', ftString, 4 );
      aADo.DataSet.FieldDefs.Add( 'SUBSCRIPTION_BEGIN_DATE', ftDateTime );
      aADo.DataSet.FieldDefs.Add( 'SUBSCRIPTION_END_DATE', ftDateTime );
      aADo.DataSet.FieldDefs.Add( 'NOTES', ftString, 2000 );
      aADo.DataSet.FieldDefs.Add( 'CMD_STATUS', ftString, 1 );
      aADo.DataSet.FieldDefs.Add( 'OPERATOR', ftString, 20 );
      aADo.DataSet.FieldDefs.Add( 'IMS_PRODUCT_ID', ftString, 12 );
      aADo.DataSet.FieldDefs.Add( 'ZIP_CODE', ftString, 5 );
      aADo.DataSet.FieldDefs.Add( 'CREDIT_MODE', ftFloat );
      aADo.DataSet.FieldDefs.Add( 'CREDIT', ftFloat );
      aADo.DataSet.FieldDefs.Add( 'CREDIT_LIMIT', ftFloat );
      aADo.DataSet.FieldDefs.Add( 'THRESHOLD_CREDIT', ftFloat );
      aADo.DataSet.FieldDefs.Add( 'EVENT_NAME', ftString, 32 );
      aADo.DataSet.FieldDefs.Add( 'PRICE', ftFloat );
      aADo.DataSet.FieldDefs.Add( 'CC_NUMBER_1', ftString, 16  );
      aADo.DataSet.FieldDefs.Add( 'IP_ADDR', ftString, 15 );
      aADo.DataSet.FieldDefs.Add( 'CC_PORT', ftInteger );
      aADo.DataSet.FieldDefs.Add( 'CALLBACK_DATE', ftString, 8 );
      aADo.DataSet.FieldDefs.Add( 'CALLBACK_TIME', ftString, 6 );
      aADo.DataSet.FieldDefs.Add( 'CALLBACK_FREQUENCY', ftString, 2 );
      aADo.DataSet.FieldDefs.Add( 'FIRST_CALLBACK_DATE',ftDateTime );
      aADo.DataSet.FieldDefs.Add( 'PHONE_NUM_1', ftString, 16 );
      aADo.DataSet.FieldDefs.Add( 'PHONE_NUM_2', ftString, 16  );
      aADo.DataSet.FieldDefs.Add( 'PHONE_NUM_3', ftString, 16  );
      aADo.DataSet.FieldDefs.Add( 'MIS_IRD_CMD_ID', ftInteger );
      aADo.DataSet.FieldDefs.Add( 'MIS_IRD_CMD_DATA', ftString, 100 );
      aADo.DataSet.FieldDefs.Add( 'SEQNO', ftInteger );
      aADo.DataSet.FieldDefs.Add( 'COMPCODE', ftInteger );
      aADo.DataSet.FieldDefs.Add( 'CLEANUP_DATE', ftDateTime );
      aADo.DataSet.FieldDefs.Add( 'CONDITION_DATE', ftDateTime );
      aADo.DataSet.FieldDefs.Add( 'COLLECT_DATE', ftDateTime );
      aADo.DataSet.FieldDefs.Add( 'RESENTTIMES', ftInteger );
      aADO.DataSet.CreateDataSet;
      FADOControlList.Add( aADO );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ConnectToDb;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FADOControlList.Count - 1 do
  begin
    FCurrentActiveSo := PSoInfo( FSoInfo.Items[aIndex] );
    ConnectToDb( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ConnectToDb(const aIndex: Integer);
var
  aSoADO: PSoADOControl;
begin
  aSoADO := PSoADOControl( FADOControlList.Items[aIndex] );
  if aSoADO.Connection.Connected then
    aSoADO.Connection.Connected := False;
  aSoADO.Connection.ConnectionString := Format(
    'Provider=MSDAORA.1;Persist Security Info=True;' +
    'Password=%s;User ID=%s;Data Source=%s;',
   [FCurrentActiveSo.LoginPass, FCurrentActiveSo.LoginUser, FCurrentActiveSo.DbAliase] );
  try
    aSoADO.Connection.Connected := True;
    FCurrentActiveSo.DbConnectStatus := dbOK;
    FCurrentActiveSo.NotifyMsg := Format( SDbConnectSuccess, [FCurrentActiveSo.CompName] );
  except
    on E: Exception do
    begin
      FCurrentActiveSo.DbConnectStatus := dbError;
      FCurrentActiveSo.NotifyMsg := Format( SDbConnectError, [FCurrentActiveSo.CompName, E.Message] );
    end;
  end;
  Synchronize( Notify );
  Synchronize( Update );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.DisconnectFromDb;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FADOControlList.Count - 1 do
  begin
    FCurrentActiveSo := PSoInfo( FSoInfo.Items[aIndex] );
    DisconnectFromDb( aIndex );
  end;  
end;

{ ---------------------------------------------------------------------------- }


procedure TSoDbThread.DisconnectFromDb(const aIndex: Integer);
begin
  try
    if PSoADOControl( FADOControlList.Items[aIndex] ).Connection.Connected then
      PSoADOControl( FADOControlList.Items[aIndex] ).Connection.Connected := False;
    FCurrentActiveSo.DbConnectStatus := dbNone;  
    FCurrentActiveSo.RecordCount := 0;
    FCurrentActiveSo.NotifyMsg :=
      Format( SDbDisConnectSuccess, [FCurrentActiveSo.CompName] );
  except
    on E: Exception do
    begin
      FCurrentActiveSo.NotifyMsg := Format( SDbDisConnectError,
        [FCurrentActiveSo.CompName, E.Message] );
    end;    
  end;
  Synchronize( Notify );
  Synchronize( Update );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Execute;
var
  aIndex: Integer;
begin
  PrepareADOControl;
  ConnectToDb;
  { 等待 Main Thread 打出開始執行的訊號 }
  WaitForPlaySignal;
  Sleep( GetWaitWhileFrquence );
  while ( not Stop ) do
  begin
    MessageSubject.RunState := rsRunning;
    try
      for aIndex := 0 to FSoInfo.Count - 1 do
      begin
        if Stop then Break;
        FCurrentActiveSo := PSoInfo( FSoInfo.Items[aIndex] );
        FCurrentActiveADO := PSoADOControl( FADOControlList.Items[aIndex] );
        BeginActive;
        try
          if Stop then Break;
          PhysicalWriteAlreadySend( aIndex );
          if Stop then Break;
          PhysicalWriteNagraAck( aIndex );
          if Stop then Break;
          GetData( aIndex );
          if Stop then Break;
          PhysicalWriteLog( aIndex );
        finally
          EndActive;
        end;
        if Stop then Break;
        Sleep( GetWaitWhileFrquence );
      end;
      if Stop then Break;
      { 每多久讀一次資料庫 }
      Sleep( GetDelayFrquence );
    except
      { ... }
    end;
  end;
  Sleep( GetWaitWhileFrquence );
  PhysicalWriteAlreadySend;
  PhysicalWriteNagraAck;
  PhysicalWriteLog;
  MessageSubject.RunState := rsStop;
  { 已先收到停止的訊號 }
  WaitForTerminalSignal;
  DisconnectFromDb;
  ReleaseADOResource;
  ReleaseSoInfoResource;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Notify;
begin
  MessageSubject.NotifyMessage := FCurrentActiveSo.NotifyMsg;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Update;
begin
  fmMain.UpdateSoTreeStatus( FCurrentActiveSo.ItemIndex );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.BeginActive;
begin
  if FCurrentActiveSo.DbConnectStatus in [dbOK, dbActive] then
    FCurrentActiveSo.DbConnectStatus := dbActive;
  Synchronize( Update );
  Sleep( GetWaitWhileFrquence );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.EndActive;
begin
  if FCurrentActiveSo.DbConnectStatus in [dbOK, dbActive] then
    FCurrentActiveSo.DbConnectStatus := dbOK;
  Synchronize( Update );
  Sleep( GetWaitWhileFrquence );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetData(aIndex: Integer);
var
  aADO: PSoADOControl;
begin
  if not CheckDbStatus( aIndex ) then Exit;
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  aADO.DataReader.Close;
  aADO.DataReader.SQL.Text := GetSelectSQL;
  try
    aADO.DataReader.Open;
    FCurrentActiveSo.RecordCount := aADO.DataReader.RecordCount;
    if aADO.DataReader.RecordCount > 0 then
    begin
      AppendRecord( aIndex );
      CustomSort( aIndex );
      AddToList( aIndex );
    end;
    aADO.DataReader.Close;
  except
    on E: Exception do
    begin
      FCurrentActiveSo.CriticalError := ( FCurrentActiveSo.CriticalError + 1 );
      FCurrentActiveSo.NotifyMsg :=
        Format( SDbGetDataError, [FCurrentActiveSo.CompName, E.Message] );
      FCurrentActiveSo.DbConnectStatus := dbError;
      Synchronize( Notify );
      Synchronize( Update );
    end;
  end;
  if FCurrentActiveSo.CriticalError >= 3 then
  begin
    FCurrentActiveSo.CriticalError := 0;
    FCurrentActiveSo.LastCriticalErrorTickCount := GetTickCount;
    FCurrentActiveSo.NotifyMsg := Format( SDbHasProblem,
      [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] );
    Synchronize( Notify );
    DisconnectFromDb( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.AppendRecord(aIndex: Integer);
var
  aADO: PSoADOControl;
begin
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  if aADO.DataSet.IndexName <> EmptyStr then
    aADO.DataSet.DeleteIndex( 'CUSTOMSORT' );
  aADO.DataSet.EmptyDataSet;
  aADO.DataReader.First;
  while not aADO.DataReader.Eof do
  begin
    { 在 List 中, 尚未送出去 }
    if not CanAppsendRecord( aIndex,
      aADO.DataReader.FieldByName( 'SEQNO' ).AsString ) then
    begin
      aADO.DataReader.Next;
      Continue;
    end;
    aADO.DataSet.AppendRecord( [
      'W',
      EmptyStr,
      0,
      0,
      aADO.DataReader.FieldByName( 'HIGH_LEVEL_CMD_ID' ).Value,
      aADO.DataReader.FieldByName( 'ICC_NO' ).Value,
      aADO.DataReader.FieldByName( 'STB_NO' ).Value,
      aADO.DataReader.FieldByName( 'SUBSCRIPTION_BEGIN_DATE' ).Value,
      aADO.DataReader.FieldByName( 'SUBSCRIPTION_END_DATE' ).Value,
      aADO.DataReader.FieldByName( 'NOTES' ).Value,
      aADO.DataReader.FieldByName( 'CMD_STATUS' ).Value,
      aADO.DataReader.FieldByName( 'OPERATOR' ).Value,
      aADO.DataReader.FieldByName( 'IMS_PRODUCT_ID' ).Value,
      aADO.DataReader.FieldByName( 'ZIP_CODE' ).Value,
      aADO.DataReader.FieldByName( 'CREDIT_MODE' ).Value,
      aADO.DataReader.FieldByName( 'CREDIT' ).Value,
      aADO.DataReader.FieldByName( 'CREDIT_LIMIT' ).Value,
      aADO.DataReader.FieldByName( 'THRESHOLD_CREDIT' ).Value,
      aADO.DataReader.FieldByName( 'EVENT_NAME' ).Value,
      aADO.DataReader.FieldByName( 'PRICE' ).Value,
      aADO.DataReader.FieldByName( 'CC_NUMBER_1' ).Value,
      aADO.DataReader.FieldByName( 'IP_ADDR' ).Value,
      aADO.DataReader.FieldByName( 'CC_PORT' ).Value,
      aADO.DataReader.FieldByName( 'CALLBACK_DATE' ).Value,
      aADO.DataReader.FieldByName( 'CALLBACK_TIME' ).Value,
      aADO.DataReader.FieldByName( 'CALLBACK_FREQUENCY' ).Value,
      aADO.DataReader.FieldByName( 'FIRST_CALLBACK_DATE' ).Value,
      aADO.DataReader.FieldByName( 'PHONE_NUM_1' ).Value,
      aADO.DataReader.FieldByName( 'PHONE_NUM_2' ).Value,
      aADO.DataReader.FieldByName( 'PHONE_NUM_3' ).Value,
      aADO.DataReader.FieldByName( 'MIS_IRD_CMD_ID' ).Value,
      aADO.DataReader.FieldByName( 'MIS_IRD_CMD_DATA' ).Value,
      aADO.DataReader.FieldByName( 'SEQNO' ).Value,
      aADO.DataReader.FieldByName( 'COMPCODE' ).Value,
      aADO.DataReader.FieldByName( 'CLEANUP_DATE' ).Value,
      aADO.DataReader.FieldByName( 'CONDITION_DATE' ).Value,
      aADO.DataReader.FieldByName( 'COLLECT_DATE' ).Value,
      aADO.DataReader.FieldByName( 'RESENTTIMES' ).Value] );
    aADO.DataReader.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.CustomSort(aIndex: Integer);
var
  aADO: PSoADOControl;
  aCmdTypePriority, aCmdPriority: Integer; 
begin
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  aADO.DataSet.First;
  while not aADO.DataSet.Eof do
  begin
    { 是否是批次的指令 }
    if ( aADO.DataSet.FieldByName( 'OPERATOR' ).AsString = CommEnv.BatchOperator ) and
       ( CommEnv.BatchOperator <> EmptyStr ) then
    begin
      aCmdTypePriority := 999;
      aCmdPriority := 999;
    end else
    if CmdEnv.Locate( 'HIGH_LEVEL_CMD',
      VarArrayOf( [aADO.DataSet.FieldByName( 'HIGH_LEVEL_CMD_ID' ).AsString] ), [] ) then
    begin
      aCmdTypePriority := CmdEnv.FieldByName( 'CMD_TYPE_PRIORITY' ).AsInteger;
      aCmdPriority := CmdEnv.FieldByName( 'CMD_PRIORITY' ).AsInteger;
    end else
    begin
      { 指令優先順序沒對應到, 一律當成批次的指令, 優先順序最低 }
      aCmdTypePriority := 999;
      aCmdPriority := 999;
    end;
    aADO.DataSet.Edit;
    aADO.DataSet.FieldByName( 'CMD_TYPE_PRIORITY' ).AsInteger := aCmdTypePriority;
    aADO.DataSet.FieldByName( 'CMD_PRIORITY' ).AsInteger := aCmdPriority;
    aADO.DataSet.Post;
    aADO.DataSet.Next;
  end;
  aADO.DataSet.AddIndex( 'CUSTOMSORT',
    'CMD_STATUS;RESENTTIMES;CMD_TYPE_PRIORITY;CMD_PRIORITY;SEQNO', [], 'CMD_STATUS' );
  aADO.DataSet.IndexName := 'CUSTOMSORT';
  {$IFDEF DEBUG }
  aADO.DataSet.First;
  while not aADO.DataSet.Eof do
  begin
    CodeSite.SendFmtMsg( 'CMD=%s, STATUS=%s, RESENTTIMES=%d, SEQNO=%s',
      [aADO.DataSet.FieldByName( 'HIGH_LEVEL_CMD_ID' ).AsString,
       aADO.DataSet.FieldByName( 'CMD_STATUS' ).AsString,
       aADO.DataSet.FieldByName( 'RESENTTIMES' ).AsInteger,
       aADO.DataSet.FieldByName( 'SEQNO' ).AsString] );
    aADO.DataSet.Next;
  end;
  {$ENDIF}
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.AddToList(aIndex: Integer);
var
  aADO: PSoADOControl;
  aObj: PSendNagra;
  aWriteCount: Integer;
begin
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  aWriteCount := 0;
  ControlSendManager.Items[aIndex].BeginWrite;
  try
    aADO.DataSet.First;
    while not aADO.DataSet.Eof do
    begin
      New( aObj );
      aObj.TransactionNumber := EmptyStr;
      aObj.HighLevelCmd := aADO.DataSet.FieldByName( 'HIGH_LEVEL_CMD_ID' ).AsString;
      aObj.IccNo := aADO.DataSet.FieldByName( 'ICC_NO' ).AsString;
      aObj.StbNo := aADO.DataSet.FieldByName( 'STB_NO' ).AsString;
      aObj.SubBeginDate := EmptyStr;
      if not VarIsNull( aADO.DataSet.FieldByName( 'SUBSCRIPTION_BEGIN_DATE' ).Value ) then
        aObj.SubBeginDate := FormatDateTime( 'yyyymmdd',
          aADO.DataSet.FieldByName( 'SUBSCRIPTION_BEGIN_DATE' ).AsDateTime );
      aObj.SubEndDate := EmptyStr;
      if not VarIsNull( aADO.DataSet.FieldByName( 'SUBSCRIPTION_END_DATE' ).Value ) then
        aObj.SubEndDate := FormatDateTime( 'yyyymmdd',
          aADO.DataSet.FieldByName( 'SUBSCRIPTION_END_DATE' ).AsDateTime );
      aObj.SubEndDate := EmptyStr;
      aObj.Notes := aADO.DataSet.FieldByName( 'STB_NO' ).AsString;
      aObj.CmdStatus := aADO.DataSet.FieldByName( 'CMD_STATUS' ).AsString;
      aObj.Operator := aADO.DataSet.FieldByName( 'OPERATOR' ).AsString;
      aObj.UpdTime := Now;
      aObj.ErrCode := EmptyStr;
      aObj.ErrMsg := EmptyStr;
      aObj.ImdProductId := aADO.DataSet.FieldByName( 'IMS_PRODUCT_ID' ).AsString;
      aObj.ZipCode := aADO.DataSet.FieldByName( 'ZIP_CODE' ).AsString;
      aObj.CreditMode := aADO.DataSet.FieldByName( 'CREDIT_MODE' ).AsString;
      aObj.Credit := aADO.DataSet.FieldByName( 'CREDIT' ).AsString;
      aObj.CreditLimit := aADO.DataSet.FieldByName( 'CREDIT_LIMIT' ).AsString;
      aObj.ThreholdCredit := aADO.DataSet.FieldByName( 'THRESHOLD_CREDIT' ).AsString;
      aObj.EventName := aADO.DataSet.FieldByName( 'EVENT_NAME' ).AsString;
      aObj.Price := aADO.DataSet.FieldByName( 'PRICE' ).AsString;
      aObj.CCNumber1 := aADO.DataSet.FieldByName( 'CC_NUMBER_1' ).AsString;
      aObj.IPAddr := aADO.DataSet.FieldByName( 'IP_ADDR' ).AsString;
      aObj.CCPort := aADO.DataSet.FieldByName( 'CC_PORT' ).AsString;
      aObj.CallbackDate := aADO.DataSet.FieldByName( 'CALLBACK_DATE' ).AsString;
      aObj.CallbackTime := aADO.DataSet.FieldByName( 'CALLBACK_TIME' ).AsString;
      aObj.CallbackFrequency := aADO.DataSet.FieldByName( 'CALLBACK_FREQUENCY' ).AsString;
      aObj.FirstCallbackDate := aADO.DataSet.FieldByName( 'FIRST_CALLBACK_DATE' ).AsString;
      aObj.PhoneNum1 := aADO.DataSet.FieldByName( 'PHONE_NUM_1' ).AsString;
      aObj.PhoneNum2 := aADO.DataSet.FieldByName( 'PHONE_NUM_2' ).AsString;
      aObj.PhoneNum3 := aADO.DataSet.FieldByName( 'PHONE_NUM_3' ).AsString;
      aObj.MisIrdCmd := aADO.DataSet.FieldByName( 'MIS_IRD_CMD_ID' ).AsString;
      aObj.MisIrdData := aADO.DataSet.FieldByName( 'MIS_IRD_CMD_DATA' ).AsString;
      aObj.SeqNo := aADO.DataSet.FieldByName( 'SEQNO' ).AsString;
      aObj.CompCode := aADO.DataSet.FieldByName( 'COMPCODE' ).AsString;
      aObj.ResentTimes := aADO.DataSet.FieldByName( 'RESENTTIMES' ).AsInteger;
      aObj.CleanupDate := aADO.DataSet.FieldByName( 'CLEANUP_DATE' ).AsString;
      aObj.ConditionDate := aADO.DataSet.FieldByName( 'CONDITION_DATE' ).AsString;
      aObj.CollectDate := aADO.DataSet.FieldByName( 'COLLECT_DATE' ).AsString;
      ControlSendManager.Items[aIndex].AddObject(
        aADO.DataSet.FieldByName( 'SEQNO' ).AsString, TObject( aObj ) );
      aADO.DataSet.Next;
      Inc( aWriteCount );
    end;
    if aWriteCount > 0 then
    begin
      FCurrentActiveSo.NotifyMsg := Format( SDbAddToControlSendList,
        [FCurrentActiveSo.CompName, aWriteCount] );
      Synchronize( Notify );
    end;  
  finally
    ControlSendManager.Items[aIndex].EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteAlreadySend(aIndex: Integer);
var
  aIndex2, aWriteDone: Integer;
  aObj: PSendNagra;
  aManager: TCommandListMagager;
begin
  if not CheckDbStatus( aIndex ) then Exit;
  aManager := ControlSendDoneManager;
  aManager.Items[aIndex].BeginWrite;
  try
    aWriteDone := 0;
    for aIndex2 := 0 to aManager.Items[aIndex].Count - 1 do
    begin
      aObj := PSendNagra( aManager.Items[aIndex].Objects[aIndex2] );
      { 只有標示為 P 才可以回填為處理中, 因為要等到 Nagra Ack 回來後,
        會將此筆高階指令從 List 中移除 }
      { 為什不用 aObj.CmdStatus ? 因為有可能有指令重送的須要 }  
      if aObj.SendStatus = 'P' then
      begin
        FCurrentActiveADO.DataWriter.SQL.Text := Format(
          ' UPDATE SEND_NAGRA                                 '#13 +
          '    SET CMD_STATUS = ''P'',                        '#13 +
          '        RESENTTIMES = ( NVL( RESENTTIMES, 0 ) + 1  '#13 +
          '  WHERE SEQNO = ''%s''                             '#13 +
          '    AND COMPCODE = ''%s''                          '#13 +
          '    AND CMD_STATUS IN ( ''W'', ''P'' )             ',
          [aObj.SeqNo,aObj.CompCode] );
        try
          FCurrentActiveADO.DataWriter.ExecSQL;
          aObj.SendStatus := 'P';
          Inc( aWriteDone );
        except
          on E: Exception do
          begin
            FCurrentActiveSo.CriticalError := ( FCurrentActiveSo.CriticalError + 1 );
            FCurrentActiveSo.NotifyMsg :=
            Format( SDbWriteAlreadySendError,
              [FCurrentActiveSo.CompName, aObj.HighLevelCmd, aObj.SeqNo, E.Message] );
            FCurrentActiveSo.DbConnectStatus := dbError;
            Synchronize( Notify );
            Synchronize( Update );
          end;
        end;
      end;
      if FCurrentActiveSo.CriticalError >= 3 then
      begin
        FCurrentActiveSo.CriticalError := 0;
        FCurrentActiveSo.LastCriticalErrorTickCount := GetTickCount;
        FCurrentActiveSo.NotifyMsg := Format( SDbHasProblem,
          [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] );
        Synchronize( Notify );
        DisconnectFromDb( aIndex );
        Break;
      end;
    end;
    FCurrentActiveSo.NotifyMsg := Format( SDbWriteAlreadySendCount,
      [FCurrentActiveSo.CompName, aWriteDone] );
  finally
    aManager.Items[aIndex].EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteAlreadySend;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoInfo.Count - 1 do
  begin
    FCurrentActiveSo := PSoInfo( FSoInfo[aIndex] );
    FCurrentActiveADO := PSoADOControl( FADOControlList[aIndex] );
    PhysicalWriteAlreadySend( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteNagraAck(aIndex: Integer);
var
  aIndex2, aWriteDone: Integer;
  aObj: PSendNagra;
  aManager: TCommandListMagager;
begin
  if not CheckDbStatus( aIndex ) then Exit;
  aManager := ControlRecvManager;
  aManager.Items[aIndex].BeginWrite;
  try
    aWriteDone := 0;
    for aIndex2 := aManager.Items[aIndex].Count - 1 downto 0 do
    begin
      aObj := PSendNagra( aManager.Items[aIndex].Objects[aIndex2] );
      { 更新狀態 }
      FCurrentActiveADO.DataWriter.SQL.Text := Format(
        ' UPDATE SEND_NAGRA                                 '#13 +
        '    SET CMD_STATUS = ''%s'',                       '#13 +
        '        ERR_CODE = ''%s'',                         '#13 +
        '        ERR_MSG = ''%s''                           '#13 +
        '  WHERE SEQNO = ''%s''                             '#13 +
        '    AND COMPCODE = ''%s''                          '#13 +
        '    AND CMD_STATUS IN ( ''W'', ''P'', ''E'' )      ',
        [aObj.CmdStatus, aObj.ErrCode, aObj.ErrMsg, aObj.SeqNo, aObj.CompCode] );
      try
        FCurrentActiveADO.DataWriter.ExecSQL;
        { 等 Nagara Ack 回來後, 就可釋放 }
        Dispose( aObj );
        aManager.Items[aIndex].Delete( aIndex2 );
        Inc( aWriteDone );
      except
        on E: Exception do
        begin
          FCurrentActiveSo.CriticalError := ( FCurrentActiveSo.CriticalError + 1 );
          FCurrentActiveSo.NotifyMsg :=
          Format( SDbWriteAckError,
            [FCurrentActiveSo.CompName, aObj.HighLevelCmd, aObj.SeqNo, E.Message] );
          FCurrentActiveSo.DbConnectStatus := dbError;
          Synchronize( Notify );
          Synchronize( Update );
        end;
      end;
      if FCurrentActiveSo.CriticalError >= 3 then
      begin
        FCurrentActiveSo.CriticalError := 0;
        FCurrentActiveSo.LastCriticalErrorTickCount := GetTickCount;
        FCurrentActiveSo.NotifyMsg := Format( SDbHasProblem,
          [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] );
        Synchronize( Notify );
        DisconnectFromDb( aIndex );
        Break;
      end;
    end;
    FCurrentActiveSo.NotifyMsg := Format( SDbWriteAckCount,
      [FCurrentActiveSo.CompName, aWriteDone] );
  finally
    aManager.Items[aIndex].EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteNagraAck;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoInfo.Count - 1 do
  begin
    FCurrentActiveSo := PSoInfo( FSoInfo[aIndex] );
    FCurrentActiveADO := PSoADOControl( FADOControlList[aIndex] );
    PhysicalWriteNagraAck( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteLog(aIndex: Integer);
var
  aIndex2, aWriteDone: Integer;
  aObj: PSendNagra;
  aManager: TCommandListMagager;
begin
  if not CheckDbStatus( aIndex ) then Exit;
  aManager := ControlSendLogManager;
  aManager.Items[aIndex].BeginWrite;
  try
    aWriteDone := 0;
    for aIndex2 := aManager.Items[aIndex].Count - 1 downto 0 do
    begin
      aObj := PSendNagra( aManager.Items[aIndex].Objects[aIndex2] );
      try
        { 傳送指令的 Log }
        FCurrentActiveADO.DataWriter.SQL.Text := Format(
          ' INSERT INTO CASENDCMDLOG (                             '#13 +
          '    COMPCODE, HIGH_LEVEL_CMD_ID, LOW_LEVEL_CMD,         '#13 +
          '    TRANSACTIONNUM, ICC_NO, STB_NO, FULLCMDTEXT,        '#13 +
          '    OPERATOR, UPDTIME  )                                '#13 +
          ' VALUES ( ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',       '#13 +
          '          ''%s'', ''%s'', SYSDATE    )                  ',
          [aObj.CompCode,
           aObj.HighLevelCmd,
           aObj.LowLevelCmd,
           aObj.TransactionNumber,
           aObj.IccNo,
           aObj.StbNo,
           aObj.FullCommandText,
           aObj.Operator] );
        FCurrentActiveADO.DataWriter.ExecSQL;
        { Ack 回來的 Log }
        FCurrentActiveADO.DataWriter.SQL.Text := Format(
          ' INSERT INTO CACMDRESPONSELOG (                         '#13 +
          '    COMPCODE, NAGRATRANSACTIONNUM, TRANSACTIONNUM,      '#13 +
          '    HIGH_LEVEL_CMD_ID, LOW_LEVEL_CMD, ICC_NO, STB_NO,   '#13 +
          '    FULLRESPONSETEXT, CMDRESPONSEID, UPDTIME   )        '#13 +
          ' VALUES ( ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',       '#13 +
          '          ''%s'', ''%s'', ''%s'', ''%s'', SYSDATE )     ',
          [aObj.CompCode,
           aObj.ResponseTransactionNumber,
           aObj.TransactionNumber,
           aObj.HighLevelCmd,
           aObj.LowLevelCmd,
           aObj.IccNo,
           aObj.StbNo,
           aObj.FullReponseCommandText,
           aObj.ResponseCommandId] );
        FCurrentActiveADO.DataWriter.ExecSQL;
        Dispose( aObj );
        aManager.Items[aIndex].Delete( aIndex2 );
        Inc( aWriteDone );
      except
        on E: Exception do
        begin
          FCurrentActiveSo.CriticalError := ( FCurrentActiveSo.CriticalError + 1 );
          FCurrentActiveSo.NotifyMsg :=
          Format( SDbWriteAckError,
            [FCurrentActiveSo.CompName, aObj.HighLevelCmd, aObj.SeqNo, E.Message] );
          FCurrentActiveSo.DbConnectStatus := dbError;
          Synchronize( Notify );
          Synchronize( Update );
        end;
      end;
      if FCurrentActiveSo.CriticalError >= 3 then
      begin
        FCurrentActiveSo.CriticalError := 0;
        FCurrentActiveSo.LastCriticalErrorTickCount := GetTickCount;
        FCurrentActiveSo.NotifyMsg := Format( SDbHasProblem,
          [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] );
        Synchronize( Notify );
        DisconnectFromDb( aIndex );
        Break;
      end;
    end;
    FCurrentActiveSo.NotifyMsg := Format( SDbWriteLogCount,
      [FCurrentActiveSo.CompName, aWriteDone] );
  finally
    aManager.Items[aIndex].EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteLog;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoInfo.Count - 1 do
  begin
    FCurrentActiveSo := PSoInfo( FSoInfo[aIndex] );
    FCurrentActiveADO := PSoADOControl( FADOControlList[aIndex] );
    PhysicalWriteLog( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
