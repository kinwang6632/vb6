unit cbSoDbThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, ADODB, DBClient,
  cbClass,Forms,SyncObjs;


type
  TSoDbThread = class(TSMSCommandThread)
  private
    { Private declarations }
    FSoInfo: TList;
    FADOControlList: TList;
    FCurrentActiveSo: PSoInfo;
    FCurrentActiveADO: PSoADOControl;
    FLastWriteLog: TDateTime;
    FGetDataLastTime: TDateTime;
    FCheckBusyTimePoint: TDateTime;
    FCanGetCommandData: Boolean;
    FUpdateGUI: Boolean;
    FSocketFail: Boolean;
    FSocketErrCode: String;
    FSocketErrMsg: String;
    FIppvCmdItem: String;
    FLock: TCriticalSection; //By Kin
    function GetIppvWhereCondition(const aType: String): String;
    function OnlineCSRSql(const aLeftCount: Integer): String;
    function BacthCSRSql(const aLeftCount: Integer): String;
    function CanDbRertyConnect: Boolean;
    function CheckDbStatus(aIndex: Integer): Boolean;
    function CheckCanWriteLog: Boolean;
    function CheckCanGetData: Boolean;
    function CheckSendListTooMany: Boolean;
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
    function GetData(aIndex: Integer): Integer;
    procedure AppendRecord(aIndex: Integer);
    procedure CustomSort(aIndex: Integer);
    procedure AddToList(aIndex: Integer);
    { 將已抓取的資料回寫 Send_Nagra 狀態成 P --> 已處理 }
    procedure PhysicalWriteAlreadyGet(aIndex: Integer); overload;
    { 將 Nagra 已回應指令取出, 回寫 Send_Nagra 狀態成 C --> 已完成或 E ---> 錯誤  }
    procedure PhysicalWriteNagraAck(aIndex: Integer); overload;
    procedure PhysicalWriteNagraAck; overload;
    { 寫 Log }
    procedure PhysicalWriteLog(aIndex: Integer); overload;
    procedure PhysicalWriteLog; overload;
    procedure SetErrorCode;
  protected
    function GetWaitWhileFrquence: Cardinal; override;
    function GetDbDelayReadFrequence: Integer;
    procedure Execute; override;
    procedure BeginActive; override;
    procedure EndActive; override;
    procedure Update; override;
    procedure UpdateEx;
    procedure WndProc(var Msg: TMessage); override;
  public
    constructor Create(const aSoInfo: TThreadList); overload;
    constructor Create(const aSoInfo: PSoInfo); overload;
    destructor Destroy; override;
    property UpdateGUI: Boolean read FUpdateGUI write FUpdateGUI;
  end;

implementation

uses cbMain, cbResStr, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSoDbThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }


var
  SendNagraSqlHeader: String =
    ' SELECT A.HIGH_LEVEL_CMD_ID,                                    ' +
    '        A.ICC_NO,                                               ' +
    '        A.STB_NO,                                               ' +
    '        TO_CHAR( A.SUBSCRIPTION_BEGIN_DATE, ''YYYYMMDD'' )      ' +
    '            AS SUBSCRIPTION_BEGIN_DATE,                         ' +
    '        TO_CHAR( A.SUBSCRIPTION_END_DATE, ''YYYYMMDD'' )        ' +
    '            AS SUBSCRIPTION_END_DATE,                           ' +
    '        A.CMD_STATUS,                                           ' +
    '        A.OPERATOR,                                             ' +
    '        NVL( A.UPDTIME, SYSDATE ) AS UPDTIME,                   ' +
    '        A.IMS_PRODUCT_ID,                                       ' +
    '        A.ZIP_CODE,                                             ' +
    '        A.CREDIT_MODE,                                          ' +
    '        A.CREDIT,                                               ' +
    '        A.CREDIT_LIMIT,                                         ' +
    '        A.THRESHOLD_CREDIT,                                     ' +
    '        A.EVENT_NAME,                                           ' +
    '        A.PRICE,                                                ' +
    '        A.CC_NUMBER_1,                                          ' +
    '        A.IP_ADDR,                                              ' +
    '        A.CC_PORT,                                              ' +
    '        A.CALLBACK_DATE,                                        ' +
    '        A.CALLBACK_TIME,                                        ' +
    '        A.CALLBACK_FREQUENCY,                                   ' +
    '        TO_CHAR( A.FIRST_CALLBACK_DATE, ''YYYYMMDD'' )          ' +
    '           AS FIRST_CALLBACK_DATE,                              ' +
    '        A.PHONE_NUM_1,                                          ' +
    '        A.PHONE_NUM_2,                                          ' +
    '        A.PHONE_NUM_3,                                          ' +
    '        A.SEQNO,                                                ' +
    '        A.MIS_IRD_CMD_ID,                                       ' +
    '        A.MIS_IRD_CMD_DATA,                                     ' +
    '        A.COMPCODE,                                             ' +
    '        A.NOTES,                                                ' +
    '        TO_CHAR( A.CLEANUP_DATE, ''YYYYMMDD'' )                 ' +
    '            AS CLEANUP_DATE,                                    ' +
    '        TO_CHAR( A.CONDITION_DATE, ''YYYYMMDD'' )               ' +
    '            AS CONDITION_DATE,                                  ' +
    '        TO_CHAR( A.COLLECT_DATE, ''YYYYMMDD'' )                 ' +
    '            AS COLLECT_DATE,                                    ' +
    '        NVL( A.RESENTTIMES, 0 ) AS RESENTTIMES,                 ' +
    '        NVL( A.STB_FLAG, 1 ) AS STB_FLAG,                       ' +
    '        A.PINCODE,                                              ' +
    '        A.DVRQUOTA                                              ' +
    '   FROM SEND_NAGRA A                                            ';


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
  FLock.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.InternalCreate;
begin
  FSoInfo := TList.Create;
  FADOControlList := TList.Create;
  FLastWriteLog := 0;
  FGetDataLastTime := 0;
  FCheckBusyTimePoint := 0;
  FCanGetCommandData := False;
  FSocketFail := False;
  FUpdateGUI := False;
  FIppvCmdItem := EmptyStr;
  FLock := TCriticalSection.Create;
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
        PSoInfo( aSourceList.Items[aIndex] ).CriticalErrorCount := 0;
        PSoInfo( aSourceList.Items[aIndex] ).LastCriticalErrorTime := 0;
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
      PSoADOControl( FADOControlList.Items[aIndex] ).DataReader.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).DataWriter.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).Connection.Free;
      PSoADOControl( FADOControlList.Items[aIndex] ).DataSet.Free;
      Dispose( PSoADOControl( FADOControlList.Items[aIndex] ) );
      FADOControlList.Items[aIndex] := nil;
    end;
  end;
  FADOControlList.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetIppvWhereCondition(const aType: String): String;
var
  aBuildConditon: Boolean;
begin
  aBuildConditon :=
    ( ( FIppvCmdItem = EmptyStr ) and ( aType = 'O' ) or
      ( FIppvCmdItem = EmptyStr ) and ( aType = 'N' ) );
  if ( aBuildConditon ) then
  begin
    HighCmdEnv.Filter := 'CMD_TYPE=''PPV''';
    HighCmdEnv.Filtered := True;
    try
      HighCmdEnv.First;
      while not HighCmdEnv.Eof do
      begin
        FIppvCmdItem := ( Result +
          Format( '''%s''', [HighCmdEnv.FieldByName( 'HIGH_LEVEL_CMD' ).AsString] ) );
        HighCmdEnv.Next;
        if ( not HighCmdEnv.Eof ) and ( Result <> EmptyStr ) then
          FIppvCmdItem := Result + ',';
      end;
    finally
      HighCmdEnv.Filtered := False;
      HighCmdEnv.Filter := EmptyStr;
    end;
  end;
  if ( FIppvCmdItem <> EmptyStr ) then
  begin
    if ( aType = 'O' )  then
      Result := Format( ' IN ( %s ) '#13, [FIppvCmdItem] )
    else
      Result := Format( ' NOT IN ( %s ) '#13, [FIppvCmdItem] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.OnlineCSRSql(const aLeftCount: Integer): String;
var
  aSql, aIppvCondition: String;
begin
  aSql := Format( SendNagraSqlHeader +
    '  WHERE A.CMD_STATUS = ''W''            ' +
    '    AND A.COMPCODE = ''%s''             ' +
    '    AND A.PROCESSINGDATE IS NULL        ',
    [FCurrentActiveSo.CompCode] );
  { 只處理 IPPV資料, 不處理 IPPV 資料 }
  if ( CommEnv.ProcessIPPV = 'O' ) or ( CommEnv.ProcessIPPV = 'N' ) then
  begin
    aIppvCondition := GetIppvWhereCondition( CommEnv.ProcessIPPV );
    if ( aIppvCondition <> EmptyStr ) then
      aSql := ( aSql +  ' AND HIGH_LEVEL_CMD_ID ' + aIppvCondition );
  end;
  aSql := ( aSql + ' ORDER BY A.SEQNO ' );
  Result := Format( ' SELECT * FROM ( %s ) WHERE ROWNUM <= %d ', [aSql, aLeftCount] );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.BacthCSRSql(const aLeftCount: Integer): String;
var
  aSql, aIppvCondition: String;
begin
  aSql := Format( SendNagraSqlHeader +
    '  WHERE A.CMD_STATUS = ''W''            ' +
    '    AND A.COMPCODE = ''%s''             ' +
    '    AND A.PROCESSINGDATE <= SYSDATE     ',
    [FCurrentActiveSo.CompCode] );
  { 只處理 IPPV資料, 不處理 IPPV 資料 }
  if ( CommEnv.ProcessIPPV = 'O' ) or ( CommEnv.ProcessIPPV = 'N' ) then
  begin
    aIppvCondition := GetIppvWhereCondition( CommEnv.ProcessIPPV );
    if ( aIppvCondition <> EmptyStr ) then
      aSql := ( aSql +  ' AND HIGH_LEVEL_CMD_ID ' + aIppvCondition );
  end;
  aSql := ( aSql + ' ORDER BY A.SEQNO ' );
  Result := Format( ' SELECT * FROM ( %s ) WHERE ROWNUM <= %d ',
    [aSql, aLeftCount] );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetDbDelayReadFrequence: Integer;
//const
//  aLastCheckTickCount: Cardinal = 0;
//  aStoredResult: Cardinal = 0;
var
//  aIsSetBusyTime, aIsBusyTime, aCanCheck: Boolean;
  aTimeStart, aTimeEnd: TDateTime;
  aHasSetBusyTime, aRightNowIsBusyTime, aWantToCheck: Boolean;
begin
//  Result := aStoredResult;
//  { 不須每次都檢查時段, 每 1 分鐘檢查一次即可 }
//  aCanCheck := ( ( GetTickCount - aLastCheckTickCount ) <= 0 ) or
//    ( aLastCheckTickCount = 0 ) or ( ( GetTickCount - aLastCheckTickCount ) > 60000 );
//  if aCanCheck then
//  begin
//    aIsBusyTime := False;
//    aIsSetBusyTime :=
//      ( CommEnv.BusyTimeStart <> EmptyStr ) and ( CommEnv.BusyTimeEnd <> EmptyStr );
//    if ( aIsSetBusyTime ) then
//    begin
//      aTimeStart := StrToTime( CommEnv.BusyTimeStart );
//      aTimeEnd := StrToTime( CommEnv.BusyTimeEnd );
//      aIsBusyTime := ( GetTime >= aTimeStart ) and ( GetTime <= aTimeEnd );
//    end;
//    if ( aIsSetBusyTime ) and ( aIsBusyTime ) then
//      Result := Trunc( CommEnv.BusyTimeDbReadFrequence * 1000 )
//    else
//      Result := Trunc( CommEnv.NormalTimeDbReadFrequence * 1000 );
//    aLastCheckTickCount := GetTickCount;
//  end;
//  aStoredResult := Result;
  Result := Trunc( CommEnv.NormalTimeDbReadFrequence );
  { 不須每次都檢查時段, 每 1 分鐘檢查一次即可 }
  aWantToCheck := ( FCheckBusyTimePoint <= 0 ) or ( SecondsBetween( Now, FCheckBusyTimePoint ) >= 60 );
  if ( aWantToCheck ) then
  begin
    aRightNowIsBusyTime := False;
    aHasSetBusyTime := ( CommEnv.BusyTimeStart <> EmptyStr ) and ( CommEnv.BusyTimeEnd <> EmptyStr );
    if ( aHasSetBusyTime ) then
    begin
      aTimeStart := StrToTime( CommEnv.BusyTimeStart );
      aTimeEnd := StrToTime( CommEnv.BusyTimeEnd );
      aRightNowIsBusyTime := ( GetTime >= aTimeStart ) and ( GetTime <= aTimeEnd );
    end;
    if ( aHasSetBusyTime ) and ( aRightNowIsBusyTime ) then
      Result := Trunc( CommEnv.BusyTimeDbReadFrequence );
    FCheckBusyTimePoint := Now;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetWaitWhileFrquence: Cardinal;
begin
  Result := inherited GetWaitWhileFrquence;
  if CommEnv.DbMultiThread then Result := ( Result * 3 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanDbRertyConnect: Boolean;
var
  //aRetry: Cardinal;
  aSecond: Int64;
begin
//  aRetry := GetBetweenFrquence( FCurrentActiveSo.LastCriticalErrorTickCount );
//  if aRetry <= 0 then aRetry := ( CommEnv.DbRetryFrequence * 1000 );
//  Result := ( aRetry >= ( CommEnv.DbRetryFrequence * 1000 ) );
  if ( FCurrentActiveSo.LastCriticalErrorTime <= 0 ) then
    aSecond := ( CommEnv.DbRetryFrequence )
  else
    aSecond := SecondsBetween( Now, FCurrentActiveSo.LastCriticalErrorTime );
  Result := ( aSecond >= CommEnv.DbRetryFrequence );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckDbStatus(aIndex: Integer): Boolean;
var
  aADO: PSoADOControl;
begin
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  if ( not aADO.Connection.Connected ) and CanDbRertyConnect then
  begin
    MessageSubject.Info( Format( SDbRertyConnect, [FCurrentActiveSo.CompName] ) );
    ConnectToDb( aIndex );
  end;
  Result := aADO.Connection.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckCanWriteLog: Boolean;
begin
  { 每 90 秒寫一次 Log }
  if ( FLastWriteLog <= 0 ) then FLastWriteLog := Now;
  Result := ( SecondsBetween( Now, FLastWriteLog ) >= 90 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckCanGetData: Boolean;
var
//  aTick, aFreq: Cardinal;
  aDbReadDelaySecond, aSecond: Int64;
begin
  { 每多久讀一次資料庫 }
//  aFreq := GetDelayFrquence;
//  aTick := GetTickCount - FGetDataLastTime;
//  if aTick <= 0 then aTick := aFreq;
//  Result := ( aTick >= aFreq );
  aDbReadDelaySecond := GetDbDelayReadFrequence;
  if ( FGetDataLastTime <= 0 ) then
    aSecond := aDbReadDelaySecond
  else
    aSecond := SecondsBetween( Now, FGetDataLastTime );
  Result := ( aSecond >= aDbReadDelaySecond );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckSendListTooMany: Boolean;
var
  aIndex, aCount: Integer;
begin
  { 時間到可以抓資料, 可是如果待傳送指令佇列大於設定的值, 則不可再抓取資料 }
  { 此規則是為了應付大量批次傳送指令 }
  aCount := 0;
  for aIndex := 0 to ControlSendManager.ItemCount - 1 do
  begin
    aCount := ( aCount + ControlSendManager.Items[aIndex].Count );
    PSoInfo( FSoInfo[aIndex] ).BufferCount := aCount;
    FCurrentActiveSo := PSoInfo( FSoInfo[aIndex] );
    Synchronize( Update );
  end;
  Result := ( aCount > CommEnv.DbWaterMark );
  if Result then
  begin
    if not Assigned( FCurrentActiveSo ) then
    begin
      if FSoInfo.Count > 0 then
        FCurrentActiveSo := PSoInfo( FSoInfo[0] );
    end;
    if Assigned( FCurrentActiveSo ) then
    begin
      MessageSubject.Warning( Format( SDbCheckSendListTooMany,
        [aCount, CommEnv.DbWaterMark] ) );
    end;
  end;
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
      aADO.Connection.LoginPrompt := False;
      aADO.DataReader.CacheSize := 1000;
      aADO.DataReader.Connection := aADO.Connection;
      aADO.DataReader.LockType := ltBatchOptimistic;
      aADO.DataWriter.CacheSize := 1000;
      aADO.DataWriter.Connection := aADO.Connection;
      aADO.DataSet.FieldDefs.Clear;
      aADO.DataSet.FieldDefs.Add( 'SENDSTATUS', ftString, 2 );
      aADO.DataSet.FieldDefs.Add( 'TRANSACTIONNUMBER', ftString, 9 );
      aADO.DataSet.FieldDefs.Add( 'CMD_TYPE_PRIORITY', ftInteger  );
      aADO.DataSet.FieldDefs.Add( 'CMD_PRIORITY', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'HIGH_LEVEL_CMD_ID', ftString, 4 );
      aADO.DataSet.FieldDefs.Add( 'ICC_NO', ftString, 20 );
      aADO.DataSet.FieldDefs.Add( 'STB_NO', ftString, 20 );
      aADO.DataSet.FieldDefs.Add( 'SUBSCRIPTION_BEGIN_DATE', ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'SUBSCRIPTION_END_DATE', ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'NOTES', ftString, 2000 );
      aADO.DataSet.FieldDefs.Add( 'CMD_STATUS', ftString, 1 );
      aADO.DataSet.FieldDefs.Add( 'OPERATOR', ftString, 20 );
      aADO.DataSet.FieldDefs.Add( 'UPDTIME', ftDateTime );
      aADO.DataSet.FieldDefs.Add( 'IMS_PRODUCT_ID', ftString, 12 );
      aADO.DataSet.FieldDefs.Add( 'ZIP_CODE', ftString, 5 );
      aADO.DataSet.FieldDefs.Add( 'CREDIT_MODE', ftFloat );
      aADO.DataSet.FieldDefs.Add( 'CREDIT', ftFloat );
      aADO.DataSet.FieldDefs.Add( 'CREDIT_LIMIT', ftFloat );
      aADO.DataSet.FieldDefs.Add( 'THRESHOLD_CREDIT', ftFloat );
      aADO.DataSet.FieldDefs.Add( 'EVENT_NAME', ftString, 32 );
      aADO.DataSet.FieldDefs.Add( 'PRICE', ftFloat );
      aADO.DataSet.FieldDefs.Add( 'CC_NUMBER_1', ftString, 16  );
      aADO.DataSet.FieldDefs.Add( 'IP_ADDR', ftString, 15 );
      aADO.DataSet.FieldDefs.Add( 'CC_PORT', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'CALLBACK_DATE', ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'CALLBACK_TIME', ftString, 6 );
      aADO.DataSet.FieldDefs.Add( 'CALLBACK_FREQUENCY', ftString, 2 );
      aADO.DataSet.FieldDefs.Add( 'FIRST_CALLBACK_DATE',ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'PHONE_NUM_1', ftString, 16 );
      aADO.DataSet.FieldDefs.Add( 'PHONE_NUM_2', ftString, 16  );
      aADO.DataSet.FieldDefs.Add( 'PHONE_NUM_3', ftString, 16  );
      aADO.DataSet.FieldDefs.Add( 'MIS_IRD_CMD_ID', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'MIS_IRD_CMD_DATA', ftString, 100 );
      aADO.DataSet.FieldDefs.Add( 'SEQNO', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'COMPCODE', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'CLEANUP_DATE', ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'CONDITION_DATE', ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'COLLECT_DATE', ftString, 8 );
      aADO.DataSet.FieldDefs.Add( 'RESENTTIMES', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'STB_FLAG', ftInteger );
      aADO.DataSet.FieldDefs.Add( 'PINCODE', ftString, 10 );
      aADO.DataSet.FieldDefs.Add( 'DVRQUOTA', ftString, 5 );
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
    FCurrentActiveSo.LastCriticalErrorTime := 0;
    FCurrentActiveSo.CriticalErrorCount := 0;
    MessageSubject.OK( Format( SDbConnectSuccess, [FCurrentActiveSo.CompName] ) );
  except
    on E: Exception do
    begin
      FCurrentActiveSo.LastCriticalErrorTime := Now;
      FCurrentActiveSo.DbConnectStatus := dbError;
      MessageSubject.Error( Format( SDbConnectError,
        [FCurrentActiveSo.CompName, E.Message, CommEnv.DbRetryFrequence] ) );
    end;
  end;
  { 通知主執行緒 }
  PostMessage( MainFormHandle, WM_DATABASE, Ord( CommandSubject.ThreadType ),
    Ord( aSoADO.Connection.Connected ) );
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
    FCurrentActiveSo.BufferCount := 0;
    MessageSubject.OK( Format( SDbDisConnectSuccess, [FCurrentActiveSo.CompName] ) );
  except
    on E: Exception do
    begin
      MessageSubject.Error( Format( SDbDisConnectError,
        [FCurrentActiveSo.CompName, E.Message] ) );
    end;
  end;
  Synchronize( Update );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Execute;
var
  aIndex: Integer;
  aCanWriteLog: Boolean;
begin
  PrepareADOControl;
  ConnectToDb;
  { 等待 Main Thread 打出執行的訊號 }
  WaitForPlaySignal;
  Sleep( GetWaitWhileFrquence );
  while ( not Stop ) do
  begin
    MessageSubject.RunState := rsRunning;
    try
      if ( FCanGetCommandData and CheckCanGetData ) then
      begin
        { 檢查傳送佇列是否有太多高階指令尚未處理 }
        if not CheckSendListTooMany then
        begin
          for aIndex := 0 to FSoInfo.Count - 1 do
          begin
            if Stop then Break;
            { 設定現正掃瞄對應的系統台 }
            FCurrentActiveSo := PSoInfo( FSoInfo.Items[aIndex] );
            FCurrentActiveADO := PSoADOControl( FADOControlList.Items[aIndex] );
            if FUpdateGUI then BeginActive;
            try
              { 抓 Send_Nagra 資料 }
              if GetData( aIndex ) = -1 then Continue; //By Kin
            finally
              if FUpdateGUI then EndActive;
            end;
            if Stop then Break;
            Sleep( GetWaitWhileFrquence );
          end;
        end;
        FGetDataLastTime := Now;
      end;
      if Stop then Break;
      Sleep( GetWaitWhileFrquence * 3 );
      { 是否已到寫 Log 時間 }
      aCanWriteLog := CheckCanWriteLog;
      aCanWriteLog := False;
      for aIndex := 0 to FSoInfo.Count - 1 do
      begin
        FCurrentActiveSo := PSoInfo( FSoInfo.Items[aIndex] );
        FCurrentActiveADO := PSoADOControl( FADOControlList.Items[aIndex] );
        if Stop then Break;
        PhysicalWriteNagraAck( aIndex );
        if Stop then Break;
        if aCanWriteLog then PhysicalWriteLog( aIndex );
        Sleep( GetWaitWhileFrquence );
      end;
      if Stop then Break;
    except
      on E: Exception do
        MessageSubject.Error( Format( SUnHandleException, [E.Message] ) );
    end;
  end;
  Sleep( GetWaitWhileFrquence );
  PhysicalWriteNagraAck;
//  PhysicalWriteLog; By Kin
  MessageSubject.RunState := rsStop;
  { 已先收到停止的訊號 }
  WaitForTerminalSignal;
  DisconnectFromDb;
  ReleaseADOResource;
  ReleaseSoInfoResource;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Update;
begin
  FLock.Enter; //By Kin
  try
    fmMain.UpdateSoTreeStatus( FCurrentActiveSo );
  finally
    FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.UpdateEx;
begin
  CommandSubject.UpdateCommandObject := UpdatedObject; 
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

function TSoDbThread.GetData(aIndex: Integer): Integer;
var
  aADO: PSoADOControl;
  aRecordCount: Integer;
  aSqlBacthCSR, aGetSql: String;

  { -------------------------------------- }

  function OpenDataSet(const aSQL: String): Integer;
  begin
    aADO.DataReader.Close;
    aADO.DataReader.SQL.Text := aSQL;
    aADO.DataReader.Open;
    Result := aADO.DataReader.RecordCount;
  end;

  { -------------------------------------- }

begin
  Result := 0;
  if not CheckDbStatus( aIndex ) then Exit;
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  { CommEnv.ProcessBatch = O  --> 只處理批次資料
    CommEnv.ProcessBatch = N  --> 不處理批次資料
    CommEnv.ProcessBatch = A  --> 處理批次資料及線上客服資料 }
  aGetSql := EmptyStr;  
  if ( CommEnv.ProcessBatch <> 'O' ) then
  begin
    { 1.處理線上客服的資料 }
    aGetSql := OnlineCSRSql( CommEnv.ProcessRecordCount );
  end;
  if ( CommEnv.ProcessBatch <> 'N' ) then
  begin
    { 2.處理批次 }
    aSqlBacthCSR := BacthCSRSql( CommEnv.ProcessRecordCount );
    if aGetSql = EmptyStr then
      aGetSql := aSqlBacthCSR
    else
      aGetSql := Format( ' %s UNION ALL %s ', [aGetSql, aSqlBacthCSR] );
  end;
  aGetSql := Format( ' SELECT * FROM ( %s ) WHERE ROWNUM <= %d ',
    [aGetSql, CommEnv.ProcessRecordCount] );
  try
    aRecordCount := OpenDataSet( aGetSql );
    { 該系統台累計本次抓取的數量 }
    FCurrentActiveSo.RecordCount := ( FCurrentActiveSo.RecordCount + aRecordCount );
    Inc( Result, aRecordCount );
    if ( aRecordCount > 0 ) then
    begin
      try
        PhysicalWriteAlreadyGet( aIndex );
      except
        on E: Exception do
        begin
          MessageSubject.Error( Format( SDbGetDataError, [FCurrentActiveSo.CompName, 'PW '+E.Message] ) );
          aRecordCount := -1; //By Kin
        end;
      end;
      try
        AppendRecord( aIndex );
      except
        on E: Exception do
        begin
          MessageSubject.Error( Format( SDbGetDataError, [FCurrentActiveSo.CompName, 'AR '+E.Message] ) );
          aRecordCount := -1; //By Kin
        end;

      end;
      try
        CustomSort( aIndex );
      except
        on E: Exception do
        begin
          MessageSubject.Error( Format( SDbGetDataError, [FCurrentActiveSo.CompName, 'CS '+E.Message] ) );
          aRecordCount := -1;  //By Kin
        end;
      end;
      try
        //將傳送的資料寫入到ControlSendManager ControlThread會負責傳送 Kin 註記
        if aRecordCount > 0 then
          AddToList( aIndex );
      except
        on E: Exception do
        begin
          MessageSubject.Error( Format( SDbGetDataError, [FCurrentActiveSo.CompName, 'AT '+E.Message] ) );
          aRecordCount := -1;
        end;
      end;
      Result := aRecordCount;
    end;
  except
    on E: Exception do
    begin
      MessageSubject.Error( Format( SDbGetDataError, [FCurrentActiveSo.CompName, E.Message] ) );
      FCurrentActiveSo.CriticalErrorCount := ( FCurrentActiveSo.CriticalErrorCount + 1 );
      FCurrentActiveSo.DbConnectStatus := dbError;
      Result := -1;
      Synchronize( Update );
    end;
  end;
  if ( FCurrentActiveSo.CriticalErrorCount >= 3 ) then
  begin
    MessageSubject.Warning( Format( SDbHasProblem, [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] ) );
    FCurrentActiveSo.CriticalErrorCount := 0;
    FCurrentActiveSo.LastCriticalErrorTime := Now;
    DisconnectFromDb( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.AppendRecord(aIndex: Integer);
var
  aADO: PSoADOControl;
begin
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  //By Kin
  try
    if ( aADO.DataSet.IndexName <> EmptyStr ) then
      aADO.DataSet.DeleteIndex( aADO.DataSet.IndexName );
    aADO.DataSet.EmptyDataSet;
    aADO.DataReader.First;
    while not aADO.DataReader.Eof do
    begin
      { 已是錯誤的, 不須傳送 }
      if aADO.DataReader.FieldByName( 'CMD_STATUS' ).AsString = 'E' then
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
        aADO.DataReader.FieldByName( 'UPDTIME' ).Value,
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
        aADO.DataReader.FieldByName( 'RESENTTIMES' ).Value,
        aADO.DataReader.FieldByName( 'STB_FLAG' ).Value,
        aADO.DataReader.FieldByName( 'PINCODE' ).Value,
        aADO.DataReader.FieldByName( 'DVRQUOTA' ).Value] );
      aADO.DataReader.Next;
    end;
  except
    on E: Exception do
    begin
      MessageSubject.Error( Format( SDbGetDataError, [FCurrentActiveSo.CompName, E.Message] ) );
    end;

  end;
  aADO.DataReader.Close;
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
    if Stop then Break;
    { 是否是批次的指令 }
    if ( aADO.DataSet.FieldByName( 'OPERATOR' ).AsString = CommEnv.BatchOperator ) and
       ( CommEnv.BatchOperator <> EmptyStr ) then
    begin
      aCmdTypePriority := 999;
      aCmdPriority := 999;
    end else
    if HighCmdEnv.Locate( 'HIGH_LEVEL_CMD',
      VarArrayOf( [aADO.DataSet.FieldByName( 'HIGH_LEVEL_CMD_ID' ).AsString] ), [] ) then
    begin
      aCmdTypePriority := HighCmdEnv.FieldByName( 'CMD_TYPE_PRIORITY' ).AsInteger;
      aCmdPriority := HighCmdEnv.FieldByName( 'CMD_PRIORITY' ).AsInteger;
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
    if Stop then Break;
  end;
  aADO.DataSet.AddIndex( 'CUSTOMSORT', 'CMD_TYPE_PRIORITY;CMD_PRIORITY;SEQNO', [] );
  aADO.DataSet.IndexName := 'CUSTOMSORT';
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
  aADO.DataSet.First;
  while not aADO.DataSet.Eof do
  begin
    if Stop then Break;
    New( aObj );
    aObj.SmsTransactionNumber := EmptyStr;
    aObj.HighCmdId := aADO.DataSet.FieldByName( 'HIGH_LEVEL_CMD_ID' ).AsString;
    aObj.IccNo := aADO.DataSet.FieldByName( 'ICC_NO' ).AsString;
    aObj.StbNo := aADO.DataSet.FieldByName( 'STB_NO' ).AsString;
    aObj.SubBeginDate :=
      aADO.DataSet.FieldByName( 'SUBSCRIPTION_BEGIN_DATE' ).AsString;
    aObj.SubEndDate :=
      aADO.DataSet.FieldByName( 'SUBSCRIPTION_END_DATE' ).AsString;
    aObj.Notes := aADO.DataSet.FieldByName( 'NOTES' ).AsString;
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
    aObj.PinCode := aADO.DataSet.FieldByName( 'PINCODE' ).AsString;
    aObj.DVRQuota := aADO.DataSet.FieldByName( 'DVRQUOTA' ).AsString;
    aObj.VodMoppId := EmptyStr;
    { 是否是全域指令, 前端不判斷, 不寫進 Send_Nagra, 判斷式寫死的, 寫在 Gateway 這邊 }
    aObj.IsGlobalCommand := False;
    if ( aObj.HighCmdId = 'E21' ) then
    begin
      aObj.IsGlobalCommand := True;
    end;
    { 低階指令才會用到的欄位 }
    aObj.SmsCommandText := EmptyStr;
    aObj.SmsFullCommandText := EmptyStr;
    aObj.LowCmdId := EmptyStr;
    aObj.CaAckCommandId := EmptyStr;
    aObj.CaAckStatus := EmptyStr;
    { 高階指令對應到總共幾筆低階指令, 剛從 Send_Nagra 取出來, 取值為 0 }
    { 低階指令用不到 }
    aObj.ReferenceLowCmdCount := 0;
    { 該筆階指令對應的 Send List 是第幾個, 方便後續 Ack 回來後快速使用該 List  }
    { 低階指令也會用到 }
    aObj.SoCompListIndex := aIndex;
    { KBRO 版本STB_FLA = 0 舊盒子, STB_FLAG = 1 新盒子, TBC 一律是1 }
    //aObj.STBFlag := aADO.DataSet.FieldByName( 'STB_FLAG' ).AsString;
    aObj.STBFlag := '1';
    { 該系統台的網路編號 }
    aObj.NetWorkId := FCurrentActiveSo.NetworkId;
    { 高階指令一起始狀態, 即標示為 P, 因為 DB 已更新為 P }
    aObj.CmdStatus := 'P';
    { 可是傳送狀態為 W }
    aObj.SmsSendStatus := 'W';
    { 起寫 Log 狀態, 此欄位高階指令不會用到, 是對應實際的低階指令才會用到 }
    { 預設填值 }
    aObj.HasbeenLog := False;
    aObj.Data := nil;
    aObj.ParentData := nil;
    aObj.Index := -1;
    aObj.ParentIndex := -1;
    if Stop then
    begin
      Dispose( aObj );
      Break;
    end; 
    { 將該筆高階指令現顯示在 TreeList 上 }
    UpdatedObject := aObj;
    Synchronize( UpdateEx );
    Sleep( 150 );
    { 寫入傳送佇列中 }
    ControlSendManager.Items[aIndex].BeginWrite;
    try
      ControlSendManager.Items[aIndex].AddObject(
        aADO.DataSet.FieldByName( 'SEQNO' ).AsString, TObject( aObj ) );
    finally
      ControlSendManager.Items[aIndex].EndWrite;
    end;
    aADO.DataSet.Next;
    Inc( aWriteCount );
  end;
  if ( aWriteCount > 0 ) then
  begin
    MessageSubject.Info( Format( SDbAddToControlSendList,
      [FCurrentActiveSo.CompName, aWriteCount] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteAlreadyGet(aIndex: Integer);
var
  aWriteSuccess, aWriteError: Integer;
  aADO: PSoADOControl;
  aCmdStatus, aErrCode, aErrMsg: String;
begin
  if not CheckDbStatus( aIndex ) then Exit;
  aWriteSuccess := 0;
  aWriteError := 0;
  aADO := PSoADOControl( FADOControlList.Items[aIndex] );
  aADO.DataReader.First;
  while not aADO.DataReader.Eof do
  begin
    { }
    if ( FSocketFail ) then
    begin
      aCmdStatus := 'E';
      aErrCode := FSocketErrCode;
      aErrMsg := FSocketErrMsg;
      Inc( aWriteError );
    end else
    begin
      aCmdStatus := 'P';
      aErrCode := EmptyStr;
      aErrMsg := EmptyStr;
      Inc( aWriteSuccess );
    end;
    { }
    aADO.DataWriter.SQL.Text := Format(
      ' UPDATE SEND_NAGRA                                 ' +
      '    SET CMD_STATUS = ''%s'',                       ' +
      '        ERR_CODE = ''%s'',                         ' +
      '        ERR_MSG = ''%s'',                          ' +  
      '        RESENTTIMES = NVL( RESENTTIMES, 0 ) + 1    ' +
      '  WHERE SEQNO = ''%s''                             ' +
      '    AND COMPCODE = ''%s''                          ' +
      '    AND CMD_STATUS IN ( ''W'', ''P'' )             ',
      [aCmdStatus, aErrCode, aErrMsg,
       aADO.DataReader.FieldByName( 'SEQNO' ).AsString,
       aADO.DataReader.FieldByName( 'COMPCODE' ).AsString] );
    try
      aADO.DataWriter.ExecSQL;
    except
      on E: Exception do
      begin
        MessageSubject.Error( Format( SDbWriteAlreadySendError,
          [FCurrentActiveSo.CompName,
           aADO.DataReader.FieldByName( 'HIGH_LEVEL_CMD_ID' ).AsString,
           aADO.DataReader.FieldByName( 'SEQNO' ).AsString, E.Message] ) );
        FCurrentActiveSo.CriticalErrorCount := ( FCurrentActiveSo.CriticalErrorCount + 1 );
        FCurrentActiveSo.DbConnectStatus := dbError;
        Synchronize( Update );
      end;
    end;
    aADO.DataReader.Edit;
    aADO.DataReader.FieldByName( 'CMD_STATUS' ).AsString := aCmdStatus;
    aADO.DataReader.Post;
    if ( FCurrentActiveSo.CriticalErrorCount >= 3 ) then
    begin
      MessageSubject.Warning( Format( SDbHasProblem,
        [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] ) );
      FCurrentActiveSo.CriticalErrorCount := 0;
      FCurrentActiveSo.LastCriticalErrorTime := Now;
      DisconnectFromDb( aIndex );
      Break;
    end;
    aADO.DataReader.Next;
    Sleep(10); // By Kin
  end;
  if ( aWriteSuccess > 0 ) then
  begin
    MessageSubject.Info( Format( SDbWriteAlreadySendCount,
      [FCurrentActiveSo.CompName, aWriteSuccess] ) );
  end;
  if ( aWriteError > 0 ) then
  begin
    MessageSubject.Warning( Format( SDbWriteNotSendCount,
      [FCurrentActiveSo.CompName, aWriteError] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PhysicalWriteNagraAck(aIndex: Integer);
var
  aWriteDone: Integer;
  aObj: PSendNagra;
begin
  if not CheckDbStatus( aIndex ) then Exit;
  aWriteDone := 0;
  while ControlRecvManager.Items[aIndex].Count > 0 do
  begin
    try // By Kin
      ControlRecvManager.Items[aIndex].BeginWrite; //一開始就先做Lock By Kin
      aObj := PSendNagra( ControlRecvManager.Items[aIndex].Objects[0] );
      FCurrentActiveADO.DataWriter.SQL.Clear;
      FCurrentActiveADO.DataWriter.SQL.Text := Format(
        ' UPDATE SEND_NAGRA                                 ' +
        '    SET CMD_STATUS = ''%s'',                       ' +
        '        ERR_CODE = ''%s'',                         ' +
        '        ERR_MSG = ''%s''                           ' +
        '  WHERE SEQNO = ''%s''                             ' +
        '    AND COMPCODE = ''%s''                          ' +
        '    AND CMD_STATUS IN ( ''W'', ''P'', ''E'' )      ',
        [aObj.SmsSendStatus, aObj.ErrCode, aObj.ErrMsg, aObj.SeqNo, aObj.CompCode] );

      try
        FCurrentActiveADO.DataWriter.ExecSQL;
        { 等 Nagara Ack 回來後, 就可釋放 }
        //ControlRecvManager.Items[aIndex].BeginWrite; //By Kin
        try
          //if aObj <> nil then Dispose( aObj ); // By Kin
          if ControlRecvManager.Items[ aIndex ].Count > 0 then
            ControlRecvManager.Items[aIndex].Delete( 0 );
        finally
          //ControlRecvManager.Items[aIndex].EndWrite;  // By Kin
        end;
        Inc( aWriteDone );
      except
        on E: Exception do
        begin
          MessageSubject.Error( Format( SDbWriteAckError,
          [FCurrentActiveSo.CompName, aObj.HighCmdId, aObj.SeqNo, E.Message] ) );
          FCurrentActiveSo.CriticalErrorCount := ( FCurrentActiveSo.CriticalErrorCount + 1 );
          FCurrentActiveSo.DbConnectStatus := dbError;
          Synchronize( Update );
        end;
      end;
    finally
      try
        ControlRecvManager.Items[aIndex].EndWrite;
        Dispose( aObj );
      except
        {}
      end;
    end;
    if ( FCurrentActiveSo.CriticalErrorCount >= 3 ) then
    begin
      MessageSubject.Error( Format( SDbHasProblem,
      [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] ) );
      FCurrentActiveSo.CriticalErrorCount := 0;
      FCurrentActiveSo.LastCriticalErrorTime := Now;
      DisconnectFromDb( aIndex );
      Break;
    end;
    Sleep(10); // By Kin
  end;

  if ( aWriteDone > 0 ) then
  begin
    MessageSubject.Info( Format( SDbWriteAckCount,
      [FCurrentActiveSo.CompName, aWriteDone] ) );
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
  aObj: PSendNagra;
  aWriteDone, aWaterMark, aObjIndex: Integer;

  { ----------------------------------------------- }

  procedure PrepareLogSql;
  begin
    FCurrentActiveADO.DataWriter.SQL.Clear;
    FCurrentActiveADO.DataWriter.SQL.Text :=
      ' INSERT INTO NAGRACMDLOG (                          ' +
      '    COMPCODE, CATRANNUM, GWTRANNUM,                 ' +
      '    HIGHCMD, LOWCMD, ICC, STB,                      ' +
      '    SENDCMDTEXT, RECVCMDTEXT, ACK,                  ' +
      '    OPERATOR, UPDTIME, SENDTIME, RECVTIME  )        ' +
      ' VALUES ( :1, :2, :3, :4, :5, :6, :7,               ' +
      '          :8, :9, :10, :11, :12, :13, :14 )         ';
  end;  

  { ----------------------------------------------- }
  
begin
  { 先記錄本次處理 Log 的時間 }
  FLastWriteLog := Now;
  if not CheckDbStatus( aIndex ) then Exit;
  if ControlSendLogManager.Items[aIndex].Count <= 0 then Exit;
  aWriteDone := 0;
  { Log 處理原則是每次只處理 Log 佇列裏的 2/3 資料 }
  aWaterMark := ControlSendLogManager.Items[aIndex].Count;
  { 非停止執行狀況下, 才取 2/3 資料寫 Log }
  if not Stop then
    aWaterMark := Trunc( ( aWaterMark / 3 ) * 2 ) + 1;
  { 處理佇列裏的位置 }  
  aObjIndex := 0;
  while aWaterMark > 0 do
  begin
    aObj := PSendNagra( ControlSendLogManager.Items[aIndex].Objects[aObjIndex] );
    try
      { 是否已 Ack 回來, 若是尚未 ACK 回來則此 Object 不可釋放 }
      { 迴圈起頭需取佇列中的下一筆 }
      if ( aObj.SmsSendStatus = 'C' ) or
         ( aObj.SmsSendStatus = 'E' ) then
      begin
        { Ack 回來的 Log }
        PrepareLogSql;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '1' ).Value :=
          aObj.CompCode;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '2' ).Value :=
          aObj.CaAckTransactionNumber;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '3' ).Value :=
          aObj.SmsTransactionNumber;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '4' ).Value :=
          aObj.HighCmdId;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '5' ).Value :=
          aObj.LowCmdId;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '6' ).Value :=
          aObj.IccNo;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '7' ).Value :=
          aObj.StbNo;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '8' ).Value :=
          aObj.SmsCommandText;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '9' ).Value :=
          aObj.CaAckCommandText;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '10' ).Value :=
          aObj.CaAckCommandId;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '11' ).Value :=
          aObj.Operator;
        FCurrentActiveADO.DataWriter.Parameters.ParamByName( '12' ).Value :=
          aObj.UpdTime;
        {}
        if ( aObj.SmsSendTime <> -1 ) then
          FCurrentActiveADO.DataWriter.Parameters.ParamByName( '13' ).Value := aObj.SmsSendTime
        else
          FCurrentActiveADO.DataWriter.Parameters.ParamByName( '13' ).Value := EmptyStr;
        {}
        if ( aObj.CaAckTime <> -1 ) then
          FCurrentActiveADO.DataWriter.Parameters.ParamByName( '14' ).Value := aObj.CaAckTime
        else
          FCurrentActiveADO.DataWriter.Parameters.ParamByName( '14' ).Value := EmptyStr;
        {}
        FCurrentActiveADO.DataWriter.ExecSQL;
        { Log 寫完後就可以釋放 }
        ControlSendLogManager.Items[aIndex].BeginWrite;
        try
          Dispose( aObj );
          ControlSendLogManager.Items[aIndex].Delete( aObjIndex );
        finally
          ControlSendLogManager.Items[aIndex].EndWrite;
        end;
        Inc( aWriteDone );
      end else
      begin
        Inc( aObjIndex );
      end;
      Dec( aWaterMark );
    except
      on E: Exception do
      begin
        MessageSubject.Error( Format( SDbWriteAckError,
          [FCurrentActiveSo.CompName, aObj.HighCmdId, aObj.SeqNo, E.Message] ) );
        FCurrentActiveSo.CriticalErrorCount := ( FCurrentActiveSo.CriticalErrorCount + 1 );
        FCurrentActiveSo.DbConnectStatus := dbError;
        Synchronize( Update );
      end;
    end;
    if ( FCurrentActiveSo.CriticalErrorCount >= 3 ) then
    begin
      MessageSubject.Error( Format( SDbHasProblem,
        [FCurrentActiveSo.CompName, CommEnv.DbRetryFrequence] ) );
      FCurrentActiveSo.CriticalErrorCount := 0;
      FCurrentActiveSo.LastCriticalErrorTime := Now;
      DisconnectFromDb( aIndex );
      Break;
    end;
  end;
  if ( aWriteDone > 0 ) then
  begin
    MessageSubject.Info( Format( SDbWriteLogCount,
      [FCurrentActiveSo.CompName, aWriteDone] ) );
  end;
  { 再記錄一次處理 }
  FLastWriteLog := Now;
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

procedure TSoDbThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_UPDATEGUI:
      begin
        if Msg.LParam = Ord( False ) then
          CommandSubject.RemoveObServer( TObServer( Msg.WParam ) )
        else
          CommandSubject.AddObServer( TObServer( Msg.WParam ) );
      end;
    WM_SOCKET:
      begin
        FSocketErrCode := EmptyStr;
        FSocketErrMsg := EmptyStr;
        if TWMSocket( Msg ).ErrorCode <> -1 then
          FSocketErrCode := IntToStr( TWMSocket( Msg ).ErrorCode );
        if TWMSocket( Msg ).ErrorMsg <> -1 then
          FSocketErrMsg := IntToStr( TWMSocket( Msg ).ErrorMsg );
        if ( FSocketErrCode <> EmptyStr ) or ( FSocketErrMsg <> EmptyStr ) then
          SetErrorCode;
        FSocketFail := ( TNotifyCommandType( TWMSocket( Msg ).NotifyCommandType )
          in [ncNack, ncSocketError] );
        if not FSocketFail then
          FCanGetCommandData := True
        else begin
          if CommEnv.DbWriteErrorWhenSocketFail then
            FCanGetCommandData := True;
        end;
      end;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.SetErrorCode;
begin
  if CmdError.Locate( 'ERROR_FLAG;ERROR_CODE',
    VarArrayOf( [0, FSocketErrCode] ), [] ) then
   FSocketErrCode := Format( '%s--%s',
     [FSocketErrCode, CmdError.FieldByName( 'ERROR_DESC' ).AsString] );
  if CmdError.Locate( 'ERROR_FLAG;ERROR_CODE',
    VarArrayOf( [1, FSocketErrMsg] ), [] ) then
   FSocketErrMsg := Format( '%s--%s',
     [FSocketErrMsg, CmdError.FieldByName( 'ERROR_DESC' ).AsString] );
end;

{ ---------------------------------------------------------------------------- }

end.
