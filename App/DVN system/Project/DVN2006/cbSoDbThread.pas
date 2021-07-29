unit cbSoDbThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, ADODB, DBClient,
  {$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF} cbAppClass;

type
  TSoDbThread = class(TAppThread)
  private
    { Private declarations }
    FSoList: TSoList;
    FGetDataTime: TDateTime;
    FTextSubj: TTextSubject;
    FSoSubj: TSoSubject;
    FCmdSubj: TSendDvnSubject;
    FCanGetData: Boolean;
    FSocketFail: Boolean;
    FLastWriteLog: TDateTime;
    FWMSocket: TWMSocket;
    procedure SetSoList(const Value: TSoList);
    procedure BeginActive(const aSo: TSo);
    procedure EndActive(const aSo: TSo);
    procedure CreateADOControl;
    procedure ReleaseADOControl;
    procedure ConnectToDb; overload;
    procedure ConnectToDb(const aSo: TSo); overload;
    procedure DisconnectFromDb; overload;
    procedure DisconnectFromDb(const aSo: TSo); overload;
    function CheckCanDbReConnect(const aSo: TSo): Boolean;
    function CheckCanGetData: Boolean;
    function CheckCanWriteLog: Boolean;
    function CheckDbState(const aSo: TSo): Boolean;
    function CheckDbWaterMarkTooHigh: Boolean;
    procedure WriteDataGet(const aSo: TSo);
    procedure WriteDataAck; overload;
    procedure WriteDataLog; overload;
    procedure GetData; overload;
    procedure GetData(const aSo: TSo); overload;
    function GetIppvWhereCondition(const aType: String): String;
    function OnlineCSRSql(const aSo: TSo; const aLeftCount: Integer): String;
    function BacthCSRSql(const aSo: TSo; const aLeftCount: Integer): String;
    procedure AddToBuffer(const aSo: TSo);
    procedure BufferSort(const aSo: TSo);
    procedure BufferToSendList(const aSo: TSo);
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property SoList: TSoList read FSoList write SetSoList;
    property TextSubj: TTextSubject read FTextSubj;
    property SoSubj: TSoSubject read FSoSubj;
    property CmdSubj: TSendDvnSubject read FCmdSubj;
  end;

implementation

uses cbMain, cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSoDbThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

var

  BufferIndexName: String = 'IDX1';

  SendDvnSqlHeader: String =
    ' select a.high_level_cmd_id,  ' +
    '        a.icc_no,             ' +
    '        a.stb_no,             ' +
    '        a.areacode,           ' +
    '        a.notes,              ' +
    '        a.cmd_status,         ' +
    '        a.operator,           ' +
    '        a.updtime,            ' +
    '        a.err_code,           ' +
    '        a.err_msg,            ' +
    '        a.startingtime,       ' +
    '        a.expirytime,         ' +
    '        a.product_id,         ' +
    '        a.credit,             ' +
    '        a.creditaction,       ' +
    '        a.creditcfcode,       ' +
    '        a.banktype,           ' +
    '        a.producttype,        ' +
    '        a.configvalue,        ' +
    '        a.mailaction,         ' +
    '        a.compcode,           ' +
    '        a.sno,                ' +
    '        a.processingdate,     ' +
    '        a.seqno               ' +
    '   from send_dvn a            ';

{ ---------------------------------------------------------------------------- }

function BuildWrieAckSql(aObj: TSendDvn): String;
begin
  Result := Format(
   ' update send_dvn                       ' +
   '    set cmd_status = ''%s'',           ' +
   '        err_code = ''%s'',             ' +
   '        err_msg = ''%s'',              ' +
   '        creditcfcode = nvl( ''%s'', creditcfcode ) ' +
   '  where seqno = ''%s''                 ' +
   '    and compcode = ''%s''              ' +
   '    and cmd_status in ( ''P'', ''W'' ) ',
   [aObj.HighCmdStatus,
    aObj.ErrCode,
    aObj.ErrMsg,
    aObj.CreditCfCode,
    aObj.HighSeqNo,
    aObj.CompCode] );
end;

{ ---------------------------------------------------------------------------- }

function BuildWriteLogSql(aObj: TSendDvn): String;
var
  aUpd, aSend, aRecv: String;
begin
  aUpd := EmptyStr;
  aSend := EmptyStr;
  aRecv := EmptyStr;
  if ( aObj.UpdTime <>  0 ) then
    aUpd := FormatDateTime( 'yyyymmdd hhnnss', aObj.UpdTime );
  if ( aObj.SendTime <> 0 ) then
    aSend := FormatDateTime( 'yyyymmdd hhnnss', aObj.SendTime );
  if ( aObj.RecvTime <> 0 ) then
    aRecv := FormatDateTime( 'yyyymmdd hhnnss', aObj.RecvTime );
  Result := Format(
    ' insert into dvncmdlog (                           ' +
    '   compcode, icc_no, stb_no, high_level_cmd_id,    ' +
    '   low_level_cmd_id, seqno, frameno, sendcmdtext,  ' +
    '   recvcmdtext, ack, operator, updtime,            ' +
    '   sendtime, recvtime )                            ' +
    ' values ( ''%s'', ''%s'', ''%s'', ''%s'',          ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'',                         ' +
    '   to_date( ''%s'', ''YYYYMMDD HH24MISS'' ),       ' +
    '   to_date( ''%s'', ''YYYYMMDD HH24MISS'' ),       ' +
    '   to_date( ''%s'', ''YYYYMMDD HH24MISS'' ) )      ',
    [aObj.CompCode, aObj.IccNo, aObj.StbNo, aObj.HighCmd,
     aObj.LowCmd, aObj.HighSeqNo, aObj.FrameNo, aObj.SendCmdText,
     aObj.RecvCmdText, aObj.Ack, aObj.Operator, aUpd, aSend, aRecv] );
end;

{ ---------------------------------------------------------------------------- }

function BuildWriteGetSql(aReader: TADOQuery): String;
begin
  Result := Format(
    ' update send_dvn                                   ' +
    '    set cmd_status = ''%s'',                       ' +
    '        err_code = ''%s'',                         ' +
    '        err_msg = ''%s''                           ' +
    '  where seqno = ''%s''                             ' +
    '    and compcode = ''%s''                          ' +
    '    and cmd_status in ( ''W'', ''P'' )             ',
    [aReader.FieldByName( 'cmd_status' ).AsString,
     aReader.FieldByName( 'err_code' ).AsString,
     aReader.FieldByName( 'err_msg' ).AsString,
     aReader.FieldByName( 'seqno' ).AsString,
     aReader.FieldByName( 'compcode' ).AsString] );
end;

{ ---------------------------------------------------------------------------- }

{ TSoDbThread }

constructor TSoDbThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FSoList := TSoList.Create;
  FGetDataTime := 0;
  FLastWriteLog := 0;
  FTextSubj := TTextSubject.Create;
  FSoSubj := TSoSubject.Create;
  FCmdSubj := TSendDvnSubject.Create;
  FCanGetData := False;
  FSocketFail := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TSoDbThread.Destroy;
begin
  StrDispose( FWMSocket.ErrorDescription );
  FSoList.Free;
  FTextSubj.Free;
  FSoSubj.Free;
  FCmdSubj.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_SOCKET:
      begin
        FWMSocket.Msg := TWMSocket( Msg ).Msg;
        FWMSocket.ThreadType := TWMSocket( Msg ).ThreadType;
        FWMSocket.SocketMsgReason := TWMSocket( Msg ).SocketMsgReason;
        FSocketFail := ( FWMSocket.SocketMsgReason in [Ord( smrSocketError )] );
        if ( FSocketFail ) then
        begin
          FWMSocket.ErrorCode := TWMSocket( Msg ).ErrorCode;
          StrDispose( FWMSocket.ErrorDescription );
          FWMSocket.ErrorDescription := StrNew( TWMSocket( Msg ).ErrorDescription );
        end;
        if ( not FSocketFail ) or ( CommEnv.DbWriteError ) then
          FCanGetData := True
        else
          FCanGetData := False;
        if ( not FSocketFail ) then
          FTextSubj.Ok( LanguageManager.Get( 'SDbWaitForSocketOk' ) )
        else begin
          if ( CommEnv.DbWriteError ) then
            FTextSubj.Warning( LanguageManager.Get( 'SDbWaitForSocketError1' ) )
          else
            FTextSubj.Warning( LanguageManager.Get( 'SDbWaitForSocketError2' ) )
        end;
      end;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.SetSoList(const Value: TSoList);
begin
  FSoList.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.CreateADOControl;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    FSoList[aIndex].Connection := TADOConnection.Create( nil );
    FSoList[aIndex].Connection.LoginPrompt := False;
    {}
    FSoList[aIndex].Reader := TADOQuery.Create( nil );
    FSoList[aIndex].Reader.CacheSize := 1000;
    FSoList[aIndex].Reader.Prepared := True;
    FSoList[aIndex].Reader.LockType := ltBatchOptimistic;
    FSoList[aIndex].Reader.Connection := FSoList[aIndex].Connection;
    {}
    FSoList[aIndex].Writer := TADOQuery.Create( nil );
    FSoList[aIndex].Writer.CacheSize := 1000;
    FSoList[aIndex].Writer.Prepared := True;
    FSoList[aIndex].Writer.LockType := ltBatchOptimistic;
    FSoList[aIndex].Writer.Connection := FSoList[aIndex].Connection;
    {}
    FSoList[aIndex].Buffer := TClientDataSet.Create( nil );
    FSoList[aIndex].Buffer.FieldDefs.Clear;
    FSoList[aIndex].Buffer.FieldDefs.Add( 'PRIORITY1', ftInteger );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'PRIORITY2', ftInteger );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'HIGH_LEVEL_CMD_ID', ftString, 4 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'ICC_NO', ftString, 20 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'STB_NO', ftString, 20 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'AREACODE', ftString, 20 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'NOTES', ftString,  2000 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'CMD_STATUS', ftString, 1 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'OPERATOR', ftString, 20 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'UPDTIME', ftDateTime );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'STARTINGTIME', ftDateTime );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'EXPIRYTIME', ftDateTime );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'PRODUCT_ID', ftString, 20 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'CREDIT', ftString, 10 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'CREDITACTION', ftString, 3 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'CREDITCFCODE', ftString, 16 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'BANKTYPE', ftString, 3 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'PRODUCTTYPE', ftString, 2 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'CONFIGVALUE', ftString, 20 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'MAILACTION', ftString, 3 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'COMPCODE', ftString, 3 ) ;
    FSoList[aIndex].Buffer.FieldDefs.Add( 'SNO', ftString, 15 );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'PROCESSINGDATE', ftDateTime );
    FSoList[aIndex].Buffer.FieldDefs.Add( 'SEQNO', ftInteger );
    FSoList[aIndex].Buffer.CreateDataSet;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ReleaseADOControl;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if Assigned( FSoList[aIndex].Writer ) then
       FSoList[aIndex].Writer.Free;
    if Assigned( FSoList[aIndex].Reader ) then
      FSoList[aIndex].Reader.Free;
    if Assigned( FSoList[aIndex].Connection ) then
      FSoList[aIndex].Connection.Free;
    if Assigned( FSoList[aIndex].Buffer ) then
      FSoList[aIndex].Buffer.Free;
    {}
    FSoList[aIndex].Writer := nil;
    FSoList[aIndex].Reader := nil;
    FSoList[aIndex].Connection := nil;
    FSoList[aIndex].Buffer := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ConnectToDb;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].Selected ) then
    begin
      FTextSubj.InfoFmt( LanguageManager.Get( 'SDbInitConnect' ),
        [FSoList[aIndex].CompName] );
      ConnectToDb( FSoList[aIndex] );
      Sleep( 150 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ConnectToDb(const aSo: TSo);
begin
  if aSo.Connection.Connected then aSo.Connection.Close;
  aSo.Connection.ConnectionString := Format(
    'Provider=MSDAORA.1;Persist Security Info=True;' +
    'Password=%s;User ID=%s;Data Source=%s;', [aSo.DbUserPass, aSo.DbUserId, aSo.DbAliase] );
  try
    aSo.Connection.Open;
    aSo.DbState := dbOpen;
    FTextSubj.OKFmt( LanguageManager.Get( 'SDbConnectSuccess' ), [aSo.CompName] );
  except
    on E: Exception do
    begin
      aSo.DbState := dbError;
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SDbConnectError' ),
        [aSo.CompName, E.Message, CommEnv.DbRetryFreq] );
    end;
  end;
  FSoSubj.Notify( aSo );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.DisconnectFromDb;
var
  aIndex: Integer;
begin
  FTextSubj.Info( LanguageManager.Get( 'SDbDisConnectStart' ) );
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].Selected ) then
    begin
      DisconnectFromDb( FSoList[aIndex] );
      Sleep( 150 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.DisconnectFromDb(const aSo: TSo);
begin
  try
    if aSo.Connection.Connected then aSo.Connection.Close;
    aSo.DbState := dbClose;
    aSo.RecordCount := 0;
    FTextSubj.OKFmt( LanguageManager.Get( 'SDbDisConnectSuccess' ), [aSo.CompName] );
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SDbDisConnectError' ),
        [aSo.CompName, E.Message] );
    end;
  end;
  FSoSubj.Notify( aSo );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckCanDbReConnect(const aSo: TSo): Boolean;
begin
  Result :=
    ( not Self.Terminated ) and
    ( SecondsBetween( Now, aSo.ErrorTime ) > CommEnv.DbRetryFreq );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckCanGetData: Boolean;
var
  aCriterion: Integer;
begin
  Result := ( FGetDataTime <= 0 );
  if ( not Result ) then
  begin
    aCriterion := CommEnv.NorTimeReadFreq;
    if ( Time >= StrToTime( CommEnv.BusyTimeStart ) ) and
       ( Time <= StrToTime( CommEnv.BusyTimeEnd ) ) then
    begin
      aCriterion := CommEnv.BusyTimeReadFreq;
    end;
    Result := ( SecondsBetween( Now, FGetDataTime ) >= aCriterion ) ;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckCanWriteLog: Boolean;
begin
  { 每 90 秒 寫一次 log }
  if ( FLastWriteLog <= 0 ) then FLastWriteLog := Now;
  Result := ( SecondsBetween( Now, FLastWriteLog ) >= 90 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckDbState(const aSo: TSo): Boolean;
begin
  if ( not aSo.Connection.Connected ) then
  begin
    if CheckCanDbReConnect( aSo ) then
    begin
      FTextSubj.InfoFmt( LanguageManager.Get( 'SDbRertyConnect' ), [aSo.CompName] );
      ConnectToDb( aSo );
    end;
  end;
  Result := aSo.Connection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.BeginActive(const aSo: TSo);
begin
  if ( aSo.DbState in [dbOpen, dbActive] ) then aSo.DbState := dbActive;
  FSoSubj.Notify( aSo );
  Sleep( 100 );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.EndActive(const aSo: TSo);
begin
  if ( aSo.DbState in [dbOpen, dbActive] ) then aSo.DbState := dbOpen;
  FSoSubj.Notify( aSo );
  Sleep( 100 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CheckDbWaterMarkTooHigh: Boolean;
var
  aCount: Integer;
begin
  aCount := HCmdBeginList.Count;
  Result := ( aCount > CommEnv.DbWaterMark );
  if ( Result ) then
    FTextSubj.WarningFmt( LanguageManager.Get( 'DbWaterMarkTooHigh' ),
     [aCount, CommEnv.DbWaterMark] );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Execute;
begin
  CreateADOControl;
  Sleep( 100 );
  ConnectToDb;
  Sleep( 100 );
  FTextSubj.Info( LanguageManager.Get( 'SDbWaitForSocketReady' ) );
  while not Self.Terminated do
  begin
    if ( FCanGetData ) and ( CheckCanGetData ) then
    begin
      if ( not CheckDbWaterMarkTooHigh ) then
      begin
        GetData;
        Sleep( 10 );
      end;
    end;
    if ( Self.Terminated ) then Break;
    WriteDataAck;
    Sleep( 10 );
    if CheckCanWriteLog then WriteDataLog;
    Sleep( 10 );
    if ( Self.Terminated ) then Break;
  end;
  Sleep( 500 );
  WriteDataAck;
  Sleep( 10 );
  WriteDataLog;
  Sleep( 10 );
  DisconnectFromDb;
  Sleep( 10 );
  ReleaseADOControl;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.WriteDataGet(const aSo: TSo);
var
  aWriteCount: Integer;
  aCmdStatus, aErrCode, aErrMsg: String;
begin
  if not CheckDbState( aSo ) then Exit;
  aWriteCount := 0;
  aSo.Reader.First;
  while not aSo.Reader.Eof do
  begin
    aCmdStatus := 'P';
    aErrCode := EmptyStr;
    aErrMsg := EmptyStr;
    if ( FSocketFail ) and ( CommEnv.DbWriteError ) then
    begin
      aCmdStatus := 'E';
      { 這邊做個轉換, 若是 Socket 有誤, 不寫 ErrorCode }
      //aErrCode := IntToStr( FWMSocket.ErrorCode );
      aErrCode := 'Socket Error';
      aErrMsg := FWMSocket.ErrorDescription;
    end;
    aSo.Reader.Edit;
    aSo.Reader.FieldByName( 'cmd_status' ).AsString := aCmdStatus;
    aSo.Reader.FieldByName( 'err_code' ).AsString := aErrCode;
    aSo.Reader.FieldByName( 'err_msg' ).AsString := aErrMsg;
    aSo.Reader.Post;
    aSo.Writer.SQL.Text := BuildWriteGetSql( aSo.Reader );
    try
      aSo.Writer.ExecSQL;
      Inc( aWriteCount );
    except
      on E: Exception do
      begin
        aSo.DbState := dbError;
        FTextSubj.ErrorFmt( LanguageManager.Get( 'SDbWriteGetError' ),
          [aSo.CompName,
           aSo.Reader.FieldByName( 'seqno' ).AsString,
           aSo.Reader.FieldByName( 'high_level_cmd_id' ).AsString, E.Message] );
        FSoSubj.Notify( aSo );
      end;
    end;
    if ( aSo.ErrorCount >= 3 ) then
    begin
      DisconnectFromDb( aSo );
      FTextSubj.WarningFmt( LanguageManager.Get( 'SDbHasProblem' ),
        [aSo.CompName, CommEnv.DbRetryFreq] );
      FSoSubj.Notify( aSo );
    end;
    if Self.Terminated then Break;
    aSo.Reader.Next;
  end;
  if ( aWriteCount > 0 ) then
  begin
    if ( not FSocketFail ) then
      FTextSubj.InfoFmt( LanguageManager.Get( 'SDbWriteGetCount1' ),
        [aSo.CompName, aWriteCount] )
    else
      FTextSubj.InfoFmt( LanguageManager.Get( 'SDbWriteGetCount2' ),
        [aSo.CompName, aWriteCount, aErrCode, aErrMsg] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.WriteDataAck;
var
  aIndex, aCount, aSuccess, aError: Integer;
  aObj: TSendDvn;
  aSo: TSo;
begin
  aSo := nil;
  aIndex := 0;
  aSuccess := 0;
  aError := 0;
  aCount := HCmdFinishList.Count;
  while ( aIndex < aCount ) do
  begin
    aObj := HCmdFinishList.Objects[aIndex];
    if not Assigned( aSo ) then
      aSo := FSoList[FSoList.IndexOf( aObj.CompCode )];
    if ( aObj.CompCode <> aSo.CompCode ) then
      aSo := FSoList[FSoList.IndexOf( aObj.CompCode )];
    if not CheckDbState( aSo ) then
    begin
      Inc( aIndex );
      Continue;
    end;
    aSo.Writer.SQL.Text := BuildWrieAckSql( aObj );
    try
      aSo.Writer.ExecSQL;
      if ( aObj.HighCmdStatus = 'C' ) then
        Inc( aSuccess )
      else
        Inc( aError );
      HCmdFinishList.FreeObject( aIndex );
      Dec( aCount );
    except
      on E: Exception do
      begin
        Inc( aIndex );
        aSo.DbState := dbError;
        FTextSubj.ErrorFmt( LanguageManager.Get( 'SDbWriteAckError' ),
          [aSo.CompName, aObj.HighSeqNo, aObj.HighCmd, E.Message] );
        FSoSubj.Notify( aSo );
      end;
    end;
    if ( aSo.ErrorCount >= 3 ) then
    begin
      DisconnectFromDb( aSo );
      FTextSubj.WarningFmt( LanguageManager.Get( 'SDbHasProblem' ),
        [aSo.CompName, CommEnv.DbRetryFreq] );
      FSoSubj.Notify( aSo );
    end;
  end;
  if ( aSuccess > 0 )  or ( aError > 0 ) then
  begin
    FTextSubj.InfoFmt( LanguageManager.Get( 'SDbWriteAckCount' ),
      [aSo.CompName, aSuccess, aError] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.WriteDataLog;
var
  aIndex, aDone, aCount: Integer;
  aObj: TSendDvn;
  aSo: TSo;
begin
  aIndex := 0;
  aDone := 0;
  aCount := LCmdLogList.Count;
  aSo := nil;
  while ( aIndex < aCount ) do
  begin
    aObj := LCmdLogList.Objects[aIndex];
    { 如果是 Thread 正在跑, 則要檢核有 Ack 的才可以寫 Log }
    { 若是 thread 已經被呼叫中斷, 則不檢核, 一律寫 log }
    if ( not Self.Terminated ) then
    begin
      { 先檢核此筆低階指令是否已 Ack 回來, 沒有 Ack 回來不寫 Log }
      if ( aObj.LowCmdStatus = 'P' ) then
      begin
        Inc( aIndex );
        Continue;
      end else
      { 若是已經是 Ack 來, 檢查 Ack 回來的時間大於 30 秒, 才可以寫 log }
      if ( aObj.LowCmdStatus = 'C' ) then
      begin
        if ( SecondsBetween( Now, aObj.RecvTime ) <= 30 ) then
        begin
          Inc( aIndex );
          Continue;
        end;
      end else
      { 錯誤指令, 若是 RecvTime 有值, 則判斷是否已大於 30 秒 }
      if ( aObj.LowCmdStatus = 'E' ) then
      begin
        if ( aObj.RecvTime > 0 ) then
        begin
          if ( SecondsBetween( Now, aObj.RecvTime ) <= 30 ) then
          begin
            Inc( aIndex );
            Continue;
          end;
        end;  
      end;
    end;  
    if not Assigned( aSo ) then
      aSo := FSoList[FSoList.IndexOf( aObj.CompCode )];
    if ( aObj.CompCode <> aSo.CompCode ) then
      aSo := FSoList[FSoList.IndexOf( aObj.CompCode )];
    if not CheckDbState( aSo ) then
    begin
      Inc( aIndex );
      Continue;
    end;
    aSo.Writer.SQL.Text := BuildWriteLogSql( aObj );
    try
      aSo.Writer.ExecSQL;
      LCmdLogList.FreeObject( aIndex );
      Inc( aDone );
      Dec( aCount );
    except
      on E: Exception do
      begin
        Inc( aIndex );
        aSo.DbState := dbError;
        FTextSubj.ErrorFmt( LanguageManager.Get( 'SDbWriteLogError' ),
          [aSo.CompName, aObj.FrameNo, aObj.LowCmd, E.Message] );
        FSoSubj.Notify( aSo );
      end;
    end;
    if ( aSo.ErrorCount >= 3 ) then
    begin
      DisconnectFromDb( aSo );
      FTextSubj.WarningFmt( LanguageManager.Get( 'SDbHasProblem' ),
        [aSo.CompName, CommEnv.DbRetryFreq] );
      FSoSubj.Notify( aSo );
    end;
  end;
  if ( aDone > 0 ) then
    FTextSubj.InfoFmt( LanguageManager.Get( 'SDbWriteLogCount' ), [aDone] );
  FLastWriteLog := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetData;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].Selected ) then
    begin
      BeginActive( FSoList[aIndex]  );
      try
        GetData( FSoList[aIndex] );
      finally
        EndActive( FSoList[aIndex] );
      end;
    end;  
    if ( Self.Terminated ) then Break;
  end;
  FGetDataTime := Now;
end;

{ ---------------------------------------------------------------------------- }
procedure TSoDbThread.GetData(const aSo: TSo);
var
  aRecordCount: Integer;
  aSqlOnlineCSR, aSqlBacthCSR, aGetSql: String;

  { -------------------------------------- }

  function OpenDataSet(const aSql: String): Integer;
  begin
    aSo.Reader.Close;
    aSo.Reader.SQL.Text := aSql;
    {$IFDEF APPDEBUG}
      CodeSite.Send( Format( 'CompCode=%s, SQL', [aSo.CompCode] ),
        aSo.Reader.SQL );
    {$ENDIF}
    aSo.Reader.Open;
    Result := aSo.Reader.RecordCount;
  end;

  { -------------------------------------- }

begin
  if not CheckDbState( aSo ) then Exit;
  { CommEnv.ProcessBatch = O  --> 只處理批次資料
    CommEnv.ProcessBatch = N  --> 不處理批次資料
    CommEnv.ProcessBatch = A  --> 處理批次資料及線上客服資料 }
  aGetSql := EmptyStr;
  if ( CommEnv.DbProcBatch <> 'O' ) then
  begin
    { 1.處理線上客服的資料 }
    aSqlOnlineCSR := OnlineCSRSql( aSo, CommEnv.DbProcRecords );
    aGetSql := aSqlOnlineCSR;
  end;
  if ( CommEnv.DbProcBatch <> 'N' ) then
  begin
    { 3. 處理批次 }
    aSqlBacthCSR := BacthCSRSql( aSo, CommEnv.DbProcRecords );
    if aGetSql = EmptyStr then
      aGetSql := aSqlBacthCSR
    else
      aGetSql := Format( ' %s union all %s ', [aGetSql, aSqlBacthCSR] );
  end;
  aGetSql := Format( ' select * from ( %s ) where rownum <= %d ',
    [aGetSql, CommEnv.DbProcRecords] );
  try
    aRecordCount := OpenDataSet( aGetSql );
    { 該系統台累計本次抓取的數量 }
    aSo.RecordCount := ( aSo.RecordCount + aRecordCount );
    if ( aRecordCount > 0 ) then
    begin
      WriteDataGet( aSo );
      if Self.Terminated then Exit;
      AddToBuffer( aSo );
      if Self.Terminated then Exit;
      BufferSort( aSo );
      if Self.Terminated then Exit;
      BufferToSendList( aSo );
    end;
  except
    on E: Exception do
    begin
      aSo.DbState := dbError;
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SDbGetDataError' ),
        [aSo.CompName, E.Message] );
      FSoSubj.Notify( aSo );
    end;
  end;
  if ( aSo.ErrorCount >= 3 ) then
  begin
    DisconnectFromDb( aSo );
    FTextSubj.WarningFmt( LanguageManager.Get( 'SDbHasProblem' ),
      [aSo.CompName, CommEnv.DbRetryFreq] );
    FSoSubj.Notify( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.BacthCSRSql(const aSo: TSo; const aLeftCount: Integer): String;
var
  aSql: String;
begin
  { 批次資料不用分 IPPV 資料或 SUB 資料, 一律全抓出來 }
  aSql := Format( SendDvnSqlHeader +
    '  where a.cmd_status = ''W''                        ' +
    '    and a.compcode = ''%s''                         ' +
    '    and a.processingdate <= sysdate ', [aSo.CompCode] );
  aSql := ( aSql + ' order by a.seqno ' );
  Result := Format( ' select * from ( %s ) where rownum <= %d ', [aSql, aLeftCount] );
end;

{ ---------------------------------------------------------------------------- }

{$WRITEABLECONST ON}

function TSoDbThread.GetIppvWhereCondition(const aType: String): String;
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
    HighCmdEnv.Filter := 'CMDTYPE=''PPV''';
    HighCmdEnv.Filtered := True;
    try
      HighCmdEnv.First;
      while not HighCmdEnv.Eof do
      begin
        Result := ( Result +
          Format( '''%s''', [HighCmdEnv.FieldByName( 'HighLevelCmd' ).AsString] ) );
        HighCmdEnv.Next;
        if ( not HighCmdEnv.Eof ) and ( Result <> EmptyStr ) then
          Result := Result + ',';
      end;
      if aType = 'O' then
        aConditionO := Format( ' in ( %s ) '#13, [Result] )
      else
        aConditionN := Format( ' not in ( %s ) '#13, [Result] );
    finally
      HighCmdEnv.Filtered := False;
      HighCmdEnv.Filter := EmptyStr;
    end;
  end;
  if aType = 'O' then
    Result :=  ' and ' + aConditionO
  else
    Result := aConditionN;
end;

{ ---------------------------------------------------------------------------- }

{$WRITEABLECONST OFF}

function TSoDbThread.OnlineCSRSql(const aSo: TSo; const aLeftCount: Integer): String;
var
  aSql: String;
begin
  aSql := Format( SendDvnSqlHeader +
    '  where a.cmd_status = ''W''          ' +
    '    and a.compcode = ''%s''           ' +
    '    and a.processingdate is null      ', [aSo.CompCode] );
  { 只處理 IPPV資料, 不處理 IPPV 資料 }
  if ( CommEnv.DbProcIPPV = 'O' ) or ( CommEnv.DbProcIPPV = 'N' ) then
  begin
    aSql := ( aSql +  ' and high_level_cmd_id ' +
      GetIppvWhereCondition( CommEnv.DbProcIPPV ) );
  end;
  aSql := ( aSql + ' order by a.seqno ' );
  Result := Format( ' select * from ( %s ) where rownum <= %d ',
    [aSql, aLeftCount] );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.AddToBuffer(const aSo: TSo);
var
  aPriority1, aPriority2: Integer;
begin
  if ( aSo.Buffer.IndexName <> EmptyStr ) then
    aSo.Buffer.DeleteIndex( BufferIndexName );
  aSo.Buffer.EmptyDataSet;
  aSo.Reader.First;
  while not aSo.Reader.Eof do
  begin
    { 若是 socket 錯誤, 已被標成錯誤, 則不須送指令, 不須加進 buffer }
    if ( aSo.Reader.FieldByName( 'cmd_status' ).AsString <> 'E' ) then
    begin
      aPriority1 := 999;
      aPriority2 := 999;
      { 判斷是否是批次, 批次指令優先權最低 }
      { 正常一般的指令 }
      if ( VarIsNull( aSo.Reader.FieldByName( 'processingdate' ).Value ) ) then
      begin
        if HighCmdEnv.Locate( 'highlevelcmd',
          aSo.Reader.FieldByName( 'high_level_cmd_id' ).AsString, [] ) then
        begin
          aPriority1 := HighCmdEnv.FieldByName( 'cmdtypepriority' ).AsInteger;
          aPriority2 := HighCmdEnv.FieldByName( 'cmdpriority' ).AsInteger;
        end;
      end;
      aSo.Buffer.AppendRecord( [
        aPriority1,
        aPriority2,
        aSo.Reader.FieldByName( 'high_level_cmd_id').Value,
        aSo.Reader.FieldByName( 'icc_no' ).Value,
        aSo.Reader.FieldByName( 'stb_no' ).Value,
        aSo.Reader.FieldByName( 'areacode' ).Value,
        aSo.Reader.FieldByName( 'notes' ).Value,
        aSo.Reader.FieldByName( 'cmd_status' ).Value,
        aSo.Reader.FieldByName( 'operator' ).Value,
        aSo.Reader.FieldByName( 'updtime' ).Value,
        aSo.Reader.FieldByName( 'startingtime' ).Value,
        aSo.Reader.FieldByName( 'expirytime' ).Value,
        aSo.Reader.FieldByName( 'product_id' ).Value,
        aSo.Reader.FieldByName( 'credit' ).Value,
        aSo.Reader.FieldByName( 'creditaction' ).Value,
        aSo.Reader.FieldByName( 'creditcfcode' ).Value,
        aSo.Reader.FieldByName( 'banktype' ).Value,
        aSo.Reader.FieldByName( 'producttype').Value,
        aSo.Reader.FieldByName( 'configvalue' ).Value,
        aSo.Reader.FieldByName( 'mailaction' ).Value,
        aSo.Reader.FieldByName( 'compcode' ).Value,
        aSo.Reader.FieldByName( 'sno' ).Value,
        aSo.Reader.FieldByName( 'processingdate' ).Value,
        aSo.Reader.FieldByName( 'seqno' ).Value] );
    end;
    aSo.Reader.Next;
    if ( Self.Terminated ) then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.BufferSort(const aSo: TSo);
begin
  if ( aSo.Buffer.IndexName <> EmptyStr ) then
    aSo.Buffer.DeleteIndex( BufferIndexName );
  aSo.Buffer.AddIndex( BufferIndexName, 'PRIORITY1;PRIORITY2;SEQNO', [] );
  aSo.Buffer.IndexName := BufferIndexName;
{$IFDEF APPDEBUG}
  aSo.Buffer.First;
  while not aSo.Buffer.Eof do
  begin
     CodeSite.SendFmtMsg( 'HighCmd=%s, SeqNo=%s', [
      aSo.Reader.FieldByName( 'high_level_cmd_id').AsString,
      aSo.Reader.FieldByName( 'seqno' ).AsString] );
    aSo.Buffer.Next;
  end;
{$ENDIF}
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.BufferToSendList(const aSo: TSo);
var
  aObj: TSendDvn;
begin
  aSo.Buffer.First;
  while not aSo.Buffer.Eof do
  begin
    if ( Self.Terminated ) then Break;
    aObj := TSendDvn.Create( ctHigh );
    aObj.CompCode := aSo.Buffer.FieldByName( 'compcode' ).AsString;
    aObj.CompName := aSo.CompName;
    {}
    { 高階指令的 UniqueId = SeqNo, 低階指令的 UniqueId = FrameNo }
    aObj.HighSeqNo := aSo.Buffer.FieldByName( 'seqno' ).AsString;
    aObj.UniqueId := aObj.HighSeqNo;
    {}
    aObj.HighCmd := aSo.Buffer.FieldByName( 'high_level_cmd_id' ).AsString;
    aObj.IccNo := aSo.Buffer.FieldByName( 'icc_no' ).AsString;
    aObj.StbNo := aSo.Buffer.FieldByName( 'stb_no' ).AsString;
    aObj.AreaCode := Nvl( aSo.Buffer.FieldByName( 'areacode' ).AsString, aSo.AreaCode );
    aObj.Notes := aSo.Buffer.FieldByName( 'notes' ).AsString;
    aObj.Operator := aSo.Buffer.FieldByName( 'operator' ).AsString;
    aObj.UpdTime := Now;
    if ( not VarIsNull( aSo.Buffer.FieldByName( 'updtime' ).Value ) ) then
      aObj.UpdTime := aSo.Buffer.FieldByName( 'updtime' ).AsDateTime;
    if ( not VarIsNull( aSo.Buffer.FieldByName( 'startingtime' ).Value )  ) then
      aObj.StartingTime := FormatDateTime( 'yyyy/mm/dd', aSo.Buffer.FieldByName( 'startingtime' ).Value ) + ' 00:00:00';
    if ( not VarIsNull( aSo.Buffer.FieldByName( 'expirytime' ).Value )  ) then
      aObj.ExpiryTime := FormatDateTime( 'yyyy/mm/dd', aSo.Buffer.FieldByName( 'expirytime' ).Value ) + ' 23:59:59';
    aObj.ProductId := aSo.Buffer.FieldByName( 'product_id' ).AsString;
    aObj.Credit := 0;
    if not VarIsNull( aSo.Buffer.FieldByName( 'credit' ).Value ) then
      aObj.Credit := aSo.Buffer.FieldByName( 'credit' ).AsFloat;
    aObj.CreditAction := aSo.Buffer.FieldByName( 'creditaction' ).AsString;
    aObj.CreditCfCode := aSo.Buffer.FieldByName( 'creditcfcode' ).AsString;
    aObj.BankType := aSo.Buffer.FieldByName( 'banktype' ).AsString;
    aObj.ProductType := aSo.Buffer.FieldByName( 'producttype' ).AsString;
    aObj.ConfigValue := aSo.Buffer.FieldByName( 'configvalue' ).AsString;
    aObj.MailAction := aSo.Buffer.FieldByName( 'mailaction' ).AsString;
    aObj.SNo := aSo.Buffer.FieldByName( 'sno' ).AsString;
    if ( not VarIsNull( aSo.Buffer.FieldByName( 'processingdate' ).Value ) ) then
      aObj.ProcessingDate := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aSo.Buffer.FieldByName( 'processingdate' ).AsDateTime );
    aObj.IsBatch := ( aObj.ProcessingDate <> EmptyStr );
    aObj.HighCmdStatus := 'P';
    FCmdSubj.Notify( aObj );
    Sleep( 50 );
    HCmdBeginList.AddObject( aObj.HighSeqNo, aObj );
    aSo.Buffer.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
