unit cbFeedbackDbThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, CommCtrl, DB, ADODB, DBClient,
  cbClass;

type
  TFeedbackDbThread = class(TSMSCommandThread)
  private
    { Private declarations }
    FSoInfo: PSoInfo;
    FADOControl: PSoADOControl;
    FUpdateGUI: Boolean;
    procedure PrepareADOControl;
    procedure ReleaseADOResource;
    procedure ConnectToDb;
    procedure DisconnectFromDb;
    function CanDbRertyConnect: Boolean;
    function CheckDbStatus: Boolean;
    function GetWriteFeedbackSQL(aObj: PRecvNagra): String;
    function GetWriteLogSQL(aObj: PRecvNagra): String;
  protected
    function GetWaitWhileFrquence: Cardinal; override;
    procedure Execute; override;
    procedure BeginActive; override;
    procedure EndActive; override;
    procedure Update; override;
    procedure UpdateEx; 
    procedure PhysicalWriteFeedbackData;
  public
    constructor Create(const aSoInfo: PSoInfo); overload;
    destructor Destroy; override;
    property UpdateGUI: Boolean read FUpdateGUI write FUpdateGUI;
  end;

implementation

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TFeedbackDbThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }


uses DateUtils, cbMain, cbResStr, cbUtilis;    

{ TFeedbackDbThread }

constructor TFeedbackDbThread.Create(const aSoInfo: PSoInfo);
begin
  inherited Create;
  FSoInfo := aSoInfo;
end;

{ ---------------------------------------------------------------------------- }

destructor TFeedbackDbThread.Destroy;
begin
  FSoInfo := nil;
  FADOControl := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.BeginActive;
begin
  if FSoInfo.DbConnectStatus in [dbOK, dbActive] then
    FSoInfo.DbConnectStatus := dbActive;
  Synchronize( Update );
  Sleep( GetWaitWhileFrquence );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.EndActive;
begin
  if FSoInfo.DbConnectStatus in [dbOK, dbActive] then
    FSoInfo.DbConnectStatus := dbOK;
  Synchronize( Update );
  Sleep( GetWaitWhileFrquence );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.Execute;
begin
  PrepareADOControl;
  ConnectToDb;
  { ???? Main Thread ???X???????T?? }
  WaitForPlaySignal;
  Sleep( GetWaitWhileFrquence );
  while ( not Stop ) do
  begin
    MessageSubject.RunState := rsRunning;
    Sleep( GetWaitWhileFrquence );
    if Stop then Break;
    try
      if FUpdateGUI then BeginActive;
      try
        PhysicalWriteFeedbackData;
      finally
        if FUpdateGUI then EndActive;
      end;
      if Stop then Break;
      Sleep( GetWaitWhileFrquence * 10 );
    except
      { ... }
    end;
  end;
  Sleep( GetWaitWhileFrquence );
  MessageSubject.RunState := rsStop;
  { ?w?????????????T?? }
  WaitForTerminalSignal;
  DisconnectFromDb;
  ReleaseADOResource;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackDbThread.GetWaitWhileFrquence: Cardinal;
begin
  Result := inherited GetWaitWhileFrquence;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.Update;
begin
  fmMain.UpdateSoTreeStatus( FSoInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.UpdateEx;
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.ReleaseADOResource;
begin
  FADOControl.DataWriter.Free;
  FADOControl.DataReader.Free;
  FADOControl.Connection.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.PrepareADOControl;
begin
  New( FADOControl );
  FADOControl.CompCode := FSoInfo.CompCode;
  FADOControl.Connection := TADOConnection.Create( nil );
  FADOControl.DataReader := TADOQuery.Create( nil );
  FADOControl.DataWriter := TADOQuery.Create( nil );
  FADOControl.DataSet := nil;
  FADOControl.Connection.LoginPrompt := False;
  FADOControl.DataReader.CacheSize := 1000;
  FADOControl.DataReader.Connection := FADOControl.Connection;
  FADOControl.DataReader.LockType := ltBatchOptimistic;
  FADOControl.DataWriter.CacheSize := 1000;
  FADOControl.DataWriter.Connection := FADOControl.Connection;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.ConnectToDb;
begin
  if FADOControl.Connection.Connected then
    FADOControl.Connection.Connected := False;
  FADOControl.Connection.ConnectionString := Format(
    'Provider=MSDAORA.1;Persist Security Info=True;' +
    'Password=%s;User ID=%s;Data Source=%s;',
   [FSoInfo.LoginPass, FSoInfo.LoginUser, FSoInfo.DbAliase] );
  try
    FADOControl.Connection.Connected := True;
    FSoInfo.DbConnectStatus := dbOK;
    MessageSubject.OK( SFeedbackDbConnectSuccess );
    FSoInfo.LastCriticalErrorTime := 0;
    FSoInfo.CriticalErrorCount := 0;
  except
    on E: Exception do
    begin
      FSoInfo.LastCriticalErrorTime := Now;
      FSoInfo.DbConnectStatus := dbError;
      MessageSubject.Error( Format( SFeedbackDbConnectError,
        [E.Message, CommEnv.DbRetryFrequence] ) );
    end;
  end;
  { ?q???D?????? }
  PostMessage( MainFormHandle, WM_DATABASE, Ord( CommandSubject.ThreadType ),
    Ord( FADOControl.Connection.Connected ) );
  Synchronize( Update );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.DisconnectFromDb;
begin
  try
    if FADOControl.Connection.Connected then
      FADOControl.Connection.Connected := False;
    FSoInfo.DbConnectStatus := dbNone;
    FSoInfo.RecordCount := 0;
    MessageSubject.OK( SFeedbackDbDisConnectSuccess );
  except
    on E: Exception do
      MessageSubject.Error( Format( SFeedbackDbDisConnectError, [E.Message] ) );
  end;
  Synchronize( Update );
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackDbThread.CanDbRertyConnect: Boolean;
var
  //aRetry: Cardinal;
  aSeocnd: Integer;
begin
//  aRetry := GetBetweenFrquence( FSoInfo.LastCriticalErrorTickCount );
//  if aRetry <= 0 then aRetry := ( CommEnv.DbRetryFrequence * 1000 );
//  Result := ( aRetry >= ( CommEnv.DbRetryFrequence * 1000 ) );
  if ( FSoInfo.LastCriticalErrorTime <= 0 ) then
    aSeocnd := ( CommEnv.DbRetryFrequence * 1000 )
  else
    aSeocnd := SecondsBetween( Now, FSoInfo.LastCriticalErrorTime );
  Result := ( aSeocnd >= ( CommEnv.DbRetryFrequence * 1000 ) );
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackDbThread.CheckDbStatus: Boolean;
begin
  if ( not FADOControl.Connection.Connected ) and CanDbRertyConnect then
  begin
    MessageSubject.Info( SFeedbackDbRertyConnect );
    ConnectToDb;
  end;
  Result := FADOControl.Connection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackDbThread.PhysicalWriteFeedbackData;
var
  aObj: PRecvNagra;
  aImmediateError, aIsErrorNotify: Boolean;
  aWriteDone, aWriteError: Integer;
begin
  if not CheckDbStatus then Exit;
  if FeedbackRecvList.Count = 0 then Exit;
  aWriteDone := 0;
  aWriteError := 0;
  aIsErrorNotify := False;
  while ( FeedbackRecvList.Count > 0 ) do
  begin
    aImmediateError := False;
    aObj := PRecvNagra( FeedbackRecvList.Objects[0] );
    FADOControl.DataWriter.Close;
    try
      FADOControl.DataWriter.SQL.Text := GetWriteFeedbackSQL( aObj );
      FADOControl.DataWriter.ExecSQL;
      FADOControl.DataWriter.SQL.Text := GetWriteLogSQL( aObj );
      FADOControl.DataWriter.ExecSQL;
      Inc( aWriteDone );
    except
      on E: Exception do
      begin
        aImmediateError := True;
        Inc( aWriteError );
        FSoInfo.CriticalErrorCount := ( FSoInfo.CriticalErrorCount + 1 );
        if not aIsErrorNotify then
        begin
          MessageSubject.Error( Format( SFeedbackDbWriteDataError, [E.Message] ) );
          FSoInfo.DbConnectStatus := dbError;
          Synchronize( Update );
        end;
        aIsErrorNotify := True;
      end;
    end;
    { ACK ?? NACK }
    aObj.CmdStatus := 'C';
    if aImmediateError then
    begin
      aObj.SendStatus := 'E';
      aObj.CmdStatus := 'E';
    end;
    { ?N?????????n ACK ?? NAGRA }
    FeedbackSendList.BeginWrite;
    try
      FeedbackSendList.AddObject( aObj.ResponseTransactionNumber, TObject( aObj ) );
    finally
      FeedbackSendList.EndWrite;
    end;
    { ?N???????q???? List ?????? }
    FeedbackRecvList.BeginWrite;
    try
      FeedbackRecvList.Delete( 0 );
    finally
      FeedbackRecvList.EndWrite;
    end;
    Sleep( 100 ); 
  end;
  if ( FSoInfo.CriticalErrorCount > 0 ) then
  begin
    FSoInfo.CriticalErrorCount := 0;
    FSoInfo.LastCriticalErrorTime := Now;
    MessageSubject.Warning( Format( SFeedbackDbHasProblem, [CommEnv.DbRetryFrequence] ) );
    DisconnectFromDb;
  end;
  FSoInfo.RecordCount := ( FSoInfo.RecordCount + aWriteDone );
  if ( aWriteDone > 0 ) then
    MessageSubject.Info( Format( SFeedbackDbWriteDataDone, [aWriteDone] ) );
  if ( aWriteError > 0 ) then
    MessageSubject.Warning( Format( SFeedbackDbWriteDataError2, [aWriteError] ) );
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackDbThread.GetWriteFeedbackSQL(aObj: PRecvNagra): String;
var
  aSQL: String;
begin
  aSQL := Format(
    ' INSERT INTO RECV_NAGRA       ' +
    '      ( HIGH_LEVEL_CMD_ID,    ' +
    '        LOW_LEVEL_CMD_ID,     ' +
    '        TRANSACTIONNUM,       ' +
    '        ICC_NO,               ' +
    '        STB_NO,               ' +
    '        CALLBACK_DATE,        ' +
    '        CALLBACK_TIME,        ' +
    '        CREDIT,               ' +
    '        DEBIT,                ' +
    '        IMS_PRODUCT_ID,       ' +
    '        PURCHASE_DATE,        ' +
    '        WATCH_STATUS,         ' +
    '        PRODUCT_SUSPENDED,    ' +
    '        ICC_SUSPENDED,        ' +
    '        PHONE_NUMBER_1,       ' +
    '        PHONE_NUMBER_2,       ' +
    '        PHONE_NUMBER_3,       ' +
    '        ABNORMAL_PHONE,       ' +
    '        STB_RESPONDING,       ' +
    '        CMD_STATUS,           ' +
    '        UPDTIME  )            ' +
    '  VALUES                      ' +
    '      ( ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''%s'',               ' +
    '        ''W'',                ' +
    '        SYSDATE  )            ',
    [aObj.HighLevelCmd, aObj.LowLevelCmd, aObj.ResponseTransactionNumber,
     aObj.IccNo, aObj.StbNo, aObj.CallbackDate, aObj.CallbackTime,
     aObj.Credit, aObj.Debit, aObj.ImdProductId, aObj.PurchaseDate,
     aObj.WatchStatus, aObj.ProductSuspended, aObj.IccSuspended,
     aObj.PhoneNum1, aObj.PhoneNum2, aObj.PhoneNum3, aObj.AbnormalPhone,
     aObj.StbResponding] );
  Result := aSQL;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackDbThread.GetWriteLogSQL(aObj: PRecvNagra): String;
var
  aSQL: String;
begin
  aSQL := Format(
    ' INSERT INTO CACALLBACKDATA   ' +
    '      ( TRANSACTIONNUM,       ' +
    '        CALLBACKDATA,         ' +
    '        UPDTIME )             ' +
    '  VALUES                      ' +
    '      ( ''%s'',               ' +
    '        ''%s'',               ' +
    '        SYSDATE  )            ',
    [aObj.ResponseTransactionNumber, aObj.ResponseCommandText] );
  Result := aSQL;  
end;

{ ---------------------------------------------------------------------------- }

end.
