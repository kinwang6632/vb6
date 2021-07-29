unit cbFeedbackThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, DBClient,
  cbClass, cbNagraCmd, cbNagraDevice,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;

type

  { 回傳機制傳送回應執行緒 }

  TFeedbackSendThread = class(TSMSSocketThread)
  private
    { Private declarations }
    FSendBuffer: PDeviceBuffer;
    FSendCommand: Boolean;
    FCommandBuilder: TNagraCommandBuilder;
    FSendError: Integer;
    FLastErrorTime: TDateTime;
    FLastCommunication: TDateTime;
    FSignal: TNotifyCommandType;
    FWMSocket: TWMSocket;
    function CanSocketRertyConnect: Boolean;
    function WaitForSendCommandSignal: Boolean;
    function FeedbackSend: Integer;
    function CanCommunicationCheck: Boolean;
    procedure EnsureCommunicationCheck(aCount: Integer);
    procedure NotifyCommandSendError;
    procedure NotifyCommunicationCheckError;
    procedure InitWMSocket(const aNotifyCommandType: TNotifyCommandType;
      const aErrorCode: Integer = -1; aErrorMsg: Integer = -1);
    procedure NotifyToThread(aThreadHandle: Integer; aType: TNotifyCommandType);  
  protected
    procedure SetCASEnv(const aValue: TCASEnv); override;
    procedure SetHighCmdEnv(aValue: TClientDataSet); override;
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
    procedure Update; override;
    function EndConnection: Boolean; override;
  public
    constructor Create(const aSocket: TIdTCPClient);
    destructor Destroy; override;
    function BeginConnection: Boolean; override;
    function CommunicationCheck: Boolean;
    property CanSendCommand: Boolean read FSendCommand;
  end;

  { 接收回傳機制資料執行緒 }

  TFeedbackRecvThread = class(TSMSSocketThread)
  private
    { Private declarations }
    FRecvBuffer: PDeviceBuffer;
    FRecvCommand: Boolean;
    FCommunicationCheckCounter: Integer;
    FCmdError: TClientDataSet;
    FWMSocket: TWMSocket;
    procedure FeedbackRecv;
    procedure ProcessFeedbackRecv;
    procedure SetObjectErrorCode(aObj: PRecvNagra);
    function ExtractBuffer(const aBuffer; const aBytes: Integer): String;
    function AddToList(aFullText: String): Integer;
    procedure InitWMSocket(const aNotifyCommandType: TNotifyCommandType;
      const aErrorCode: Integer = -1; aErrorMsg: Integer = -1);
    procedure NotifyToThread(aThreadHandle: Integer; aType: TNotifyCommandType);
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
    procedure Update; override;
  public
    constructor Create(const aSocket: TIdTCPClient);
    destructor Destroy; override;
    property CanRecvCommand: Boolean read FRecvCommand;
  end;


implementation


uses cbMain, cbResStr, cbUtilis, DateUtils;


var
  FeedbackSendThread: TFeedbackSendThread;
  FeedbackRecvThread: TFeedbackRecvThread;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TFeedbackThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }



{ TFeedbackSendThread }

constructor TFeedbackSendThread.Create(const aSocket: TIdTCPClient);
begin
  inherited Create( aSocket );
  FeedbackSendThread := Self;
  CommandSubject.ThreadType := tdFeedbackSend;
  New( FSendBuffer );
  FCommandBuilder := TNagraCommandBuilder.Create;
  FSendCommand := False;
  FSendError := 0;
  FLastErrorTime := 0;
  FLastCommunication := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TFeedbackSendThread.Destroy;
begin
  Dispose( FSendBuffer );
  FCommandBuilder.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.BeginConnection: Boolean;
begin
  Result := False;
  MessageSubject.Info( Format( SSocketBeginConnection,
    [SFeedbackSend, CASEnv.IP, CASEnv.FeedbackPort] ) );
  try
    Result := inherited BeginConnection;
  except
    on E: Exception do
      MessageSubject.Error( Format( SSocketConnectionError, [SFeedbackSend, E.Message] ) );
  end;
  { 仍然有錯, 記錄最後一次錯誤時間, 用來比對重試連線用 }
  FLastErrorTime := Now;
  { 連線成功 }
  if Result then
  begin
    { 錯誤計數及最後錯誤時間歸零 }
    FSendError := 0;
    FLastErrorTime := 0;
  end;
  FSendCommand := False;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.CommunicationCheck: Boolean;
var
  aObj: PRecvNagra;
  aCmdText: String;
  aIndex: Integer;
begin
  New( aObj );
  aObj.HighLevelCmd := 'Z1';
  aObj.LowLevelCmd := '1002';
  aObj.CompCode := '99';
  aObj.CmdStatus := 'W';
  aObj.SendStatus := 'W';
  aObj.Data := nil;
  aObj.ParentData := nil;
  aObj.Index := -1;
  aObj.ParentIndex := -1;
  aCmdText := FCommandBuilder.BuildCommunicationCheck( aObj );
  { 先 Store 進 List 裏, 發生錯誤的話, 再從 List 裏刪除 }
  FeedbackCommunicationCheck.BeginWrite;
  try
    FeedbackCommunicationCheck.AddObject( aObj.TransactionNumber,
      TObject( aObj ) );
  finally
    FeedbackCommunicationCheck.EndWrite;
  end;
  { 顯示 CommunicationCheck 指令 }
  aObj.SendStatus := 'P';
  aObj.CmdStatus := 'P';
  UpdatedObject := aObj;
  Synchronize( Update );
  try
    GateWaySocket.Write( aCmdText );
    Result := True;
  except
    on E: Exception do
    begin
      Result := False;
      Inc( FSendError );
      aObj.SendStatus := 'E';
      aObj.CmdStatus := 'E';
      aObj.ErrCode := '-1000';
      aObj.ErrMsg := Copy( E.Message, 1, 10 );
      MessageSubject.Error( Format( SSendCommandError, [SFeedbackSend,
        aObj.HighLevelCmd, aObj.LowLevelCmd, E.Message] ) );
    end;
  end;
  if not Result then
  begin
    FeedbackCommunicationCheck.BeginWrite;
    try
      aIndex := FeedbackCommunicationCheck.IndexOf( aObj.TransactionNumber );
      if aIndex >= 0 then
        FeedbackCommunicationCheck.Delete( aIndex );
      Dispose( aObj );
    finally
       FeedbackCommunicationCheck.EndWrite;
    end;
  end;
  FLastCommunication := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_SOCKET:
      FSignal := TNotifyCommandType( TWMSocket( Msg ).NotifyCommandType );
  else
    inherited WndProc( Msg );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.Execute;
begin
  { 等待 Main Thread 打出執行的訊號 }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
    { 須要連線, 且已達到重試連線時間, 第一次進來一定是【已達重試連線時間】 }
    if ( not GateWaySocket.Connected ) and ( CanSocketRertyConnect ) then
    begin
      { 通知接收指令回應執行緒, 先停止接收資料, 不打出此信號,
        會跟 ControlRecv 搶接收資料, 導致 Dead Lock }
      NotifyToThread( FeedbackRecvThread.WndHandle, ncNack );
      Sleep( 300 );
      { 重試連線 }
      if not BeginConnection then
      begin
        { 仍然連不上, 顯示重試訊息 }
        MessageSubject.Warning( Format( SWaitRetry, [SFeedbackSend, CommEnv.CASRetryFrequence] ) );
        EndConnection;
        NotifyToThread( MainFormHandle, ncSocketError );
      end;
    end;
    if Stop then Break;
    if ( GateWaySocket.Connected ) and ( not FSendCommand ) then
    begin
      { 連上了, 但是不見得可以傳送指令, 先送出 Communication Check,
        確認是否可以丟指令 }
      { 通知接收指令回應執行緒, 接收資料, 不打出此信號會導致無法接收 ACK }
      NotifyToThread( FeedbackRecvThread.WndHandle, ncCommunication );
      Sleep( 300 );
      FSignal := ncCommunication;
      MessageSubject.Info( Format( SEnsureCandSendCommand, [SFeedbackSend] ) );
      EnsureCommunicationCheck( 3 );
      FSendCommand := WaitForSendCommandSignal;
      if ( FSendCommand ) then
      begin
        MessageSubject.OK( Format( SEnsureOK, [SFeedbackSend, SFeedbackSend] ) );
        { 通知主執行緒 SOCKET 連線正常, 並且可以接收回傳資料 }
        NotifyToThread( MainFormHandle, ncAck );
      end else
      begin
        MessageSubject.Error( Format( SEnsureError, [SFeedbackSend, SFeedbackSend] ) );
        MessageSubject.Warning( Format( SWaitRetry, [SFeedbackSend, CommEnv.CASRetryFrequence] ) );
        EndConnection;
        { 通知接收指令回應執行緒, 停止接收資料 }
        NotifyToThread( FeedbackRecvThread.WndHandle, ncNack );
        { 通知主執行緒 NAGRA 拒絕回傳機制 }
        NotifyToThread( MainFormHandle, ncNack );
      end;
    end;
    if Stop then Break;
    if ( FSendCommand ) then
    begin
      { 傳送回應 }
      FeedbackSend;
      { 錯誤次數大於設定值 }
      if ( FSendError >= CommEnv.CASSendErrMax ) then
      begin
        NotifyCommandSendError;
        EndConnection;
        { 通知接收回傳資料執行緒, 停止接收資料 }
        NotifyToThread( FeedbackRecvThread.WndHandle, ncSocketError );
        { 通知主執行緒 SOCKET 連線有誤 }
        NotifyToThread( MainFormHandle, ncSocketError );
      end;
    end;
    if Stop then Break;
    if ( FSendCommand ) then
    begin
      if ( CanCommunicationCheck ) then
        CommunicationCheck;
      { 錯誤次數大於設定值或是 Communication Check Nack 回來  }
      if ( FSendError >= CommEnv.CASSendErrMax ) or
         ( FSignal in [ncNack, ncSocketError] ) then
      begin
        if ( FSendError >= CommEnv.CASSendErrMax ) then
          NotifyCommandSendError
        else
          NotifyCommunicationCheckError;
        EndConnection;
        { 通知接收指令回應執行緒, 停止接收資料 }
        NotifyToThread( FeedbackRecvThread.WndHandle, ncNack );
        { 通知主執行緒 SOCKET 連線有誤 }
        NotifyToThread( MainFormHandle, ncNack );
      end;
    end;
    if Stop then Break;
    Sleep( GetWaitWhileFrquence * 3 );
    if Stop then Break;
  end;
  Sleep( GetWaitWhileFrquence );
  WaitForTerminalSignal;
  EndConnection;
  MessageSubject.Info( Format( SRunStop, [SFeedbackRecv] ) );
  MessageSubject.RunState := rsStop;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.CanSocketRertyConnect: Boolean;
var
  aSecond: Int64;
begin
  if ( FLastErrorTime <= 0 ) then
  begin
    FLastErrorTime := Now;
    aSecond := ( CommEnv.CASRetryFrequence );
  end else
  begin
    aSecond := SecondsBetween( Now, FLastErrorTime );
  end;  
  Result := ( aSecond >= ( CommEnv.CASRetryFrequence ) );
//  if FLastErrorTime <= 0 then
//    FLastErrorTime := Now;
//  aRetry := GetBetweenFrquence( FLastErrorTime );
//  if aRetry <= 0 then aRetry := ( CommEnv.CASRetryFrequence * 1000 );
//  Result := ( aRetry >= ( CommEnv.CASRetryFrequence * 1000 ) );
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.WaitForSendCommandSignal: Boolean;
var
  aTickCount: Cardinal;
  aTimeout: Boolean;
begin
  aTimeout := False;
  aTickCount := GetTickCount;
  while True do
  begin
    Sleep( 100 );
    if ( FSignal in [ncAck,ncNack] ) then
      Break;
    if Stop then Break;
    if ( GetTickCount - aTickCount ) <= 0 then
      aTickCount := GetTickCount
    else begin
      { 超過 10 秒仍然等不到 OK 信號 }
      aTimeout := ( ( GetTickCount - aTickCount ) div 1000 ) >= 10;
      if aTimeout then Break;
    end;
  end;
  if aTimeout then
    FSignal := ncSocketError
  else begin
    if FSignal <> ncAck then FSignal := ncNack;
  end;
  Result := ( FSignal = ncAck );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.EnsureCommunicationCheck(aCount: Integer);
var
  aIndex: Integer;
begin
  { 一連上 Nagra, 必須先送 n 次 Communication Check Command }
  { 連續 n 個 Ack 回來都是 OK, 才可以丟 Command 給 Nagra }
  for aIndex := 1 to aCount do
  begin
    CommunicationCheck;
    { 一連上 Nagra, 每1秒送一次 Communication Check Command }
    Sleep( 1000 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.FeedbackSend: Integer;
var
  aObj: PRecvNagra;

  { ------------------------------------------------ }

  procedure UpdateCommandStatus(aCmd: PRecvNagra);
  begin
    UpdatedObject := aCmd;
    Synchronize( Update );
    Sleep( 100 );
  end;

  { ------------------------------------------------ }

var
  aText, aErrorCode, aErrorMsg: String;
  aImmediateError: Boolean;
begin
  Result := 0;
  { TCP/IP 傳送錯誤 Flag }
  aImmediateError := False;
  while ( FeedbackSendList.Count > 0 ) and
        ( Result <= CommEnv.ProcessRecordCount ) do
  begin
    aObj := PRecvNagra( FeedbackSendList.Objects[0] );
    FCommandBuilder.BuildCommand( aObj );
    FCommandBuilder.ClearCommand;
    Inc( Result );
    aText := aObj.FullCommandText;
    try
      { 沒有錯誤發生, 則傳送指令 }
      if not aImmediateError then
      begin
        ZeroMemory( FSendBuffer, SizeOf( FSendBuffer^ ) );
        Move( PChar( aText )^, FSendBuffer^, Length( aText ) );
        GateWaySocket.WriteBuffer( FSendBuffer^, Length( aText ), True );
        { 回應正常 }
        aObj.SendStatus := 'C';
      end;
    except
      on E: Exception do
      begin
        Inc( FSendError );
        aImmediateError := True;
        aErrorCode := '-1000';
        aErrorMsg := Copy( E.Message, 1, 40 );
        MessageSubject.Error( Format( SSendCommandError,
          [SFeedbackSend, aObj.HighLevelCmd, aObj.LowLevelCmd, E.Message] ) );
      end;
    end;
    { 該ACK指令錯誤 }
    if aImmediateError then
    begin
      { 將高階指令及低階標示成錯誤 }
      aObj.SendStatus := 'E';
      aObj.ErrCode := aErrorCode;
      aObj.ErrMsg := aErrorMsg;
    end;
    { 更新該筆指令 ACK 狀態 }
    UpdateCommandStatus( aObj );
    FeedbackSendList.BeginWrite;
    try
      Dispose( aObj );
      FeedbackSendList.Delete( 0 );
    finally
      FeedbackSendList.EndWrite;
    end;
    if not aImmediateError then
    begin
      { 延遲 }
      Sleep( Trunc( CmdTransEnv.SendCommandDelay * 1000 ) );
    end;  
  end;
  if ( Result > 0 ) then
    MessageSubject.Info( Format( SFeedbackSendRecord, [SFeedbackSend, Result] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.NotifyCommandSendError;
begin
  MessageSubject.Error( Format( SSendErrors, [SFeedbackSend, CommEnv.CASSendErrMax] ) );
  MessageSubject.Warning( Format( SWaitRetry, [SFeedbackSend, CommEnv.CASRetryFrequence] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.NotifyCommunicationCheckError;
begin
  MessageSubject.Error( Format( SCommunicationCheckError,
    [CASEnv.IP, CASEnv.ControlPort, SFeedbackSend, SFeedbackSend] ) );
  MessageSubject.Warning( Format( SWaitRetry, [SFeedbackSend, CommEnv.CASRetryFrequence] ) );
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.CanCommunicationCheck: Boolean;
begin
  if ( FLastCommunication <= 0 ) then
    FLastCommunication := Now;
  Result := ( SecondsBetween( Now, FLastCommunication ) >= CommEnv.CASCommCheck );
//  aRetry := GetBetweenFrquence( FLastCommunicationTickCount );
//  if aRetry <= 0 then aRetry := ( CommEnv.CASCommCheck * 1000 );
//  Result := ( aRetry >= ( CommEnv.CASCommCheck * 1000 ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.SetCASEnv(const aValue: TCASEnv);
begin
  inherited SetCASEnv( aValue );
  GateWaySocket.Host := CASEnv.IP;
  GateWaySocket.Port := CASEnv.FeedbackPort;
  FCommandBuilder.CASEnv := aValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.SetHighCmdEnv(aValue: TClientDataSet);
begin
  inherited SetHighCmdEnv( aValue );
  FCommandBuilder.HighCmdEnv := aValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.Update;
begin
  CommandSubject.UpdateCommandObject := UpdatedObject;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackSendThread.EndConnection: Boolean;
begin
  Result := inherited EndConnection;
  FSendCommand := False;
  FLastErrorTime := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.InitWMSocket(
  const aNotifyCommandType: TNotifyCommandType; const aErrorCode: Integer;
  aErrorMsg: Integer);
begin
  FWMSocket.Msg := WM_SOCKET;
  FWMSocket.ThreadType := Ord( Self.CommandSubject.ThreadType );
  FWMSocket.NotifyCommandType := Ord( aNotifyCommandType );
  FWMSocket.ErrorCode := aErrorCode;
  FWMSocket.ErrorMsg := aErrorMsg;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackSendThread.NotifyToThread(aThreadHandle: Integer;
  aType: TNotifyCommandType);
begin
  if aType = ncNack then
    InitWMSocket( aType, -1001, -1001 )
  else if aType = ncSocketError then
    InitWMSocket( aType, -1001, -1002 )
  else
    InitWMSocket( aType );
  PostMessage( aThreadHandle, WM_SOCKET,
    TMessage( FWMSocket ).WParam, TMessage( FWMSocket ).LParam );
end;

{ ---------------------------------------------------------------------------- }

{ TFeedbackRecvThread }

constructor TFeedbackRecvThread.Create(const aSocket: TIdTCPClient);
begin
  inherited Create( aSocket );
  FeedbackRecvThread := Self;
  CommandSubject.ThreadType := tdFeedbackRecv;
  New( FRecvBuffer );
  FCmdError := TClientDataSet.Create( nil );
  FRecvCommand := False;
  FCommunicationCheckCounter := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TFeedbackRecvThread.Destroy;
begin
  Dispose( FRecvBuffer );
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.Execute;
begin
  { 等待 Main Thread 打出執行的訊號 }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
    if ( GateWaySocket.Connected and FRecvCommand ) then
      FeedbackRecv;
    if Stop then Break;
    ProcessFeedbackRecv;
    if Stop then Break;
    Sleep( GetWaitWhileFrquence );
  end;
  Sleep( GetWaitWhileFrquence );
  MessageSubject.RunState := rsStop;
  WaitForTerminalSignal;
  EndConnection;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.FeedbackRecv;
var
  aByte: Integer;
begin
  aByte := -1;
  try
    aByte := GateWaySocket.ReadFromStack( False, -1, False );
  except
    { ... }
  end;
  if aByte > 0 then
  begin
    ZeroMemory( FRecvBuffer, SizeOf( FRecvBuffer^ ) );
    Move( GateWaySocket.InputBuffer.Memory^, FRecvBuffer^, aByte );
    GateWaySocket.InputBuffer.Remove( aByte );
    AddToList( ExtractBuffer( FRecvBuffer, aByte ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackRecvThread.ExtractBuffer(const aBuffer; const aBytes: Integer): String;
var
  aDoubleChr, aFullText: String;
  aPos, aLength: Integer;
begin
  { Buffer 內容, 頭 2 個 Byte 內容為長度 }
  { Buffer 內容, 會有 1 筆以上資料須要拆解, 大部份時候只有一筆 }
  Result := EmptyStr;
  aPos := 0;
  while ( aPos < aBytes ) do
  begin
    aDoubleChr := PChar( aBuffer )[aPos] + PChar( aBuffer )[aPos+1];
    aLength := HexToLength( aDoubleChr );
    { 跳過起頭 2 個 Byte  }
    SetString( aFullText, PChar( aBuffer ) + ( aPos + 2 ), aLength );
    Result := ( Result + aFullText );
    aPos := ( aPos + aLength + 2 );
    if ( aPos < aBytes ) then Result := Result + ',';
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFeedbackRecvThread.AddToList(aFullText: String): Integer;
var
  aAckData: String;
  aObj: PRecvNagra;
begin
  Result := 0;
  if ( aFullText = EmptyStr ) or ( Length( aFullText ) <= 20 ) then Exit;
  repeat
    aAckData := ExtractValue( aFullText );
    aObj := TNagraAckParser.ParseFeedbackData( aAckData );
    if Assigned( aObj ) then
    begin
      if ( aObj.LowLevelCmd = '1000' ) or
         ( aObj.LowLevelCmd = '1001' ) then
      begin
        if ( aObj.LowLevelCmd = '1001' ) then SetObjectErrorCode( aObj );
        FeedbackAck.BeginWrite;
        try
          FeedbackAck.AddObject( aObj.ResponseTransactionNumber,
            TObject( aObj ) );
        finally
          FeedbackAck.EndWrite;
        end;
      end else
      begin
        aObj.SendStatus := 'W';
        aObj.CmdStatus := 'W';
        UpdatedObject := aObj;
        Synchronize( Update );
        FeedbackRecvList.BeginWrite;
        try
          FeedbackRecvList.AddObject( aObj.ResponseTransactionNumber,
            TObject( aObj ) );
        finally
          FeedbackRecvList.EndWrite;
        end;
        Sleep( 150 );
      end;
    end;
  until ( aFullText = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.ProcessFeedbackRecv;
var
  aObj, aObj2: PRecvNagra;
  aCommIndex: Integer;
begin
  while FeedbackAck.Count > 0 do
  begin
    aObj := PRecvNagra( FeedbackAck.Objects[0] );
    FeedbackCommunicationCheck.BeginWrite;
    try
      aCommIndex := FeedbackCommunicationCheck.IndexOf( aObj.TransactionNumber );
      if aCommIndex >= 0 then
      begin
        aObj2 := PRecvNagra( FeedbackCommunicationCheck.Objects[aCommIndex] );
        aObj2.SendStatus := aObj.SendStatus;
        aObj2.ResponseTransactionNumber := aObj.ResponseTransactionNumber;
        UpdatedObject := aObj2;
        Synchronize( Update );
        { 只要 Ack 回來任何一個是錯誤的, 就通知傳送指令執行緒,
          順便將計數器歸零, PS: 只通知傳送指令執行緒 }
        if ( aObj2.SendStatus = 'E' ) then
        begin
          NotifyToThread( FeedbackSendThread.WndHandle, ncNack );
          FCommunicationCheckCounter := 0;
        end else
          Inc( FCommunicationCheckCounter );
        { 打出訊號通知傳送指令的執行緒 }
        if ( FCommunicationCheckCounter >= 3 ) then
        begin
          NotifyToThread( FeedbackSendThread.WndHandle, ncAck );        
          FCommunicationCheckCounter := 0;
        end;
        Dispose( aObj2 );
        FeedbackCommunicationCheck.Delete( aCommIndex );
      end;
    finally
      FeedbackCommunicationCheck.EndWrite;
    end;
    FeedbackAck.BeginWrite;
    try
      Dispose( aObj );
      FeedbackAck.Delete( 0 );
    finally
      FeedbackAck.EndWrite;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.Update;
begin
  CommandSubject.UpdateCommandObject := UpdatedObject;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_PLAY:
      FRecvCommand := True;
    WM_STOP:
      FRecvCommand := False;
    WM_SOCKET:
      FRecvCommand := ( TNotifyCommandType( TWMSocket( Msg ).NotifyCommandType )
        in [ncCommand, ncCommunication, ncAck] );
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.SetObjectErrorCode(aObj: PRecvNagra);
begin
  if CmdError.Locate( 'ERROR_FLAG;ERROR_CODE',
    VarArrayOf( [0,aObj.ErrCode] ), [] ) then
   aObj.ErrCode := Format( '%s--%s',
     [aObj.ErrCode, CmdError.FieldByName( 'ERROR_DESC' ).AsString] );
  if CmdError.Locate( 'ERROR_FLAG;ERROR_CODE',
    VarArrayOf( [1,aObj.ErrMsg] ), [] ) then
   aObj.ErrMsg := Format( '%s--%s',
     [aObj.ErrMsg, CmdError.FieldByName( 'ERROR_DESC' ).AsString] );
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.InitWMSocket(
  const aNotifyCommandType: TNotifyCommandType; const aErrorCode: Integer;
  aErrorMsg: Integer);
begin
  FWMSocket.Msg := WM_SOCKET;
  FWMSocket.ThreadType := Ord( Self.CommandSubject.ThreadType );
  FWMSocket.NotifyCommandType := Ord( aNotifyCommandType );
  FWMSocket.ErrorCode := aErrorCode;
  FWMSocket.ErrorMsg := aErrorMsg;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedbackRecvThread.NotifyToThread(aThreadHandle: Integer;
  aType: TNotifyCommandType);
begin
  if aType = ncNack then
    InitWMSocket( aType, -1001, -1001 )
  else if aType = ncSocketError then
    InitWMSocket( aType, -1001, -1002 )
  else
    InitWMSocket( aType );  
  PostMessage( aThreadHandle, WM_SOCKET,
    TMessage( FWMSocket ).WParam, TMessage( FWMSocket ).LParam );
end;

{ ---------------------------------------------------------------------------- }

end.
