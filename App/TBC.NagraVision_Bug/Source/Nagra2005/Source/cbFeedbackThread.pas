unit cbFeedbackThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, DBClient,
  cbClass, cbNagraCmd, cbNagraDevice,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;

type

  { �^�Ǿ���ǰe�^������� }

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

  { �����^�Ǿ����ư���� }

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
  { ���M����, �O���̫�@�����~�ɶ�, �ΨӤ�ﭫ�ճs�u�� }
  FLastErrorTime := Now;
  { �s�u���\ }
  if Result then
  begin
    { ���~�p�Ƥγ̫���~�ɶ��k�s }
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
  { �� Store �i List ��, �o�Ϳ��~����, �A�q List �اR�� }
  FeedbackCommunicationCheck.BeginWrite;
  try
    FeedbackCommunicationCheck.AddObject( aObj.TransactionNumber,
      TObject( aObj ) );
  finally
    FeedbackCommunicationCheck.EndWrite;
  end;
  { ��� CommunicationCheck ���O }
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
  { ���� Main Thread ���X���檺�T�� }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
    { ���n�s�u, �B�w�F�쭫�ճs�u�ɶ�, �Ĥ@���i�Ӥ@�w�O�i�w�F���ճs�u�ɶ��j }
    if ( not GateWaySocket.Connected ) and ( CanSocketRertyConnect ) then
    begin
      { �q���������O�^�������, ����������, �����X���H��,
        �|�� ControlRecv �m�������, �ɭP Dead Lock }
      NotifyToThread( FeedbackRecvThread.WndHandle, ncNack );
      Sleep( 300 );
      { ���ճs�u }
      if not BeginConnection then
      begin
        { ���M�s���W, ��ܭ��հT�� }
        MessageSubject.Warning( Format( SWaitRetry, [SFeedbackSend, CommEnv.CASRetryFrequence] ) );
        EndConnection;
        NotifyToThread( MainFormHandle, ncSocketError );
      end;
    end;
    if Stop then Break;
    if ( GateWaySocket.Connected ) and ( not FSendCommand ) then
    begin
      { �s�W�F, ���O�����o�i�H�ǰe���O, ���e�X Communication Check,
        �T�{�O�_�i�H����O }
      { �q���������O�^�������, �������, �����X���H���|�ɭP�L�k���� ACK }
      NotifyToThread( FeedbackRecvThread.WndHandle, ncCommunication );
      Sleep( 300 );
      FSignal := ncCommunication;
      MessageSubject.Info( Format( SEnsureCandSendCommand, [SFeedbackSend] ) );
      EnsureCommunicationCheck( 3 );
      FSendCommand := WaitForSendCommandSignal;
      if ( FSendCommand ) then
      begin
        MessageSubject.OK( Format( SEnsureOK, [SFeedbackSend, SFeedbackSend] ) );
        { �q���D����� SOCKET �s�u���`, �åB�i�H�����^�Ǹ�� }
        NotifyToThread( MainFormHandle, ncAck );
      end else
      begin
        MessageSubject.Error( Format( SEnsureError, [SFeedbackSend, SFeedbackSend] ) );
        MessageSubject.Warning( Format( SWaitRetry, [SFeedbackSend, CommEnv.CASRetryFrequence] ) );
        EndConnection;
        { �q���������O�^�������, �������� }
        NotifyToThread( FeedbackRecvThread.WndHandle, ncNack );
        { �q���D����� NAGRA �ڵ��^�Ǿ��� }
        NotifyToThread( MainFormHandle, ncNack );
      end;
    end;
    if Stop then Break;
    if ( FSendCommand ) then
    begin
      { �ǰe�^�� }
      FeedbackSend;
      { ���~���Ƥj��]�w�� }
      if ( FSendError >= CommEnv.CASSendErrMax ) then
      begin
        NotifyCommandSendError;
        EndConnection;
        { �q�������^�Ǹ�ư����, �������� }
        NotifyToThread( FeedbackRecvThread.WndHandle, ncSocketError );
        { �q���D����� SOCKET �s�u���~ }
        NotifyToThread( MainFormHandle, ncSocketError );
      end;
    end;
    if Stop then Break;
    if ( FSendCommand ) then
    begin
      if ( CanCommunicationCheck ) then
        CommunicationCheck;
      { ���~���Ƥj��]�w�ȩάO Communication Check Nack �^��  }
      if ( FSendError >= CommEnv.CASSendErrMax ) or
         ( FSignal in [ncNack, ncSocketError] ) then
      begin
        if ( FSendError >= CommEnv.CASSendErrMax ) then
          NotifyCommandSendError
        else
          NotifyCommunicationCheckError;
        EndConnection;
        { �q���������O�^�������, �������� }
        NotifyToThread( FeedbackRecvThread.WndHandle, ncNack );
        { �q���D����� SOCKET �s�u���~ }
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
      { �W�L 10 ���M������ OK �H�� }
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
  { �@�s�W Nagra, �������e n �� Communication Check Command }
  { �s�� n �� Ack �^�ӳ��O OK, �~�i�H�� Command �� Nagra }
  for aIndex := 1 to aCount do
  begin
    CommunicationCheck;
    { �@�s�W Nagra, �C1��e�@�� Communication Check Command }
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
  { TCP/IP �ǰe���~ Flag }
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
      { �S�����~�o��, �h�ǰe���O }
      if not aImmediateError then
      begin
        ZeroMemory( FSendBuffer, SizeOf( FSendBuffer^ ) );
        Move( PChar( aText )^, FSendBuffer^, Length( aText ) );
        GateWaySocket.WriteBuffer( FSendBuffer^, Length( aText ), True );
        { �^�����` }
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
    { ��ACK���O���~ }
    if aImmediateError then
    begin
      { �N�������O�ΧC���Хܦ����~ }
      aObj.SendStatus := 'E';
      aObj.ErrCode := aErrorCode;
      aObj.ErrMsg := aErrorMsg;
    end;
    { ��s�ӵ����O ACK ���A }
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
      { ���� }
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
  { ���� Main Thread ���X���檺�T�� }
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
  { Buffer ���e, �Y 2 �� Byte ���e������ }
  { Buffer ���e, �|�� 1 ���H�W��ƶ��n���, �j�����ɭԥu���@�� }
  Result := EmptyStr;
  aPos := 0;
  while ( aPos < aBytes ) do
  begin
    aDoubleChr := PChar( aBuffer )[aPos] + PChar( aBuffer )[aPos+1];
    aLength := HexToLength( aDoubleChr );
    { ���L�_�Y 2 �� Byte  }
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
        { �u�n Ack �^�ӥ���@�ӬO���~��, �N�q���ǰe���O�����,
          ���K�N�p�ƾ��k�s, PS: �u�q���ǰe���O����� }
        if ( aObj2.SendStatus = 'E' ) then
        begin
          NotifyToThread( FeedbackSendThread.WndHandle, ncNack );
          FCommunicationCheckCounter := 0;
        end else
          Inc( FCommunicationCheckCounter );
        { ���X�T���q���ǰe���O������� }
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
