unit cbControlThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, DBClient,
  {$IFDEF DEBUG} CodeSiteLogging, {$ENDIF}
  cbClass, cbNagraCmd, cbNagraDevice,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;

type

  { �ǰe���O����� }

  TControlSendThread = class(TSMSSocketThread)
  private
    { Private declarations }
    FSendBuffer: PDeviceBuffer;
    FSocketErrorCount: Integer;
    FLastErrorTickCount: Cardinal;
    FLastCommunication: TDateTime;
    FCommandBuilder: TNagraCommandBuilder;
    FSendCommand: Boolean;
    FSignal: TNotifyCommandType;
    FWMSocket: TWMSocket;
    function ControlSend: Integer;
    procedure NotifyCommandSendError;
    procedure NotifyCommunicationCheckError;
    function CanSocketRertyConnect: Boolean;
    function CanCommunicationCheck: Boolean;
    function WaitForSendCommandSignal: Boolean;
    procedure EnsureCommunicationCheck(aCount: Integer);
    procedure InitWMSocket(const aNotifyCommandType: TNotifyCommandType;
      const aErrorCode: Integer = -1; aErrorMsg: Integer = -1);
    procedure NotifyToThread(aThreadHandle: Cardinal; aType: TNotifyCommandType);
  protected
    function GetWaitWhileFrquence: Cardinal; override;
    function EndConnection: Boolean; override;
    procedure SetCASEnv(const aValue: TCASEnv); override;
    procedure SetHighCmdEnv(aValue: TClientDataSet); override;
    procedure SetCmdTransEnv(const aValue: TCmdTransEnv); override;
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
    procedure Update; override;
  public
    constructor Create(const aSocket: TIdTCPClient);
    destructor Destroy; override;
    function BeginConnection: Boolean; override;
    function CommunicationCheck: Boolean;
    property CanSendCommand: Boolean read FSendCommand;
  end;


  { �����ǰe�^������� }

  TControlRecvThread = class(TSMSSocketThread)
  private
    { Private declarations }
    FRecvBuffer: PDeviceBuffer;
    FRecvCommand: Boolean;
    FCommunicationCheckCounter: Integer;
    FCmdError: TClientDataSet;
    FWMSocket: TWMSocket;
    procedure ControlRecv;
    procedure ProcessControlRecv;
    procedure SetObjectErrorCode(aObj: PSendNagra);
    function ExtractBuffer(const aBuffer; const aBytes: Integer): String;
    function AddToList(aFullText: String): Integer;
    procedure InitWMSocket(const aNotifyCommandType: TNotifyCommandType;
      const aErrorCode: Integer = -1; aErrorMsg: Integer = -1);
    procedure NotifyToThread(aThreadHandle: Cardinal; aType: TNotifyCommandType);
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

uses cbMain, cbResStr, cbUtilis, DateUtils, Math;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TControlSendThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }


var
  ControlSendThread: TControlSendThread;
  ControlRecvThread: TControlRecvThread;
  aTestCount: Integer = 0;

{ ---------------------------------------------------------------------------- }

procedure IncrementCasCounter;
begin
  {$IFDEF CASCOUNTER}
    if ( RandomRange( 1, 10 ) in [6..10] ) then
      Inc( ControlSendMaxCounter );
    CodeSite.SendReminder( 'IncrementCasCounter.Counter=%d', [ControlSendMaxCounter] );  
  {$ELSE}
      Inc( ControlSendMaxCounter );
  {$ENDIF}
end;

{ ---------------------------------------------------------------------------- }

procedure DecrementCasCounter;
begin
  if ( ControlSendMaxCounter > 0 ) then
    Dec( ControlSendMaxCounter );
  {$IFDEF CASCOUNTER}
    CodeSite.SendReminder( 'DecrementCasCounter.Counter=%d', [ControlSendMaxCounter] );
  {$ENDIF}
end;

{ ---------------------------------------------------------------------------- }


{ TControlSendThread }

constructor TControlSendThread.Create(const aSocket: TIdTCPClient);
begin
  inherited Create( aSocket );
  New( FSendBuffer );
  ControlSendThread := Self;
  CommandSubject.ThreadType := tdControlSend;
  FCommandBuilder := TNagraCommandBuilder.Create;
  FSocketErrorCount := 0;
  FLastErrorTickCount := 0;
  FLastCommunication := 0;
  FSendCommand := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TControlSendThread.Destroy;
begin
  Dispose( FSendBuffer );
  FCommandBuilder.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.GetWaitWhileFrquence: Cardinal;
begin
  Result := inherited GetWaitWhileFrquence * 3;
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.BeginConnection: Boolean;
begin
  Result := False;
  MessageSubject.Info( Format( SSocketBeginConnection,
    [SControlSend, CASEnv.IP, CASEnv.ControlPort] ) );
  try
    Result := inherited BeginConnection;
  except
    on E: Exception do
    begin
      MessageSubject.Error( Format( SSocketConnectionError,
        [SControlSend, E.Message] ) );
    end;
  end;
  FLastErrorTickCount := GetTickCount;
  { �s�u���\ }
  if Result then
  begin
    { ���~�p�Ƥγ̫���~�ɶ��k�s }
    FSocketErrorCount := 0;
    FLastErrorTickCount := 0;
  end;
  FSendCommand := False;
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.EndConnection: Boolean;
begin
  Result := inherited EndConnection;
  FSendCommand := False;
  FLastErrorTickCount := GetTickCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.SetCASEnv(const aValue: TCASEnv);
begin
  inherited SetCASEnv( aValue );
  GateWaySocket.Host := CASEnv.IP;
  GateWaySocket.Port := CASEnv.ControlPort;
  FCommandBuilder.CASEnv := aValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.SetHighCmdEnv(aValue: TClientDataSet);
begin
  inherited SetHighCmdEnv( aValue );
  FCommandBuilder.HighCmdEnv := aValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.SetCmdTransEnv(const aValue: TCmdTransEnv);
begin
  inherited SetCmdTransEnv( aValue );
  FCommandBuilder.CmdTransEnv := aValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_SOCKET:
      FSignal := TNotifyCommandType( TWMSocket( Msg ).NotifyCommandType );
  else
    inherited WndProc( Msg );
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.Execute;
begin
  { ���� Main Thread ���X���檺�T�� }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
    { �O�_�w�F�쭫�ճs�u�ɶ� }
    if ( not GateWaySocket.Connected ) and ( CanSocketRertyConnect ) then
    begin
      { �q���������O�^�������, ����������, �����X���H��,
        �|�� ControlRecv �m�������, �ɭP Dead Lock }
      NotifyToThread( ControlRecvThread.WndHandle, ncNack );
      Sleep( 300 );
      { ���ճs�u }
      if not BeginConnection then
      begin
        { ���M�s���W, ��ܭ��հT�� }
        MessageSubject.Info( Format( SWaitRetry, [SControlSend, CommEnv.CASRetryFrequence] ) );
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
      NotifyToThread( ControlRecvThread.WndHandle, ncCommunication );
      Sleep( 300 );
      MessageSubject.Info( Format( SEnsureCandSendCommand, [SControlSend] ) );
      FSignal := ncCommunication;
      EnsureCommunicationCheck( 3 );
      FSendCommand := WaitForSendCommandSignal;
      if ( FSendCommand ) then
      begin
        MessageSubject.OK( Format( SEnsureOK, [SControlSend,SControlSend] ) );
        { �q���D����� SOCKET �s�u���`, �åB�i�H�ǰe���O }
        NotifyToThread( MainFormHandle, ncAck );
      end else
      begin
        MessageSubject.Error( Format( SEnsureError, [SControlSend, SControlSend] ) );
        MessageSubject.Warning( Format( SWaitRetry, [SControlSend, CommEnv.CASRetryFrequence] ) );
        EndConnection;
        { �q���������O�^�������, �������� }
        NotifyToThread( ControlRecvThread.WndHandle, ncNack );
        { �q���D����� NAGRA �ڵ����O�ǰe }
        NotifyToThread( MainFormHandle, ncNack );
      end;
    end;
    if Stop then Break;
    if ( FSendCommand ) then
    begin
      { �ǰe���O }
      ControlSend;
      if Stop then Break;
      { ���~���Ƥj��]�w�ȩάO Communication Check Nack �^��  }
      if ( FSocketErrorCount >= CommEnv.CASSendErrMax ) or
         ( FSignal in [ncNack, ncSocketError] ) then
      begin
        EndConnection;
        if ( FSocketErrorCount >= CommEnv.CASSendErrMax ) then
        begin
          NotifyCommandSendError;
          { �q���������O�^�������, �������� }
          NotifyToThread( ControlRecvThread.WndHandle, ncSocketError );
          { �q���D����� SOCKET �s�u���~ }
          NotifyToThread( MainFormHandle, ncSocketError );
        end else
        begin
          NotifyCommunicationCheckError;
          { �q���������O�^�������, �������� }
          NotifyToThread( ControlRecvThread.WndHandle, ncNack );
          { �q���D����� SOCKET �s�u���~ }
          NotifyToThread( MainFormHandle, ncNack );
        end;
      end;
    end;
    if Stop then Break;
    Sleep( GetWaitWhileFrquence );
    if Stop then Break;
  end;
  Sleep( GetWaitWhileFrquence );
  WaitForTerminalSignal;
  EndConnection;
  MessageSubject.Info( Format( SRunStop, [SControlSend] ) );
  MessageSubject.RunState := rsStop;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.Update;
begin
  CommandSubject.UpdateCommandObject := UpdatedObject;
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.ControlSend: Integer;
var
  aIndex, aCmdIndex: Integer;
  aObj, aCmdObj: PSendNagra;
  aText, aErrorCode, aErrorMsg: String;
  aSocketError, aCounterError, aCanSend: Boolean;
  aHigCmdCount, aLowCmdCount: Integer;
  aAlreadyShowCasMaxCounterWarning: Boolean;

  { ------------------------------------------------ }

  procedure UpdateCommandStatus(aCmd: PSendNagra);
  begin
    UpdatedObject := aCmd;
    Synchronize( Update );
    Sleep( 100 );
  end;

  { ------------------------------------------------ }

  procedure WaitForCasMaxCounterReady;
  var
    aCounterTime: TDateTime;
    aTimeOut: Boolean;
  begin
    aTimeOut := False;
    aAlreadyShowCasMaxCounterWarning := False;
    repeat
      aCounterError := ( ControlSendMaxCounter <= 0 );
      if ( aCounterError ) then
      begin
        aCounterTime := Now;
        if ( not aAlreadyShowCasMaxCounterWarning ) then
        begin
          aAlreadyShowCasMaxCounterWarning := True;
          MessageSubject.Warning( Format( SControlSendCasMaxCounterWarning1,
            [SControlSend] ) );
        end;
        Sleep( 10 );
        aTimeOut := ( SecondsBetween( Now, aCounterTime ) >= 3 );
      end;
      if Stop then Break;
    until ( not aCounterError ) or ( aTimeOut );
  end;

  { ------------------------------------------------ }

begin
  Result := 0;
  aLowCmdCount := 0;
  { �o�@�ӳv�@���˩Ҧ��t�Υx���ǰe���O��C }
  for aIndex := 0 to ControlSendManager.ItemCount - 1 do
  begin
    aHigCmdCount := 0;
    if Stop then Break;
    { �Өt�Υx���������O���n�ǰe, �άO�ǰe���Ƥw�j��]�w���� }
    while ( ControlSendManager.Items[aIndex].Count > 0 ) and
          ( aHigCmdCount <= CommEnv.ProcessRecordCount ) do
    begin
      aObj := PSendNagra( ControlSendManager.Items[aIndex].Objects[0] );
      { �N�ӵ��������O�ন�C�����O }
      FCommandBuilder.BuildCommand( aObj );
      { �N�������ͪ��C�����O���ưO���b�찪�����O��,
        �� Ack �^�ӫ�, �n�P�_�O�_�찪�����O�������C�����O�O�_���B�z���� }
      aObj.ReferenceLowCmdCount := FCommandBuilder.CommandCount;
      Inc( aHigCmdCount );
      { �����i�w�ǰe��C }
      ControlSendDoneManager.Items[aIndex].BeginWrite;
      try
        ControlSendDoneManager.Items[aIndex].AddObject( aObj.SeqNo,
          TObject( aObj ) );
      finally
        ControlSendDoneManager.Items[aIndex].EndWrite;
      end;
      { �A�����ݶǰe��C }
      ControlSendManager.Items[aIndex].BeginWrite;
      try
        ControlSendManager.Items[aIndex].Delete( 0 );
      finally
        ControlSendManager.Items[aIndex].EndWrite;
      end;
      {}
      aSocketError := False;
      aCounterError := False;
      { ��ڶǰe�����O, ���i��O�h�ӧC�����O, �Ҧp A2 --> ��� }
      for aCmdIndex := 0 to FCommandBuilder.CommandCount - 1 do
      begin
        Inc( aLowCmdCount );
        aCmdObj := FCommandBuilder.CommandObject[aCmdIndex];
        aCmdObj.SmsSendTime := Now;
        aCmdObj.ParentData := aObj.Data;
        aText := aCmdObj.SmsFullCommandText;
        try
          if ( not aCounterError ) then WaitForCasMaxCounterReady;
          if Stop then Break;
          { �O�_�i�H�ǰe���O }
          aCanSend := ( not aSocketError ) and ( not aCounterError );
          if ( aCanSend ) then
          begin
            ZeroMemory( FSendBuffer, SizeOf( FSendBuffer^ ) );
            Move( PChar( aText )^, FSendBuffer^, Length( aText ) );
            GateWaySocket.WriteBuffer( FSendBuffer^, Length( aText ), True );
            aCmdObj.SmsSendStatus := 'P';
            aCmdObj.CmdStatus := 'P';
            DecrementCasCounter;
          end;
        except
          on E: Exception do
          begin
            Inc( FSocketErrorCount );
            aSocketError := True;
            aErrorCode := '-1000';
            aErrorMsg := Copy( E.Message, 1, 40 );
            MessageSubject.Error( Format( SSendCommandError,
              [SControlSend, aCmdObj.HighCmdId, aCmdObj.LowCmdId, E.Message] ) );
          end;
        end;
        {}
        if ( aCounterError ) and ( aErrorCode = EmptyStr ) then
        begin
          aErrorCode := '-1002';
          aErrorMsg := 'CA_BUFFER_COUNTER_ERROR';
        end;
        { �u�n�b�ǰe�C�����O�ɵo�� TCP/IP �s�u����@�Ӧ���, �ѤU�|���ǰe���N���e,
          �j���~��], ���O��ѤU���C�����O�Хܦ����~ }
        if ( aSocketError ) or ( aCounterError ) then
        begin
          { �C�����O�Хܦ����~ }
          aCmdObj.SmsSendStatus := 'E';
          aCmdObj.CmdStatus := 'E';
          aCmdObj.ErrCode := aErrorCode;
          aCmdObj.ErrMsg := aErrorMsg;
        end;
        { �N�C�����O���� Log List �� }
        ControlSendLogManager.Items[aIndex].BeginWrite;
        try
          ControlSendLogManager.Items[aIndex].AddObject(
            aCmdObj.SmsTransactionNumber, TObject( aCmdObj ) );
        finally
          ControlSendLogManager.Items[aIndex].EndWrite;
        end;
        { ��s�ӵ��C�����O�ǰe���A }
        UpdateCommandStatus( aCmdObj );
        {$IFDEF DEBUG}
           CodeSite.Send( Format( 'HighCmd=%s, LowCmd=%s',
            [aCmdObj.HighCmdId, aCmdObj.LowCmdId] ), aCmdObj.SmsCommandText );
        {$ENDIF}
        { �C�e�@���C�����O, �N���� }
        Sleep( Trunc( CmdTransEnv.SendCommandDelay * 1000 ) );
      end;
      FCommandBuilder.Clear;
      if ( aSocketError ) or ( aCounterError ) then
      begin
        { �������O�Хܦ����~ }
        aObj.SmsSendStatus := 'E';
        aObj.CmdStatus := 'E';
        aObj.ErrCode := aErrorCode;
        aObj.ErrMsg := aErrorMsg;
        { ��s�������O���A }
        UpdateCommandStatus( aObj );
        { ��������w�������G��C }
        ControlRecvManager.Items[aIndex].BeginWrite;
        try
          ControlRecvManager.Items[aIndex].AddObject( aObj.SeqNo,
            TObject( aObj ) );
        finally
          ControlRecvManager.Items[aIndex].EndWrite;
        end;
        { �����w�ǰe��C }
        ControlSendDoneManager.Items[aIndex].BeginWrite;
        try
          aCmdIndex := ControlSendDoneManager.Items[aIndex].IndexOf(
            aObj.SeqNo );
          ControlSendDoneManager.Items[aIndex].Delete( aCmdIndex );
        finally
          ControlSendDoneManager.Items[aIndex].EndWrite;
        end;
      end;
      { �P�_�ǰe���~�w�j�� 3 ��, TCP/IP �s�u�����D, �����۰��_�u,
        �o�إ����_�Өt�Υx�j�� }
      if ( FSocketErrorCount >= CommEnv.CASSendErrMax ) then Break;
      { ���p Communication Check NACK �^�� }
      if ( FSignal in [ncNack, ncSocketError] ) then Break;
      if Stop then Break;
      { �O�_���n�e Communication Check }
      if ( CanCommunicationCheck ) then
      begin
        CommunicationCheck;
        Sleep( 100 );
      end;
    end;
    Result := ( Result + aHigCmdCount );
    { �P�_�ǰe���~�w�j�� 3 ��, TCP/IP �s�u�����D, �����۰��_�u }
    if ( FSocketErrorCount >= CommEnv.CASSendErrMax ) then Break;
    if ( FSignal in [ncNack, ncSocketError] ) then Break;
    if Stop then Break;
    { �O�_���n�e Communication Check }
    if ( CanCommunicationCheck ) then
    begin
      CommunicationCheck;
      Sleep( 100 );
    end;
  end;
  if ( Result > 0 ) then
  begin
    MessageSubject.Info( Format( SControlSendRecord, [SControlSend,
      Result, aLowCmdCount] ) );
  end;
end;


{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.NotifyCommandSendError;
begin
  MessageSubject.Error( Format( SSendErrors, [SControlSend, CommEnv.CASSendErrMax] ) );
  MessageSubject.Warning( Format( SWaitRetry, [SControlSend, CommEnv.CASRetryFrequence] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.NotifyCommunicationCheckError;
begin
  MessageSubject.Error( Format( SCommunicationCheckError,
    [CASEnv.IP, CASEnv.ControlPort, SControlSend, SControlSend] ) );
  MessageSubject.Warning( Format( SWaitRetry, [SControlSend, CommEnv.CASRetryFrequence] ) );
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.CommunicationCheck: Boolean;
var
  aObj: PSendNagra;
  aCmdText: String;
  aIndex: Integer;
begin
  New( aObj );
  aObj.HighCmdId := 'Z1';
  aObj.LowCmdId := '1002';
  aObj.CompCode := CONTROL_COMPCODE;
  aObj.CmdStatus := 'W';
  aObj.SmsSendStatus := 'W';
  aObj.Data := nil;
  aObj.ParentData := nil;
  aObj.Index := -1;
  aObj.ParentIndex := -1;
  aCmdText := FCommandBuilder.BuildCommunicationCheck( aObj );
  { �� Store �i List ��, �o�Ϳ��~����, �A�q List �اR�� }
  ControlCommunicationCheck.BeginWrite;
  try
    ControlCommunicationCheck.AddObject( aObj.SmsTransactionNumber,
      TObject( aObj ) );
  finally
    ControlCommunicationCheck.EndWrite;
  end;
  { ����� CommunicationCheck ���O }
  aObj.SmsSendStatus := 'P';
  aObj.CmdStatus := 'P';
  UpdatedObject := aObj;
  Synchronize( Update );
  try
    GateWaySocket.Write( aCmdText );
    DecrementCasCounter;
    Result := True;
  except
    on E: Exception do
    begin
      Result := False;
      Inc( FSocketErrorCount );
      aObj.SmsSendStatus := 'E';
      aObj.CmdStatus := 'E';
      aObj.ErrCode := '-1000';                          
      aObj.ErrMsg := Copy( E.Message, 1, 10 );
      MessageSubject.Error( Format( SSendCommandError, [SControlSend,
        aObj.HighCmdId, aObj.LowCmdId, E.Message] ) );
    end;
  end;
  if not Result then
  begin
    ControlCommunicationCheck.BeginWrite;
    try
      aIndex := ControlCommunicationCheck.IndexOf( aObj.SmsTransactionNumber );
      if aIndex >= 0 then
        ControlCommunicationCheck.Delete( aIndex );
      Dispose( aObj );
    finally
       ControlCommunicationCheck.EndWrite;
    end;
  end;
  FCommandBuilder.Clear;
  FLastCommunication := Now;
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.CanSocketRertyConnect: Boolean;
var
  aRetry: Cardinal;
begin
  if FLastErrorTickCount <= 0 then
    FLastErrorTickCount := GetTickCount;
  aRetry := GetBetweenFrquence( FLastErrorTickCount );
  if aRetry <= 0 then aRetry := ( CommEnv.CASRetryFrequence * 1000 );
  Result := ( aRetry >= ( CommEnv.CASRetryFrequence * 1000 ) );
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.CanCommunicationCheck: Boolean;
begin
  if ( FLastCommunication <= 0 ) then FLastCommunication := Now;
  Result :=
    ( SecondsBetween( Now, FLastCommunication ) >= CommEnv.CASCommCheck ) and
    ( ControlSendMaxCounter > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.EnsureCommunicationCheck(aCount: Integer);
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

function TControlSendThread.WaitForSendCommandSignal: Boolean;
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


procedure TControlSendThread.InitWMSocket(
  const aNotifyCommandType: TNotifyCommandType;
  const aErrorCode: Integer; aErrorMsg: Integer);
begin
  FWMSocket.Msg := WM_SOCKET;
  FWMSocket.ThreadType := Ord( Self.CommandSubject.ThreadType );
  FWMSocket.NotifyCommandType := Ord( aNotifyCommandType );
  FWMSocket.ErrorCode := aErrorCode;
  FWMSocket.ErrorMsg := aErrorMsg;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.NotifyToThread(aThreadHandle: Cardinal;
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

{ TControlRecvThread }

constructor TControlRecvThread.Create(const aSocket: TIdTCPClient);
begin
  inherited Create( aSocket );
  ControlRecvThread := Self;
  CommandSubject.ThreadType := tdControlRecv;
  New( FRecvBuffer );
  FCmdError := TClientDataSet.Create( nil );
  FRecvCommand := False;
  FCommunicationCheckCounter := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TControlRecvThread.Destroy;
begin
  Dispose( FRecvBuffer );
  FCmdError.Free;
  inherited Destroy;
end;                        

{ ---------------------------------------------------------------------------- }

procedure TControlRecvThread.WndProc(var Msg: TMessage);
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

procedure TControlRecvThread.Execute;
begin
  { ���� Main Thread ���X���檺�T�� }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
    if ( GateWaySocket.Connected and FRecvCommand ) then
      ControlRecv;
    if Stop then Break;
    ProcessControlRecv;
    Sleep( GetWaitWhileFrquence );
    {$IFDEF CASCOUNTER }
      if ( RandomRange( 1, 10 ) in [4..8] ) then
        if ( ControlSendMaxCounter < CASEnv.CmdMaxCounter ) then
          Inc( ControlSendMaxCounter );

      CodeSite.SendReminder( 'MainRecvThread.Counter=%d', [ControlSendMaxCounter] );    
    {$ENDIF}
  end;

  Sleep( GetWaitWhileFrquence );
  MessageSubject.RunState := rsStop;
  WaitForTerminalSignal;
  EndConnection;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlRecvThread.Update;
begin
  CommandSubject.UpdateCommandObject := UpdatedObject;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlRecvThread.ControlRecv;
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

procedure TControlRecvThread.ProcessControlRecv;
var
  aLowObjIndex, aHighObjIndex, aSoIndex, aIndex: Integer;
  aObj, aLowCmdObj, aHighCmdObj: PSendNagra;
  aAlreadyProcess: Boolean;
  aManager, aHighCmdManager: TCommandListMagager;
begin
  { Ack �^�Ӫ��O���@�ث��O ? }
  ControlAck.BeginWrite;
  try
    while ControlAck.Count > 0 do
    begin
      aAlreadyProcess := False;
      aObj := PSendNagra( ControlAck.Objects[0] );
      ControlCommunicationCheck.BeginRead;
      try
        { Communication Check ? Command 1002 }
        aLowObjIndex := ControlCommunicationCheck.IndexOf( aObj.SmsTransactionNumber );
      finally
        ControlCommunicationCheck.EndRead;
      end;
      if aLowObjIndex >= 0 then
      begin
        aLowCmdObj := PSendNagra( ControlCommunicationCheck.Objects[aLowObjIndex] );
        aLowCmdObj.SmsSendStatus := aObj.SmsSendStatus;
        aLowCmdObj.CaAckStatus := aObj.CaAckStatus;
        aLowCmdObj.ErrCode := aObj.ErrCode;
        aLowCmdObj.ErrMsg := aObj.ErrMsg;
        aLowCmdObj.CaAckTransactionNumber := aObj.CaAckTransactionNumber;
        aLowCmdObj.CaAckCommandId := aObj.CaAckCommandId;
        aLowCmdObj.CaAckCommandText := aObj.CaAckCommandText;
        UpdatedObject := aLowCmdObj;
        Synchronize( Update );
        aAlreadyProcess := True;
        { �u�n Ack �^�ӥ���@�ӬO���~��, �N�q���ǰe���O�����,
          ���K�N�p�ƾ��k�s, PS: �u�q���ǰe���O����� }
        if ( aLowCmdObj.SmsSendStatus = 'E' ) then
        begin
          NotifyToThread( ControlSendThread.WndHandle, ncNack );
          FCommunicationCheckCounter := 0;
        end else
          Inc( FCommunicationCheckCounter );
        { ���X�T���q���ǰe���O������� }
        if ( FCommunicationCheckCounter >= 3 ) then
        begin
          NotifyToThread( ControlSendThread.WndHandle, ncAck );
          FCommunicationCheckCounter := 0;
        end;
        { ���� Ack �^�Ӫ� Obj }
        Dispose( aObj );
        ControlAck.Delete( 0 );
        ControlCommunicationCheck.BeginWrite;
        try
          { ���� Communication Check Object }
          Dispose( aLowCmdObj );
          ControlCommunicationCheck.Delete( aLowObjIndex );
        finally
          ControlCommunicationCheck.EndWrite;
        end;
        IncrementCasCounter;
      end;
      { ���椣�i�ק�, �Y Ack �^�Ӫ��O Communication Check, �h���i����U�� }
      if aAlreadyProcess then Continue;
      {}
      aManager := ControlSendLogManager;
      aLowCmdObj := nil;
      for aIndex := 0 to aManager.ItemCount - 1 do
      begin
        aManager.Items[aIndex].BeginRead;
        try
          aLowObjIndex := aManager.Items[aIndex].IndexOf( aObj.SmsTransactionNumber );
          if aLowObjIndex >= 0 then
            aLowCmdObj := PSendNagra( aManager.Items[aIndex].Objects[aLowObjIndex] );
        finally
          aManager.Items[aIndex].EndRead;
        end;
        if not Assigned( aLowCmdObj ) then
          Continue
        else
          Break;
      end;
      if Assigned( aLowCmdObj ) then
      begin
        aManager.Items[aIndex].BeginWrite;
        try
          aLowCmdObj.SmsSendStatus := aObj.SmsSendStatus;
          aLowCmdObj.CaAckStatus := aObj.CaAckStatus;
          aLowCmdObj.CaAckTime := aObj.CaAckTime;
          aLowCmdObj.ErrCode := aObj.ErrCode;
          aLowCmdObj.ErrMsg := aObj.ErrMsg;
          aLowCmdObj.CaAckCommandId := aObj.CaAckCommandId;
          aLowCmdObj.CaAckTransactionNumber := aObj.CaAckTransactionNumber;
          aLowCmdObj.CaAckCommandText := aObj.CaAckCommandText;
          UpdatedObject := aLowCmdObj;
          Synchronize( Update );
          Sleep( 10 );
          { �^�Y��ӧC�������쪺�찪�����O, �Y�Ӱ������O�U���C�����O
            �Ҥw�g Ack �^��, �B OK, �h�Ӱ������O�h�⧹��, �Y�O Ack �^
            �Ӫ��C�����O���@�����~, ����������찪�����O�����~ }
          aHighCmdManager := ControlSendDoneManager;
          aSoIndex := aLowCmdObj.SoCompListIndex;
          { �������w�����쪺�t�Υx List �� }
          aHighCmdObj := nil;
          aHighCmdManager.Items[aSoIndex].BeginRead;
          try
            aHighObjIndex := aHighCmdManager.Items[aSoIndex].IndexOf( aLowCmdObj.SeqNo );
            if ( aHighObjIndex >= 0 ) then
              aHighCmdObj := PSendNagra(
                aHighCmdManager.Items[aSoIndex].Objects[aHighObjIndex] );
          finally
            aHighCmdManager.Items[aSoIndex].EndRead;
          end;
          if Assigned( aHighCmdObj ) then
          begin
            if ( aHighCmdObj.SmsSendStatus <> 'E' ) then
            begin
              aHighCmdObj.SmsSendStatus := aLowCmdObj.SmsSendStatus;
              aHighCmdObj.ErrCode := aLowCmdObj.ErrCode;
              aHighCmdObj.ErrMsg := aLowCmdObj.ErrMsg;
            end;
            { �N�ӵ������������C�����O�Ʀ���, �Y���� 0, ��ܹ���
              ���C�����O�� Ack �^�ӤF }
            Dec( aHighCmdObj.ReferenceLowCmdCount );
            { �ӵ����������~, �άO�Ҧ����C���w�ǰe�� Ack �^��, �h��s���A }
            if ( aHighCmdObj.SmsSendStatus = 'E' ) or
               ( aHighCmdObj.ReferenceLowCmdCount = 0 ) then
            begin
              UpdatedObject := aHighCmdObj;
              Synchronize( Update );
            end;
            if ( aHighCmdObj.ReferenceLowCmdCount = 0 ) then
            begin
              { �R�� }
              aHighCmdManager.Items[aSoIndex].BeginWrite;
              try
               aHighObjIndex := aHighCmdManager.Items[aSoIndex].IndexOf(
                 aLowCmdObj.SeqNo );
                aHighCmdManager.Items[aSoIndex].Delete( aHighObjIndex );
              finally
                aHighCmdManager.Items[aSoIndex].EndWrite;
              end;
              { ���� ControlRev List �� }
              ControlRecvManager.Items[aSoIndex].BeginWrite;
              try
                ControlRecvManager.Items[aSoIndex].AddObject(
                  aHighCmdObj.SeqNo, TObject( aHighCmdObj ) );
              finally
                ControlRecvManager.Items[aSoIndex].EndWrite;
              end;
            end;
          end;
        finally
          aManager.Items[aIndex].EndWrite;
        end;
        IncrementCasCounter;
      end else
      begin
        {$IFDEF DEBUG}
          CodeSite.SendError( 'aObj.TransactionNumber=' + aObj.SmsTransactionNumber );
        {$ENDIF}
      end;
      { �u�i���� Ack �^�Ӫ� Obj, ���i�H����]���i�R����C�����O,
        �n����g Log ��~�i���� }
      { �]�� Ack �^�ӥu�|���@��, �ҥH���|�����Ч�s���A�����D }
      if aObj <> nil then
         Dispose( aObj );
      ControlAck.Delete( 0 );
    end;  
  finally
    ControlAck.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TControlRecvThread.ExtractBuffer(const aBuffer; const aBytes: Integer): String;
var
  aDoubleChr, aFullText: String;
  aPos, aLength: Integer;
begin
  { Buffer ���e, �Y 2 �� Byte ���e������ }
  { Buffer ���e, �|�� 1 ���H�W��ƶ��n���, �j�����ɭԷ|�u���@�� }
  Result := EmptyStr;
  aPos := 0;
  {$IFDEF DEBUG}
    CodeSite.AddSeparator;
    CodeSite.EnterMethod( 'ControlRecv-ExtractBuffer' );
  {$ENDIF}
  while ( aPos < aBytes ) do
  begin
    aDoubleChr := PChar( aBuffer )[aPos] + PChar( aBuffer )[aPos+1];
    aLength := HexToLength( aDoubleChr );
    { ���L�_�Y 2 �Ӧ줸 }
    SetString( aFullText, PChar( aBuffer ) + ( aPos + 2 ), aLength );
    {$IFDEF DEBUG}
      CodeSite.SendMsg( 'Text=' + aFullText );
    {$ENDIF}
    Result := ( Result + aFullText );
    aPos := ( aPos + aLength + 2 );
    if ( aPos < aBytes ) then Result := Result + ',';
  end;
  {$IFDEF DEBUG}
    CodeSite.ExitMethod( 'ControlRecv-ExtractBuffer' );
    CodeSite.AddSeparator;
  {$ENDIF}
end;

{ #E1234567890#F987654321#E1234567890 }

{ ---------------------------------------------------------------------------- }

function TControlRecvThread.AddToList(aFullText: String): Integer;
var
  aAckData: String;
  aObj: PSendNagra;
begin
  Result := 0;
  if ( aFullText = EmptyStr ) or ( Length( aFullText ) <= 20 ) then Exit;
  repeat
    aAckData := ExtractValue( aFullText );
    aObj := TNagraAckParser.ParserAcknowledge( aAckData );
    if Assigned( aObj ) then
    begin
      if aObj.CaAckCommandId = '1001' then SetObjectErrorCode( aObj );
      ControlAck.BeginWrite;
      try
        ControlAck.AddObject( aObj.CaAckTransactionNumber, TObject(aObj ) );
      finally
        ControlAck.EndWrite;
      end;
    end;
  until ( aFullText = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TControlRecvThread.SetObjectErrorCode(aObj: PSendNagra);
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

procedure TControlRecvThread.InitWMSocket(
  const aNotifyCommandType: TNotifyCommandType;
  const aErrorCode: Integer; aErrorMsg: Integer);
begin
  FWMSocket.Msg := WM_SOCKET;
  FWMSocket.ThreadType := Ord( Self.CommandSubject.ThreadType );
  FWMSocket.NotifyCommandType := Ord( aNotifyCommandType );
  FWMSocket.ErrorCode := aErrorCode;
  FWMSocket.ErrorMsg := aErrorMsg;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlRecvThread.NotifyToThread(aThreadHandle: Cardinal;
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

initialization
  Randomize;

end.
