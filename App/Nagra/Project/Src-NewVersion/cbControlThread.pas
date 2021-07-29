unit cbControlThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, DBClient,
  {$IFDEF DEBUG} CsIntf, {$ENDIF} cbClass,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;

type
  TControlSendThread = class(TGateWaySocketThread)
  private
    { Private declarations }
  protected
    procedure SetCASEnv(const aValue: TCASEnv); override;
    procedure Execute; override;
    procedure Update; override;
  public
    constructor Create(const aSocket: TIdTCPClient);
    destructor Destroy; override;
    function BeginConnection: Boolean; override;
  end;

implementation

uses cbMain, cbResStr, cbUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TControlSendThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TControlSendThread }

constructor TControlSendThread.Create(const aSocket: TIdTCPClient);
begin
  inherited Create( aSocket );
end;

{ ---------------------------------------------------------------------------- }

destructor TControlSendThread.Destroy;
begin
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.SetCASEnv(const aValue: TCASEnv);
begin
  inherited SetCASEnv( aValue );
  GateWaySocket.Host := CASEnv.IP;
  GateWaySocket.Port := CASEnv.ControlPort;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.Execute;
begin
  { 等待 Main Thread 打出開始執行的訊號 }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
  end;
  Sleep( GetWaitWhileFrquence );
  MessageSubject.RunState := rsStop;
  WaitForTerminalSignal;
  EndConnection;
end;

{ ---------------------------------------------------------------------------- }

procedure TControlSendThread.Update;
begin
  { ... }
end;

{ ---------------------------------------------------------------------------- }

function TControlSendThread.BeginConnection: Boolean;
begin
  Result := False;
  SocketMessage := Format( SControlSendBeginConnection, [CASEnv.IP, CASEnv.ControlPort] );
  Synchronize( Notify );
  try
    Result := inherited BeginConnection;
    if Result then
      SocketMessage := SControlSendConnectionSuccess
    else
      SocketMessage := SControlSendConnectionReject;
  except
    on E: Exception do
    begin
      SocketMessage := Format( SControlSendConnectionError, [E.Message] );
    end;
  end;
  Synchronize( Notify );
end;

{ ---------------------------------------------------------------------------- }

end.
