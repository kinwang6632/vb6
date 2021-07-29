unit cbClientThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbSyncObjs,
{$IFDEF APPDEBUG}
  CodeSiteLogging,
{$ENDIF}
{ App: }
  cbThread, cbDesignPattern, cbROClientModule, cbClientClass;

type
  TClientThread = class(TBaseThread)
  private
    { Private declarations }
    FCompCode: String;
    FUserId: String;
    FUserPass: String;
    FHost: String;
    FMsgSubject: TMsgSubject;
    FLastReConnectTime: TDateTime;
    FLastReLoginTime: TDateTime;
    FLastCheckCSMISProcessTime: TDateTime;
    FLogin: Boolean;
    FReadyEvent: TSimpleEvent;
    FMainFormHandle: THandle;
    FCSMISProcessExists: Boolean;
    function IsConnected: Boolean;
    function WantToReconnect: Boolean;
    function WantToReLogin: Boolean;
    function CheckCSMISProcessExists: Boolean;
    procedure InitROChannel;
    procedure CheckRemoteConnection;
    procedure CheckServiceLogin;
    procedure CheckServiceLogout;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property CompCode: String read FCompCode write FCompCode;
    property UserId: String read FUserId write FUserId;
    property UserPass: String read FUserPass write FUserPass;
    property Host: String read FHost write FHost;
    property ReadyEvent: TSimpleEvent read FReadyEvent;
    property MsgSubject: TMsgSubject read FMsgSubject;
    property MainFormHandle: THandle read FMainFormHandle write FMainFormHandle;
  end;

implementation

uses cbUtilis, DateUtils;

{ ---------------------------------------------------------------------------- }

procedure ExtractHostAndPort(const AValue: String; var AHost, APort: String);
var
  aIdx: Integer;
begin
  AHost := AValue;
  APort := EmptyStr;
  aIdx := Pos( ':', AValue );
  if ( aIdx > 0 ) then
  begin
    AHost := Copy( AValue, 1, aIdx - 1 );
    APort := Copy( AValue, aIdx + 1, Length( AValue ) - aIdx );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TClientThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ---------------------------------------------------------------------------- }

{ TClientThread }

constructor TClientThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FMsgSubject := TMsgSubject.Create;
  FReadyEvent := TSimpleEvent.Create( nil, True, False, EmptyStr );
  FLastReConnectTime := 0;
  FLastReLoginTime := 0;
  FLogin := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TClientThread.Destroy;
begin
  FMsgSubject.Free;
  FReadyEvent.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TClientThread.IsConnected: Boolean;
begin
  Result := (
    ROClientModule.ROSuperTcpChannel.Connected and
    ROClientModule.ROSuperTcpChannel.Active );
end;

{ ---------------------------------------------------------------------------- }

function TClientThread.WantToReconnect: Boolean;
var
  aTryReconnectRate: Integer;
begin
  Result := ( FLastReConnectTime = 0 );
  if not Result then
  begin
    aTryReconnectRate := 10;
    ROClientModule.BeginLockClientParam;
    try
      if not VarIsNull( ROClientModule.ClientParam.FieldByName( 'TryReconnectRate' ).Value ) then
        aTryReconnectRate := ROClientModule.ClientParam.FieldByName( 'TryReconnectRate' ).AsInteger;
    finally
      ROClientModule.EndLockClientParam;
    end;
    Result :=
      ( not Self.Terminated ) and
      ( SecondsBetween( Now, FLastReConnectTime ) >= aTryReconnectRate );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TClientThread.WantToReLogin: Boolean;
begin
  Result := ( FLastReLoginTime = 0 );
  if not Result then
  begin
    Result :=
      ( not Self.Terminated ) and
      ( SecondsBetween( Now, FLastReLoginTime ) >= 15 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TClientThread.CheckCSMISProcessExists: Boolean;
begin
  Result := True;
  if ( not Self.Terminated ) and
     ( SecondsBetween( Now, FLastCheckCSMISProcessTime ) >= 5 ) then
  begin
    Result := IsProcessExists( 'CSMIS.exe' );
    FLastCheckCSMISProcessTime := Now;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TClientThread.InitROChannel;
var
  aHost, aPort: String;
begin
  ExtractHostAndPort( FHost, aHost, aPort );
  ROClientModule.ROSuperTcpChannel.Host := aHost;
  if ( aPort <> EmptyStr ) then
    ROClientModule.ROSuperTcpChannel.Port := StrToInt( aPort );
end;

{ ---------------------------------------------------------------------------- }

procedure TClientThread.CheckRemoteConnection;
begin
  if ( not IsConnected ) and ( WantToReconnect ) and ( not Self.Terminated ) then
  begin
    FReadyEvent.ResetEvent;
    FLogin := False;
    ROClientModule.ROSuperTcpChannel.Active := False;
    Sleep( 100 );
    ROClientModule.ROSuperTcpChannel.Active := True;
    { 與控制台連線中 }
    FMsgSubject.Info( 'C1' );
    Sleep( 1000 );
    if ( not IsConnected ) then
    begin
      FMsgSubject.Info( 'E0,0x0002' );
      Sleep( 500 );
    end;
    FLastReConnectTime := Now;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TClientThread.CheckServiceLogin;
var
  aMsg: String;
begin
  if ( IsConnected ) and ( not FLogin ) and ( WantToReLogin ) and ( not Self.Terminated ) then
  begin
    { 登入中 }
    FMsgSubject.Info( 'A1' );
    Sleep( 1000 );
    FLogin := ROClientModule.Login( FCompCode, FUserId, FUserPass, aMsg );
    Sleep( 500 );
    if FLogin then
    begin
      FReadyEvent.SetEvent;
      { 已登入 }
      FMsgSubject.Info( 'A2' );
      Sleep( 1000 );
    end else
    begin
      FMsgSubject.Info( Format( 'E1,0x0001:%s', [aMsg] ) );
    end;
    FLastReLoginTime := Now;
  end;
end;

{ ---------------------------------------------------------------------------- }
                                            
procedure TClientThread.CheckServiceLogout;
begin
  if ( IsConnected ) and ( FLogin ) then
  begin
    FReadyEvent.ResetEvent;
    { 登出中 }
    FMsgSubject.Info( 'A3' );
    Sleep( 1000 );
    ROClientModule.Logout;
    { 已登出 }
    FMsgSubject.Info( 'A0' );
    Sleep( 500 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TClientThread.Execute;
begin
  InitROChannel;
  FLastCheckCSMISProcessTime := Now;
  FCSMISProcessExists := True;
  while not Self.Terminated do
  begin
    try
      Sleep( 200 );
      CheckRemoteConnection;
      Sleep( 200 );
      if Self.Terminated then Break;
      CheckServiceLogin;
      Sleep( 200 );
      {$IFDEF APPDEBUG}
      {$ELSE}
      if FCSMISProcessExists then
      begin
        if not CheckCSMISProcessExists then
        begin
          PostMessage( FMainFormHandle, WM_CSMISCLOSE, 0, 0 );
          FCSMISProcessExists := False;
        end;
      end;
      {$ENDIF}
    except
      on E: Exception do
      begin
        {$IFDEF APPDEBUG}
          CodeSite.SendError( '%s=%s', [Self.ClassName,E.Message] );
        {$ENDIF}
      end;
    end;
  end;
  CheckServiceLogout;
  Sleep( 1000 );
  ROClientModule.ROSuperTcpChannel.Active := False;
end;

{ ---------------------------------------------------------------------------- }

end.
