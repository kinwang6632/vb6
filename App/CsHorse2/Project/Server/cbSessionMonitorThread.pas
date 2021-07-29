unit cbSessionMonitorThread;

interface

uses
  SysUtils, Classes, Windows, Messages, cbSyncObjs,
{ App }
  cbSo, cbSrvClass, cbDesignPattern, cbThread, cbLanguage;

type
  TSessionMonitorThread = class(TBaseThread)
  private
    { Private declarations }
    FMsgSub: TMsgSubject;
    FSoSub: TSoSubject;
    FWarningSub: TMsgSubject;
    FCommEnv: TCommEnv;
    FEvent: TSimpleEvent;
    FLastVdSessionTime: TDateTime;
    FDbSessionManagerList: TDbSessionManagerList;
    function CheckCanValidateSession: Boolean;
    procedure OpenSession(ASession: TDbSession; const ASilent: Boolean = False);
    procedure OpenAllSession; overload;
    procedure OpenAllSession(AMgr: TDbSessionManager); overload;
    procedure CloseSession(ASession: TDbSession);
    procedure CloseAllSession; overload;
    procedure CloseAllSession(AMgr: TDbSessionManager); overload;
    procedure ValidateSession(AMgr: TDbSessionManager); overload;
    procedure ValidateSession; overload;
    procedure UpdateSessionState;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property CommEnv: TCommEnv read FCommEnv;
    property Event: TSimpleEvent read FEvent;
    property MsgSubject: TMsgSubject read FMsgSub;
    property SoSubject: TSoSubject read FSoSub;
    property WarningSubject: TMsgSubject read FWarningSub;
    property DbSessionManagerList: TDbSessionManagerList read FDbSessionManagerList write FDbSessionManagerList;
  end;

implementation

uses cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSessionMonitor.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ---------------------------------------------------------------------------- }

{ TSessionMonitor }

constructor TSessionMonitorThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FMsgSub := TMsgSubject.Create;
  FSoSub := TSoSubject.Create;
  FWarningSub := TMsgSubject.Create;
  FCommEnv := TCommEnv.Create;
  FEvent := TSimpleEvent.Create( nil, True, False, EmptyStr );
  FLastVdSessionTime := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TSessionMonitorThread.Destroy;
begin
  FMsgSub.Free;
  FSoSub.Free;
  FWarningSub.Free;
  CommEnv.Free;
  FDbSessionManagerList := nil;
  FEvent.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.Execute;
begin
  FMsgSub.OK( LanguageManager.GetFmt( 'SDbCreateSessionMonitorThreadEnd',
    [HexThreadId] ) );
  Sleep( 500 );
  OpenAllSession;
  Sleep( 500 );
  FEvent.SetEvent;
  while not Self.Terminated do
  begin
    if CheckCanValidateSession then
      ValidateSession;
    if Self.Terminated then Break;
    UpdateSessionState;
    if Self.Terminated then Break;
    Sleep( 200 );
  end;
  Sleep( 500 );
  CloseAllSession;
  Sleep( 500 );
  FMsgSub.Warning( LanguageManager.GetFmt( 'SDbSessionMonitorThreadStop',
    [HexThreadId] ) );
end;

{ ---------------------------------------------------------------------------- }

function TSessionMonitorThread.CheckCanValidateSession: Boolean;
begin
  if ( FLastVdSessionTime <= 0 ) then FLastVdSessionTime := Now;
  Result := ( not Self.Terminated ) and ( SecondsBetween( Now,
    FLastVdSessionTime ) >= FCommEnv.DbVaildateSessionFreq );
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.OpenSession(ASession: TDbSession; const ASilent: Boolean = False);
begin
  ASession.Connection.Close;
  try
    ASession.Connection.Open;
  except
    on E: Exception do
    begin
      if not ASilent then
        FMsgSub.Error( LanguageManager.GetFmt( 'SDbOpenSessionConnection',
          [ASession.AppSo.CompName, ASession.Index, E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.OpenAllSession(AMgr: TDbSessionManager);
var
  aIndex: Integer;
begin
  Sleep( 500 );
  for aIndex := 0 to AMgr.SessionList.Count - 1 do
  begin
    if AMgr.SessionList.GetErrorSessionLock( aIndex ) then
    begin
      try
        OpenSession( AMgr.SessionList.DbSession[aIndex] );
      finally
        AMgr.SessionList.ReleaseSessionLock( aIndex );
      end;
    end;
    Sleep( 30 );
    if Self.Terminated then Break;
  end;
  Sleep( 500 );
  if ( AMgr.SessionList.FreeSessionCount = 0 ) then
  begin
    FMsgSub.Error( LanguageManager.GetFmt( 'SDbSoOpenSessionError',
      [AMgr.AppSo.CompName, AMgr.SessionList.ErrSessionCount] ) )
  end else
  if ( AMgr.SessionList.ErrSessionCount = 0 ) then
  begin
    FMsgSub.OK( LanguageManager.GetFmt( 'SDbSoOpenSessionEnd',
      [AMgr.AppSo.CompName, AMgr.SessionList.FreeSessionCount] ) );
  end else
  begin
    FMsgSub.Warning( LanguageManager.GetFmt( 'SDbSoOpenSessionWarning1',
      [AMgr.AppSo.CompName, AMgr.SessionList.FreeSessionCount, AMgr.SessionList.ErrSessionCount] ) );
  end;
  if ( AMgr.SessionList.FreeSessionCount <> AMgr.AppSo.MaxSession ) then
  begin
    FMsgSub.Warning( LanguageManager.GetFmt( 'SDbSoOpenSessionWarning2',
      [AMgr.AppSo.CompName, FCommEnv.DbRetryFreq] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.OpenAllSession;
var
  aIndex: Integer;
  aMgr: TDbSessionManager;
begin
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    aMgr := FDbSessionManagerList.Manager[aIndex];
    OpenAllSession( aMgr );
    aMgr.AppSo.DbState := dbOpen;
    if ( aMgr.AppSo.ErrSession > 0 ) then aMgr.AppSo.DbState := dbError;
    FSoSub.Notify( FDbSessionManagerList.Manager[aIndex].AppSo );
    Sleep( 500 );
    if Self.Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.CloseSession(ASession: TDbSession);
begin
  try
    ASession.Connection.Close;
  except
    on E: Exception do
    begin
      FMsgSub.Error( LanguageManager.GetFmt( 'SDbCloseSessionConnection',
        [ASession.AppSo.CompName, ASession.Index, E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.CloseAllSession(AMgr: TDbSessionManager);
var
  aIndex: Integer;
begin
  Sleep( 500 );
  for aIndex := 0 to AMgr.SessionList.Count - 1 do
  begin
    if AMgr.SessionList.GetFreeSessionLock( aIndex ) then
    begin
      try
        CloseSession( AMgr.SessionList.DbSession[aIndex] );
      finally
        AMgr.SessionList.ReleaseSessionLock( aIndex );
      end;
    end;
    Sleep( 30 );
  end;
  Sleep( 500 );
  if ( AMgr.SessionList.ErrSessionCount = 0 ) then
  begin
    FMsgSub.Error( LanguageManager.GetFmt( 'SDbSoCloseSessionError',
      [AMgr.AppSo.CompName, AMgr.SessionList.ErrSessionCount] ) )
  end else
  if ( AMgr.SessionList.FreeSessionCount = 0 ) then
  begin
    FMsgSub.OK( LanguageManager.GetFmt( 'SDbSoCloseSessionEnd',
      [AMgr.AppSo.CompName, AMgr.SessionList.ErrSessionCount] ) );
  end;
  if ( AMgr.SessionList.ErrSessionCount <> AMgr.AppSo.MaxSession ) then
  begin
    FMsgSub.Warning( LanguageManager.GetFmt( 'SDbSoCloseSessionWarning',
      [AMgr.AppSo.CompName, AMgr.SessionList.FreeSessionCount, AMgr.SessionList.ErrSessionCount] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.CloseAllSession;
var
  aIndex: Integer;
  aMgr: TDbSessionManager;
begin
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    aMgr := FDbSessionManagerList.Manager[aIndex];
    CloseAllSession( aMgr );
    aMgr.AppSo.DbState := dbClose;
    aMgr.AppSo.BusySession := 0;
    aMgr.AppSo.ErrSession := 0;
    aMgr.AppSo.FreeSession := 0;
    FSoSub.Notify( aMgr.AppSo );
    Sleep( 200 );
  end;
end;


{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.ValidateSession(AMgr: TDbSessionManager);
var
  aIndex: Integer;
begin
  for aIndex := 0 to AMgr.SessionList.Count - 1 do
  begin
    if AMgr.SessionList.GetErrorSessionLock( aIndex ) then
    begin
      try
        OpenSession( AMgr.SessionList.DbSession[aIndex] );
      finally
        AMgr.SessionList.ReleaseSessionLock( aIndex );
      end;
    end;
    Sleep( 30 );
    if Self.Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.ValidateSession;
var
  aIndex, aIndex2, aErrCount: Integer;
begin
  FMsgSub.Info( LanguageManager.Get( 'SDbValidateSessionStart' ) );
  Sleep( 500 );
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    ValidateSession( FDbSessionManagerList.Manager[aIndex] );
    FSoSub.Notify( FDbSessionManagerList.Manager[aIndex].AppSo );
    Sleep( 100 );
    if Self.Terminated then Break;
  end;
  {}
  if Self.Terminated then Exit;
  {}
  aErrCount := 0;
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    for aIndex2 := 0 to FDbSessionManagerList.Manager[aIndex].SessionList.Count - 1 do
    begin
      if FDbSessionManagerList.Manager[aIndex].SessionList.DbSession[aIndex].IsError then
        Inc( aErrCount );
    end;
  end;
  {}
  if Self.Terminated then Exit;
  {}
  if ( aErrCount > 0 ) then
    FMsgSub.Warning( LanguageManager.GetFmt( 'SDbValidateSessionWarning', [aErrCount] ) );
  FLastVdSessionTime := Now;
  Sleep( 500 );
  FMsgSub.Info( LanguageManager.Get( 'SDbValidateSessionEnd' ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TSessionMonitorThread.UpdateSessionState;
var
  aIndex, aDbErrCount: Integer;
  aMgr: TDbSessionManager;
begin
  aDbErrCount := 0;
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    aMgr := FDbSessionManagerList.Manager[aIndex];
    aMgr.AppSo.DbState := dbOpen;
    if ( aMgr.SessionList.ErrSessionCount > 0 ) then
    begin
      aMgr.AppSo.DbState := dbError;
      Inc( aDbErrCount );
    end;
    FSoSub.Notify( aMgr.AppSo );
    if Self.Terminated then Break;
    Sleep( 200 );
  end;
  if Self.Terminated then Exit;
  if ( aDbErrCount <= 0 ) then
    FWarningSub.OK( EmptyStr )
  else
    FWarningSub.Warning( EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

end.
