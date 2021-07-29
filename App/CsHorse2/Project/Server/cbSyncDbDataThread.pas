unit cbSyncDbDataThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, ADODB, 
  {$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF} cbSyncObjs,
{ App }
  cbSo, cbSrvClass, cbDesignPattern, cbThread, cbLanguage,
{ ODAC }
  Ora;


type
  TSyncDbDataThread = class(TBaseThread)
  private
    { Private declarations }
    FCommEnv: TCommEnv;
    FMsgSub: TMsgSubject;
    FSoSub: TSoSubject;
    FSyncUserTime: TDateTime;
    FOraSql: TOraSQL;
    FDbSession: TDbSession;
    FDbSessionManagerList: TDbSessionManagerList;
    FEvent: TSimpleEvent;
    function CheckCanSyncUser: Boolean;
    function GetDbSession(AMgr: TDbSessionManager): Boolean;
    procedure ReleaseDbSession(AMgr: TDbSessionManager);
    procedure SyncUser(AMgr: TDbSessionManager); overload;
    procedure SyncUser; overload;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property CommEnv: TCommEnv read FCommEnv;
    property Event: TSimpleEvent read FEvent write FEvent;
    property MsgSubject: TMsgSubject read FMsgSub;
    property SoSubject: TSoSubject read FSoSub;
    property DbSessionManagerList: TDbSessionManagerList read FDbSessionManagerList write FDbSessionManagerList;
  end;


implementation

uses cbUtilis, DateUtils;

const SYNC_USER = 'SyncUser.sql';

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

constructor TSyncDbDataThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FCommEnv := TCommEnv.Create;
  FMsgSub := TMsgSubject.Create;
  FSoSub := TSoSubject.Create;
  FOraSql := TOraSQL.Create( nil );
  FOraSql.AutoCommit := False;
  FDbSession := nil;
  FSyncUserTime := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TSyncDbDataThread.Destroy;
begin
  FEvent := nil;
  FOraSql.Session := nil;
  FOraSql.Free;
  FMsgSub.Free;
  FSoSub.Free;
  FCommEnv.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TSyncDbDataThread.GetDbSession(AMgr: TDbSessionManager): Boolean;
begin
  Result := False;
  FDbSession := AMgr.LockFreeSession;
  if Assigned( FDbSession ) then
  begin
    FOraSql.Session := FDbSession.Connection;
    Result := True;
  end else
  begin
    FMsgSub.Warning( LanguageManager.GetFmt( 'SDbLockSessionError', [AMgr.AppSo.CompName] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncDbDataThread.ReleaseDbSession(AMgr: TDbSessionManager);
begin
  if Assigned( FDbSession ) then
  begin
    AMgr.UnLockFreeSession( FDbSession.Index );
    FOraSql.Session := nil;
    FDbSession := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncDbDataThread.CheckCanSyncUser: Boolean;
begin
  Result := ( FSyncUserTime <= 0 );
  if ( not Result ) then
  begin
    Result :=
      ( not Self.Terminated ) and
      ( SecondsBetween( Now, FSyncUserTime ) >= FCommEnv.DbSyncUser );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncDbDataThread.SyncUser(AMgr: TDbSessionManager);
var
  aPath: String;
  aDelete, aInsert, aUpdate: Integer;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    Format( '%s\%s', [SQLFOLDER, SYNC_USER] );
  if not FileExists( aPath ) then
  begin
    FMsgSub.Error( LanguageManager.GetFmt( 'SDbSyncUserError', [AMgr.AppSo.CompName,
      LanguageManager.GetFmt( 'SDbSqlScriptError', [aPath] )] ) );
    Exit;
  end;
  FOraSql.SQL.Clear;
  FOraSql.SQL.LoadFromFile( aPath );
  try
    FOraSql.ParamByName( 'ADelete' ).DataType := ftInteger;
    FOraSql.ParamByName( 'AInsert' ).DataType := ftInteger;
    FOraSql.ParamByName( 'AUpdate' ).DataType := ftInteger;
    {}
    FOraSql.Execute;
    FOraSql.Session.Commit;
    {}
    aDelete := FOraSQL.ParamByName( 'ADelete' ).AsInteger;
    aInsert := FOraSQL.ParamByName( 'AInsert' ).AsInteger;
    aUpdate := FOraSQL.ParamByName( 'AUpdate' ).AsInteger;
    {}
    if ( aDelete + aInsert + aUpdate ) > 0 then
    begin
      FMsgSub.OK( LanguageManager.GetFmt( 'SDbSyncUserSuccessWithRecords',
        [AMgr.AppSo.CompName, aDelete, aInsert, aUpdate] ) );
    end;
  except
    on E: Exception do
    begin
      FOraSql.Session.Close;
      FMsgSub.Error( LanguageManager.GetFmt( 'SDbSyncUserError',
        [AMgr.AppSo.CompName, E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncDbDataThread.SyncUser;
var
  aIndex: Integer;
  aMgr: TDbSessionManager;
begin
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    aMgr := FDbSessionManagerList.Manager[aIndex];
    if ( aMgr.AppSo.Selected ) then
    begin
      try
        if GetDbSession( aMgr ) then
          SyncUser( aMgr )
      finally
        ReleaseDbSession( aMgr );
      end;
    end;
    if Self.Terminated then Break;
    Sleep( 100 );
  end;
  FSyncUserTime := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncDbDataThread.Execute;
begin
  FMsgSub.OK( LanguageManager.GetFmt( 'SDbCreateSyncThreadEnd', [HexThreadId] ) );
  Sleep( 1000 );
  while not Self.Terminated do
  begin
    if ( FEvent.WaitFor( 1000 ) = wrSignaled ) then
    begin
      if ( CheckCanSyncUser ) then
        SyncUser;
    end;
    Sleep( 100 );
    if ( Self.Terminated ) then Break;
  end;
  Sleep( 500 );
  FMsgSub.Warning( LanguageManager.GetFmt( 'SDbSyncThreadStop', [HexThreadId] ) );
end;

{ ---------------------------------------------------------------------------- }

end.
