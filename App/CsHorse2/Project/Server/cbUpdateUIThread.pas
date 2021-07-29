unit cbUpdateUIThread;

interface

uses
   SysUtils, Classes, Windows, Messages, DB, cbSyncObjs,
{App:}
  cbSo, cbSrvClass, cbDesignPattern, cbThread, cbLanguage, cbHrHelper,
{ODAC:}
  Ora, MemDS, VirtualTable,
{RemObject SDK:}
  CsHorse2Library_Intf;


type
  TUpdateUIThread = class(TBaseThread)
  private
    { Private declarations }
    FMsgSub: TMsgSubject;
    FUserSub: TDataSubject;
    FCommEnv: TCommEnv;
    FEvent: TSimpleEvent;
    FReader: TOraQuery;
    FWriter: TOraQuery;
    FOraSQL: TOraSQL;
    FDbSession: TDbSession;
    FDbSessionManagerList: TDbSessionManagerList;
    FUserBuffer: TVirtualTable;
    FGetUserTime: TDateTime;
    FSqlLoad: TStringList;
    function CheckCanGetUser: Boolean;
    function PrepareSql(const ASqlFileName: String): String;
    procedure PrepareSqlParam(const AIndex: Integer);
    function GetDbSession(AMgr: TDbSessionManager): Boolean;
    procedure ReleaseDbSession(AMgr: TDbSessionManager);
    procedure UpdateUserBuffer;
    procedure GetUserList;
    procedure ResetUserDbState;
    procedure UpdateUserState; overload;
    procedure Cleanup;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property CommEnv: TCommEnv read FCommEnv;
    property Event: TSimpleEvent read FEvent write FEvent;
    property MsgSubject: TMsgSubject read FMsgSub;
    property UserSubject: TDataSubject read FUserSub;
    property DbSessionManagerList: TDbSessionManagerList read FDbSessionManagerList write FDbSessionManagerList;
  end;


implementation

uses cbUtilis, DateUtils, cbROServerModule;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TUIThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ---------------------------------------------------------------------------- }

{ TUIThread }

constructor TUpdateUIThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FMsgSub := TMsgSubject.Create;
  FUserSub := TDataSubject.Create;
  FCommEnv := TCommEnv.Create;
  FReader := TOraQuery.Create( nil );
  FReader.FetchAll := True;
  FWriter := TOraQuery.Create( nil );
  FWriter.FetchAll := True;
  FOraSQL := TOraSQL.Create( nil );
  FUserBuffer := TVirtualTable.Create( nil );
  TBufferHelper.CreateFieldDefs( biCH005, FUserBuffer );
  FGetUserTime := 0;
  FSqlLoad := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TUpdateUIThread.Destroy;
begin
  FSqlLoad.Free;
  FMsgSub.Free;
  FUserSub.Free;
  FCommEnv.Free;
  FReader.Free;
  FWriter.Free;
  FOraSQL.Free;
  FEvent := nil;
  FUserBuffer.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TUpdateUIThread.CheckCanGetUser: Boolean;
begin
  Result := ( FGetUserTime <= 0 );
  if ( not Result ) then
  begin
    Result :=
      ( not Self.Terminated ) and
      ( SecondsBetween( Now, FGetUserTime ) >= FCommEnv.UIGetUserListFreq );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TUpdateUIThread.PrepareSql(const ASqlFileName: String): String;
var
  aPath: String;
begin
  Result := EmptyStr;
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
   Format( '%s\%s', [SQLFOLDER, ASqlFileName] );
  if not FileExists( aPath ) then
  begin
    FMsgSub.Error( LanguageManager.GetFmt( 'SDbSyncUserError', [EmptyStr,
      LanguageManager.GetFmt( 'SDbSqlScriptError', [aPath] )] ) );
    Exit;
  end;
  FSqlLoad.LoadFromFile( aPath );
  Result := FSqlLoad.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.PrepareSqlParam(const AIndex: Integer);
begin
  case AIndex of
    2:
      begin
        FOraSQL.ParamByName( 'ACompCode' ).DataType := ftInteger;
        FOraSQL.ParamByName( 'AUserId' ).DataType := ftString;
        FOraSQL.ParamByName( 'ASessionId' ).DataType := ftString;
        FOraSQL.ParamByName( 'AStatus' ).DataType := ftInteger;
        FOraSQL.ParamByName( 'AHostName' ).DataType := ftString;
        FOraSQL.ParamByName( 'AIp' ).DataType := ftString;
        FOraSQL.ParamByName( 'ATermSId' ).DataType := ftString;
        FOraSQL.ParamByName( 'ATermSName' ).DataType := ftString;
        FOraSQL.ParamByName( 'ATermSPC' ).DataType := ftString;
        FOraSQL.ParamByName( 'ATermSIp' ).DataType := ftString;
        FOraSQL.ParamByName( 'ATermState' ).DataType := ftString;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TUpdateUIThread.GetDbSession(AMgr: TDbSessionManager): Boolean;
begin
  Result := False;
  FDbSession := AMgr.LockFreeSession;
  if Assigned( FDbSession ) then
  begin
    FReader.Session := FDbSession.Connection;
    FWriter.Session := FDbSession.Connection;
    FOraSQL.Session := FDbSession.Connection;
    Result := True;
  end else
  begin
    FMsgSub.Warning( LanguageManager.GetFmt( 'SDbLockSessionError', [AMgr.AppSo.CompName] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.ReleaseDbSession(AMgr: TDbSessionManager);
begin
  if Assigned( FDbSession ) then
  begin
    AMgr.UnLockFreeSession( FDbSession.Index );
    FReader.Session := nil;
    FWriter.Session := nil;
    FOraSQL.Session := nil;
    FDbSession := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.UpdateUserBuffer;
begin
  FReader.First;
  while not FReader.Eof do
  begin
    FUserBuffer.Append;
    FUserBuffer.FieldByName( 'CompCode' ).AsString := FReader.FieldByName( 'CompCode' ).AsString;
    FUserBuffer.FieldByName( 'CompName' ).AsString := FReader.FieldByName( 'CompName' ).AsString;
    FUserBuffer.FieldByName( 'SessionId' ).AsString := FReader.FieldByName( 'SessionId' ).AsString;
    FUserBuffer.FieldByName( 'UserId' ).AsString := FReader.FieldByName( 'UserId' ).AsString;
    FUserBuffer.FieldByName( 'UserName' ).AsString := FReader.FieldByName( 'UserName' ).AsString;
    FUserBuffer.FieldByName( 'CompStr' ).AsString := FReader.FieldByName( 'CompStr' ).AsString;
    FUserBuffer.FieldByName( 'HostName' ).AsString := FReader.FieldByName( 'HostName' ).AsString;
    FUserBuffer.FieldByName( 'IP' ).AsString := FReader.FieldByName( 'IP' ).AsString;
    FUserBuffer.FieldByName( 'WorkClassCode' ).AsString := FReader.FieldByName( 'WorkClass' ).AsString;
    FUserBuffer.FieldByName( 'WorkClassName' ).AsString := FReader.FieldByName( 'WorkClassName' ).AsString;
    FUserBuffer.FieldByName( 'Status' ).AsString := FReader.FieldByName( 'Status' ).AsString;
    FUserBuffer.FieldByName( 'TermSId' ).AsString := FReader.FieldByName( 'TermSId' ).AsString;
    FUserBuffer.FieldByName( 'TermSName' ).AsString := FReader.FieldByName( 'TermSName' ).AsString;
    FUserBuffer.FieldByName( 'TermSPC' ).AsString := FReader.FieldByName( 'TermSPC' ).AsString;
    FUserBuffer.FieldByName( 'TermSIP' ).AsString := FReader.FieldByName( 'TermSIP' ).AsString;
    FUserBuffer.FieldByName( 'TermState' ).AsString := FReader.FieldByName( 'TermState' ).AsString;
    FUserBuffer.FieldByName( 'OnTime' ).AsString := FReader.FieldByName( 'OnTime' ).AsString;
    FUserBuffer.FieldByName( 'OffTime' ).AsString := FReader.FieldByName( 'OffTime' ).AsString;
    FUserBuffer.Post;
    FReader.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.GetUserList;
var
  aSql: String;
  aIndex: Integer;
  aMgr: TDbSessionManager;
begin
  TBufferHelper.CreateSqlText( biCH005, aSql );
  FReader.Sql.Text := aSql;
  FUserBuffer.Clear;
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    aMgr := FDbSessionManagerList.Manager[aIndex];
    if ( not aMgr.AppSo.Selected ) then Continue;
    if ( not GetDbSession( aMgr ) ) then Continue;
    try
      FReader.Open;
      UpdateUserBuffer;
    except
      on E: Exception do
      begin
        FReader.Session.Close;
        FMsgSub.Error( LanguageManager.GetFmt( 'SUIGetUserError', [AMgr.AppSo.CompName,
          E.Message] ) );
      end;
    end;
    FReader.Close;
    ReleaseDbSession( aMgr );
    if Self.Terminated then Break;
    Sleep( 100 );
  end;
  FGetUserTime := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.ResetUserDbState;
var
  aIndex: Integer;
  aMgr: TDbSessionManager;

  { ------------------------------------ }

  procedure DbUpdate;
  begin
    try
      FWriter.Close;
      FWriter.SQL.Text :=
        ' update ch005 set     ' +
        '   hostname = null,   ' +
        '   ip = null,         ' +
        '   sessionid = null,  ' +
        '   status = 0,        ' +
        '   termsid = null,    ' +
        '   termsname = null,  ' +
        '   termspc = null,    ' +
        '   termsip = null,    ' +
        '   termstate = null  ';
      FWriter.ExecSQL;
      FWriter.Session.Commit;
    except
      on E: Exception do
      begin
        FOraSQL.Session.Close;
        FMsgSub.Error( LanguageManager.GetFmt( 'SUIInitUserStateError', [aMgr.AppSo.CompName,
          E.Message] ) );
      end;
    end;
  end;

  { ------------------------------------ }

begin
  for aIndex := 0 to FDbSessionManagerList.Count - 1 do
  begin
    aMgr := FDbSessionManagerList.Manager[aIndex];
    try
      if GetDbSession( aMgr ) then DbUpdate;
    finally
      ReleaseDbSession( aMgr );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.UpdateUserState;
var
  aExecCount: Integer;
  aInfo: TLoginInfo;
  aMgr: TDbSessionManager;

  { -------------------------------------------------------------- }

  procedure DbUpdate;
  begin
    try
      FOraSQL.ParamByName( 'ACompCode' ).AsString := aInfo.CompCode;
      FOraSQL.ParamByName( 'AUserId' ).AsString := aInfo.UserId;
      FOraSQL.ParamByName( 'ASessionId' ).AsString := aInfo.SessionId;
      FOraSQL.ParamByName( 'AStatus' ).AsString := aInfo.Status;
      FOraSQL.ParamByName( 'AHostName' ).AsString := aInfo.HostName;
      FOraSQL.ParamByName( 'AIp' ).AsString := aInfo.IP;
      FOraSQL.ParamByName( 'ATermSId' ).AsString := aInfo.TermSId;
      FOraSQL.ParamByName( 'ATermSName' ).AsString := aInfo.TermSName;
      FOraSQL.ParamByName( 'ATermSPC' ).AsString := aInfo.TermSPC;
      FOraSQL.ParamByName( 'ATermSIp' ).AsString := aInfo.TermSIP;
      FOraSQL.ParamByName( 'ATermState' ).AsString := aInfo.TermState;
      FOraSQL.Execute;
      FOraSQL.Session.Commit;
    except
      on E: Exception do
      begin
        FOraSQL.Session.Close;
        FMsgSub.Error( LanguageManager.GetFmt( 'SUIGetUserError', [aMgr.AppSo.CompName,
          E.Message] ) );
      end;
    end;
  end;

  { -------------------------------------------------------------- }

begin
  aExecCount := 0;
  FOraSQL.Sql.Text := PrepareSql( 'UpdateUserSync.sql' );
  if ( FOraSQL.SQL.Text = EmptyStr ) then Exit;
  PrepareSqlParam( 2 ); 
  while ( LogUserBufferList.Count > 0 ) do
  begin
    aInfo := TLoginInfo( LogUserBufferList.Objects[0] );
    if Assigned( aInfo ) then
    begin
      aMgr := FDbSessionManagerList.Manager[FDbSessionManagerList.IndexOf( aInfo.CompCode )];
      try
        if GetDbSession( aMgr ) then DbUpdate;
      finally
        ReleaseDbSession( aMgr );
      end;
      aInfo.Free;
    end;
    LogUserBufferList.Delete( 0 );
    Inc( aExecCount );
    if ( aExecCount >= 30 ) then Break;
    if ( Self.Terminated ) then Break;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.Execute;
var
  aIsFirstRun: Boolean;
begin
  aIsFirstRun := True;
  FMsgSub.OK( LanguageManager.GetFmt( 'SUICreateThreadEnd', [HexThreadId] ) );
  Sleep( 1000 );
  while not Self.Terminated do
  begin
    if ( FEvent.WaitFor( 1000 ) = wrSignaled ) then
    begin
      if aIsFirstRun then
      begin
        ResetUserDbState;
        aIsFirstRun := False;
      end;
      Sleep( 50 );
      if Self.Terminated then Break;
      {}
      if ( CheckCanGetUser ) then
      begin
        GetUserList;
        FUserSub.Notify( FUserBuffer );
      end;
      Sleep( 50 );
      if Self.Terminated then Break;
      {}
      UpdateUserState;
      Sleep( 50 );
      if Self.Terminated then Break;
      {}
    end;
    Sleep( 50 );
    if Self.Terminated then Break;
  end;
  ResetUserDbState;
  GetUserList;
  FUserSub.Notify( FUserBuffer );
  Sleep( 100 );
  Cleanup;
  Sleep( 500 );
  FMsgSub.Warning( LanguageManager.GetFmt( 'SUIThreadStop', [HexThreadId] ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TUpdateUIThread.Cleanup;
begin
  while ( LogUserBufferList.Count > 0 ) do
  begin
    TLoginInfo( LogUserBufferList.Objects[0] ).Free;
    LogUserBufferList.Delete( 0 );
  end;
end;
{ ---------------------------------------------------------------------------- }

end.
