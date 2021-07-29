unit cbROServerModule;

interface

uses
  SysUtils, Classes, Variants, IniFiles,
{App:}
  cbDesignPattern, cbSrvClass, cbLanguage,
{RemObjects SDK:}
  uROSessions, uROClient, uROBinMessage, uROServer, uROSuperTCPServer,
  uROClassFactories, uROEventRepository;

type
  TROServerModule = class(TDataModule)
    ROSuperTcpServer: TROSuperTcpServer;
    ROSessionManager: TROInMemorySessionManager;
    ROServerBinMessage: TROBinMessage;
    ROSrvCallbackEventRepository: TROInMemoryEventRepository;
    ROCallbackEventBinMessage: TROBinMessage;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FLoginMsgSubject: TMsgSubject;
    FLoginSrvFactory: TROClassFactory;
    FAnnSrvFactory: TROClassFactory;
    FCallbackSrvFactory: TROClassFactory;
    //FROPooledClassFactory: TROPooledClassFactory;
    FSoList: TAppSoList;
    FDbSessionManagerList: TDbSessionManagerList;
    FClientEnv: TClientEnv;
    FMsgIdGenerator: TMsgIdGenerator;
    procedure SetSoList(const ASource: TAppSoList);
    procedure Cleanup;
  public
    { Public declarations }
    function GetCompName(const ACompCode: String): String;
    function GetDbSession(const ACompCode: String): TDbSession;
    procedure ReleaseDbSession(const ACompCode: String; const ASession: TDbSession);
    procedure ServiceStart;
    procedure ServiceStop;
    property LoginMsgSubject: TMsgSubject read FLoginMsgSubject;
    property SoList: TAppSoList read FSoList write SetSoList;
    property DbSessionManagerList: TDbSessionManagerList read FDbSessionManagerList write FDbSessionManagerList;
    property ClientEnv: TClientEnv read FClientEnv;
    property MsgIdGenerator: TMsgIdGenerator read FMsgIdGenerator;
  end;

var
  ROServerModule: TROServerModule;

  LogUserBufferList: TSyncList;
  UserList: TSyncList;
  AppIni: TIniFile;


implementation

uses
  cbUtilis,
  LoginService_Impl, AnnService_Impl, CallbackService_Impl,
  CsHorse2Library_Invk, CsHorse2Library_Intf;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure Create_LoginService(out anInstance : IUnknown);
begin
  anInstance := TLoginService.Create(nil);
end;

{ ---------------------------------------------------------------------------- }

procedure Create_AnnService(out anInstance : IUnknown);
begin
  anInstance := TAnnService.Create(nil);
end;

{ ---------------------------------------------------------------------------- }

procedure Create_CallbackService(out anInstance : IUnknown);
begin
  anInstance := TCallbackService.Create(nil);
end;

{ ---------------------------------------------------------------------------- }

function GenSavedId(ASavedId: String): String;
begin
  if ( ASavedId = EmptyStr ) or ( Length( ASavedId ) <> 12 ) then
    Result := GetSysDate + IntToHex( 1, 4 )
  else
    Result := ASavedId;
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.DataModuleCreate(Sender: TObject);
var
  aPrefix, aFullId, aId: String;
begin
  FLoginMsgSubject := TMsgSubject.Create;
  LogUserBufferList := TSyncList.Create;
  UserList := TSyncList.Create;
  FSoList := TAppSoList.Create;
  FClientEnv := TClientEnv.Create;
  aFullId := GenSavedId( AppIni.ReadString( 'IdGenerator', 'MsgId', EmptyStr ) );
  aPrefix := Copy( aFullId, 1, 8 );
  aId :=  Copy( aFullId, 9, 4 );
  FMsgIdGenerator := TMsgIdGenerator.Create( True,
    TIdRange.Create( 1, 65536,  StrToInt( '$' + aId ) ) );
  FMsgIdGenerator.IdPrefix := aPrefix;  
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.DataModuleDestroy(Sender: TObject);
begin
  Cleanup;
  FLoginMsgSubject.Free;
  LogUserBufferList.Free;
  UserList.Free;
  FSoList.Free;
  FClientEnv.Free;
  AppIni.WriteString( 'IdGenerator', 'MsgId', FMsgIdGenerator.CurrentValue );
  FMsgIdGenerator.Free;
end;

{ ---------------------------------------------------------------------------- }

function TROServerModule.GetCompName(const ACompCode: String): String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  aIndex := FSoList.IndexOf( ACompCode );
  if ( aIndex >= 0 ) then Result := FSoList.Items[aIndex].CompName;
end;

{ ---------------------------------------------------------------------------- }

function TROServerModule.GetDbSession(const ACompCode: String): TDbSession;
var
  AMgr: TDbSessionManager;
begin
  Result := nil;
  AMgr := FDbSessionManagerList.Manager[FDbSessionManagerList.IndexOf( ACompCode )];
  if Assigned( AMgr ) then
  begin
    Result := AMgr.LockFreeSession;
    if not Assigned( Result ) then
      LoginMsgSubject.Warning( LanguageManager.GetFmt( 'SDbLockSessionError', [AMgr.AppSo.CompName] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.ReleaseDbSession(const ACompCode: String; const ASession: TDbSession);
var
  AMgr: TDbSessionManager;
begin
  AMgr := FDbSessionManagerList.Manager[FDbSessionManagerList.IndexOf( ACompCode )];
  AMgr.UnLockFreeSession( ASession.Index );
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.ServiceStart;
begin
  //FROPooledClassFactory := TROPooledClassFactory.Create( 'LoginService', Create_LoginService,
  //  TLoginService_Invoker, 100, pbWait, True );
  Cleanup;
  FLoginSrvFactory := TROClassFactory.Create( 'LoginService', Create_LoginService, TLoginService_Invoker );
  FAnnSrvFactory := TROClassFactory.Create( 'AnnService', Create_AnnService, TAnnService_Invoker );
  FCallbackSrvFactory := TROClassFactory.Create( 'CallbackService', Create_CallbackService, TCallbackService_Invoker );
  ROSuperTcpServer.Active := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.ServiceStop;
begin
  Cleanup;
  ROSuperTcpServer.Active := False;
  //FROPooledClassFactory.ClearPool;
  //UnRegisterClassFactory( FROPooledClassFactory );
  UnRegisterClassFactory( FLoginSrvFactory );
  UnRegisterClassFactory( FAnnSrvFactory );
  UnRegisterClassFactory( FCallbackSrvFactory );
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.SetSoList(const ASource: TAppSoList);
begin
  FSoList.Clear;
  FSoList.Assign( ASource );
end;

{ ---------------------------------------------------------------------------- }

procedure TROServerModule.Cleanup;
var
  aIndex: Integer;
begin
  for aIndex := 0 to LogUserBufferList.Count - 1 do
  begin
    if Assigned( LogUserBufferList.Objects[aIndex] ) then
      TLoginInfo( LogUserBufferList.Objects[aIndex] ).Free;
  end;
  LogUserBufferList.Clear;
  {}
  for aIndex := 0 to UserList.Count - 1 do
  begin
    if Assigned( UserList.Objects[aIndex] ) then
      TLoginInfo( UserList.Objects[aIndex] ).Free;
  end;
  UserList.Clear;
end;

{ ---------------------------------------------------------------------------- }

end.
