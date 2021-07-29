unit cbROClientModule;

interface

uses
  Windows, Messages, SysUtils, Classes,
{ App: }
  cbSo, cbHrHelper,
{ ODAC: }
  MemDS, VirtualTable,
{ RemObject: }
  uROSuperTCPChannel, uRORemoteService, uROTypes, uROBinMessage, uROClient,
  CsHorse2Library_Intf;

type
  TROClientModule = class(TDataModule)
    ROSuperTcpChannel: TROSuperTcpChannel;
    ROLoginSrvBinMessage: TROBinMessage;
    ROAnnSrvBinMessage: TROBinMessage;
    ROCallbackSrvBinMessage: TROBinMessage;
    ROCallbackSrv2BinMessage: TROBinMessage;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FLoginService: ILoginService;
    FAnnService: IAnnService;
    FCallbackService: ICallbackService;
    FCallbackService2: ICallbackService;
    FLoginInfo: TLoginInfo;
    FClientParam: TVirtualTable;
    FClientParamLock: TMREWSync;
    FGroupList: TVirtualTable;
    FUserList: TVirtualTable;
    procedure InitLoginInfo;
  public
    { Public declarations }
    procedure Logout;
    procedure UpdateClientParam;
    procedure BeginLockClientParam;
    procedure EndLockClientParam;
    function Login(const ACompCode, AUserId, APassword: String; var aMsg: String): Boolean;
    function GetSoList: TSoList;
    function GetSO021: Binary;
    function GetCD042: Binary;
    function GetSO022: Binary;
    function GetSO023: Binary;
    function GetGroupList: Binary;
    function GetUserList: Binary;
    function SendMsg(ARecver: TVirtualTable; AMsg: Binary; AMsgInfo: TMsgInfo; var AErrMsg: String): Boolean;
    function GetMsgList: Binary;
    function GetMsg(const AMsgInfo: TMsgInfo): Binary;
    procedure SetMsgRead(const AMsgInfo: TMsgInfo);
    property LoginInfo: TLoginInfo read FLoginInfo;
    property ClientParam: TVirtualTable read FClientParam;
    property GroupList: TVirtualTable read FGroupList;
    property UserList: TVirtualTable read FUserList;
  end;

var
  ROClientModule: TROClientModule;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.DataModuleCreate(Sender: TObject);
begin
  FLoginInfo := TLoginInfo.Create;
  InitLoginInfo;
  try
    FLoginService := CoLoginService.Create( ROLoginSrvBinMessage, ROSuperTcpChannel );
  except
    {}
  end;
  try
    FAnnService := CoAnnService.Create( ROAnnSrvBinMessage, ROSuperTcpChannel );
  except
    {}
  end;
  try
    FCallbackService := CoCallbackService.Create( ROCallbackSrvBinMessage, ROSuperTcpChannel );
  except
    {}
  end;
  try
    FCallbackService2 := CoCallbackService.Create( ROCallbackSrv2BinMessage, ROSuperTcpChannel );
  except
    {}
  end;
  FClientParam := TVirtualTable.Create( nil );
  TBufferHelper.CreateFieldDefs( biClientParam, FClientParam );
  FClientParamLock := TMREWSync.Create;
  FGroupList := TVirtualTable.Create( nil );
  FGroupList.Open;
  FUserList := TVirtualTable.Create( nil );
  FUserList.Open;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.DataModuleDestroy(Sender: TObject);
begin
  FLoginInfo.Free;
  FClientParam.Free;
  FClientParamLock.Free;
  FGroupList.Free;
  FUserList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.InitLoginInfo;
var
  aTermId, aTermIp, aTermName, aTermPC, aTermState: String;
begin
  FLoginInfo.SessionId := EmptyStr;
  FLoginInfo.HostName := GetComputerName;
  FLoginInfo.IP := GetIpAddress;
  if IsRunInTerminalServer then
  begin
    GetTerminalServerInfo( aTermId, aTermIp, aTermName, aTermPC, aTermState );
    FLoginInfo.TermSId := aTermId;
    FLoginInfo.TermSIP := aTermIP;
    FLoginInfo.TermSName := aTermName;
    FLoginInfo.TermSPC := aTermPC;
    FLoginInfo.TermState := aTermState;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.Login(const ACompCode, AUserId, APassword: String; var aMsg: String): Boolean;
begin
  aMsg := EmptyStr;
  FLoginInfo.CompCode := ACompCode;
  FLoginInfo.UserId := AUserId;
  FLoginInfo.Password := APassword;
  FLoginInfo.Status := '1';
  try
    Result := FLoginService.Login( FLoginInfo, aMsg );
  except
    on E: Exception do
    begin
      aMsg := E.Message;
      Result := False;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.Logout;
begin
  FLoginInfo.Status := '0';
  FLoginService.Logout( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetSoList: TSoList;
var
  aValue, aSoText: String;
  aPos: Integer;
  aSo: TSo;
begin
  aValue := FAnnService.GetSOListText( FLoginInfo );
  Result := TSoList.Create;
  while ( aValue <> EmptyStr ) do
  begin
    aSoText := ExtractValue( aValue );
    aPos := Pos( '@', aSoText );
    aSo := TSo.Create;
    aSo.CompCode := Copy( aSoText, 1, aPos - 1 );
    aSo.CompName := Copy( aSoText, aPos + 1, Length( aSoText ) - aPos );
    Result.Add( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetSO021: Binary;
begin
  Result := FAnnService.GetSO021( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetCD042: Binary;
begin
  Result := FAnnService.GetCD042( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetSO022: Binary;
begin
  Result := FAnnService.GetSO022( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetSO023: Binary;
begin
  Result := FAnnService.GetSO023( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.BeginLockClientParam;
begin
  FClientParamLock.BeginRead;
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.EndLockClientParam;
begin
  FClientParamLock.EndRead;
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.UpdateClientParam;
var
  aBinary: Binary;
begin
  aBinary := FLoginService.GetClientParam;
  if Assigned( aBinary ) then
  begin
    FClientParamLock.BeginWrite;
    try
      FClientParam.LoadFromStream( aBinary );
    finally
      FClientParamLock.EndWrite;
    end;
    aBinary.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetGroupList: Binary;
begin
  Result := FCallbackService.GetGroupList( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetUserList: Binary;
begin
  Result := FCallbackService.GetUserList( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.SendMsg(ARecver: TVirtualTable; AMsg: Binary; AMsgInfo: TMsgInfo;
  var AErrMsg: String): Boolean;
var
  aRecvBin: Binary;
begin
  aRecvBin := Binary.Create;
  try
    ARecver.SaveToStream( aRecvBin );
    FCallbackService.SendMsg( ROClientModule.LoginInfo, aRecvBin, AMsg, AMsgInfo, AErrMsg );
  finally
    aRecvBin.Free;
    AMsg.Free;
    AMsgInfo.Free;
  end;
  Result := ( AErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetMsgList: Binary;
begin
  Result := FCallbackService.GetMsgList( ROClientModule.LoginInfo );
end;

{ ---------------------------------------------------------------------------- }

function TROClientModule.GetMsg(const AMsgInfo: TMsgInfo): Binary;
begin
  Result := FCallbackService2.GetMsg( ROClientModule.LoginInfo, AMsgInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TROClientModule.SetMsgRead(const AMsgInfo: TMsgInfo);
begin
  FCallbackService2.SetMsgRead( ROClientModule.LoginInfo, AMsgInfo );
end;

{ ---------------------------------------------------------------------------- }

end.
