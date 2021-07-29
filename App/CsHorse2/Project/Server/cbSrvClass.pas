unit cbSrvClass;

interface

uses
  Windows, SysUtils, Variants, Classes, Controls, Forms, DB, cbSyncObjs,
{ App }
  cbSo, cbDesignPattern,
{ ODAC }
  DBAccess, Ora, MemDS, VirtualTable;

type

  { TAppSo }

  TAppSo = class(TSo)
  private
    FData: Pointer;
    FOnlines: Integer;
    FOfflines: Integer;
    FSynData: Boolean;
    FMaxSession: Integer;
    FFreeSession: Integer;
    FBusySession: Integer;
    FErrSession: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    property Onlines: Integer read FOnlines write FOnlines;
    property Offlines: Integer read FOfflines write FOfflines;
    property SynData: Boolean read FSynData write FSynData;
    property MaxSession: Integer read FMaxSession write FMaxSession;
    property FreeSession: Integer read FFreeSession write FFreeSession;
    property BusySession: Integer read FBusySession write FBusySession;
    property ErrSession: Integer read FErrSession write FErrSession;
    property Data: Pointer read FData write FData;
  end;

  { TAppSoList }

  TAppSoList = class(TSoList)
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TAppSo;
    procedure SetItem(Index: Integer; aObj: TAppSo);
  public
    procedure Assign(ASource: TAppSoList);
    property Items[Index: Integer]: TAppSo read GetItem write SetItem; default;
  end;


  { TSoSubject }

  TSoSubject = class(TSubject)
  private
    FAppSo: TAppSo;
  public
    procedure Notify(AAppSo: TAppSo);
    property AppSo: TAppSo read FAppSo;
  end;


  { TCommEnv }

  TCommEnv = class(TPersistent)
  private
    FDbRetryFreq: Integer;
    FDbSyncUser: Integer;
    FDbValidateSessionFreq: Integer;
    FDbGetPoolObjectRertyCount: Integer;
    FUIGetUserListFreq: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    { 當資料庫斷線時, 每多少秒重試連線 }
    property DbRetryFreq: Integer read FDbRetryFreq write FDbRetryFreq;
    { 多久同步使用者列表 SO026 --> CH005 }
    property DbSyncUser: Integer read FDbSyncUser write FDbSyncUser;
    { 多久檢查一次 Session Pool 裏的資料庫連線是否正常 }
    property DbVaildateSessionFreq: Integer read FDbValidateSessionFreq write FDbValidateSessionFreq;
    { 嘗試鎖定資料庫集區物件重試次數 }
    property DbGetPoolObjectRertyCount: Integer read FDbGetPoolObjectRertyCount write FDbGetPoolObjectRertyCount;
    { 多久同步一次 Sevrer 程式 UI 顯示的始用者列表 }
    property UIGetUserListFreq: Integer read FUIGetUserListFreq write FUIGetUserListFreq;
  end;


  { TClientEnv }

  TClientEnv = class(TPersistent)
  private
    FAutoRefresh: Boolean;
    FAuthorizeRefreshRate: Integer;
    FAnnRefreshRate: Integer;
    FUserRefreshRate: Integer;
    FTryReconnectRate: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    { 客戶端是自動或手動更新資料 }
    property AutoRefresh: Boolean read FAutoRefresh write FAutoRefresh;
    { 當為自動更新資料時, 多久更新一次權限資料 }
    property AuthorizeRefreshRate: Integer read FAuthorizeRefreshRate write FAuthorizeRefreshRate;
    { 當為自動更新資料時, 多久更新一次公佈欄資料 }
    property AnnRefreshRate: Integer read FAnnRefreshRate write FAnnRefreshRate;
    { 當為自動更新資料時, 多久更新一次線上使用者清單 }
    property UserRefreshRate: Integer read FUserRefreshRate write FUserRefreshRate;
    { 客戶端當斷線時, 每多少秒自動重試連線 }
    property TryReconnectRate: Integer read FTryReconnectRate write FTryReconnectRate;
  end;


  { TSessionInfo }

  TDbSession = class(TObject)
  private
    FAppSo: TAppSo;
    FConnection: TOraSession;
    FLastActivity: TDateTime;
    FIsBusy: Boolean;
    FIsError: Boolean;
    FIndex: Integer;
    function GetIdleTime: Int64;
    function GetIsFree: Boolean;
  public
    constructor Create(AAppSo: TAppSo);
    destructor Destroy; override;
    property AppSo: TAppSo read FAppSo;
    property Connection: TOraSession read FConnection;
    property LastActivity: TDateTime read FLastActivity write FLastActivity;
    property IdleTime: Int64 read GetIdleTime;
    property IsFree: Boolean read GetIsFree;
    property IsBusy: Boolean read FIsBusy write FIsBusy;
    property IsError: Boolean read FIsError write FIsError;
    property Index: Integer read FIndex write FIndex;
  end;


  { TSessionInfoList }

  TDbSessionList = class(TList)
  private
    FLock: TCriticalSection;
    FAppSo: TAppSo;
    FBusySessionCount: Integer;
    FErrSessionCount: Integer;
    FFreeSessionCount: Integer;
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetDbSession(Index: Integer): TDbSession;
  public
    constructor Create(AAppSo: TAppSo);
    destructor Destroy; override;
    procedure UpdateSessionCount;
    procedure Insert(Index: Integer; ASession: TDbSession);
    function Add(ASession: TDbSession): Integer;
    function Extract(ASession: TDbSession): TDbSession;
    function Remove(ASession: TDbSession): Integer;
    function IndexOf(ASession: TDbSession): Integer; overload;
    function First: TDbSession;
    function Last: TDbSession;
    function GetFreeSessionLock(AIndex: Integer): Boolean;
    function GetErrorSessionLock(AIndex: Integer): Boolean;
    procedure ReleaseSessionLock(AInfo: TDbSession); overload;
    procedure ReleaseSessionLock(AIndex: Integer); overload;
    property DbSession[Index: Integer]: TDbSession read GetDbSession;
    property BusySessionCount: Integer read FBusySessionCount;
    property ErrSessionCount: Integer read FErrSessionCount;
    property FreeSessionCount: Integer read FFreeSessionCount;
  end;


  { TSessionManager }

  TDbSessionManager = class(TObject)
  private
    FAppSo: TAppSo;
    FCommEnv: TCommEnv;
    FSemaphore: THandle;
    FSessionList: TDbSessionList;
    procedure SetCommEnv(const ACommEnv: TCommEnv);
  public
    constructor Create(AAppSo: TAppSo);
    destructor Destroy; override;
    function LockFreeSession: TDbSession;
    procedure UnLockFreeSession(ASessionInfo: TDbSession); overload;
    procedure UnLockFreeSession(AIndex: Integer); overload;
    property AppSo: TAppSo read FAppSo;
    property CommEnv: TCommEnv read FCommEnv write SetCommEnv;
    property SessionList: TDbSessionList read FSessionList;
  end;


  { TSessionManagerList }

  TDbSessionManagerList = class(TList)
  private
    FCommEnv: TCommEnv;
    function GetManager(Index: Integer): TDbSessionManager;
    procedure SetCommEnv(const ACommEnv: TCommEnv);
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(AMgr: TDbSessionManager): Integer; overload;
    function Add(AAppSo: TAppSo): Integer; overload;
    function IndexOf(AMgr: TDbSessionManager): Integer; overload;
    function IndexOf(ACompCode: String): Integer; overload;
    property Manager[Index: Integer]: TDbSessionManager read GetManager;
    property CommEnv: TCommEnv read FCommEnv write SetCommEnv;
  end;


  { TIdRule }

  TIdRange = class(TObject)
  private
    FMinValue: Integer;
    FMaxValue: Integer;
    FCurValue: Integer;
  protected
    function GetMinValue: Variant; virtual;
    function GetMaxValue: Variant; virtual;
    function GetNextValue: Variant; virtual;
    function GetCurValue: Variant; virtual;
    procedure ResetValue; virtual;
  public
    constructor Create(AMinValue, AMaxValue, ACurValue: Integer); overload;
    destructor Destroy; override;
  end;


  { TIdGenerator }

  TCustomIdGenerator = class(TObject)
  protected
    FLock: TCriticalSection;
    FThreadSafe: Boolean;
    FRange: TIdRange;
    function GetMinValue: Variant; virtual;
    function GetMaxValue: Variant; virtual;
    function GetNextValue: Variant; virtual;
    function GetCurValue: Variant; virtual;
    procedure Reset;
  public
    constructor Create(const AThreadSafe: Boolean; ARange: TIdRange);
    destructor Destroy; override;
    property MinValue: Variant read GetMinValue;
    property MaxValue: Variant read GetMaxValue;
    property NextValue: Variant read GetNextValue;
    property CurrentValue: Variant read GetCurValue;
  end;

  { TMsgIdGenerator }

  TMsgIdGenerator = class(TCustomIdGenerator)
  private
    FIdPrefix: String;
    procedure SetIdPrefix(const AValue: String);
  protected
    function GetMinValue: Variant; override;
    function GetMaxValue: Variant; override;
    function GetNextValue: Variant; override;
    function GetCurValue: Variant; override;
  public
    property IdPrefix: String read FIdPrefix write SetIdPrefix;  
  end;


  TAppRunState = ( arIdle, arRun );
  var AppRunState: TAppRunState;
  var DbSessionManagerList: TDbSessionManagerList;


  const
    SQLFOLDER = 'SqlScript';


implementation


uses cbUtilis, DateUtils;

{ ---------------------------------------------------------------------------- }

{ TAppSo }

constructor TAppSo.Create;
begin
  inherited Create;
  FSynData := False;
  FOnlines := 0;
  FOfflines := 0;
  FMaxSession := 0;
  FFreeSession := 0;
  FBusySession := 0;
  FErrSession := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSo.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TAppSo ) then
  begin
    TAppSo( Dest ).Data := FData;
    TAppSo( Dest ).Onlines := FOnlines;
    TAppSo( Dest ).Offlines := FOfflines;
    TAppSo( Dest ).SynData := FSynData;
    TAppSo( Dest ).MaxSession := FMaxSession;
    TAppSo( Dest ).FreeSession := FFreeSession;
    TAppSo( Dest ).BusySession := FBusySession;
    TAppSo( Dest ).ErrSession := FErrSession;
  end;
  inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TAppSoList }

procedure TAppSoList.Assign(ASource: TAppSoList);
var
  aIndex: Integer;
  aSo: TAppSo;
begin
  if not Assigned( ASource ) then Exit;
  Self.Clear;
  for aIndex := 0 to ASource.Count - 1 do
  begin
    aSo := TAppSo.Create;
    aSo.Assign( ASource[aIndex] );
    Self.Add( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSoList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
  begin
    TAppSo( Ptr ).Free;
    Ptr := nil;
  end;
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TAppSoList.GetItem(Index: Integer): TAppSo;
begin
  Result := TAppSo( inherited Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

procedure TAppSoList.SetItem(Index: Integer; aObj: TAppSo);
begin
  inherited Items[Index] := aObj;
end;

{ ---------------------------------------------------------------------------- }

{ TSoSubject }

procedure TSoSubject.Notify(AAppSo: TAppSo);
begin
  FAppSo := AAppSo;
  if Assigned( FAppSo ) then inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TCommEnv }

procedure TCommEnv.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TCommEnv ) then
  begin
    TCommEnv( Dest ).DbRetryFreq := FDbRetryFreq;
    TCommEnv( Dest ).DbSyncUser := FDbSyncUser;
    TCommEnv( Dest ).DbVaildateSessionFreq := FDbValidateSessionFreq;
    TCommEnv( Dest ).UIGetUserListFreq := FUIGetUserListFreq;
    TCommEnv( Dest ).DbGetPoolObjectRertyCount := FDbGetPoolObjectRertyCount;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TClientEnv }

procedure TClientEnv.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TClientEnv ) then
  begin
    TClientEnv( Dest ).AutoRefresh := FAutoRefresh;
    TClientEnv( Dest ).AuthorizeRefreshRate := FAuthorizeRefreshRate;
    TClientEnv( Dest ).AnnRefreshRate := FAnnRefreshRate;
    TClientEnv( Dest ).UserRefreshRate := FUserRefreshRate;
    TClientEnv( Dest ).TryReconnectRate := FTryReconnectRate;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TSessionInfo }

constructor TDbSession.Create(AAppSo: TAppSo);
begin
  FAppSo := AAppSo;
  FConnection := TOraSession.Create( nil );
  FConnection.LoginPrompt := False;
  FConnection.ConnectString := Format( '%s/%s@%s',
    [FAppSo.DbUserId, FAppSo.DbUserPass, FAppSo.DbAliase] );
  FLastActivity := 0;
  FIsError := True;
  FIsBusy := False;
  FIndex := -1;
end;

{ ---------------------------------------------------------------------------- }

destructor TDbSession.Destroy;
begin
  FAppSo := nil;
  if FConnection.Connected then FConnection.Close;
  FConnection.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TDbSession.GetIdleTime: Int64;
begin
  Result := SecondsBetween( Now, FLastActivity );
end;

{ ---------------------------------------------------------------------------- }

function TDbSession.GetIsFree: Boolean;
begin
  Result := ( FIsError = False ) and ( FIsBusy = False ) and ( FConnection.Connected );
end;

{ ---------------------------------------------------------------------------- }

{ TSessionList }

constructor TDbSessionList.Create(AAppSo: TAppSo);
var
  aIndex, aIdx: Integer;
begin
  FAppSo := AAppSo;
  for aIndex := 1 to FAppSo.MaxSession do
  begin
    aIdx := Self.Add( TDbSession.Create( FAppSo ) );
    Self.DbSession[aIdx].Index := aIdx;
  end;
  FLock := TCriticalSection.Create;
  FBusySessionCount := 0;
  FErrSessionCount := 0;
  FFreeSessionCount := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TDbSessionList.Destroy;
begin
  FAppSo := nil;
  FLock.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
  begin
    TDbSession( Ptr ).Free;
    Ptr := nil;
  end;
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.GetDbSession(Index: Integer): TDbSession;
begin
  Result := TDbSession( inherited Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionList.Insert(Index: Integer; ASession: TDbSession);
begin
  inherited Insert( Index, ASession );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.Add(ASession: TDbSession): Integer;
begin
  Result := inherited Add( ASession );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.Extract(ASession: TDbSession): TDbSession;
begin
  Result := TDbSession( inherited Extract( ASession ) );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.Remove(ASession: TDbSession): Integer;
begin
  Result := inherited Remove( ASession );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.IndexOf(ASession: TDbSession): Integer;
begin
  Result := inherited IndexOf( ASession );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.First: TDbSession;
begin
  Result := TDbSession( inherited First );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.Last: TDbSession;
begin
  Result := TDbSession( inherited Last );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.GetFreeSessionLock(AIndex: Integer): Boolean;
begin
  FLock.Enter;
  try
    Result := False;
    if ( Self.DbSession[AIndex].IsFree ) then
    begin
      Self.DbSession[aIndex].IsBusy := True;
      Result := True;
      UpdateSessionCount;
    end;
  finally
    FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionList.GetErrorSessionLock(AIndex: Integer): Boolean;
begin
  FLock.Enter;
  try
    Result := False;
    if ( not Self.DbSession[AIndex].IsBusy ) and
       ( ( Self.DbSession[AIndex].IsError ) or
         ( not Self.DbSession[AIndex].Connection.Connected ) ) then
    begin
      Result := True;
    end;
  finally
    FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionList.ReleaseSessionLock(AInfo: TDbSession);
var
  aIndex: Integer;
begin
  FLock.Enter;
  try
    aIndex := Self.IndexOf( AInfo );
    if ( aIndex in [0..Self.Count-1] ) then
    begin
      Self.DbSession[aIndex].IsBusy := False;
      Self.DbSession[aIndex].IsError := not Self.DbSession[aIndex].Connection.Connected;
      UpdateSessionCount;
    end;
  finally
    FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionList.ReleaseSessionLock(AIndex: Integer);
begin
  FLock.Enter;
  try
    if ( aIndex in [0..Self.Count-1] ) then
    begin
      Self.DbSession[AIndex].IsBusy := False;
      Self.DbSession[aIndex].IsError := not Self.DbSession[aIndex].Connection.Connected;
      UpdateSessionCount;
    end;
  finally
    FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionList.UpdateSessionCount;
var
  aIndex: Integer;
begin
  FFreeSessionCount := 0;
  FBusySessionCount := 0;
  FErrSessionCount := 0;
  for aIndex := 0 to Self.Count - 1 do
  begin
    if ( Self.DbSession[aIndex].IsFree ) then
      Inc( FFreeSessionCount )
    else if ( Self.DbSession[aIndex].IsBusy ) then
      Inc( FBusySessionCount )
    else
      Inc( FErrSessionCount );
  end;
  FAppSo.FreeSession := FFreeSessionCount;
  FAppSo.BusySession := FBusySessionCount;
  FAppSo.ErrSession := FErrSessionCount;
end;

{ ---------------------------------------------------------------------------- }

{ TSessionManager }

constructor TDbSessionManager.Create(AAppSo: TAppSo);
begin
  FAppSo := AAppSo;
  FCommEnv := TCommEnv.Create;
  FSessionList := TDbSessionList.Create( FAppSo );
  FSemaphore := CreateSemaphore( nil, FAppSo.MaxSession, FAppSo.MaxSession, nil );
  Randomize;
end;

{ ---------------------------------------------------------------------------- }

destructor TDbSessionManager.Destroy;
begin
  FSessionList.Clear;
  FSessionList.Free;
  FCommEnv.Free;
  FAppSo := nil;
  CloseHandle( FSemaphore );
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionManager.SetCommEnv(const ACommEnv: TCommEnv);
begin
  FCommEnv.Assign( ACommEnv );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionManager.LockFreeSession: TDbSession;
var
  aIndex, aRetrys: Integer;
begin
  Result := nil;
  aRetrys := 0;
  if ( WaitForSingleObject( FSemaphore, 3000 )  = WAIT_OBJECT_0 ) then
  begin
    try
      repeat
        aIndex := Random( FSessionList.Count );
        if FSessionList.GetFreeSessionLock( aIndex ) then
          Result := FSessionList.DbSession[aIndex];
        Inc( aRetrys )
      until ( Assigned( Result ) or ( aRetrys >= FCommEnv.DbGetPoolObjectRertyCount ) );
    finally
      if not Assigned( Result ) then ReleaseSemaphore( FSemaphore, 1, nil );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionManager.UnLockFreeSession(ASessionInfo: TDbSession);
begin
  FSessionList.ReleaseSessionLock( ASessionInfo );
  ReleaseSemaphore( FSemaphore, 1, nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionManager.UnLockFreeSession(AIndex: Integer);
begin
  FSessionList.ReleaseSessionLock( AIndex );
  ReleaseSemaphore( FSemaphore, 1, nil );
end;

{ ---------------------------------------------------------------------------- }

{ TDbSessionManagerList }

constructor TDbSessionManagerList.Create;
begin
  inherited Create;
  FCommEnv := TCommEnv.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TDbSessionManagerList.Destroy;
begin
  FCommEnv.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionManagerList.SetCommEnv(const ACommEnv: TCommEnv);
var
  aIndex: Integer;
begin
  FCommEnv.Assign( ACommEnv );
  for aIndex := 0 to Count - 1 do
    Manager[aIndex].CommEnv := FCommEnv;
end;

{ ---------------------------------------------------------------------------- }

procedure TDbSessionManagerList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
  begin
    TDbSessionManager( Ptr ).Free;
    Ptr := nil;
  end;
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionManagerList.Add(AMgr: TDbSessionManager): Integer;
begin
  Result := inherited Add( AMgr );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionManagerList.Add(AAppSo: TAppSo): Integer;
begin
  Result := Add( TDbSessionManager.Create( AAppSo ) );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionManagerList.GetManager(Index: Integer): TDbSessionManager;
begin
  Result := nil;
  if Index in [0..( inherited Count - 1 )] then
    Result := TDbSessionManager( inherited Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionManagerList.IndexOf(ACompCode: String): Integer;
begin
  Result := 0;
  while ( Result < Count ) and ( TDbSessionManager( List^[Result] ).AppSo.CompCode <> ACompCode ) do
    Inc( Result );
  if Result = Count then
    Result := -1;
end;

{ ---------------------------------------------------------------------------- }

function TDbSessionManagerList.IndexOf(AMgr: TDbSessionManager): Integer;
begin
  Result := inherited IndexOf( AMgr );
end;

{ ---------------------------------------------------------------------------- }

{ TIdGenerateRule }

constructor TIdRange.Create(AMinValue, AMaxValue, ACurValue: Integer);
begin
  inherited Create;
  FMinValue := AMinValue;
  FMaxValue := AMaxValue;
  FCurValue := ACurValue;
end;

{ ---------------------------------------------------------------------------- }

destructor TIdRange.Destroy;
begin
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TIdRange.GetCurValue: Variant;
begin
  Result := FCurValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdRange.GetMaxValue: Variant;
begin
  Result := FMaxValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdRange.GetMinValue: Variant;
begin
  Result := FMinValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdRange.GetNextValue: Variant;
begin
  if ( FCurValue >= FMaxValue ) then
    ResetValue
  else begin
    Inc( FCurValue );
  end;
  Result := FCurValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TIdRange.ResetValue;
begin
  FCurValue := FMinValue;
end;

{ ---------------------------------------------------------------------------- }

{ TCustomIdGenerator }

constructor TCustomIdGenerator.Create(const AThreadSafe: Boolean;
  ARange: TIdRange);
begin
  inherited Create;
  FThreadSafe := AThreadSafe;
  if ( FThreadSafe ) then FLock := TCriticalSection.Create;
  FRange := ARange;
end;

{ ---------------------------------------------------------------------------- }

destructor TCustomIdGenerator.Destroy;
begin
  if ( FThreadSafe ) then FLock.Free;
  FRange.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TCustomIdGenerator.GetMaxValue: Variant;
begin
  Result := FRange.GetMaxValue;
end;

{ ---------------------------------------------------------------------------- }

function TCustomIdGenerator.GetMinValue: Variant;
begin
  Result := FRange.GetMinValue;
end;

{ ---------------------------------------------------------------------------- }

function TCustomIdGenerator.GetNextValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FRange.GetNextValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCustomIdGenerator.Reset;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    FRange.ResetValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCustomIdGenerator.GetCurValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FRange.GetCurValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TMsgIdGenerator }

function TMsgIdGenerator.GetMaxValue: Variant;
begin
  Result := FIdPrefix + IntToHex( FRange.GetMaxValue, 4 );
end;

{ ---------------------------------------------------------------------------- }

function TMsgIdGenerator.GetMinValue: Variant;
begin
  Result := FIdPrefix + IntToHex( FRange.GetCurValue, 4 );
end;

{ ---------------------------------------------------------------------------- }

function TMsgIdGenerator.GetCurValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FIdPrefix + IntToHex( FRange.GetCurValue, 4 );
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMsgIdGenerator.GetNextValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    if ( GetSysDate <> FIdPrefix ) then
    begin
      FIdPrefix := GetSysDate;
      FRange.ResetValue;
    end;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FIdPrefix + IntToHex( FRange.GetNextValue, 4 );
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgIdGenerator.SetIdPrefix(const AValue: String);
begin
  FIdPrefix := AValue;
end;

{ ---------------------------------------------------------------------------- }

end.
