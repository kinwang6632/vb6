unit cbClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, SyncObjs, ComObj, Contnrs;

const
  WM_PLAY = WM_USER + 1;
  WM_STOP = WM_USER + 2;

type

  TObServer = class;

  { TSubject }

  TSubject = class(TObject)
  private
    FList: TList;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddObServer(aObServer: TObServer);
    procedure RemoveObServer; overload;
    procedure RemoveObServer(aObServer: TObServer); overload;
    procedure NotifyObServer; overload;
    procedure NotifyObServer(aObServer: TObServer); overload;
  end;

  { TMessageSubject }

  TMessageSubject = class(TSubject)
  private
    FMessageType: Longint;
    FMessageText: String;
  protected
    procedure InternalNotifyObserver(const aMsg: String; const aType: Longint);
  public
    procedure OK(const aMsg: String);
    procedure Warning(const aMsg: String);
    procedure Error(const aMsg: String);
    procedure Normal(const aMsg: String);
    property MessageType: Longint read FMessageType;
    property MessageText: String read FMessageText;
  end;

  { TObjectSubject }

  TObjectSubject = class(TSubject)
  private
    FObject: TObject;
  public
    procedure Notify(aObj: TObject);
    property NotifyObject: TObject read FObject;
  end;


  TSubjectNotifyEvent = procedure (aSubject: TSubject) of object;

  { TObServer }

  TObServer = class(TObject)
  private
    FPause: Boolean;
    FOnUpdate: TSubjectNotifyEvent;
    FTrickedSubject: TSubject;
  private
    procedure CallOnUpdate;  
  protected
    procedure Update(aSubject: TSubject); virtual;
  public
    constructor Create;
    property Pause: Boolean read FPause write FPause;
    property OnUpdate: TSubjectNotifyEvent read FOnUpdate write FOnUpdate;
  end;

  { TMessageQueueThread }

  TMessageQueueThread = class(TThread)
  private
    FMsgHnd: THandle;
    FPlay: Boolean;
    FStop: Boolean;
  protected
    procedure WndProc(var Msg: TMessage); virtual;
    procedure WaitForPlaySignal; virtual;
    procedure WaitForStopSignal; virtual;
    property Play: Boolean read FPlay;
    property Stop: Boolean read FStop;
  public
    constructor Create(CreateSuspended: Boolean);
    destructor Destroy; override;
    property MsgHandle: THandle read FMsgHnd;
  end;


  { TGenerateRule }

  TGenerateRule = class(TObject)
  protected
    function GetMinValue: Variant; virtual; abstract;
    function GetMaxValue: Variant; virtual; abstract;
    function GetNextValue: Variant; virtual; abstract;
    function GetCurValue: Variant; virtual; abstract;
  end;


  { TNumberGenerateRule }

  TNumberGenerateRule = class(TGenerateRule)
  private
    FMinValue: Integer;
    FMaxValue: Integer;
    FCurValue: Integer;
  protected
    function GetMinValue: Variant; override;
    function GetMaxValue: Variant; override;
    function GetNextValue: Variant; override;
    function GetCurValue: Variant; override;
    procedure ResetValue; virtual;
  public
    constructor Create(aMinValue, aMaxValue, aCurValue: Integer);
  end;


  TGrPrefixType = (gpYear, gpYearMonth, gpYearMonthDay, gpFix);

  { TStringGenerateRule }

  TStringGenerateRule = class(TNumberGenerateRule)
  private
    FPrefixValue: String;
    FPrefixType: TGrPrefixType;
  protected
    function GetNextValue: Variant; override;
    function GetCurValue: Variant; override;
  public
    constructor Create(aMinValue, aMaxValue, aCurValue: Integer; aPrefix: TGrPrefixType;
     aPrefixValue: String = '');
  end;


  { TIdGenerator }

  TIdGenerator = class(TObject)
  private
    FLock: TCriticalSection;
    FThreadSafe: Boolean;
    FRule: TGenerateRule;
    function GetMinValue: Variant;
    function GetMaxValue: Variant;
    function GetNextValue: Variant;
    function GetCurValue: Variant;
  public
    constructor Create(const aThreadSafe: Boolean; aGenerateRule: TGenerateRule);
    destructor Destroy; override;
    property MinValue: Variant read GetMinValue;
    property MaxValue: Variant read GetMaxValue;
    property NextValue: Variant read GetNextValue;
    property CurrentValue: Variant read GetCurValue;
  end;


  { TSo }

  (*
  TSoInfo = record
    Selected: Boolean;                 { 是否選擇  }
    CompCode: String;                  { 系統台代碼 }
    CompName: String;                  { 系統台名稱 }
    DbLoginUser: String;               { 登入資料庫帳號 }
    DbLoginPass: String;               { 登入資料庫密碼 }
    DbAliase: String;                  { 登入資料庫 }
  end;
  *)

  TSo = class(TPersistent)
  private
    FSelected: Boolean;
    FCompCode: String;
    FCompName: String;
    FDbLoginUser: String;
    FDbLoginPass: String;
    FDbAliase: String;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    property Selected: Boolean read FSelected write FSelected;
    property CompCode: String read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property DbLoginUser: String read FDbLoginUser write FDbLoginUser;
    property DbLoginPass: String read FDbLoginPass write FDbLoginPass;
    property DbAliase: String read FDbAliase write FDbAliase;
  end;


  { TSoList }

  TSoList = class(TList)
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TSo;
    procedure SetItem(Index: Integer; aObj: TSo);
  public
    procedure Insert(Index: Integer; aObj: TSo);
    function Add(aObj: TSo): Integer;
    function Extract(aObj: TSo): TSo;
    function Remove(aObj: TSo): Integer;
    function IndexOf(aObj: TSo): Integer; overload;
    function IndexOf(aCompCode: String): Integer; overload;
    function First: TSo;
    function Last: TSo;
    property Items[Index: Integer]: TSo read GetItem write SetItem; default;
  end;


implementation

{ ---------------------------------------------------------------------------- }

{ TSubject }

constructor TSubject.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TSubject.Destroy;
begin
  RemoveObServer;
  FList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.AddObServer(aObServer: TObServer);
begin
  if FList.IndexOf( aObServer ) < 0 then
    FList.Add( aObServer );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer(aObServer: TObServer);
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aObServer );
  if aIndex >= 0 then
    FList.Delete( aIndex );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer;
begin
  while FList.Count > 0 do
  begin
    FList.Delete( 0 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.NotifyObServer(aObServer: TObServer);
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aObServer );
  if aIndex >= 0 then
  begin
    if not TObServer( FList[aIndex] ).Pause then
      TObServer( FList[aIndex] ).Update( Self );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.NotifyObServer;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if not TObServer( FList[aIndex] ).Pause then
      TObServer( FList[aIndex] ).Update( Self );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TMessageSubject }

procedure TMessageSubject.InternalNotifyObserver(const aMsg: String;
  const aType: Longint);
begin
  FMessageType := aType;
  FMessageText := aMsg;
  inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageSubject.Error(const aMsg: String);
begin
  InternalNotifyObserver( aMsg, MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageSubject.Normal(const aMsg: String);
begin
  InternalNotifyObserver( aMsg, MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageSubject.OK(const aMsg: String);
begin
  InternalNotifyObserver( aMsg, MB_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageSubject.Warning(const aMsg: String);
begin
  InternalNotifyObserver( aMsg, MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

{ TObjectSubject }

procedure TObjectSubject.Notify(aObj: TObject);
begin
  FObject := aObj;
  if Assigned( aObj ) then inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TObServer }

constructor TObServer.Create;
begin
  inherited Create;
  FPause := False;
  FTrickedSubject := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.CallOnUpdate;
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
    FOnUpdate( FTrickedSubject );
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.Update(aSubject: TSubject);
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
  begin
    FTrickedSubject := aSubject;
    if GetCurrentThreadId <> MainThreadID then
      TThread.Synchronize( nil, CallOnUpdate )
    else
      FOnUpdate( FTrickedSubject );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TMessageQueueThread }

constructor TMessageQueueThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FMsgHnd := Classes.AllocateHWnd( WndProc );
  FPlay := False;
  FStop := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TMessageQueueThread.Destroy;
begin
  Classes.DeallocateHWnd( FMsgHnd );
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageQueueThread.WndProc(var Msg: TMessage);
begin
  if ( Msg.Msg = WM_PLAY ) then
  begin
    FPlay := True;
    FStop := False;
  end else
  if ( Msg.Msg = WM_STOP ) then
  begin
    FStop := True;
    FPlay := False;
  end else
    Msg.Result := DefWindowProc( FMsgHnd, Msg.Msg, Msg.wParam, Msg.lParam );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageQueueThread.WaitForPlaySignal;
begin
  while not FPlay do
  begin
    Sleep( 300 );
    if Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageQueueThread.WaitForStopSignal;
begin
  while not FStop do
  begin
    Sleep( 300 );
    if Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TNumberGenerateRule }

constructor TNumberGenerateRule.Create(aMinValue, aMaxValue, aCurValue: Integer);
begin
  inherited Create;
  FMinValue := aMinValue;
  FMaxValue := aMaxValue;
  FCurValue := aCurValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetCurValue: Variant;
begin
  Result := FCurValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetMaxValue: Variant;
begin
  Result := FMaxValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetMinValue: Variant;
begin
  Result := FMinValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetNextValue: Variant;
begin
  if ( FCurValue >= FMaxValue ) then
    ResetValue
  else begin
    Inc( FCurValue );
  end;
  Result := FCurValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TNumberGenerateRule.ResetValue;
begin
  FCurValue := FMinValue;
end;

{ ---------------------------------------------------------------------------- }

{ TStringGenerateRule }

constructor TStringGenerateRule.Create(aMinValue, aMaxValue, aCurValue: Integer;
  aPrefix: TGrPrefixType; aPrefixValue: String);
begin
  inherited Create( aMinValue, aMaxValue, aCurValue );
  FPrefixType := aPrefix;
  FPrefixValue := aPrefixValue;
end;

{ ---------------------------------------------------------------------------- }

function TStringGenerateRule.GetCurValue: Variant;
var
  aFormatStr: String;
begin
  aFormatStr := '%.' + IntToStr( Length( IntToStr( FMaxValue ) ) ) + 'd';
  case FPrefixType of
    gpYear: Result := FormatDateTime( 'yyyy', Date ) +
      Format( aFormatStr, [FCurValue] );
    gpYearMonth: Result := FormatDateTime( 'yyyymm', Date ) +
      Format( aFormatStr, [FCurValue] );
    gpYearMonthDay: Result := FormatDateTime( 'yyyymmdd', Date ) +
      Format( aFormatStr, [FCurValue] );
  else
    Result := FPrefixValue + Format( aFormatStr, [FCurValue] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStringGenerateRule.GetNextValue: Variant;
begin
  inherited GetNextValue;
  Result := GetCurValue;
end;

{ ---------------------------------------------------------------------------- }

{ TIdGenerator }

constructor TIdGenerator.Create(const aThreadSafe: Boolean;
  aGenerateRule: TGenerateRule);
begin
  inherited Create;
  FThreadSafe := aThreadSafe;
  if ( FThreadSafe ) then FLock := TCriticalSection.Create;
  FRule := aGenerateRule;
end;

{ ---------------------------------------------------------------------------- }

destructor TIdGenerator.Destroy;
begin
  if ( FThreadSafe ) then FLock.Free;
  FRule.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetCurValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FRule.GetCurValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetMaxValue: Variant;
begin
  Result := FRule.GetMaxValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetMinValue: Variant;
begin
  Result := FRule.GetMinValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetNextValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FRule.GetNextValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSo }

constructor TSo.Create;
begin
  inherited Create;
  FSelected := False;
  FCompCode := EmptyStr;
  FCompName := EmptyStr;
  FDbLoginUser := EmptyStr;
  FDbLoginPass := EmptyStr;
  FDbAliase := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

destructor TSo.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSo.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TSo ) then
  begin
    TSo( Dest ).Selected := FSelected;
    TSo( Dest ).CompCode := FCompCode;
    TSo( Dest ).CompName := FCompName;
    TSo( Dest ).DbLoginUser := FDbLoginUser;
    TSo( Dest ).DbLoginPass := FDbLoginPass;
    TSo( Dest ).DbAliase := FDbAliase;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TSoList }

procedure TSoList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
    TSo( Ptr ).Free;
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Add(aObj: TSo): Integer;
begin
  Result := inherited Add( aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Extract(aObj: TSo): TSo;
begin
  Result := TSo( inherited Extract( aObj ) );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.GetItem(Index: Integer): TSo;
begin
  Result := TSo( inherited Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoList.SetItem(Index: Integer; aObj: TSo);
begin
  inherited Items[Index] := aObj;
end;

{ ---------------------------------------------------------------------------- }

function TSoList.IndexOf(aObj: TSo): Integer;
begin
  Result := inherited IndexOf( aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.IndexOf(aCompCode: String): Integer;
begin
  Result := 0;
  while ( Result < Count ) and ( TSo( List^[Result] ).CompCode <> aCompCode ) do
    Inc( Result );
  if Result = Count then
    Result := -1;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoList.Insert(Index: Integer; aObj: TSo);
begin
  inherited Insert( Index, aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Remove(aObj: TSo): Integer;
begin
  Result := inherited Remove( aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.First: TSo;
begin
  Result := TSo( inherited First );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Last: TSo;
begin
  Result := TSo( inherited Last );
end;

{ ---------------------------------------------------------------------------- }

end.
