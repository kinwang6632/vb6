unit cbDesignPattern;

interface

uses Windows, SysUtils, Variants, Classes, cbSyncObjs;

type

  { TNotify }

  TMsgText = class(TPersistent)
  private
    FText: String;
    FFlag: Integer;
    FData: Pointer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    property Text: String read FText write FText;
    property Flag: Integer read FFlag write FFlag;
    property Data: Pointer read FData write FData;
  end;


  TObServer = class;

  { TSubject }

  TSubject = class(TObject)
  private
    FList: TList;
    function GetObserverCount: Integer;
    function GetObserver(AIndex: Integer): TObServer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddObServer(AObSrv: TObServer);
    procedure RemoveObServer; overload;
    procedure RemoveObServer(AObSrv: TObServer); overload;
    procedure NotifyObServer; overload;
    procedure NotifyObServer(AObSrv: TObServer); overload;
    property Count: Integer read GetObserverCount;
    property Observer[Index: Integer]: TObServer read GetObserver;
  end;

  { TTextSubject }

  TMsgSubject = class(TSubject)
  private
    FNotify: TMsgText;
    FPreData: Pointer;
    procedure InternalNotify(const aMsg: String; const AFlag: Integer);
  public
    constructor Create;
    destructor Destroy; override;
    procedure OK(const AMsg: String);
    procedure Warning(const AMsg: String);
    procedure Error(const AMsg: String);
    procedure Info(const AMsg: String);
    procedure OKFmt(const AMsg: String; const Args: array of const);
    procedure WarningFmt(const AMsg: string; const Args: array of const);
    procedure ErrorFmt(const AMsg: string; const Args: array of const);
    procedure InfoFmt(const AMsg: string; const Args: array of const);
    property MsgText: TMsgText read FNotify;
  end;

  { TDataSubject  }
  
  TDataSubject = class(TSubject)
  private
    FData: Pointer;
  public
    procedure Notify(AData: Pointer);
    property Data: Pointer read FData;
  end;


  TSubjectEvent = procedure (ASub: TSubject) of object;

  { TObServer }

  TObServer = class(TObject)
  private
    FPause: Boolean;
    FOnUpdate: TSubjectEvent;
    FSource: TSubject;
    FLock: TCriticalSection;
  private
    procedure CallOnUpdate;
  protected
    procedure Update(ASub: TSubject); virtual;
  public
    constructor Create;
    destructor Destroy; override;
    property Pause: Boolean read FPause write FPause;
    property OnUpdate: TSubjectEvent read FOnUpdate write FOnUpdate;
  end;

  { TSyncList }

  TSyncList = class(TObject)
  private
    FLock: TMREWSync;
    FList: TStringList;
  protected
    function Get(AIndex: Integer): string;
    procedure Put(AIndex: Integer; const AId: string);
    function GetObject(AIndex: Integer): TObject;
    procedure PutObject(AIndex: Integer; AObj: TObject);
    function GetSorted: Boolean;
    procedure SetSorted(const Value: Boolean);
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(const AId: string): Integer;
    function AddObject(const AId: string; AObj: TObject): Integer;
    procedure Clear;
    procedure Delete(AIndex: Integer); overload;
    procedure FreeObject(AIndex: Integer); overload;
    procedure FreeObject(AObj: TObject); overload;
    function IndexOf(const AId: string): Integer; overload;
    function IndexOf(const AObj: TObject): Integer; overload;
    procedure BeginWriteLock;
    procedure EndWriteLock;
    property Id[Index: Integer]: string read Get write Put; default;
    property Objects[Index: Integer]: TObject read GetObject write PutObject;
    property Sorted: Boolean read GetSorted write SetSorted;
    property Count: Integer read GetCount;
  end;



implementation


{ ---------------------------------------------------------------------------- }

{ TNotify }

constructor TMsgText.Create;
begin
  FFlag := 0;
  FText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgText.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TMsgText ) then
  begin
    TMsgText( Dest ).Text := FText;
    TMsgText( Dest ).Flag := FFlag;
    TMsgText( Dest ).Data := FData;
  end else
    inherited AssignTo( Dest );
end;

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

procedure TSubject.AddObServer(AObSrv: TObServer);
begin
  if FList.IndexOf( AObSrv ) < 0 then
    FList.Add( AObSrv );
end;

{ ---------------------------------------------------------------------------- }

function TSubject.GetObserverCount: Integer;
begin
  Result := FList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TSubject.GetObserver(AIndex: Integer): TObServer;
begin
  Result := TObServer( FList[aIndex] );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer(AObSrv: TObServer);
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( AObSrv );
  if aIndex >= 0 then
    FList.Delete( aIndex );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer;
begin
  while FList.Count > 0 do
    FList.Delete( 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.NotifyObServer(AObSrv: TObServer);
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( AObSrv );
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

{ TTextSubject }

constructor TMsgSubject.Create;
begin
  inherited Create;
  FNotify := TMsgText.Create;
  FPreData := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TMsgSubject.Destroy;
begin
  FNotify.Free;
  FPreData := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.InternalNotify(const AMsg: String; const AFlag: Integer);
begin
  FNotify.Text := AMsg;
  FNotify.Flag := AFlag;
  inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.Error(const AMsg: String);
begin
  InternalNotify( AMsg, MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.Info(const AMsg: String);
begin
  InternalNotify( AMsg, MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.OK(const AMsg: String);
begin
  InternalNotify( AMsg, MB_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.Warning(const AMsg: String);
begin
  InternalNotify( AMsg, MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.ErrorFmt(const AMsg: string; const Args: array of const);
begin
  InternalNotify( Format( AMsg, Args ), MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.InfoFmt(const AMsg: string; const Args: array of const);
begin
  InternalNotify( Format( AMsg, Args ), MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.OKFmt(const AMsg: String; const Args: array of const) ;
begin
  InternalNotify( Format( AMsg, Args ), MB_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure TMsgSubject.WarningFmt(const AMsg: string; const Args: array of const);
begin
  InternalNotify( Format( AMsg, Args ), MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

{ TDataSubject }

procedure TDataSubject.Notify(AData: Pointer);
begin
  FData := AData;
  if Assigned( FData ) then inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TObServer }

constructor TObServer.Create;
begin
  inherited Create;
  FPause := False;
  FSource := nil;
  FLock := TCriticalSection.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TObServer.Destroy;
begin
  FLock.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.CallOnUpdate;
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
    FOnUpdate( FSource );
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.Update(ASub: TSubject);
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
  begin
    if ( GetCurrentThreadId = MainThreadID ) then
      FOnUpdate( ASub ) //主Form自行指定的Update事件
    else begin
      FLock.Enter;
      try
        FSource := ASub;
        TThread.Synchronize( nil, CallOnUpdate )
      finally
        FLock.Leave;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSyncList }

constructor TSyncList.Create;
begin
  inherited Create;
  FLock := TMREWSync.Create;
  FList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TSyncList.Destroy;
begin
  FLock.BeginWrite;
  try
    FList.Free;
    inherited Destroy;
  finally
    FLock.EndWrite;
    FLock.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.Add(const AId: string): Integer;
begin
  FLock.BeginWrite;
  try
    Result := FList.Add( AId );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.AddObject(const AId: string; AObj: TObject): Integer;
begin
  FLock.BeginWrite;
  try
    Result := FList.AddObject( AId, AObj );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Clear;
begin
  FLock.BeginWrite;
  try
    FList.Clear;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Delete(AIndex: Integer);
begin
  FLock.BeginWrite;
  try
    FList.Delete( AIndex );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.IndexOf(const AId: string): Integer;
begin
  FLock.BeginRead;
  try
    Result := FList.IndexOf( AId );
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.IndexOf(const AObj: TObject): Integer;
begin
  FLock.BeginRead;
  try
    Result := FList.IndexOfObject( AObj );
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.GetSorted: Boolean;
begin
  Result := FList.Sorted;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.SetSorted(const Value: Boolean);
begin
  FLock.BeginWrite;
  try
    FList.Sorted := Value;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.Get(AIndex: Integer): string;
begin
  FLock.BeginRead;
  try
    Result := FList[AIndex];
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.GetObject(AIndex: Integer): TObject;
begin
  FLock.BeginRead;
  try
    Result := FList.Objects[AIndex];
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Put(AIndex: Integer; const AId: string);
begin
  FLock.BeginWrite;
  try
    FList[AIndex] := AId;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.PutObject(AIndex: Integer; AObj: TObject);
begin
  FLock.BeginWrite;
  try
    FList.Objects[AIndex] := AObj;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.FreeObject(AIndex: Integer);
begin
  FLock.BeginWrite;
  try
    TObject( FList.Objects[AIndex] ).Free;
    FList.Delete( AIndex );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.FreeObject(AObj: TObject);
begin
  FLock.BeginWrite;
  try
    FList.Delete( FList.IndexOfObject( AObj ) );
    FreeAndNil( AObj );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.GetCount: Integer;
begin
  FLock.BeginRead;
  try
    Result := FList.Count;
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.BeginWriteLock;
begin
  FLock.BeginWrite;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.EndWriteLock;
begin
  FLock.EndWrite;
end;

{ ---------------------------------------------------------------------------- }

end.
