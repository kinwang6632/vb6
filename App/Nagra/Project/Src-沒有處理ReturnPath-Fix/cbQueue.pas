unit cbQueue;

interface

uses
  Windows, SysUtils, Classes, Contnrs, SyncObjs, CbUtils;


  { TCommandStringList }

type
  TCommandStringList = class(TStringList)
  private
    FLock: TCriticalSection;
    procedure LockList;
    procedure UnlockList;
  protected
    function GetCount: Integer; override;
    function GetTextStr: string; virtual;
    function GetObject(Index: Integer): TObject; override;
    function Get(Index: Integer): string; override;
    procedure Put(Index: Integer; const S: string); override;
    procedure PutObject(Index: Integer; AObject: TObject); virtual;
    procedure SetTextStr(const Value: string); virtual;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(const S: string): Integer; override;
    function AddObject(const S: string; AObject: TObject): Integer; override;
    procedure Clear; override;
    procedure Delete(Index: Integer); override;
    procedure Exchange(Index1, Index2: Integer); override;
    function Find(const S: string; var Index: Integer): Boolean; virtual;
    function IndexOf(const S: string): Integer; override;
    procedure Insert(Index: Integer; const S: string); override;
    procedure InsertObject(Index: Integer; const S: string; AObject: TObject); override;
    procedure CustomSort(Compare: TStringListSortCompare); virtual;
  end;


  { TTransactionNumberGenerator }

  TTransactionNumberGenerator = class(TObject)
  private
    FLock: TCriticalSection;
    FThreadSafe: Boolean;
    FInitValue: Integer;
    FMaxValue: Integer;
    FSequence: Integer;
    FMaxSequenceStringLen: Integer;
    FCompCode: String;
    function GetNextValue: String;
    procedure LockGenerator;
    procedure UnlockGenerator;
    function GetCurrentValue: String;
  protected
    function Generate: Integer; virtual;
  public
    constructor Create(const aThreadSafe: Boolean; const aInitValue,
      aMaxValue: Integer);
    destructor Destroy; override;
    property InitValue: Integer read FInitValue;
    property MaxValue: Integer read FMaxValue;
    property CurrentValue: String read GetCurrentValue;
    property NextValue: String read GetNextValue;
    property CompCode: String read FCompCode write FCompCode; 
  end;


implementation


{ TCommandStringList }

constructor TCommandStringList.Create;
begin
  inherited Create;
  FLock := TCriticalSection.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TCommandStringList.Destroy;
begin
  LockList; 
  try
    inherited Destroy;
  finally
    UnlockList;
    FLock.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.LockList;
begin
  FLock.Enter;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.UnlockList;
begin
  FLock.Leave;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.GetCount: Integer;
begin
  LockList;
  try
    Result := inherited GetCount;
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.GetObject(Index: Integer): TObject;
begin
  LockList;
  try
    Result := inherited GetObject( Index );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.Get(Index: Integer): string;
begin
  LockList;
  try
    Result := inherited Get( Index );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.Put(Index: Integer; const S: string);
begin
  LockList;
  try
    inherited Put( Index, S );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TCommandStringList.PutObject(Index: Integer; AObject: TObject);
begin
  LockList;
  try
    inherited PutObject( Index, AObject );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.GetTextStr: string;
begin
  LockList;
  try
    Result := inherited GetTextStr;
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.SetTextStr(const Value: string);
begin
  LockList;
  try
    inherited SetTextStr( Value );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.Delete(Index: Integer);
begin
  LockList;
  try
    inherited Delete( Index );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.Add(const S: string): Integer;
begin
  LockList;
  try
    Result := inherited Add( S );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.AddObject(const S: string; AObject: TObject): Integer;
begin
  LockList;
  try
    Result := inherited  AddObject( S, AObject );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.Clear;
begin
  LockList;
  try
    inherited Clear;
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.Exchange(Index1, Index2: Integer);
begin
  LockList;
  try
    inherited Exchange( Index1, Index2 );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.Find(const S: string; var Index: Integer): Boolean;
begin
  LockList;
  try
    Result := inherited Find( S, Index );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommandStringList.IndexOf(const S: string): Integer;
begin
  LockList;
  try
    Result := inherited IndexOf( S );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.Insert(Index: Integer; const S: string);
begin
  LockList;
  try
    inherited Insert( Index, S );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.InsertObject(Index: Integer; const S: string;
  AObject: TObject);
begin
  LockList;
  try
    inherited InsertObject( Index, S, AObject );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandStringList.CustomSort(Compare: TStringListSortCompare);
begin
  LockList;
  try
    inherited CustomSort( Compare );
  finally
    UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TTransactionNumberGenerator }

constructor TTransactionNumberGenerator.Create(const aThreadSafe: Boolean;
  const aInitValue, aMaxValue: Integer);
begin
  inherited Create;
  FThreadSafe := aThreadSafe;
  if FThreadSafe then FLock := TCriticalSection.Create;
  FInitValue := aInitValue;
  FMaxValue := aMaxValue;
  FSequence := FInitValue;
  FCompCode := EmptyStr;
  FMaxSequenceStringLen := Length( IntToStr( FMaxValue ) );
end;

{ ---------------------------------------------------------------------------- }

destructor TTransactionNumberGenerator.Destroy;
begin
  LockGenerator;
  try
    inherited Destroy;
  finally
    UnlockGenerator;
    if FThreadSafe then FLock.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.LockGenerator;
begin
  if FThreadSafe then FLock.Enter;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.UnlockGenerator;
begin
  if FThreadSafe then FLock.Leave;
end;

{ ---------------------------------------------------------------------------- }

function TTransactionNumberGenerator.Generate: Integer;
begin
  LockGenerator;
  try
    if FSequence >= FMaxValue then
      FSequence := FInitValue
    else
      Inc( FSequence );
    Result := FSequence;
  finally
    UnlockGenerator;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TTransactionNumberGenerator.GetCurrentValue: String;
begin
  Result := Lpad( FSequence, FMaxSequenceStringLen, '0' );
end;

{ ---------------------------------------------------------------------------- }

function TTransactionNumberGenerator.GetNextValue: String;
begin
  Generate;
  Result := GetCurrentValue;
end;

{ ---------------------------------------------------------------------------- }

end.
