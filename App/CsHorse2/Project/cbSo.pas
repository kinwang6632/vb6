unit cbSo;

interface

uses Windows, SysUtils, Variants, Classes;

type

  TDbState = ( dbClose, dbOpen, dbActive, dbError, dbWarning );
  
  { TSo }

  TSo = class(TPersistent)
  private
    FSelected: Boolean;
    FCompCode: String;
    FCompName: String;
    FDbAliase: String;
    FDbUserId: String;
    FDbUserPass: String;
    FDbErrorCount: Integer;
    FDbErrorTime: TDateTime;
    FDbState: TDbState;
    procedure SetDbState(const Value: TDbState);
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    property Selected: Boolean read FSelected write FSelected;
    property CompCode: String read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property DbUserId: String read FDbUserId write FDbUserId;
    property DbUserPass: String read FDbUserPass write FDbUserPass;
    property DbAliase: String read FDbAliase write FDbAliase;
    property DbState: TDbState read FDbState write SetDbState;
    property DbErrorCount: Integer read FDbErrorCount;
    property DbErrorTime: TDateTime read FDbErrorTime;
  end;

  { TSoList }

  TSoList = class(TList)
  private
    function GetSelectCount: Integer;
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TSo;
    procedure SetItem(Index: Integer; aObj: TSo);
  public
    procedure Insert(Index: Integer; aObj: TSo);
    procedure Assign(aSource: TSoList);
    function Add(aObj: TSo): Integer;
    function Extract(aObj: TSo): TSo;
    function Remove(aObj: TSo): Integer;
    function IndexOf(aObj: TSo): Integer; overload;
    function IndexOf(aCompCode: String): Integer; overload;
    function First: TSo;
    function Last: TSo;
    property Items[Index: Integer]: TSo read GetItem write SetItem; default;
    property SelectCount: Integer read GetSelectCount;
  end;



implementation

{ TSo }

constructor TSo.Create;
begin
  FSelected := False;
  FDbErrorCount := 0;
  FDbErrorTime := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TSo.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TSo ) then
  begin
    TSo( Dest ).Selected := FSelected;
    TSo( Dest ).CompCode := FCompCode;
    TSo( Dest ).CompName := FCompName;
    TSo( Dest ).DbUserId := FDbUserId;
    TSo( Dest ).DbUserPass := FDbUserPass;
    TSo( Dest ).DbAliase := FDbAliase;
    TSo( Dest ).DbState := FDbState;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

procedure TSo.SetDbState(const Value: TDbState);
begin
  FDbState := Value;
  case FDbState of
    dbClose:
      begin
        FDbErrorCount := 0;
      end;
    dbOpen, dbActive:
      begin
        FDbErrorCount := 0;
        FDbErrorTime := 0;
      end;
    dbError:
      begin
        Inc( FDbErrorCount );
        FDbErrorTime := Now;
      end;
  end;
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

procedure TSoList.Insert(Index: Integer; aObj: TSo);
begin
  inherited Insert( Index, aObj );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoList.Assign(aSource: TSoList);
var
  aIndex: Integer;
  aSo: TSo;
begin
  if not Assigned( aSource ) then Exit;
  Self.Clear;
  for aIndex := 0 to aSource.Count - 1 do
  begin
    aSo := TSo.Create;
    aSo.Assign( aSource[aIndex] );
    Self.Add( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Remove(aObj: TSo): Integer;
begin
  Result := inherited Remove( aObj );
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

function TSoList.IndexOf(aObj: TSo): Integer;
begin
  Result := inherited IndexOf( aObj );
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

function TSoList.GetSelectCount: Integer;
var
  aIndex: Integer;
begin
  Result:= 0;
  for aIndex := 0 to Self.Count - 1 do
  begin
    if ( Self.Items[aIndex].Selected ) then
      Inc( Result );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
