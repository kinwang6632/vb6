unit cbLanguage;

interface

uses Windows, SysUtils, Variants, Classes;

type

  { TLanguageManager }

  TLanguageManager = class(TObject)
  private
    FResourceNames: TStringList;
    FResourceValues: TStringList;
  public
    constructor Create;
    destructor Destroy; override;
    function Get(const AName: String): String;
    function GetFmt(const AName: String; const Args: array of const): String;
    procedure Add(const AName, AValue: String);
    procedure Clear;
    function LoadFromFile(const AFileName: String): Boolean; overload;
    function LoadFromFile: Boolean; overload;
  end;

  var LanguageManager: TLanguageManager;


implementation

uses cbUtilis;

{ TLanguageManager }

constructor TLanguageManager.Create;
begin
  inherited;
  FResourceNames := TStringList.Create;
  FResourceValues := TStringList.Create;
  FResourceNames.Sorted := True;
end;

{ ---------------------------------------------------------------------------- }

destructor TLanguageManager.Destroy;
begin
  FResourceNames.Free;
  FResourceValues.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TLanguageManager.Add(const AName, AValue: String);
var
  aIndex: Integer;
begin
  aIndex := FResourceNames.IndexOf( AName );
  if ( aIndex < 0 ) then
  begin
    FResourceNames.AddObject( AName, Pointer( FResourceValues.Count ) );
    FResourceValues.Add( aValue );
  end else
    FResourceValues[Integer( FResourceNames.Objects[aIndex] )] := aValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TLanguageManager.Clear;
begin
  FResourceNames.Clear;
  FResourceValues.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TLanguageManager.Get(const AName: String): String;
var
  aIndex: Integer;
begin
  aIndex := FResourceNames.IndexOf( AName );
  if ( aIndex >= 0 ) then
    Result := FResourceValues[Integer( FResourceNames.Objects[aIndex] )]
  else
    Result := AName;
end;

{ ---------------------------------------------------------------------------- }

function TLanguageManager.GetFmt(const AName: String;
  const Args: array of const): String;
begin
  Result := Format( Get( AName ), Args );
end;

{ ---------------------------------------------------------------------------- }

function TLanguageManager.LoadFromFile(const AFileName: String): Boolean;
var
  aList: TStringList;
  aIndex, aPos: Integer;
  aName, aValue: String;
begin
  Self.Clear;
  try
    aList := TStringList.Create;
    try
      aList.LoadFromFile( AFileName );
      for aIndex := 0 to aList.Count - 1 do
      begin
        aPos := AnsiPos( '=', aList[aIndex] );
        if ( aPos > 0 ) then
        begin
          aValue := Copy( aList[aIndex], aPos + 1, MaxInt );
          aName := Copy( aList[aIndex], 1, aPos - 1 );
          if ( aValue <> EmptyStr ) and ( aName <> EmptyStr ) then
            Self.Add( aName, aValue );
        end;
      end;
    finally
      aList.Free;
    end;
  except
    on E: Exception do ErrorMsg( E.Message );
  end;
  Result := ( FResourceValues.Count > 0 ) and ( FResourceNames.Count > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TLanguageManager.LoadFromFile: Boolean;
var
  aFileName, aAppName: String;
begin
  aAppName := ExtractFileName( ParamStr( 0 ) );
  aAppName := Copy( aAppName, 1, LastDelimiter( '.', aAppName ) - 1 );
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + ( aAppName + '.Lang' );
  Result := Self.LoadFromFile( aFileName );
end;

{ ---------------------------------------------------------------------------- }

initialization
  LanguageManager := TLanguageManager.Create;

finalization
  FreeAndNil( LanguageManager );
end.
