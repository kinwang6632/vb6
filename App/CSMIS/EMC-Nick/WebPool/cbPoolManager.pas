unit cbPoolManager;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComServ, ComObj, VCLCom, StdVcl, bdemts, DataBkr, DBClient, IniFiles, DB,
  ADODB, ADOInt, MtsRdm, Mtx, WebPool_TLB, Encryption_TLB, CodeSiteLogging;


type
  TDecrypt = class(TObject)
  private
    { Private declarations }
    FKey: String;
    FFileName: String;
  public
    { Public declarations }
    constructor Create(const aKey, aFileName: String);
    function GetDecryptString: String;
  end;


type
  TPoolManager = class(TMtsDataModule, IPoolManager)
    ADOConnection: TADOConnection;
    ADODataSet: TADODataSet;
    ADOCommand: TADOCommand;
    procedure MtsDataModuleCreate(Sender: TObject);
    procedure MtsDataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  protected
    class procedure UpdateRegistry(Register: Boolean; const ClassID, ProgID: string); override;
    function GetRecordSet(aSQL: OleVariant): OleVariant; safecall;
    function GetConnection: OleVariant; safecall;
    function GetCommand: OleVariant; safecall;
  public
    { Public declarations }
  end;

var
  PoolManager: TPoolManager;

implementation

uses Variants;

{$R *.DFM}

{ ---------------------------------------------------------------------------- }

{ TDecrypt }

constructor TDecrypt.Create(const aKey, aFileName: String);
begin
  inherited Create;
  FKey := aKey;
  FFileName := aFileName;
end;

{ ---------------------------------------------------------------------------- }

function TDecrypt.GetDecryptString: String;
var
  aPassObj: _Password;
  aIniFile: TStringList;
  aIndex: Integer;
  aKey, aContent: WideString;
begin
  Result := EmptyStr;
  if not FileExists( FFileName ) then Exit;
  aPassObj := CoPassword.Create;
  aIniFile := TStringList.Create;
  try
    aIniFile.LoadFromFile( FFileName );
    for aIndex := 0 to aIniFile.Count - 1 do
      aContent := ( aContent + aIniFile.Strings[aIndex] );
    aKey := FKey;
    Result := aPassObj.Decrypt( aContent, aKey );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

class procedure TPoolManager.UpdateRegistry(Register: Boolean; const ClassID, ProgID: string);
begin
  if Register then
  begin
    inherited UpdateRegistry(Register, ClassID, ProgID);
    EnableSocketTransport(ClassID);
    EnableWebTransport(ClassID);
  end else
  begin
    DisableSocketTransport(ClassID);
    DisableWebTransport(ClassID);
    inherited UpdateRegistry(Register, ClassID, ProgID);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TPoolManager.MtsDataModuleCreate(Sender: TObject);
var
  aDecryObj: TDecrypt;
  aPath: String;
//  aModulePath: String;
begin
  (*
  try
    aModulePath := ExtractFilePath( SysUtils.GetModuleName( HInstance ) );
  except
    {}
  end;
  aDecryObj := TDecrypt.Create( 'CS', IncludeTrailingPathDelimiter(
     aModulePath ) + 'DBInfo.ini' );
  *)   
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    'DBInfo.ini';
  CodeSite.Send( 'DBInfo Path=' + aPath );
  aDecryObj := TDecrypt.Create( 'CS', aPath );
  try
    ADOConnection.ConnectionString := aDecryObj.GetDecryptString;
    CodeSite.Send( ADOConnection.ConnectionString );
  finally
    aDecryObj.Free;
  end;
  if ( ADOConnection.ConnectionString = EmptyStr ) then
    raise Exception.Create( '須要資料庫連線設定字串。' );
  ADOConnection.Connected := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TPoolManager.MtsDataModuleDestroy(Sender: TObject);
begin
  ADOConnection.Connected := False;
end;

{ ---------------------------------------------------------------------------- }

function TPoolManager.GetConnection: OleVariant;
begin
  Result := ADOConnection.ConnectionObject;
  SetComplete;
end;

{ ---------------------------------------------------------------------------- }

function TPoolManager.GetRecordSet(aSQL: OleVariant): OleVariant;
var
  aString: String;
begin
  aString := VarToStrDef( aSQL, EmptyStr );
  if ( aString = EmptyStr ) then
  begin
    SetAbort;
    raise Exception.Create( 'SQL 是空值。' );
  end; 
  ADODataSet.Close;
  ADODataSet.CommandText := aString;
  try
    ADODataSet.Open;
  except
    SetAbort;
    raise;
  end;
  Result := ADODataSet.Recordset;
  ADODataSet.Close;
  SetComplete;
end;

{ ---------------------------------------------------------------------------- }

function TPoolManager.GetCommand: OleVariant;
begin
  try
    ADOCommand.CommandObject.Set_ActiveConnection( ADOCommand.Connection.ConnectionObject );
    Result := ADOCommand.CommandObject;
  finally
    CallComplete( True );
  end;
end;

{ ---------------------------------------------------------------------------- }


initialization
  TComponentFactory.Create(ComServer, TPoolManager, Class_PoolManager, ciMultiInstance, tmBoth);
end.