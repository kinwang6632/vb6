unit cbDataControler;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DB, ADODB, IniFiles;

type

  PSoInfoEx = ^TSoInfoEx;

  TSoInfoEx = record
    Sid: String;
    UserName: String;
    Password: String;
    Owner: String;
    CompCode: String;
    CompName: String;
    DataConnection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    DataCache: TADOQuery;
  end;

  PSoInfo = ^TSoInfo;

  TSoInfo = record
    Sid: String;
    UserName: String;
    Password: String;
    Owner: String;
    CompCode: String;
    CompName: String;
  end;

  TDataControler = class(TDataModule)
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FList: TList;
    FIniFile: TIniFile;
    procedure CreateSoInfo(aFileName: String);
    procedure CreateSoInfoEx(aFileName: String);
    procedure ReleaseSoInfo;
    procedure ReleaseSoInfoEx;
  public
    { Public declarations }
  end;

var
  DataControler: TDataControler;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TDataControler.DataModuleCreate(Sender: TObject);
begin
  FList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.DataModuleDestroy(Sender: TObject);
begin
  ReleaseSoInfo;
  if Assigned( FIniFile ) then FIniFile.Free;
  FList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.ReleaseSoInfo;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if Assigned( FList[aIndex] ) then
    begin
      Dispose( PSoInfo( FList[aIndex] ) );
      FList[aIndex] := nil;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.CreateSoInfo(aFileName: String);
var
  aCount, aIndex: Integer;
  aSoInfo: PSoInfo;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + aFileName;
  if not FileExists( aFileName ) then
    raise Exception.CreateFmt( '檔案%s不存在。', [aFileName] );
  FIniFile.Create( aFileName );
  aCount := FIniFile.ReadInteger( 'DBINFO', 'DB_COUNT', 0 );
  for aIndex := 0 to aCount - 1 do
  begin
    New( aSoInfo );
    aSoInfo.Sid := FIniFile.ReadString( 'DBINFO', Format( 'ALIAS_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.UserName := FIniFile.ReadString( 'DBINFO', Format( 'USERID_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.Password := FIniFile.ReadString( 'DBINFO', Format( 'PASSWORD_%d', [aIndex+1] ), EmptyStr );
    if FIniFile.ValueExists( 'DBINFO', Format( 'OWNER_%d', [aIndex+1] ) ) then
      aSoInfo.Owner := FIniFile.ReadString( 'DBINFO', Format( 'OWNER_%d', [aIndex+1] ), EmptyStr )
    else
      aSoInfo.Owner := aSoInfo.UserName;
    aSoInfo.CompCode := FIniFile.ReadString( 'DBINFO', Format( 'COMPCODE_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.CompName := EmptyStr;
    if FIniFile.ValueExists( 'DBINFO', Format( 'COMPNAME_%d', [aIndex+1] ) ) then
      aSoInfo.CompName := FIniFile.ReadString( 'DBINFO', Format( 'COMPNAME_%d', [aIndex+1] ), EmptyStr );
    FList.Add( aSoInfo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.CreateSoInfoEx(aFileName: String);
var
  aCount, aIndex: Integer;
  aSoInfo: PSoInfoEx;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + aFileName;
  if not FileExists( aFileName ) then
    raise Exception.CreateFmt( '檔案%s不存在。', [aFileName] );
  FIniFile.Create( aFileName );
  aCount := FIniFile.ReadInteger( 'DBINFO', 'DB_COUNT', 0 );
  for aIndex := 0 to aCount - 1 do
  begin
    New( aSoInfo );
    aSoInfo.Sid := FIniFile.ReadString( 'DBINFO', Format( 'ALIAS_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.UserName := FIniFile.ReadString( 'DBINFO', Format( 'USERID_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.Password := FIniFile.ReadString( 'DBINFO', Format( 'PASSWORD_%d', [aIndex+1] ), EmptyStr );
    if FIniFile.ValueExists( 'DBINFO', Format( 'OWNER_%d', [aIndex+1] ) ) then
      aSoInfo.Owner := FIniFile.ReadString( 'DBINFO', Format( 'OWNER_%d', [aIndex+1] ), EmptyStr )
    else
      aSoInfo.Owner := aSoInfo.UserName;
    aSoInfo.CompCode := FIniFile.ReadString( 'DBINFO', Format( 'COMPCODE_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.CompName := EmptyStr;
    if FIniFile.ValueExists( 'DBINFO', Format( 'COMPNAME_%d', [aIndex+1] ) ) then
      aSoInfo.CompName := FIniFile.ReadString( 'DBINFO', Format( 'COMPNAME_%d', [aIndex+1] ), EmptyStr );
    aSoInfo.DataConnection := TADOConnection.Create( nil );
    aSoInfo.DataReader := TADOQuery.Create( nil );
    aSoInfo.DataWriter := TADOQuery.Create( nil );
    aSoInfo.DataCache := TADOQuery.Create( nil );
    aSoInfo.DataConnection.LoginPrompt := False;
    aSoInfo.DataReader.Connection := aSoInfo.DataConnection;
    aSoInfo.DataReader.CacheSize := 1000;
    aSoInfo.DataWriter.Connection := aSoInfo.DataConnection;
    aSoInfo.DataWriter.CacheSize := 1000;
    aSoInfo.DataCache.Connection := aSoInfo.DataConnection;
    aSoInfo.DataCache.CacheSize := 1000;
    aSoInfo.
    FList.Add( aSoInfo );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
