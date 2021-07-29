unit cbDataControler;

interface

uses
  Windows, SysUtils, Classes, Variants, ADODB, DB, IniFiles, DBClient, ComObj,
  ActiveX, cbClass;

type

  { 系統台資訊 }

  PSoInfo = ^TSoInfo;

  TSoInfo = record
    Selected: Boolean;                 { 是否選擇  }
    CompCode: String;                  { 系統台代碼 }
    CompName: String;                  { 系統台名稱 }
    LoginUser: String;                 { 登入資料庫帳號 }
    LoginPass: String;                 { 登入資料庫密碼 }
    DbAliase: String;                  { 登入資料庫 }
  end;

  TDataControler = class(TDataModule)
    DataConnection: TADOConnection;
    DataReader: TADOQuery;
    ExportDataSet: TClientDataSet;
    DataWriter: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FList: TList;
    FExportList: TStringList;
    FNotify: TNotify;
    FSelectedComps: String;    FImportList: TStringList;
    procedure CreateSoInfo;
    procedure FreeSoInfo;
    procedure PrepareExportDataSet;
    procedure UnPrepareDataSet;
    procedure UpdateDataBack(const aKind: TExpoterKind);
    function GetSqlText(const aKind: TExpoterKind; aSoInfo: PSoInfo): String;
    function GetExportFileName(const aKind: TExpoterKind): String;
    function GetSoInfo(const aCompCode: String): PSoInfo;
    function PrepareDataConnection(aSoInfo: PSoInfo): Boolean;
    function PrepareDataSource(aSoInfo: PSoInfo): Boolean;
    function DataBind(const aKind: TExpoterKind; aSoInfo: PSoInfo): Integer;
    function ExportToFile(const aKind: TExpoterKind): Integer;
    function InernalExportCA: Boolean;
    function InernalExportSA: Boolean;
    function InernalExportRA: Boolean;
    procedure SynchronizeNotify;
    procedure Notify;
  public
    { Public declarations }
    constructor Create(aSelectedComps: String); reintroduce;
    destructor Destroy; override;
    procedure DoDataExport(const aKind: TExpoterKind);
    function DoDataImport(aFileName: String): Boolean;
    class procedure CheckFolder;
    class procedure CreateTemporyIniFile(aSource, aDest: String);
  end;

var
  DataControler: TDataControler;

implementation

uses cbMain, cbDataThread, Encryption_TLB;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function GetCompName(const aCompCode: Integer): String;
begin
    case aCompCode of
      { 總倉 }
      0: begin
           Result := '總倉';
         end;
      { 觀昇 }
      1: begin
           Result := '觀昇';
         end;
      { 屏南 }
      2: begin
           Result := '屏南';
         end;
      { 南天 }
      3: begin
           Result := '南天';
         end;
      { 新頻道 }
      5: begin
           Result := '新頻道';
         end;
      { 豐盟 }
      6: begin
           Result := '豐盟';
         end;
      { 振道 }
      7: begin
           Result := '振道';
         end;
      { 全聯 }
      8: begin
           Result := '全聯';
         end;
      { 陽明山 }
      9: begin
           Result := '陽明山';
         end;
     { 新台北 }
     10: begin
           Result := '新台北';
         end;
     { 金頻道 }
     11: begin
           Result := '金頻道';
         end;
     { 大安文山 }
     12: begin
           Result := '大安文山';
         end;
     { 新唐城 }
     13: begin
           Result := '新唐城';
         end;
     { 大新店 }
     14: begin
           Result := '大新店';
         end;
     { 北桃園 }
     16: begin
           Result := '北桃園';
         end;
     { 新唐城, CM 業務獨立一公司別 17 作業, 舊的 13 不使用 }
     17: begin
           Result := '新唐城';
         end;
     { 倉管測試資料庫 }
     20: begin
           Result := '倉管測試資料庫';
         end;
    else
      Result := EmptyStr;
    end;
end;

{ ---------------------------------------------------------------------------- }

function ExtractValue(var aValue: String; aSeparator: String = ','): String;
var
  aPos: Integer;
begin
  aPos := AnsiPos( aSeparator, aValue );
  if aPos = 0 then
  begin
    Result := aValue;
    aValue := EmptyStr;
  end else
  begin
    Result := Copy( aValue, 1, aPos - 1 );
    Delete( aValue, 1, aPos );
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsSelectedComp(const aCode: String; aComps: String): Boolean;
var
  aValue, aTemp: String;
begin
  aTemp := aComps;
  repeat
     aValue := ExtractValue( aTemp );
     Result := ( aCode = aValue );
     if Result then Break;      
  until ( aTemp = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

class procedure TDataControler.CheckFolder;
begin
  { Temp }
  if not DirectoryExists( LocalTempDir ) then
    ForceDirectories( LocalTempDir );
  { Uplaod }
  if not DirectoryExists( LocalUploadDir ) then
    ForceDirectories( LocalUploadDir );
  { Backup }
  if not DirectoryExists( LocalBackupDir ) then
    ForceDirectories( LocalBackupDir );
  { Download }
  if not DirectoryExists( LocalDownloadDir ) then
    ForceDirectories( LocalDownloadDir );
  { Error }
  if not DirectoryExists( LocalErrorDir ) then
    ForceDirectories( LocalErrorDir );
end;

{ ---------------------------------------------------------------------------- }

constructor TDataControler.Create(aSelectedComps: String);
begin
  CoInitialize( nil );
  inherited Create( nil );
  FSelectedComps := aSelectedComps;
end;

{ ---------------------------------------------------------------------------- }

destructor TDataControler.Destroy;
begin
  inherited;
  CoUninitialize;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.DataModuleCreate(Sender: TObject);
begin
  FList := TList.Create;
  FExportList := TStringList.Create;
  FImportList := TStringList.Create;
  CreateSoInfo;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.DataModuleDestroy(Sender: TObject);
var
  aFileName: string;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
     ParamStr( 0 ) ) )+ 'TMPCSIS.INI' ;
  if FileExists( aFileName ) then DeleteFile( aFileName );
  UnPrepareDataSet;
  FreeSoInfo;
  FExportList.Free;
  FImportList.Free;
  FList.Free;
end;

{ ---------------------------------------------------------------------------- }

{ 新裝機完工 }

function TDataControler.InernalExportCA: Boolean;
var
  aSoInfo: PSoInfo;
  aIndex, aCount: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if Assigned( Main.DataThread ) then
      if TDataThread( Main.DataThread ).Terminated then Exit;
    aSoInfo := PSoInfo( FList[aIndex] );
    if not aSoInfo.Selected then Continue;      
    if not PrepareDataConnection( aSoInfo ) then Continue;
    DataReader.Close;
    DataReader.SQL.Text := GetSqlText( ekCA, aSoInfo );
    if not PrepareDataSource( aSoInfo ) then Continue;
    {}
    aCount := DataBind( ekCA, aSoInfo );
    if ( aCount > 0 ) then
    begin
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( '系統台:%s, 裝機完工資料共有%d筆。', [aSoInfo.CompName, aCount] );
      SynchronizeNotify;
    end;
    if Assigned( Main.DataThread ) then
      if TDataThread( Main.DataThread ).Terminated then Exit;
  end;
  if not ( ExportDataSet.IsEmpty ) then
  begin
    aCount := ExportToFile( ekCA );
    if ( aCount > 0 ) then UpdateDataBack( ekCA );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ 停權 }

function TDataControler.InernalExportSA: Boolean;
var
  aSoInfo: PSoInfo;
  aIndex, aCount: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if Assigned( Main.DataThread ) then
      if TDataThread( Main.DataThread ).Terminated then Exit;
    aSoInfo := PSoInfo( FList[aIndex] );
    if not aSoInfo.Selected then Continue;
    if not PrepareDataConnection( aSoInfo ) then Continue;
    DataReader.Close;
    DataReader.SQL.Text := GetSqlText( ekSA, aSoInfo );
    if not PrepareDataSource( aSoInfo ) then Continue;
    {}
    aCount := DataBind( ekSA, aSoInfo );
    if ( aCount > 0 ) then
    begin
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( '系統台:%s, 停權資料共有%d筆。', [aSoInfo.CompName, aCount] );
      Notify;
    end;
    if Assigned( Main.DataThread ) then
      if TDataThread( Main.DataThread ).Terminated then Exit;
  end;
  if not ( ExportDataSet.IsEmpty ) then
  begin
    aCount := ExportToFile( ekSA );
    if ( aCount > 0 ) then UpdateDataBack( ekSA );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ 覆權 }

function TDataControler.InernalExportRA: Boolean;
var
  aSoInfo: PSoInfo;
  aIndex, aCount: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if Assigned( Main.DataThread ) then
      if TDataThread( Main.DataThread ).Terminated then Exit;
    aSoInfo := PSoInfo( FList[aIndex] );
    if not aSoInfo.Selected then Continue;
    if not PrepareDataConnection( aSoInfo ) then Continue;
    DataReader.Close;
    DataReader.SQL.Text := GetSqlText( ekRA, aSoInfo );
    if not PrepareDataSource( aSoInfo ) then Continue;
    {}
    aCount := DataBind( ekRA, aSoInfo );
    if ( aCount > 0 ) then
    begin
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( '系統台:%s, 覆權資料共有%d筆。', [aSoInfo.CompName, aCount] );
      Notify;
    end;
    if Assigned( Main.DataThread ) then
      if TDataThread( Main.DataThread ).Terminated then Exit;
  end;
  if not ( ExportDataSet.IsEmpty ) then
  begin
    aCount := ExportToFile( ekRA );
    if ( aCount > 0 ) then UpdateDataBack( ekRA );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.CreateSoInfo;
var
  aIndex, aCount: Integer;
  aSo: PSoInfo;
  aIni: TIniFile;
begin
  CreateTemporyIniFile(
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + 'CSIS.INI',
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCSIS.INI' );
  aIni := TIniFile.Create( IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) )+ 'TMPCSIS.INI' );
  try
    aCount := aIni.ReadInteger( 'DBINFO', 'DB_COUNT', 0 );
    for aIndex := 0 to aCount - 1 do
    begin
      New( aSo );
      aSo.DbAliase := aIni.ReadString( 'DBINFO', Format( 'ALIAS_%d', [aIndex+1] ), EmptyStr );
      aSo.LoginUser := aIni.ReadString( 'DBINFO', Format( 'USERID_%d', [aIndex+1] ), EmptyStr );
      aSo.LoginPass := aIni.ReadString( 'DBINFO', Format( 'PASSWORD_%d', [aIndex+1] ), EmptyStr );
      aSo.CompCode := aIni.ReadString( 'DBINFO', Format( 'COMPCODE_%d', [aIndex+1] ), EmptyStr );
      aSo.CompName := GetCompName( StrToInt( aSo.CompCode ) );
      aSo.Selected := IsSelectedComp( aSo.CompCode, FSelectedComps );
      FList.Add( aSo );
    end;
  finally
    aIni.Free;
  end;
  if ( FileExists( IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) )+
    'TMPCSIS.INI' ) ) then
    DeleteFile( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + 'TMPCSIS.INI' );
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.FreeSoInfo;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if Assigned( FList.Items[aIndex] ) then
      Dispose( PSoInfo( FList.Items[aIndex] ) );
    FList.Items[aIndex] := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.PrepareExportDataSet;
begin
  UnPrepareDataSet;
  ExportDataSet.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.UnPrepareDataSet;
begin
  if not VarIsNull( ExportDataSet.Data ) then
    ExportDataSet.EmptyDataSet;
  ExportDataSet.Data := Null;  
end;

{ ---------------------------------------------------------------------------- }

class procedure TDataControler.CreateTemporyIniFile(aSource, aDest: String);
var
  aList: TStringList;
  aKey: WideString;
  aObj: _Password;
  aIndex: Integer;
begin
  aKey := 'CS';
  aObj := CoPassword.Create;
  try
    aList := TStringList.Create;
    try
       aList.LoadFromFile(  aSource );
      for aIndex := 0 to aList.Count - 1 do
      begin
        if Pos( 'COMPCODE', UpperCase( aList.Strings[aIndex] ) ) <= 0 then
          aList.Strings[aIndex] := aObj.Decrypt( aList.Strings[aIndex], aKey )
        else
          aList.Strings[aIndex] := aList.Strings[aIndex];
      end;
      aList.SaveToFile( aDest );
    finally
      aList.Free;
    end;
  finally
    aObj := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.DoDataExport(const aKind: TExpoterKind);
begin
  PrepareExportDataSet;
  case aKind of
    ekCA: InernalExportCA;
    ekSA: InernalExportSA;
    ekRA: InernalExportRA;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.GetSqlText(const aKind: TExpoterKind; aSoInfo: PSoInfo): String;
begin
  if ( aKind = ekCA ) then
  begin
    Result := Format(
      ' SELECT A.COMPCODE,              ' +
      '        A.COMMANDTYPE,           ' +
      '        A.CUSTID,                ' +
      '        A.CMMAC,                 ' +
      '        A.DIALACCOUNT,           ' +
      '        A.DIALPASSWORD,          ' +
      '        A.CUSTNAME,              ' +
      '        A.TELDAY,                ' +
      '        A.TELNIGHT,              ' +
      '        A.FINTIME,               ' +
      '        A.SENDDATE,              ' +
      '        A.FTPDATE,               ' +
      '        A.CMDSEQNO,              ' +
      '        A.ID                     ' +
      '   FROM %s.SO312 A               ' +
      '  WHERE A.COMPCODE = ''%s''      ' +
      '    AND UPPER( A.COMMANDTYPE ) = ''SYNC'' ' +
      '    AND A.FTPDATE IS NULL        ' +
      '  ORDER BY A.CMDSEQNO            ',
      [aSoInfo.LoginUser, aSoInfo.CompCode] );
  end else
  if ( aKind = ekSA ) then
  begin
    Result := Format(
      ' SELECT A.COMPCODE,              ' +
      '        A.COMMANDTYPE,           ' +
      '        A.CUSTID,                ' +
      '        A.CMMAC,                 ' +
      '        A.DIALACCOUNT,           ' +
      '        A.DIALPASSWORD,          ' +
      '        A.CUSTNAME,              ' +
      '        A.TELDAY,                ' +
      '        A.TELNIGHT,              ' +
      '        A.FINTIME,               ' +
      '        A.SENDDATE,              ' +
      '        A.FTPDATE,               ' +
      '        A.CMDSEQNO,              ' +
      '        A.ID                     ' +
      '   FROM %s.SO312 A               ' +
      '  WHERE A.COMPCODE = ''%s''      ' +
      '    AND UPPER( A.COMMANDTYPE ) = ''DISABLE'' ' +
      '    AND A.FTPDATE IS NULL        ' +
      '  ORDER BY A.CMDSEQNO            ',
      [aSoInfo.LoginUser, aSoInfo.CompCode] );
  end else
  if ( aKind = ekRA ) then
  begin
    Result := Format( 
      ' SELECT A.COMPCODE,              ' +
      '        A.COMMANDTYPE,           ' +
      '        A.CUSTID,                ' +
      '        A.CMMAC,                 ' +
      '        A.DIALACCOUNT,           ' +
      '        A.DIALPASSWORD,          ' +
      '        A.CUSTNAME,              ' +
      '        A.TELDAY,                ' +
      '        A.TELNIGHT,              ' +
      '        A.FINTIME,               ' +
      '        A.SENDDATE,              ' +
      '        A.FTPDATE,               ' +
      '        A.CMDSEQNO,              ' +
      '        A.ID                     ' +      
      '   FROM %s.SO312 A               ' +
      '  WHERE A.COMPCODE = ''%s''      ' +
      '    AND UPPER( A.COMMANDTYPE ) = ''ENABLE'' ' +
      '    AND A.FTPDATE IS NULL        ' +
      '  ORDER BY A.CMDSEQNO            ',
      [aSoInfo.LoginUser, aSoInfo.CompCode] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.GetExportFileName(const aKind: TExpoterKind): String;
begin
  Result := EmptyStr;
  if ( aKind = ekCA ) then
  begin
    Result := Format( 'CA-%s.txt', [FormatDateTime( 'yyyymmddhhnn', Now )] );
  end else
  if ( aKind = ekSA ) then
  begin
    Result := Format( 'SA-%s.txt', [FormatDateTime( 'yyyymmddhhnn', Now )] );
  end else
  if ( aKind = ekRA ) then
  begin
    Result := Format( 'RA-%s.txt', [FormatDateTime( 'yyyymmddhhnn', Now )] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.GetSoInfo(const aCompCode: String): PSoInfo;
var
  aIndex: Integer;
begin
  Result := nil;
  for aIndex := 0 to FList.Count - 1 do
  begin
    Result := PSoInfo( FList[aIndex] );
    if Result.CompCode = IntToStr( StrToInt( aCompCode ) ) then
      Break
    else
      Result := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.PrepareDataConnection(aSoInfo: PSoInfo): Boolean;
begin
  DataConnection.Close;
  DataConnection.ConnectionString := Format(
    'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
    [aSoInfo.LoginPass, aSoInfo.LoginUser, aSoInfo.DbAliase] );
  try
    DataConnection.Connected := True;
  except
    on E: Exception do
    begin
      FNotify.MsgType := MB_ICONERROR;
      FNotify.MsgText := Format( '系統台:%s, 資料庫連線有誤, 原因:%s。',
       [aSoInfo.CompName, E.Message] );
      Notify;
    end;
  end;
  Result := DataConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.DataBind(const aKind: TExpoterKind; aSoInfo: PSoInfo): Integer;
var
  aSeqNo, aFileName: String;
begin
  Result := 0;

  if ( aKind = ekCA ) then
  begin
    aFileName := Format( 'CA-%s.txt', [FormatDateTime( 'yyyymmddhhnn', Now )] );
  end else
  if ( aKind = ekSA ) then
  begin
    aFileName := Format( 'SA-%s.txt', [FormatDateTime( 'yyyymmddhhnn', Now )] );
  end else
  if ( aKind = ekRA ) then
  begin
    aFileName := Format( 'RA-%s.txt', [FormatDateTime( 'yyyymmddhhnn', Now )] );
  end;

  
  DataReader.First;
  while not DataReader.Eof do
  begin
    try
      aSeqNo := DataReader.FieldByName( 'CMDSEQNO' ).AsString;
      ExportDataSet.Append;
      ExportDataSet.FieldByName( 'COMPCODE' ).AsString := DataReader.FieldByName( 'COMPCODE' ).AsString;
      ExportDataSet.FieldByName( 'COMPNAME' ).AsString := GetCompName( StrToInt( aSoInfo.CompCode ) );
      ExportDataSet.FieldByName( 'COMMANDTYPE' ).AsString := DataReader.FieldByName( 'COMMANDTYPE' ).AsString;
      ExportDataSet.FieldByName( 'CMMAC' ).AsString := DataReader.FieldByName( 'CMMAC' ).AsString;
      ExportDataSet.FieldByName( 'DIALACCOUNT' ).AsString := DataReader.FieldByName( 'DIALACCOUNT' ).AsString;
      ExportDataSet.FieldByName( 'DIALPASSWORD' ).AsString := DataReader.FieldByName( 'DIALPASSWORD' ).AsString;
      ExportDataSet.FieldByName( 'CUSTNAME' ).AsString := DataReader.FieldByName( 'CUSTNAME' ).AsString;
      ExportDataSet.FieldByName( 'TELDAY' ).AsString := DataReader.FieldByName( 'TELDAY' ).AsString;
      ExportDataSet.FieldByName( 'TELNIGHT' ).AsString := DataReader.FieldByName( 'TELNIGHT' ).AsString;
      ExportDataSet.FieldByName( 'FINTIME' ).Value := DataReader.FieldByName( 'FINTIME' ).Value;
      ExportDataSet.FieldByName( 'SENDDATE' ).Value := DataReader.FieldByName( 'SENDDATE' ).Value;
      ExportDataSet.FieldByName( 'FTPDATE' ).Value := DataReader.FieldByName( 'FTPDATE' ).Value;
      ExportDataSet.FieldByName( 'CMDSEQNO' ).AsString := DataReader.FieldByName( 'CMDSEQNO' ).AsString;
      ExportDataSet.FieldByName( 'ID' ).AsString := DataReader.FieldByName( 'ID' ).AsString;
      ExportDataSet.FieldByName( 'FILENAME' ).AsString := aFileName;
      ExportDataSet.Post;
      Inc( Result );
    except
      on E: Exception do
      begin
        ExportDataSet.Cancel;          
        FNotify.MsgType := MB_ICONERROR;
        case aKind of
          ekCA: FNotify.MsgText := '系統台:%s, 裝機完工資料合併有誤, 序號:%s, 原因:%s。';
          ekSA: FNotify.MsgText := '系統台:%s, 裝機完工資料合併有誤, 序號:%s, 原因:%s。';
          ekRA: FNotify.MsgText := '系統台:%s, 裝機完工資料合併有誤, 序號:%s, 原因:%s。';
        end;
        FNotify.MsgText := Format( FNotify.MsgText, [aSoInfo.CompName, aSeqNo,
           E.Message] );
        Notify;
      end;
    end;
    DataReader.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.PrepareDataSource(aSoInfo: PSoInfo): Boolean;
begin
  try
    DataReader.Active := True;
  except
    on E: Exception do
    begin
      FNotify.MsgType := MB_ICONERROR;
      FNotify.MsgText := Format( '系統台:%s, 抓取裝機完工資料有誤, 原因:%s。',
        [aSoInfo .CompName, E.Message] );
      Notify;
    end;
  end;
  Result := DataReader.Active;
end;

{ ---------------------------------------------------------------------------- }

function TDataControler.ExportToFile(const aKind: TExpoterKind): Integer;
var
  aText, aFileName, aFinTime, aSource, aDest: String;
begin
  FExportList.Clear;
  ExportDataSet.First;
  aFileName := ExportDataSet.FieldByName( 'FILENAME' ).AsString;
  {}
  FNotify.MsgType := MB_ICONINFORMATION;
  case aKind of
    ekCA: FNotify.MsgText := Format( '裝機完工資料匯出中, 匯出檔案名稱:%s。', [aFileName] );
    ekSA: FNotify.MsgText := Format( '停權資料匯出中, 匯出檔案名稱:%s。', [aFileName] );
    ekRA: FNotify.MsgText := Format( '覆權資料匯出中, 匯出檔案名稱:%s。', [aFileName] );
  end;
  Notify;
  {}
  while not ExportDataSet.Eof do
  begin
    if ( ExportDataSet.FieldByName( 'FILENAME' ).AsString <> EmptyStr ) then
    begin
      if ( aKind = ekCA ) then
      begin
        aFinTime := EmptyStr;
        if not VarIsNull( ExportDataSet.FieldByName( 'FINTIME' ).Value ) then
          aFinTime := FormatDateTime( 'yyyy-mm-dd hh:nn:ss',
            ExportDataSet.FieldByName( 'FINTIME' ).AsDateTime );
        aText :=
          ExportDataSet.FieldByName( 'DIALACCOUNT' ).AsString + aDelimiter +
          ExportDataSet.FieldByName( 'DIALPASSWORD' ).AsString + aDelimiter +
          ExportDataSet.FieldByName( 'CUSTNAME' ).AsString + aDelimiter +
          ExportDataSet.FieldByName( 'TELDAY' ).AsString + aDelimiter +
          ExportDataSet.FieldByName( 'TELNIGHT' ).AsString + aDelimiter +
          aFinTime  + aDelimiter +
          ExportDataSet.FieldByName( 'CMDSEQNO' ).AsString + aDelimiter +
          ExportDataSet.FieldByName( 'ID' ).AsString;
      end else
      if ( aKind = ekSA ) then
      begin
        aText :=
          ExportDataSet.FieldByName( 'DIALACCOUNT' ).AsString + aDelimiter +
          '1' + aDelimiter +
          ExportDataSet.FieldByName( 'CMDSEQNO' ).AsString;
      end else
      if ( aKind = ekRA ) then
      begin
        aText :=
          ExportDataSet.FieldByName( 'DIALACCOUNT' ).AsString + aDelimiter +
          '0' + aDelimiter +
          ExportDataSet.FieldByName( 'CMDSEQNO' ).AsString;
      end;
      FExportList.Add( aText );
    end;        
    ExportDataSet.Next;
  end;
  if ( FExportList.Count > 0 ) and ( aFileName <> EmptyStr ) then
  begin
    try
      aSource := IncludeTrailingPathDelimiter( LocalTempDir ) + aFileName;
      aDest := IncludeTrailingPathDelimiter( LocalUploadDir ) + aFileName;
      FExportList.SaveToFile( aSource );
      MoveFile( PChar( aSource ), PChar( aDest ) );
      {}
      FNotify.MsgType := MB_ICONINFORMATION;
      case aKind of
        ekCA: FNotify.MsgText := Format( '裝機完工資料匯出完成, 共計%d筆資料。', [FExportList.Count] );
        ekSA: FNotify.MsgText := Format( '停權資料匯出完成, 共計%d筆資料。', [FExportList.Count] );
        ekRA: FNotify.MsgText := Format( '覆權資料匯出完成, 共計%d筆資料。', [FExportList.Count] );
      end;
      Notify;
      {}
    except
      on E: Exception do
      begin
        FNotify.MsgType := MB_ICONERROR;
        case aKind of
          ekCA: FNotify.MsgText := '裝機完工資料匯出發生錯誤, 原因:%s。' ;
          ekSA: FNotify.MsgText := '停權資料匯出發生錯誤, 原因:%s。' ;
          ekRA: FNotify.MsgText := '覆權資料匯出發生錯誤, 原因:%s。' ;
        end;
        FNotify.MsgText := Format( FNotify.MsgText, [E.Message] );
        Notify;
        FExportList.Clear;
      end;
    end;
  end;
  Result := FExportList.Count;
  if Result > 0 then FExportList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.UpdateDataBack(const aKind: TExpoterKind);
var
  aCompCode, aDate: String;
  aSoInfo: PSoInfo;
  aPos1, aPos2: Integer;
begin
  aCompCode := EmptyStr;
  ExportDataSet.First;
  while not ExportDataSet.Eof do
  begin
    if ( aCompCode <> ExportDataSet.FieldByName( 'COMPCODE' ).AsString ) then
    begin
      aSoInfo := GetSoInfo( ExportDataSet.FieldByName( 'COMPCODE' ).AsString );
      aCompCode := ExportDataSet.FieldByName( 'COMPCODE' ).AsString;
      PrepareDataConnection( aSoInfo );
    end;
    if ( ExportDataSet.FieldByName( 'FILENAME' ).AsString <> EmptyStr ) then
    begin
      aPos1 := Pos( '-', ExportDataSet.FieldByName( 'FILENAME' ).AsString );
      aPos2 := Pos( '.', ExportDataSet.FieldByName( 'FILENAME' ).AsString );
      aDate := Copy( ExportDataSet.FieldByName( 'FILENAME' ).AsString,
        aPos1 + 1, aPos2 - ( aPos1 + 1 ) );
      DataWriter.SQL.Text := Format(
        ' UPDATE %s.SO312                                         ' +
        '    SET FTPDATE = TO_DATE( ''%s'', ''YYYYMMDDHH24MI'' )  ' +
        '  WHERE COMPCODE = ''%s''                                ' +
        '    AND CMDSEQNO = ''%s''                                ',
        [aSoInfo.LoginUser, aDate,
         ExportDataSet.FieldByName( 'COMPCODE' ).AsString,
         ExportDataSet.FieldByName( 'CMDSEQNO' ).AsString] );
       try
         DataWriter.ExecSQL;
       except
         on E: Exception do
         begin
           FNotify.MsgType := MB_ICONINFORMATION;
           case aKind of
             ekCA: FNotify.MsgText := '裝機完工資料更新上傳狀態發生錯誤, 原因:%s。' ;
             ekSA: FNotify.MsgText := '停權資料更新上傳狀態發生錯誤, 原因:%s。' ;
             ekRA: FNotify.MsgText := '覆權資料更新上傳狀態發生錯誤, 原因:%s。' ;
           end;
           FNotify.MsgText := Format( FNotify.MsgText, [E.Message] );
           Notify;
         end;
       end;
    end;
    ExportDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

type
  TImportField = record
    Account: String;
    CmdSeqNo: String;
    Success: String;
    ErrorText: String;
    CommandType: String;
    FaciSno: String;
    CustId: String;
  end;

function TDataControler.DoDataImport(aFileName: String): Boolean;
var
  aIndex: Integer;
  aText: String;
  aRecord: TImportField;
  aSoInfo: PSoInfo;
  aHasError: Boolean;
begin
  Result := True;
  if not FileExists( aFileName ) then Exit;
  aHasError := False;
  FImportList.Clear;
  FImportList.LoadFromFile( aFileName );
  for aIndex := 0 to FImportList.Count - 1 do
  begin
    try
      aText := FImportList[aIndex];
      {}
      aRecord.Account := EmptyStr;
      aRecord.CmdSeqNo := EmptyStr;
      aRecord.Success := EmptyStr;
      aRecord.ErrorText := EmptyStr;
      {}
      aRecord.Account := ExtractValue( aText, aDelimiter );
      aRecord.CmdSeqNo := ExtractValue( aText, aDelimiter );
      aRecord.Success := ExtractValue( aText, aDelimiter );

      if ( aRecord.Success = '0' ) then
        aRecord.Success := '1'
      else
        aRecord.Success := '2';
      aRecord.ErrorText := Copy( ExtractValue( aText, aDelimiter ), 1, 30 );
      aSoInfo := GetSoInfo( Copy( aRecord.CmdSeqNo, 1, 2 ) );
      if not Assigned( aSoInfo ) then
      begin
        aHasError := True;
        FNotify.MsgType := MB_ICONERROR;
        FNotify.MsgText := '處理回覆檔時發生錯誤, 對應不到系統台, 無法更新資料。';
        SynchronizeNotify;
        Continue;
      end;
      if not PrepareDataConnection( aSoInfo ) then Continue;
      {}
      DataWriter.SQL.Text := Format(
        ' SELECT UPPER( A.COMMANDTYPE ) AS COMMANDTYPE, ' +
        '        A.CUSTID                               ' + 
        '   FROM %s.SO312 A                             ' +
        '  WHERE A.CMDSEQNO = ''%s''                    ',
        [aSoInfo.LoginUser, aRecord.CmdSeqNo] );
      DataWriter.Open;
      aRecord.CommandType := DataWriter.FieldByName( 'COMMANDTYPE' ).AsString;
      aRecord.CustId := DataWriter.FieldByName( 'CUSTID' ).AsString;
      DataWriter.Close;
      {}
      DataWriter.SQL.Text := Format(
        ' SELECT A.CMMAC FROM %s.SO311 A           ' +
        '  WHERE A.CMDSEQNO = ''%s''               ',
        [aSoInfo.LoginUser, aRecord.CmdSeqNo] );
      DataWriter.Open;
      aRecord.FaciSno := DataWriter.FieldByName( 'CMMAC' ).AsString;
      DataWriter.Close;
      {}
      DataWriter.SQL.Text := Format(
        '  UPDATE %s.SO311                 ' +
        '     SET QUERYRESULT = ''%s'',    ' +
        '         FAULTREASON = ''%s''     ' +
        '   WHERE CMDSEQNO = ''%s''        ',
        [aSoInfo.LoginUser, aRecord.Success, aRecord.ErrorText, aRecord.CmdSeqNo] );
      DataWriter.ExecSQL;
      {}
      if ( aRecord.CommandType = 'ENABLE' ) or ( aRecord.CommandType = 'SYNC' ) then
      begin
        DataWriter.SQL.Text := Format(
          '  UPDATE %s.SO004                   ' +
          '     SET ENABLEACCOUNT = SYSDATE,   ' +
          '         UPDEN = ''APTG'',          ' +
          '         UPDTIME = TO_CHAR( TO_NUMBER( TO_CHAR( SYSDATE, ''YYYY'' ) ) - 1911 ) || ''/'' || ' +
          '                   TO_CHAR( SYSDATE, ''MM/DD HH24:MI:SS'' )                                ' +
          '   WHERE CUSTID = ''%s''            ' +
          '     AND FACISNO = ''%s''           ' +
          '     AND DIALACCOUNT = ''%s''       ',
          [aSoInfo.LoginUser, aRecord.CustId, aRecord.FaciSno, aRecord.Account] );
      end else
      begin
        DataWriter.SQL.Text := Format(
          '  UPDATE %s.SO004                   ' +
          '     SET DISABLEACCOUNT = SYSDATE,  ' +
          '         UPDEN = ''APTG'',          ' +
          '         UPDTIME = TO_CHAR( TO_NUMBER( TO_CHAR( SYSDATE, ''YYYY'' ) ) - 1911 ) || ''/'' || ' +
          '                   TO_CHAR( SYSDATE, ''MM/DD HH24:MI:SS'' )                                ' +
          '   WHERE CUSTID = ''%s''            ' +
          '     AND FACISNO = ''%s''           ' +
          '     AND DIALACCOUNT = ''%s''       ', 
          [aSoInfo.LoginUser, aRecord.CustId, aRecord.FaciSno, aRecord.Account] );
      end;
      DataWriter.ExecSQL;
    except
      on E: Exception do
      begin
        aHasError := True;
        FNotify.MsgType := MB_ICONERROR;
        FNotify.MsgText := Format( '處理回覆檔時發生錯誤, 行號:%d, 帳號:%s, 序號:%s, 原因:%s。',
          [aIndex, aRecord.Account, aRecord.CmdSeqNo, E.Message] );
        SynchronizeNotify;
      end;
    end;
  end;
  Result := not aHasError;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.SynchronizeNotify;
begin
  if Assigned( Main.DataThread ) then
  begin
    if not TDataThread( Main.DataThread ).Terminated then
    begin
      TDataThread( Main.DataThread ).Synchronize( Notify );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataControler.Notify;
begin
  Main.AddMsg( FNotify );
end;

{ ---------------------------------------------------------------------------- }

end.
