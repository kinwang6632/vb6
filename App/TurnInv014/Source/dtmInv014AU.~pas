unit dtmInv014AU;

interface

uses
  SysUtils, Classes, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, cxLookAndFeels, IniFiles, DB, ADODB
  , Dialogs, ImgList, Controls, DBClient, DBTables;
const
  INV_SYS_INFO_INI_FILE = 'Connection.ini';
  TMP_INV_SYS_INFO_INI_FILE = 'Tmp.ini';
  TurnLog = 'Turn_Log.Txt';
  IDENTIFYID1 = '1';
  IDENTIFYID2 = '0';

type
  PSMSDb = ^TSMSDb;
  TSMSDb = record
    ConnSeq: Integer;
    OracleSid: String;
    Description: String;
    UserId: String;
    Password: String;
    ReportPasth: String;
    MisDbOwner: String;
    MisDbPassword: String;
    MisDbCompCode: String;
    MisDbCompName: String;
    MisDbSid: String;
    MisDbLink: String;
  end;
type
  PCompany = ^TCompany;
  TCompany = record
    CompanyId: String;
    CompanyName: String;
    Tel: String;
    CompanyLongName: String;
    ServiceType: String;
    LinkToMIS: Boolean;
    ConnSeq: Integer;
    AutoCreateNum: Integer;
    ReportName1: String;
    ReportName2: String;
    ReportName3: String;
    ReportName4: String;
    ReportName5: String;
    ReportName6: String;
    IfPrintTitle: Boolean;
    IfPrintAddr: Boolean;
    IfPrintCheck: Boolean;
    FaciCombine: Integer;
    ExpAddrType: Integer;
  end;

type
  TdtmInv014 = class(TDataModule)
    cxLookAndFeelController1: TcxLookAndFeelController;
    InvConnection: TADOConnection;
    adoComm: TADOQuery;
    MainImageList: TImageList;
    adoInv014: TADOQuery;
    adoInv014A: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FSMSDbList: TList;
    FMisDbOwner: String;     // 客服系統的 Owner
    FMisDbCompCode : String;
    sG_DbSID: String;
    sG_DbUserID: String;
    sG_DbPassword: String;
    sG_ReportPath : String;
  public
    procedure GetSMSDb(const aFileName: String); overload;
    property SMSDbList: TList read FSMSDbList;
    function ConnectToINV(aObj: PSMSDb): Boolean;
    function getMisDbOwner : string;
    procedure GetInvoiceCompany(aList: TList);
    function GetSoInfo(const aConnSeq: Integer): PSMSDb;
    function ConnectToSO(aObj: PSMSDb): Boolean;
    { Public declarations }
  end;

var
  dtmInv014: TdtmInv014;

implementation
uses dtmSOU;

{$R *.dfm}

{ TdtmInv014 }

procedure TdtmInv014.GetSMSDb(const aFileName: String);
var
  aIndex: Integer;
  aIniFile: TIniFile;
  aStrList: TStringList;
  aSMSDbId: String;
  aObj: PSMSDb;
begin
  aStrList := TStringList.Create;
  try
    aIniFile := TIniFile.Create( aFileName );
    try
      aIniFile.ReadSection( 'DATAAREA', aStrList );
      if not Assigned( FSMSDbList ) then FSMSDbList := TList.Create;
      FSMSDbList.Clear;
      for aIndex :=0 to aStrList.Count-1 do
      begin
        aSMSDbId := aStrList.Strings[aIndex];
        New( aObj );
        aObj.ConnSeq := StrToInt( aSMSDbId );
        aObj.OracleSid := aIniFile.ReadString( aSMSDbId, 'SID', EmptyStr );
        aObj.Description := aIniFile.ReadString( 'DATAAREA', aSMSDbId, EmptyStr );
        aObj.UserId := aIniFile.ReadString( aSMSDbId, 'DB_USER', EmptyStr );
        aObj.Password := aIniFile.ReadString( aSMSDbId, 'DB_USER_PASSWORD', EmptyStr );
        aObj.ReportPasth := aIniFile.ReadString( aSMSDbId, 'REPORT_PATH', EmptyStr );
        aObj.MisDbOwner := aIniFile.ReadString( aSMSDbId, 'MISDB_OWNER', EmptyStr );
        aObj.MisDbPassword := aIniFile.ReadString( aSMSDbId, 'MISDB_PASSWORD', EmptyStr );
        aObj.MisDbCompCode := aIniFile.ReadString( aSMSDbId, 'MISDB_COMPCODE', EmptyStr );
        aObj.MisDbCompName := aIniFile.ReadString( aSMSDbId, 'MISDB_COMPNAME', EmptyStr );
        aObj.MisDbSid := aIniFile.ReadString( aSMSDbId, 'MISDB_SID', EmptyStr );
        aObj.MisDbLink := aIniFile.ReadString( aSMSDbId, 'MISDB_DBLINK', EmptyStr );
        FSMSDbList.Add( aObj );
      end;
    finally
      aIniFile.Free;
    end;
  finally
    aStrList.Free;
  end;
end;

procedure TdtmInv014.DataModuleCreate(Sender: TObject);
begin
  FSMSDbList := TList.Create;
end;

function TdtmInv014.ConnectToINV(aObj: PSMSDb): Boolean;
begin
  Result := False;
  try
    if not Assigned( aObj ) then
      raise Exception.Create( '對應不到發票資料區。' );
    InvConnection.Connected := False;
    InvConnection.ConnectionString := Format(
       'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
       'Data Source=%s;Persist Security Info=True',
       [aObj.Password, aObj.UserId, aObj.OracleSid] );
    InvConnection.Connected := True;
    sG_DbSID := aObj.OracleSid;
    sG_DbUserID := aObj.UserId;
    sG_DbPassword:= aObj.Password;
    sG_ReportPath:= aObj.ReportPasth;
    Result := True;
  except
    on E: Exception do
    begin
      MessageDlg( Format( '【發票】資料區連結失敗, 訊息:%s', [E.Message] ),mtError,[mbOK],0 );
      Exit;
    end;
  end;
end;

procedure TdtmInv014.GetInvoiceCompany(aList: TList);
var
  aIndex: Integer;
  aObj: PCompany;
begin
  if not Assigned( aList ) then Exit;
  for aIndex := 0 to aList.Count - 1 do
  begin
    if Assigned( aList.Items[aIndex] ) then
    begin
      Dispose( PCompany( aList.Items[aIndex] ) );
      aList.Items[aIndex] := nil;
    end;
  end;
  aList.Clear;
  adoComm.Close;
  adoComm.SQL.Text :=
    ' SELECT COMPSNAME, COMPID, RPTNAME1, RPTNAME2,   ' +
    '        RPTNAME3, RPTNAME4, RPTNAME5, RPTNAME6,  ' +
    '        SERVICETYPESTR, LINKTOMIS, CONNSEQ,      ' +
    '        AUTOCREATENUM, COMPNAME, FACICOMBINE,    ' +
    '        TEL                                      ' +
    '   FROM INV001         ' +
    '  WHERE IDENTIFYID1 =  ' + QuotedStr( IDENTIFYID1 ) +
    '    AND IDENTIFYID2 =  ' + QuotedStr( IDENTIFYID2 ) +
    '  ORDER BY TO_NUMBER( COMPID ) ';
  adoComm.Open;
  adoComm.First;
  while not adoComm.Eof do
  begin
    New( aObj );
    aObj.CompanyId := dtmInv014.adoComm.FieldByName( 'COMPID' ).AsString;
    aObj.CompanyName := dtmInv014.adoComm.FieldByName( 'COMPSNAME' ).AsString;
    aObj.Tel := dtmInv014.adoComm.FieldByName( 'TEL' ).AsString;
    aObj.CompanyLongName := dtmInv014.adoComm.FieldByName( 'COMPNAME' ).AsString;
    aObj.ServiceType := dtmInv014.adoComm.FieldByName( 'SERVICETYPESTR' ).AsString;
    aObj.LinkToMIS := ( dtmInv014.adoComm.FieldByName( 'LINKTOMIS' ).AsString = 'Y' );
    aObj.ConnSeq := dtmInv014.adoComm.FieldByName( 'CONNSEQ' ).AsInteger;
    aObj.AutoCreateNum := dtmInv014.adoComm.FieldByName( 'AUTOCREATENUM' ).AsInteger;
    aObj.ReportName1 := dtmInv014.adoComm.FieldByName( 'RPTNAME1' ).AsString;
    aObj.ReportName2 := dtmInv014.adoComm.FieldByName( 'RPTNAME2' ).AsString;
    aObj.ReportName3 := dtmInv014.adoComm.FieldByName( 'RPTNAME3' ).AsString;
    aObj.ReportName4 := dtmInv014.adoComm.FieldByName( 'RPTNAME4' ).AsString;
    aObj.ReportName5 := dtmInv014.adoComm.FieldByName( 'RPTNAME5' ).AsString;
    aObj.ReportName6 := dtmInv014.adoComm.FieldByName( 'RPTNAME6' ).AsString;
    aObj.FaciCombine := dtmInv014.adoComm.FieldByName( 'FACICOMBINE' ).AsInteger;
    aObj.IfPrintTitle := False;
    aObj.IfPrintAddr := False;
    aObj.IfPrintCheck := True;
    aList.Add( aObj );
    adoComm.Next;
  end;
  adoComm.Close;
  { 讀取發票檔參數檔 }
  for aIndex := 0 to aList.Count - 1 do
  begin
    adoComm.Close;
    adoComm.SQL.Text :=
      ' SELECT IFPRINTTITLE, IFPRINTADDR, IFPRINTCHECK,  ' +
      '        EXPADDRTYPE    ' +                           
      '   FROM INV003         ' +
      '  WHERE IDENTIFYID1 =  ' + QuotedStr( IDENTIFYID1 ) +
      '    AND IDENTIFYID2 =  ' + QuotedStr( IDENTIFYID2 ) +
      '    AND COMPID =       ' + QuotedStr( PCompany( aList[aIndex] ).CompanyId ) +
      '  ORDER BY TO_NUMBER( COMPID ) ';
    adoComm.Open;
    if not adoComm.IsEmpty then
    begin
      PCompany( aList[aIndex] ).IfPrintTitle := (
        adoComm.FieldByName( 'IFPRINTTITLE' ).AsString = 'Y' );
      PCompany( aList[aIndex] ).IfPrintAddr := (
        adoComm.FieldByName( 'IFPRINTADDR' ).AsString = 'Y' );
      PCompany( aList[aIndex] ).IfPrintCheck := (
        adoComm.FieldByName( 'IFPRINTCHECK' ).AsString = 'Y' );
      PCompany( aList[aIndex] ).ExpAddrType :=
        adoComm.FieldByName( 'EXPADDRTYPE' ).AsInteger;
    end;
    adoComm.Close;
  end;
end;

function TdtmInv014.getMisDbOwner: string;
begin
  Result := UpperCase(FMisDbOwner);
end;

function TdtmInv014.GetSoInfo(const aConnSeq: Integer): PSMSDb;
var
  aIndex: Integer;
begin
  Result := nil;
  for aIndex := 0 to FSMSDbList.Count - 1 do
  begin
    if PSMSDb( FSMSDbList[aIndex] ).ConnSeq = aConnSeq then
    begin
      Result := PSMSDb( FSMSDbList[aIndex] );
      Break;
    end;
  end;
end;

function TdtmInv014.ConnectToSO(aObj: PSMSDb): Boolean;
begin
  Result := False;
  try
   if not Assigned( aObj ) then
     raise Exception.Create( '對應不到系統台資料區。' );
   InvConnection.Connected := False;
   InvConnection.ConnectionString :=Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [aObj.MisDbPassword, aObj.MisDbOwner, aObj.MisDbSid] );
   InvConnection.Connected := True;
   FMisDbOwner := aObj.MisDbOwner;
   FMisDbCompCode := aObj.MisDbCompCode;
   Result := True;
  except
    on E: Exception do
    begin
      MessageDlg( Format( '【客服系統】資料區連結失敗, 訊息:%s', [E.Message] ),mtWarning,[mbOK],0);
      Exit;
    end;
  end;
end;

procedure TdtmInv014.DataModuleDestroy(Sender: TObject);
begin
  FSMSDbList.Free;
end;

end.
