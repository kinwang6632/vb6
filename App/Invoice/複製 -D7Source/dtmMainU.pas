unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB ,Forms ,IniFiles, Dialogs, DBClient,Controls,
  Provider, xmldom, XMLIntf, msxmldom, XMLDoc, ImgList, 
  cxLookAndFeels, cxContainer, cxEdit, cxStyles, dxBar, cxPropertiesStore,
  cxDropDownEdit, dxSkinsCore, dxSkinsDefaultPainters, dxSkinsdxBarPainter,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee, ShellAPI, Windows;

const

    // 版本
    VERSION = 'v.20060417';

    // 自訂 Grid 儲存路徑
    GRID_FOLDER = 'GridColumn';

    // FastReport 報表 Template 資料夾名稱
    REPORT_FOLDER = 'ReportTemplate';

    CHARGESOURCE33 = 'SO033';
    CHARGESOURCE34 = 'SO034';

    INV_SYS_INFO_INI_FILE = 'Connection.ini';
    TMP_INV_SYS_INFO_INI_FILE = 'Tmp.ini';
    TMP_INV_COM_INFO_INI_FILE = 'Tmp2.ini';
    Date_Format = 'YYYY/MM/DD';
    ERRORMSGFILENAME = '發票開立之錯誤訊息.txt';
    STR_SEP = '''';
    PRINT_TYPE_FILE = 'PrintType.txt';
    EXPIRE_SECTION = 'EXPIRE';
    EXPIRE_DATE = 'DATE';
    INV_COM_INI_FILE = 'Com.ini';
    //開立發票異常資料查詢 Type
    CONDITION_TYPE_BILLNO = 0;
    CONDITION_TYPE_CUSTID = 1;
    CONDITION_TYPE_CHARGEDATE = 2;
    CONDITION_TYPE_LOGDATE = 3;
    CONDITION_TYPE_ASSIGN_LOGDATE = 4;   //整批開立自動產生查詢異常資料

    INV_ELECTRON_UPLOAD_EXE = 'ELECTRONINV.EXE'; //執行產生上傳電子發票檔(B07)
    INV_ELECTRON_PROClIST_EXT = 'CableSoft.Invoice.Win.B08.exe'; //執行產生電子清單(B08)

    //未開發票資料查詢模式
    QUERY_TYPE_ALL = '0';
    QUERY_TYPE_CUSTID = '1';
    QUERY_TYPE_BILLID = '2';
    QUERY_TYPE_CHARGEDATE = '3';

    IDENTIFYID1 = '1';
    IDENTIFYID2 = '0';

    DEFAULT_ALL_DATA_ITEM_CODE_VALUE = '999';
    DEFAULT_ALL_DATA_ITEM_DESC_VALUE = '全部';

    CODE_VALUE_SEP = '-';
    EMPTY_DATE_STR = '/  /';
    EMPTY_YM_STR = '/';

    BUSINESS_ID_LENGTH = 8;
    BUSINESS_ID_CHECK_NUM = '12121241';//統編邏輯乘數


    USER_INFO_FILENAME = 'user.dat';

    FORM_INI_FILE_NAME = 'LanguageSetting.ini';

    FORM_INI_SIMPLE_CHINESE = 'FORM_1';

    SYS_INI_FILE_NAME = 'GateWayParam.ini';
    TMP_SYS_INI_FILE_NAME = 'TempGateWayParam.ini';

    DATA_AREA_HEADER = 'COMPINFO';

    SELECT_MODE = 1;
    IUD_MODE = 2;

    SGS_UI_MAX_COUNT=500;
    LISTVIEW_SUB_ITEM_COUNT = 9;

    TESTING_CMD_COMP_CODE = '99';


    EXCHANGE_DATE_QUERY = 'ExchangeDateQuery';
    IC_CARD_QUERY = 'ICCardQuery';
    CA_PRODUCT_QUERY = 'CAProductQuery';
    PROD_PURCHASE_QUERY = 'ProdPurchaseQuery';
    ENTITLEMENT_QUERY ='EntitlementQuery';


    XML_ENCODING_SC = 'GB2312';//簡體
    XML_ENCODING_TC = 'BIG5';//繁體
    XML_STANDALONE = 'yes';


    MAX_FILE_SIZE = 1000;  //大於1MB壓縮
    IIS_HISTORY ='SGS1';
    IIS_IMMEDIATELY ='SGS2';



    SHOW_UI_LINE_COUNTS = 1000;
    SHOW_UI_COUNTS = 1000;
    MAX_STRING_LIST_COUNTS = 10000001;
    FLUSH_XML_WRITE = 100;

    ENCRYPTION_SEP = 'H';
    ENCRYPTION_KEY = '1234';
    REGISTRY_ROOT = 'SOFTWARE\Windows\Sys';
    SYS4_Y = 'A';
    SYS4_N = '5';
    SYS5_Y = '9';
    SYS5_N = 'B';

    TAB_STRING = ''#9'';

var
    PASSKEY: WideString = 'CS';    

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
    EInvDb:string;
    EInvUser:string;
    EInvPassword: String;
    USEPrizeCom : Boolean;
  end;

type
  PPRIZEDb = ^TPRIZEDb;
  TPRIZEDb = record
    ConnSeq: Integer;
    DBSid: string;
    UserId: String;
    Password:String;
    CompTitle:String;
  end;
type
  PRIZECOMDb = ^TPRIZECOMDb;
  TPRIZECOMDb = record
    DBSid : String;
    UserId: String;
    Password:String;
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
    ShowFaci: Integer;
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
    StarEInvoice: Boolean;
    StarEmail: Boolean;
    StarMessage: Boolean;
    StarTVMail: Boolean;
    StarCMSend: Boolean;
    StartAutoNotify : Boolean;
  end;

type
  TOracleDataBaseError = record
    ErrorType: String;
    ErrorCode: Integer;
    ErrorMsg: String;
  end;

  var OracleErros: TOracleDataBaseError;

type
  TQueryParam = record
    Param1: String;
    Param2: String;
    Param3: String;
    Param4: String;
    Param5: String;
    Param6: String;
  end;

type
  TdtmMain = class(TDataModule)
    InvConnection: TADOConnection;
    adoComm: TADOQuery;
    adoComm2: TADOQuery;
    adoComm3: TADOQuery;
    adoCompetence: TADOQuery;
    cdsCompetence: TClientDataSet;
    getXMLDocument: TXMLDocument;
    cdsCompetenceID: TStringField;
    cdsCompetenceQuery: TStringField;
    cdsCompetenceInsert: TStringField;
    cdsCompetenceUpdate: TStringField;
    cdsCompetenceDelete: TStringField;
    cdsCompetenceExecute: TStringField;
    cxEditStyle: TcxEditStyleController;
    cxStyleRepository1: TcxStyleRepository;
    cxBandHeaderStyle: TcxStyle;
    cxGridBackGroundStyle: TcxStyle;
    cxLabelStyle: TcxEditStyleController;
    cxLabelNotNullStyle: TcxEditStyleController;
    cxLookAndFeel: TcxLookAndFeelController;
    cxGridInActiveStyle: TcxStyle;
    cxGridNotEditStyle: TcxStyle;
    cxGridActiveStyle: TcxStyle;
    ActionImageList: TImageList;
    MsgImageList: TImageList;
    PageControlImageList: TImageList;
    EInvConnection: TADOConnection;
    adoEComm: TADOQuery;
    EInvConnectComm: TADOConnection;
    cdsCompetenceVerify: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    sG_DbSID: String;
    sG_DbUserID: String;
    sG_DbPassword: String;
    sG_ReportPath : String;
    FMisDbOwner: String;     // 客服系統的 Owner
    FMisDbCompCode: String;  // 客服系統的系統台代碼
    FMisDbCompName: String;  // 客服系統的系統台名稱
    FMisDbLink: String;
    FSMSDbList: TList;
    FPrizeDbList : TList;
    FCOMDbList : TList;
    procedure ReleaseSoInfo;
    procedure ReleasePrizeInfo;
    procedure ReleaseComInfo;
  public
    { Public declarations }
    sG_CompID: String;
    sG_CompName: String;
    sG_CompLongName: String;
    sG_CompTel: String;
    sG_ServiceTypeStr: String;
    sG_LoginUser: String;
    sG_LoginUserName: String;
    sG_GroupId : String;
    sG_LinkToMIS: Boolean;
    sG_ConnSeq: Integer;
    sG_AutoCreateNum: Integer;
    sG_ShowFaci:Integer;
    sG_IfPrintTitle: Boolean;
    sG_IfPrintAddr: Boolean;
    sG_IfPrintCheck: Boolean;
    sG_FaciCombine: Integer;
    sG_ExpAddrType: Integer;
    sG_StarEInvoice: Boolean;
    sG_StarEmail: Boolean;
    sG_StarTVmail: Boolean;
    sG_StarCMSend: Boolean;
    sG_StarMessage: Boolean;
    sG_StarAutoNotify : Boolean;
    sG_EInvDb : String;
    sG_EInvUser: String;
    sG_EInvPassword : String;
    sG_UseCOM : Boolean;

    procedure SetWaitingCursor;
    procedure SetDefaultCursor;
    procedure SaveToFile(aContent: OleVariant; aFileName:String);
    procedure GetCompetence(aId: String; var aQuery, aInsert, aDelete,
      aUpdate, aExecute,aVerify: String);
    function ConnectToINV(aObj: PSMSDb): Boolean;
    function ConnectToEINV(aObj: PRIZECOMDb): Boolean; overload;
    function ConnectToEINV(aSid,aUserId,aPassword : String):Boolean; overload;
    function ConnectToEINVComm(aSid,aUserId,aPassword : String):Boolean;
    function ConnectToSO(aObj: PSMSDb): Boolean;
    function checkBusinessID(sI_BusinessID: String): boolean;
    function CheckPasswd(const aUserID, aPasswd: String;
      var aGroupId, aUserName: String): Boolean;
    function LoadCompetence(aGroupId: String): String;
    function getDbSID: String;
    function getDbUserID: String;
    function getDbPassword: String;
    function getReportPath: String;
    function getMisDbOwner: String;
    function getMisDbCompCode : String;
    function getCompID: String;
    function getCompIDEx: String;
    function getCompIdAndName: String;
    function getLoginUser: String;
    function getLoginUserName: String;
    function getCompName: String;
    function getCompNameEx: String;
    function getCompLongName: String;
    function getCompTel: String;
    function getGroupId: String;
    function GetServiceTypeStr: String;
    function GetServiceTypeStrEx: String;
    function GetLinkToMIS: Boolean;
    function GetConnSeq: Integer;
    function GetAutoCreateNum: Integer;
    function GetShowFaci:Integer;
    function GetMisDbLink: String;
    function GetIfPrintTitle: Boolean;
    function GetIfPrintAddr: Boolean;
    function GetIfPrintCheck: Boolean;
    function GetFaciCombine: Integer;
    function GetExpAddrType: Integer;
    function GetStarEInvoice: Boolean;
    function GetStarEmail: Boolean;
    function GetStarMessage: Boolean;
    function GetStarTVmail:Boolean;
    function GetStarCMSend: Boolean;
    function GetSoInfo(const aConnSeq: Integer): PSMSDb;
    function GetEInvDb: String;
    function GetEInvUser: String;
    function GetEInvPassword: String;
    function GetUseCOM:Boolean;
    procedure ShellApplication(AppName, Caption : String; Param : array of String);
    procedure GetInvoiceCompany(aList: TList);
    procedure GetSMSDb(const aFileName: String); overload;
    procedure GetSMSDb;overload;
    procedure GetCOMDb(const aFileName: String) ;
    property SMSDbList: TList read FSMSDbList;
    property PRIZEDbList: TList read FPrizeDbList;
    property ComDbList : TList read FCOMDbList;

  end;


function ParseStrings(Strs:string; Sep:Char) : TStringList;
function getInvoiceYearMonth(sI_Date : String ) :String;
function getYearMonthDay(sI_Date : String ):String ;
function getYearMonthDay2(sI_Date : String ):String ;
function getYearMonthDay3(sI_Date : String; nI_Index: Integer ):String ;
function getYearMonthDay4(sI_Date : String ):String ;
function getYearMonthDay5(sI_Date : String; nI_Index: Integer ):String ;
function getYearMonthDay6(dI_Date : TDate ):String ;
function getYearMonthDay7(dI_Date : TDate ;Di_Integer:Integer):String ;
function getYearMonthDay8(sI_Date : String; nI_Index: Integer ):String ;
function adjString(nI_FieldLength:Integer;sI_Value:String;bI_Right,bI_isInteger:boolean):String;
function adjString2(nI_FieldLength: Integer; sI_Value: Double): String;
function getOracleSQLDateTimeStr(dI_Date: TDateTime): String;
function CutAllChar(S:string;C:char):string;
function AssignSpace(nI_Count: integer;bL_isInteger:boolean):String;

function _CutLeft(S:string;L:integer):string; {減掉字串中左邊若干個字元}
function _Right(S:string;L:integer):string; {同 Clipper RIGHT() }
function _Left(S:string;L:integer):string; {同 Clipper Left() }
function _monthLastDay(yyyymm:string ; onlyday:integer):string; {求西元年月之該月最後一天日期字串 ; onlyday=1 表只傳回該月天數字串}
function _StrToInt(S:string):LONGINT; {與 DELPHI 之 strtoint() 不同,格式不對其傳回 0 , 而不是錯誤訊息}
function _Int(R:REAL):LONGINT; {與 DELPHI 之 INT() 不同,其傳回 LONGINT , 而非 REAL ( 如123.0 ) }
function _Value(S:string):Extended;
function _isLeapYear(yyyy:string):Boolean; {求西元年字串是否為閏年}
function _SubStrNum(S,SUBS:string):integer; {取字串中出現某一子字串之次數 }
function _AllTrim(cText: string): string; { 將字串中左右所有的空白字元刪除 }
function _LTrim(S:string):string;  { 去字串 左空白}
function _RTrim(S:string):string; {同 Clipper RTRIM() 去掉字串右邊空白}

function GetGridStoragePath(const aGridName: String): String;
function TranslateError(const aErrMsg: String): Boolean;
function MappingChineseNumber(const aNumber: Integer): String;
function FacisNoAddMask(aText: String): String;
function MapingInvoiceYearMonth(const aDate: TDate): String; { 日期對應到所屬發票年月 }

type
  TComboBoxCreateParam = record
    Sql: String;
    KeyField: String;
    DescField: String;
    DelimiterText: String;
    AddAllText: Boolean;
    AllTextCode: String;
  end;

function CreateCxComboBoxItem(aComboBox: TcxComboBox;
  aParam:TComboBoxCreateParam): Integer;

type
  TComboBoxItemParam = record
    KeyValue: String;
    DesValue: String;
    DelimiterText: String;
  end;

function GetCxComboBoxItemValue(aComboBox: TcxComboBox;
  var aParam: TComboBoxItemParam): Boolean;

procedure SetCxComboBoxItemValue(aComboBox: TcxComboBox;
  aValue: String);

type
  TCommonParam = record
    InStr1: String;
    InStr2: String;
    InStr3: String;
    InStr4: String;
    InStr5: String;
    InStr6: String;
    InStr7: String;
    InStr8: String;
    InStr9: String;
    InStr10: String;
    InInt1: Integer;
    InInt2: Integer;
    InInt3: Integer;
    InInt4: Integer;
    InInt5: Integer;
    InInt6: Integer;
    OutStr1: String;
    OutStr2: String;
    OutStr3: String;
    OutStr4: String;
    OutStr5: String;
    OutStr6: String;
    OutStr7: String;
    OutStr8: String;
    OutInt1: Integer;
    OutInt2: Integer;
    OutInt3: Integer;
    OutInt4: Integer;
    OutInt5: Integer;
    Msg: String;
  end;

var
  dtmMain: TdtmMain;

implementation

uses cbUtilis, dtmSOU, Encryption_TLB, xmlU, frmMainU ;

{$R *.dfm}

{ TdtmMain }

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
  DateSeparator := '/';
  TimeSeparator := ':';
  LongDateFormat := 'yyyy/mm/dd';
  ShortDateFormat := 'yyyy/mm/dd';
  ShortTimeFormat :='hh:nn:ss';
  LongTimeFormat :='hh:nn:ss';
  TimeAMString := '';
  TimePMString := '';
  FSMSDbList := TList.Create;
  FPrizeDbList := TList.Create;
  FCOMDbList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
  ReleaseSoInfo;
  ReleasePrizeInfo;
  ReleaseComInfo;
//  FSMSDbList.Free;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.ConnectToINV(aObj: PSMSDb): Boolean;
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
   sG_DbSID:= aObj.OracleSid;
   sG_DbUserID:= aObj.UserId;
   sG_DbPassword:= aObj.Password;
   sG_ReportPath:= aObj.ReportPasth;
   Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '【發票】資料區連結失敗, 訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.ConnectToSO(aObj: PSMSDb): Boolean;
begin

  Result := False;
  try
   if not Assigned( aObj ) then
     raise Exception.Create( '對應不到系統台資料區。' );
   dtmSO.SOConnection.Connected := False;
   dtmSO.SOConnection.ConnectionString := Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [aObj.MisDbPassword, aObj.MisDbOwner, aObj.MisDbSid] );
   dtmSO.SOConnection.Connected := True;
   FMisDbOwner := aObj.MisDbOwner;
   FMisDbCompCode := aObj.MisDbCompCode;
   FMisDbCompName := aObj.MisDbCompName;
   FMisDbLink := aObj.MisDbLink;
   Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '【客服系統】資料區連結失敗, 訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.CheckPasswd(const aUserID, aPasswd: String;
  var aGroupId, aUserName: String): Boolean;
var
  aPassObj: _Password;
  aComparePass: String;
begin
  aPassObj := CoPassword.Create;
  adoComm.Close;
  adoComm.SQL.Text :=
     ' SELECT GROUPID, USERNAME, PASSWORD FROM INV004 ' +
     '  WHERE IDENTIFYID1 = ' + QuotedStr( IDENTIFYID1 ) +
     '    AND IDENTIFYID2 = ' + IDENTIFYID2 +
     '    AND USERID =      ' + QuotedStr( aUserID );
  adoComm.Open;
  try
    Result := not adoComm.IsEmpty;
    if Result then
    begin
      aComparePass := aPassObj.Decrypt(
        adoComm.FieldByName('PASSWORD').AsString, PASSKEY );
      Result := ( aComparePass = aPasswd );
    end;
    if Result then
    begin
      aGroupId := adoComm.FieldByName( 'GROUPID' ).AsString;
      aUserName := adoComm.FieldByName( 'USERNAME' ).AsString;
    end;
  finally
    adoComm.Close;
  end;
  aPassObj := nil;
end;


{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompID: String;
begin
  Result := sG_CompID;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompIDEx: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  { 取所有建立在 INV001 的公司別代碼 }
  for aIndex := 0 to frmMain.CompanyList.Count - 1 do
  begin
    Result := ( Result + PCompany(
      frmMain.CompanyList[aIndex] ).CompanyId ) + ',';
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompName: String;
begin
  Result := sG_CompName;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompNameEx: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  { 取所有建立在 INV001 的公司別代碼 }
  for aIndex := 0 to frmMain.CompanyList.Count - 1 do
  begin
    Result := ( Result + PCompany(
      frmMain.CompanyList[aIndex] ).CompanyName ) + ',';
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompLongName: String;
begin
  Result := sG_CompLongName;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompIdAndName: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  for aIndex := 0 to frmMain.CompanyList.Count - 1 do
  begin
    Result := ( Result + PCompany( frmMain.CompanyList[aIndex] ).CompanyId + '-' +
      PCompany( frmMain.CompanyList[aIndex] ).CompanyName ) + ',';
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getCompTel: String;
begin
  Result := sG_CompTel;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getDbSID: String;
begin
  Result := sG_DbSID;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getLoginUser: String;
begin
  Result := sG_LoginUser;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getLoginUserName: String;
begin
  Result := sG_LoginUserName;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getMisDbCompCode: String;
begin
  Result := FMisDbCompCode;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getMisDbOwner: String;
begin
    Result := UpperCase( FMisDbOwner );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getDbPassword: String;
begin
    Result := sG_DbPassword;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getReportPath: String;
begin
  Result := sG_ReportPath;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getDbUserID: String;
begin
    Result := sG_DbUserID;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.getGroupId: String;
begin
    Result := sG_GroupId;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.LoadCompetence(aGroupId : String) : String;
var
  aXmlText, aId, aQuery, aInsert, aDelete, aUpdate, aExecute, aVerify: String;
  aNodeList : IDOMNodeList;
  aIndex,jj : Integer;
begin
  Result := EmptyStr;
  adoCompetence.Close;
  adoCompetence.SQL.Text :=
     ' SELECT COMPETENCE FROM INV013 ' +
     '  WHERE IDENTIFYID1 = ' + QuotedStr( IDENTIFYID1 ) +
     '    AND IDENTIFYID2 = ' + QuotedStr( IDENTIFYID2 ) +
     '    AND GROUPID = ' + QuotedStr( aGroupId );
  adoCompetence.Open;
  aXmlText := adoCompetence.FieldByName( 'Competence' ).AsString;
  adoCompetence.Close;
  if ( aXmlText <> EmptyStr ) then
  begin
    if not cdsCompetence.Active then cdsCompetence.CreateDataSet;
    cdsCompetence.EmptyDataSet;
    getXMLDocument.Active := False;
    getXMLDocument.XML.Text := aXmlText;
    try
      getXMLDocument.Active := True;
      aNodeList := getXMLDocument.DOMDocument.getElementsByTagName( 'GROUP' );
      for aIndex := 0 to aNodeList.length - 1 do
      begin
        for jj:=0 to aNodeList.item[aIndex].childNodes.length - 1 do
        begin
          aId := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'ID');
          aQuery := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'Q');
          aInsert := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'I');
          aDelete := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'D');
          aUpdate := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'U');
          aExecute := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'E');
          aVerify := TMyXML.getAttributeValue(aNodeList.item[aIndex].childNodes.item[jj],'V');
          cdsCompetence.Append;
          cdsCompetence.FieldByName( 'ID' ).AsString := aId;
          cdsCompetence.FieldByName( 'Query' ).AsString := aQuery;
          cdsCompetence.FieldByName( 'Insert' ).AsString := aInsert;
          cdsCompetence.FieldByName( 'Delete' ).AsString := aDelete;
          cdsCompetence.FieldByName( 'Update' ).AsString := aUpdate;
          cdsCompetence.FieldByName( 'Execute' ).AsString := aExecute;
          //#5760 增加一個審核的權限 By Kin 2010/09/08
          if aVerify = EmptyStr then
            aVerify := 'N';
          cdsCompetence.FieldByName( 'Verify' ).AsString := aVerify;
          cdsCompetence.Post;
        end;
      end;
    finally
      getXMLDocument.Active := False;
    end;
    Result := aXmlText;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.SaveToFile(aContent: OleVariant; aFileName: String);
var
  aList : TStringList;
begin
  aList := TStringList.Create;;
  try
    aList.Add( aContent );
    aList.SaveToFile( aFileName );
  finally
    aList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.GetCompetence(aId: String; var aQuery, aInsert, aDelete,
  aUpdate, aExecute,aVerify: String);
begin
  if cdsCompetence.Locate('ID', aID,[] ) then
  begin
    aQuery := cdsCompetence.FieldByName( 'Query' ).AsString;
    aInsert := cdsCompetence.FieldByName( 'Insert' ).AsString;
    aDelete := cdsCompetence.FieldByName( 'Delete' ).AsString;
    aUpdate := cdsCompetence.FieldByName( 'Update' ).AsString;
    aExecute := cdsCompetence.FieldByName( 'Execute' ).AsString;
    aVerify := cdsCompetence.FieldByName( 'Verify' ).AsString;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.SetWaitingCursor;
begin
  Screen.Cursor := crSQLWait;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.SetDefaultCursor;
begin
  Screen.Cursor := crDefault;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.checkBusinessID(sI_BusinessID: String): boolean;
var
    bL_Result, bL_IsSpecial : boolean ;
    nL_CurCheckNum, nL_CurNum, nL_Time : Integer;
    L_Time : array[0..BUSINESS_ID_LENGTH-1] of Integer;
    i,j, nL_Tmp, nL_Total : Integer;

begin
    bL_IsSpecial := false;
    //down,若統編不是數字,則return false
    try
      StrToInt64(sI_BusinessID);
    except
      bL_Result := false;
      Result := bL_Result;
      exit;
    end;
    //up,若統編不是數字,則return false

    //down,若統編字數不足,則return false
    if (length(sI_BusinessID)<>BUSINESS_ID_LENGTH) then
    begin
      bL_Result := false;
      Result := bL_Result;
      exit;
    end;
    //up,若統編字數不足,則return false

    //down,統編之第七位數字是否為7
    if (copy(sI_BusinessID,7,1)='7') then
      bL_IsSpecial := true;
    //up,統編之第七位數字是否為7

    for i:=0 to BUSINESS_ID_LENGTH-1 do
    begin
      nL_CurNum := StrToInt(Copy(sI_BusinessID,i+1,1));//取出統編數字
      nL_CurCheckNum := StrToInt(Copy(BUSINESS_ID_CHECK_NUM,i+1,1)); //取出邏輯乘數
      nL_Tmp:=0;
      nL_Time := nL_CurNum*nL_CurCheckNum;//統編數字*邏輯乘數
      //down,算出乘積之和
      for j:=0 to length(IntToStr(nL_Time))-1 do
      begin
        if (j=6) then
        begin
          if (bL_IsSpecial) then
            nL_Tmp := nL_Tmp + 0
          else
            nL_Tmp := nL_Tmp + StrToInt(Copy(IntToStr(nL_Time), j+1,1));
        end
        else
          nL_Tmp := nL_Tmp + StrToInt(Copy(IntToStr(nL_Time), j+1,1));
//          nL_Tmp = nL_Tmp + Integer.parseInt(Integer.toString(nL_Time).substring(j,j+1));
      //up,算出乘積之和
        L_Time[i] := nL_Tmp;
      end;
    end;
    //down,把乘積之和加起來,存入nL_Total
    nL_Total := 0;
    for i:=0 to High(L_Time) do
      nL_Total := nL_Total + L_Time[i];
    //up,把乘積之和加起來,存入nL_Total


    //down,檢查nL_Total是否能被10整除
    if (nL_Total mod 10=0) then
        bL_Result := true
    else
    begin
      bL_Result := false;
      //down,若統編之倒數第二位數字是7,則檢查(nL_Total+1)是否能被10整除
      if (bL_IsSpecial) then
        if ((nL_Total+1) mod 10=0) then
          bL_Result := true;

      //up,若統編之倒數第二位數字是7,則檢查(nL_Total+1)是否能被10整除
    end;
    //up,檢查nL_Total是否能被10整除

    result := bL_Result;

end;


{ ---------------------------------------------------------------------------- }

function TdtmMain.getServiceTypeStr: String;
begin
  Result := sG_ServiceTypeStr;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetServiceTypeStrEx: String;
var
  aIndex: Integer;
  aSplit: array of string;
  aSource, aString: String;
begin
  aSource := getServiceTypeStr;
  if Pos( ',', aSource ) = 0 then
  begin
    SetLength( aSplit, Length( aSplit ) + 1 );
    aSplit[ Length( aSplit ) - 1 ] := aSource;
  end
  else begin
    aString := '';
    for aIndex := 1 to Length( aSource ) do
    begin
      if ( aSource[aIndex] <> ',' ) then
      begin
        aString := ( aString + aSource[aIndex] );
      end
      else begin
        SetLength( aSplit, Length( aSplit ) + 1 );
        aSplit[ Length( aSplit ) - 1 ] := aString;
        aString := '';
      end;
    end;
    if ( aString <> '' ) then
    begin
      SetLength( aSplit, Length( aSplit ) + 1 );
      aSplit[ Length( aSplit ) - 1 ] := aString;
    end;
  end;
  for aIndex := Low( aSplit ) to High( aSplit ) do
    Result := ( Result + '''' + aSplit[aIndex] + '''' + ',' );
  if ( Length( Result ) > 1 ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.GetInvoiceCompany(aList: TList);
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
  //#5358 增加ShowFaci 判斷是否顯示設備 By Kin 2009/12/14
  aList.Clear;
  adoComm.Close;
  adoComm.SQL.Text :=
    ' SELECT COMPSNAME, COMPID, RPTNAME1, RPTNAME2,   ' +
    '        RPTNAME3, RPTNAME4, RPTNAME5, RPTNAME6,  ' +
    '        SERVICETYPESTR, LINKTOMIS, CONNSEQ,      ' +
    '        AUTOCREATENUM, COMPNAME, FACICOMBINE,    ' +
    '        TEL, NVL( SHOWFACI, 0 ) SHOWFACI,        ' +
    '        NVL( STARTEINVOICE, 0 ) STARTEINVOICE,   ' +
    '        NVL( STARTEMAIL, 0 ) STARTEMAIL,         ' +
    '        NVL( STARTMESSAGE, 0 ) STARTMESSAGE,     ' +
    '        NVL( STARTTVMAIL, 0 ) STARTTVMAIL,       ' +
    '        NVL( STARTCMSEND, 0 ) STARTCMSEND,       ' +
    '        NVL( AutoNotify, 0 ) AutoNotify          ' +
    '   FROM INV001         ' +
    '  WHERE IDENTIFYID1 =  ' + QuotedStr( IDENTIFYID1 ) +
    '    AND IDENTIFYID2 =  ' + QuotedStr( IDENTIFYID2 ) +
    '  ORDER BY TO_NUMBER( COMPID ) ';
  adoComm.Open;
  adoComm.First;
  //#5358 增加ShowFaci 判斷在開立發票時是否要將設備序號寫入Memo1 By Kin 2009/12/14
  while not adoComm.Eof do
  begin
    New( aObj );
    aObj.CompanyId := dtmMain.adoComm.FieldByName( 'COMPID' ).AsString;
    aObj.CompanyName := dtmMain.adoComm.FieldByName( 'COMPSNAME' ).AsString;
    aObj.Tel := dtmMain.adoComm.FieldByName( 'TEL' ).AsString;
    aObj.CompanyLongName := dtmMain.adoComm.FieldByName( 'COMPNAME' ).AsString;
    aObj.ServiceType := dtmMain.adoComm.FieldByName( 'SERVICETYPESTR' ).AsString;
    aObj.LinkToMIS := ( dtmMain.adoComm.FieldByName( 'LINKTOMIS' ).AsString = 'Y' );
    aObj.ConnSeq := dtmMain.adoComm.FieldByName( 'CONNSEQ' ).AsInteger;
    aObj.AutoCreateNum := dtmMain.adoComm.FieldByName( 'AUTOCREATENUM' ).AsInteger;
    aObj.ReportName1 := dtmMain.adoComm.FieldByName( 'RPTNAME1' ).AsString;
    aObj.ReportName2 := dtmMain.adoComm.FieldByName( 'RPTNAME2' ).AsString;
    aObj.ReportName3 := dtmMain.adoComm.FieldByName( 'RPTNAME3' ).AsString;
    aObj.ReportName4 := dtmMain.adoComm.FieldByName( 'RPTNAME4' ).AsString;
    aObj.ReportName5 := dtmMain.adoComm.FieldByName( 'RPTNAME5' ).AsString;
    aObj.ReportName6 := dtmMain.adoComm.FieldByName( 'RPTNAME6' ).AsString;
    aObj.FaciCombine := dtmMain.adoComm.FieldByName( 'FACICOMBINE' ).AsInteger;
    aObj.ShowFaci := dtmMain.adoComm.FieldByName( 'SHOWFACI' ).AsInteger;
    aObj.IfPrintTitle := False;
    aObj.IfPrintAddr := False;
    aObj.IfPrintCheck := True;
    aObj.StarEInvoice := ( dtmMain.adoComm.FieldByName( 'STARTEINVOICE' ).AsInteger = 1 );
    aObj.StarEmail := ( dtmMain.adoComm.FieldByName( 'STARTEMAIL' ).AsInteger = 1 );
    aObj.StarMessage := ( dtmMain.adoComm.FieldByName( 'STARTMESSAGE' ).AsInteger = 1 );
    aObj.StarTVMail := ( dtmMain.adoComm.FieldByName( 'STARTTVMAIL' ).AsInteger = 1 );
    aObj.StarCMSend := ( dtmMain.adoComm.FieldByName( 'STARTCMSEND' ).AsInteger = 1 );
    aObj.StartAutoNotify := ( dtmMain.adoComm.FieldByName( 'AutoNotify' ).AsInteger = 1 );
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

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetConnSeq: Integer;
begin
  Result := sG_ConnSeq;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetAutoCreateNum: Integer;
begin
  Result := sG_AutoCreateNum;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetLinkToMIS: Boolean;
begin
  Result := sG_LinkToMIS;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetMisDbLink: String;
begin
  if FMisDbLink <> EmptyStr then
    Result := '@' + FMisDbLink;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetIfPrintAddr: Boolean;
begin
  Result := sG_IfPrintAddr;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetIfPrintCheck: Boolean;
begin
  Result := sG_IfPrintCheck;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetIfPrintTitle: Boolean;
begin
  Result := sG_IfPrintTitle;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetFaciCombine: Integer;
begin
  Result := sG_FaciCombine;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetExpAddrType: Integer;
begin
  if not ( sG_ExpAddrType in [1..2] ) then
    sG_ExpAddrType := 1;
  Result := sG_ExpAddrType
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetShowFaci: Integer;
begin
  Result := sG_ShowFaci;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetStarEInvoice: Boolean;
begin
  Result := sG_StarEInvoice;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetStarEmail: Boolean;
begin
  Result := sG_StarEmail;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetStarMessage: Boolean;
begin
  Result := sG_StarMessage;
end;

{ ---------------------------------------------------------------------------- }

function ParseStrings(Strs:string; Sep:Char) : TStringList ;
var
  i : Integer ;
  slst : TStringList ;
  sL_Tmp : String;
begin
  if Strs = '' then
  begin
    Result := nil ;
    Exit ;
  end ;

  slst := TStringList.Create ;
  sL_Tmp := '';

  for i:= 1 to Length(Strs) do   // 1,1528,4587721 ,,范宇宏
  begin
    if Strs[i] = Sep then
    begin
      slst.Add(sL_Tmp);
      sL_Tmp := '';
    end
    else
      sL_Tmp := sL_Tmp + Strs[i];
  end;
  slst.Add(sL_Tmp); //存入最後一筆
  Result := slst ;
end ;

{ ---------------------------------------------------------------------------- }

function getInvoiceYearMonth(sI_Date : String ) :String ;
//ex:傳入2000/01/01 傳回 200001的字串
var
  sL_StrOfYearMonth,sL_Temp  : String ;
begin
  sL_StrOfYearMonth := CutAllChar(sI_Date,'/') ;
  sL_Temp := copy(sL_StrOfYearMonth,5,2) ;
  if StrToInt(sL_Temp) mod 2 = 0 then sL_Temp := IntToStr(StrToInt(sL_Temp)-1) ;
  if Length(sL_Temp) < 2 then sL_Temp := '0'+ sL_Temp ;
  sL_StrOfYearMonth := copy(sL_StrOfYearMonth,0,4) + sL_Temp ;
  Result := sL_StrOfYearMonth ;
end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay(sI_Date: String): String;
//ex:傳入 2001/8/21 傳回20010821
var
  slst : TStringList ;
  sL_Day,sL_Month :String ;
begin
  slst := ParseStrings(sI_Date,'/');
  try
    if Length(slst.Strings[1]) < 2 then
      sL_Month := '0'+ slst.Strings[1]
    else
      sL_Month := slst.Strings[1];

    if Length(slst.Strings[2]) < 2 then
      sL_Day := '0'+ slst.Strings[2]
    else
      sL_Day := slst.Strings[2];

    Result := slst.Strings[0] + sL_Month + sL_Day ;
  finally
    slst.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }
function getYearMonthDay2(sI_Date: String): String;
//ex:傳入 20010821 傳回 2001/08/21
begin
  if sI_Date ='' then
    Result := ''
  else
  Result := copy(sI_Date,1,4)+'/'+ copy(sI_Date,5,2) +'/'+copy(sI_Date,7,2);
end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay3(sI_Date: String ; nI_Index: Integer): String;
//傳入 200109 傳回 2001/09/01 及 2001/11/01
var
  sL_Month,sL_Year,sL_EndMonth : String ;
  sL_StrResult,sL_EndResult : String ;
begin

  sL_Year := copy(sI_Date,1,4);
  sL_Month := copy(sI_Date,5,2);

  if StrToInt(sL_Month) mod 2 = 0 then
  begin
    if StrToInt(sL_Month)=12 then
    begin
      sL_Year := IntToStr(StrToInt(sL_Year)+1);
      sL_EndMonth := '01';
      sL_Month := IntToStr(StrToInt(sL_Month)-1)
    end
    else
    begin
      sL_EndMonth := IntToStr(StrToInt(sL_Month)+1);
      sL_Month := IntToStr(StrToInt(sL_Month)-1)
    end
  end
  else
  begin
    if StrToInt(sL_Month)=11 then
    begin
      sL_Year := IntToStr(StrToInt(sL_Year)+1);
      sL_EndMonth := '01';
    end
    else
      sL_EndMonth := IntToStr(StrToInt(sL_Month)+2)
  end;

  if Length(sL_Month) < 2 then
    sL_Month := '0'+ sL_Month ;
  if Length(sL_EndMonth) < 2 then
    sL_EndMonth := '0'+ sL_EndMonth ;


  sL_StrResult := sL_Year +'/'+sL_Month+'/01';
  sL_EndResult := sL_Year +'/'+sL_EndMonth+'/01';

  if nI_Index =1 then
    Result := sL_StrResult
  else
    Result := sL_EndResult;

end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay4(sI_Date: String): String;
//輸入 200110 傳回200109 或輸入 200109 傳回200109
var
  sL_Year,sL_Month : String;
begin
  sL_Year := copy(sI_Date,1,4);
  sL_Month := copy(sI_Date,5,2);

  if StrToInt(sL_Month) mod 2 = 0 then
    sL_Month := IntToStr(StrToInt(sL_Month)-1) ;
  if Length(sL_Month) < 2 then
    sL_Month := '0'+ sL_Month ;

  Result :=sL_Year+sL_Month;

end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay5(sI_Date: String; nI_Index: Integer): String;
//傳入 200109 傳回 2001/09/01 及 2001/10/01
var
  sL_Month,sL_Year,sL_EndMonth : String ;
  sL_StrResult,sL_EndResult : String ;
  bL_isExceed : Boolean;
begin
  bL_isExceed := false ;
  sL_Year := copy(sI_Date,1,4);
  sL_Month := copy(sI_Date,5,2);

  if StrToInt(sL_Month)=12 then
  begin
    bL_isExceed := true ;
    sL_EndMonth := '01';
  end
  else
    sL_EndMonth := IntToStr(StrToInt(sL_Month)+1);

  if Length(sL_Month) < 2 then
    sL_Month := '0'+ sL_Month ;
  if Length(sL_EndMonth) < 2 then
    sL_EndMonth := '0'+ sL_EndMonth ;

  sL_StrResult := sL_Year +'/'+sL_Month+'/01';
  if bL_isExceed then
    sL_EndResult := IntToStr(StrToInt(sL_Year)+1) +'/'+sL_EndMonth+'/01'
  else
    sL_EndResult := sL_Year +'/'+sL_EndMonth+'/01';

  if nI_Index =1 then
    Result := sL_StrResult
  else
    Result := sL_EndResult;

end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay6(dI_Date: TDate): String;
var
  wL_Year, wL_Month, wL_Day : word;
begin
  DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
  result := Format('%.4d%.2d%.2d', [wL_Year, wL_Month, wL_Day]);
end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay7(dI_Date: TDate;Di_Integer:Integer): String;
var
  wL_Year, wL_Month, wL_Day : word;
begin
  DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
  if wL_Year > 1000 then //表為西元年  要回傳中華民國曆
    wL_Year := wL_Year - 1911 ;

  if Di_Integer = 2 then
  begin
    if wL_Year<100 then
      Result := '0'+Format('%.2d%.2d',[wL_Year,wL_Month ])
    else
      Result := Format('%3d%.2d%',[wL_Year,wL_Month ]);
    Exit;
  end;

  if Di_Integer = 0 then
  begin
    if wL_Year<100 then
      result :='0'+ Format('%.2d%.2d%.2d', [wL_Year, wL_Month, wL_Day]) //4348 因應民國100年,所以增加0 By Kin 2009/02/04
    else
      result := Format('%.2d%.2d%.2d', [wL_Year, wL_Month, wL_Day]);
  end
  else
    result := Format('%.2d%.2d', [wL_Year, wL_Month]);
end;

{ ---------------------------------------------------------------------------- }

function getYearMonthDay8(sI_Date: String ; nI_Index: Integer): String;
//傳入 200109 傳回 2001/09/01 及 2001/09/30 及 2001/10/31
var
  sL_Month,sL_Year,sL_EndMonth : String ;
  sL_StrResult,sL_EndResult : String ;
  sL_StrResult2, sL_EndResult2 : String;
  sL_EndDay1, sL_EndDay2 : String;
begin

  sL_Year := copy(sI_Date,1,4);
  sL_Month := copy(sI_Date,5,2);
  sL_EndMonth := IntToStr(StrToInt(sL_Month)+1);

  if Length(sL_Month) < 2 then
    sL_Month := '0'+ sL_Month ;
  if Length(sL_EndMonth) < 2 then
    sL_EndMonth := '0'+ sL_EndMonth ;

  sL_EndDay1 := _monthLastDay(SL_Year+sL_Month,1);
  sL_EndDay2 := _monthLastDay(SL_Year+sL_endMonth,1);

  if Length(sL_Endday1) < 2 then
    sL_EndDay1 := '0'+ sL_EndDay1 ;
  if Length(sL_EndDay2) < 2 then
    sL_EndDay2 := '0'+ sL_EndDay2 ;

  sL_StrResult := sL_Year +'/'+sL_Month+'/01';
  sL_EndResult := sL_Year +'/'+sL_Month+'/'+sL_EndDay1;
  sL_StrResult2 := sL_Year +'/'+sL_EndMonth+'/01';
  sL_EndResult2:= sL_Year +'/'+sL_EndMonth+'/'+sL_EndDay2;

  if nI_Index =1 then
    Result := sL_StrResult
  else if nI_Index =2 then
    Result := sL_EndResult
  else if nI_Index =3 then
    Result := sL_StrResult2
  else
    Result := sL_EndResult2;
end;

{ ---------------------------------------------------------------------------- }

function adjString(nI_FieldLength: Integer; sI_Value: String;
  bI_Right,bI_isInteger: boolean): String;
//填滿零或填滿空白，nI_FieldLength表總長度
//bI_Right為TRUE時將空白或零填在左邊
//bI_isInteger為TRUE時填入零，否則填入空白
var
  nL_ValueLen : Integer ;
  sL_Space,sL_Result : String ;
begin
  sI_Value := Trim(sI_Value);
  nL_ValueLen := Length(sI_Value);
  if bI_isInteger then
    sL_Space := AssignSpace(nI_FieldLength-nL_ValueLen,true)
  else
    sL_Space := AssignSpace(nI_FieldLength-nL_ValueLen,false);
  if bI_Right then
    sL_Result := sL_Space+ sI_Value
  else
    sL_Result := sI_Value+sL_Space;

  Result := sL_Result;

end;

{ ---------------------------------------------------------------------------- }

function adjString2(nI_FieldLength: Integer; sI_Value: Double): String;
//填滿零或填滿空白
//nI_FieldLength表總長度
//sI_Value 為卻轉換之字串
var
  sL_Value : String;
  nL_ValueLen : Integer ;
  i : Integer;
  sL_Result : String;
begin
  sL_Value := formatfloat('###,###,###,###',sI_Value);
  nL_ValueLen := nI_FieldLength - length(sL_Value);
  if nL_ValueLen > 0 then
    for I := 1 to nL_ValueLen do
      sL_Result:=sL_Result+'0';

  Result := sL_Result+ sL_Value;
end;

{ ---------------------------------------------------------------------------- }

function getOracleSQLDateTimeStr(dI_Date: TDateTime): String;
var
    wL_Year, wL_Month, wL_Day : Word;
    wL_Hour, wL_Min, wL_Sec, wL_MSec : Word;
    sL_Result, sL_DateTimeStr : String;
begin
    if dI_Date=0 then
      sL_Result := 'null'
    else
    begin
      DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
      DecodeTime(dI_Date, wL_Hour, wL_Min, wL_Sec, wL_MSec);
      sL_DateTimeStr := format('%.4d/%.2d/%.2d %.2d:%.2d:%.2d', [wL_Year, wL_Month, wL_Day,wL_Hour, wL_Min, wL_Sec ]);
      sL_Result := 'to_date(' + '''' + sL_DateTimeStr + '''' + ',' + '''' + 'YYYY/MM/DD HH24:MI:SS' + '''' + ')';
    end;
    result := sL_Result;
end;

{ ---------------------------------------------------------------------------- }

function CutAllChar(S:string;C:char):string;
  //[功用]將字串 S 中某單一字元 C 全數惕除
  //[範例]例 T:=_CutAllChar('A-B-C','-'); 即 T='ABC'
var I:integer;
  SS,R:string;
begin
  R:='';
  if Length(S)=0 then
    begin
      Result:='';
      exit;
    end;
  if Length(S)=1 then
    begin
      if S[1]=C then Result:=''
      else Result:=S;
        exit;
    end;
  for I:=1 to Length(S) do
    begin
      SS:=Copy(S,I,1);
      if SS[1]=C then Continue
      else R:=R+SS;
    end;
  Result:=R;
end;

{ ---------------------------------------------------------------------------- }

function AssignSpace(nI_Count: integer;bL_isInteger:boolean): String;
var
  I:integer;
  sL_Result :string;
begin
  sL_Result:='';
  if nI_Count<>0 then
  begin
    if bL_isInteger then
    begin
      for I := 1 to nI_Count do
        sL_Result:=sL_Result+'0';
    end
    else
    begin
      for I := 1 to nI_Count do
        sL_Result:=sL_Result+' ';
    end
  end;
  Result:=sL_Result;

end;

{ ---------------------------------------------------------------------------- }

function _CutLeft(S:string;L:integer):string; {減掉字串中左邊若干個字元}
var R:string;				      {例 _CUTRIGHT('TEST',1) 即傳回'EST'}
begin
  {R:=S;}
  R:='';
  if (S<>'') and (L<=Length(S)) and (L>0) then
     begin
      R:=_RIGHT(S,(Length(S)-L));
     end;

  if L=0 then R:=S;

  Result:=R;
end;

{ ---------------------------------------------------------------------------- }

function _Right(S:string;L:integer):string; {同 Clipper RIGHT() }
var R:string;
begin
  R:='';

  if (S<>'') and (L<Length(S)) and (L>0) then
     begin
      R:=COPY(S,(Length(S)-L+1),L);
     end;

  if L>=Length(S) then R:=S;

  Result:=R;
end;

{ ---------------------------------------------------------------------------- }

function _Left(S:string;L:integer):string; {同 Clipper Left() }
var R:string;
begin
  R:='';

  if (S<>'') and (L<Length(S)) and (L>0) then
     begin
      R:=COPY(S,1,L);
     end;

  if L>=Length(S) then R:=S;

  Result:=R;
end;

{ ---------------------------------------------------------------------------- }

{求西元年月之該月最後一天日期字串 ; onlyday=1 表只傳回該月天數字串}
function _monthLastDay(yyyymm:string ; onlyday:integer):string;
var y:string;
    m:integer;
    d:string;
begin

  if Length(yyyymm)<>6 then
     begin
       Result:='';
       exit;
     end;

  y:=_Left(yyyymm,4);
  m:=_strtoint(_right(yyyymm,2));

  d:='';
  case m of
     1,3,5,7,8,10,12 : d:='31';
     4,6,9,11	       : d:='30';
     2		           : if _isLeapYear(y) then d:='29' else d:='28';
  end;

  if d='' then
     begin
       Result:='';
       exit;
     end;

 if onlyday=0 then
    Result:=yyyymm+d
 else
    Result:=d;
end;

{ ---------------------------------------------------------------------------- }

function _Int(R:REAL):LONGINT;
{與 DELPHI 之 INT() 不同,其傳回 LONGINT , 而非 REAL ( 如123.0 ) }
var RR:string;
begin
  RR:=FLOATTOSTR(INT(R));
  Result:=strtoint(RR);
end;

{ ---------------------------------------------------------------------------- }

function _Value(S:string):Extended;
var I:LONGINT;     {88.01.20 修改}
    CODE:integer;
    SS:string;
begin
S:=_AllTrim(S);  {將字串前後空白去掉}

if _SubStrNum(S,'.')=0 then
   begin
     VAL(S,I,CODE);
     Result:=I;
     exit;
   end
else
   begin
     if _SubStrNum(S,'.')>1 then
	begin
	Result:=0;
	exit;
	end;

     SS:=S;
     Delete(SS,(POS('.',SS)),1);

     Val(SS,I,CODE);

     if I=0 then
	begin
	Result:=0;
	exit;
	end;

    Result:=STRTOFLOAT(S);

   end;
end;

{ ---------------------------------------------------------------------------- }

function _StrToInt(S:string):LONGINT;
{與 DELPHI 之 strtoint() 不同,格式不對其傳回 0 , 而不是錯誤訊息}
begin
  Result:=_Int(_Value(S));
end;

{ ---------------------------------------------------------------------------- }

function _isLeapYear(yyyy:string):Boolean; {求西元年字串是否為閏年}
var y:integer;
    r:Boolean;
begin

  y:=_strtoint(yyyy);

  if y=2000 then
     begin
       Result:=True;
       exit;
     end;

  r:=False;

  if y>2000 then
     begin
       if ((y-2000) mod 4)=0 then r:=True;
     end
  else
     begin
       if ((2000-y) mod 4)=0 then r:=True;
     end;

  Result:=r;
end;

{ ---------------------------------------------------------------------------- }

function _SubStrNum(S,SUBS:string):integer; {取字串中出現某一子字串之次數 }
var L,N,I : integer;
begin

  if (Length(S)=0) or (Length(SUBS)=0) then
      begin
	Result:=0;
	exit;
      end;

  N:=0;
  L:=Length(SUBS);
  {
  for I:=1 to (Length(S)-L+1) do
      begin
      if copy(S,I,L)=SUBS then N:=N+1;
      end;}

  i:=1;  {1999/06/21 修正,如此才不會將_SubStrNum('EEE','EE') 當成2,應為1才對(算過的不算) }
  while i<=(Length(S)-L+1) do
    begin
      if copy(S,I,L)=SUBS then
         begin
           N:=N+1;
           i:=i+L;
         end
      else
        i:=i+1;
    end;

  Result:=N;

end;

{ ---------------------------------------------------------------------------- }

function _AllTrim(cText: string): string; { 將字串中左右所有的空白字元刪除 }
begin
  Result := _LTrim(_RTrim(cTEXT));
end;

{ ---------------------------------------------------------------------------- }

function _LTrim(S:string):string;  { 去字串 左空白}
var I:integer;
    R:string;
begin
  R:='';
  if (S<>'') and (S<>' ') then
     begin
       for I :=1 to Length(S) do
	   if copy(S,I,1)<>' ' then break;

     R:=COPY(S,I,(Length(S)-I+1));
     if R=' ' then R:='';

     end;
  Result:=R;
end;

{ ---------------------------------------------------------------------------- }

function _RTrim(S:string):string; {同 Clipper RTRIM() 去掉字串右邊空白}
var I:integer;
    R:string;
begin
  R:='';
  if (S<>'') and (S<>' ') then
     begin
       for I :=Length(S) downto 1 do
	   if copy(S,I,1)<>' ' then break;
     R:=COPY(S,1,I);
     if R=' ' then R:='';
     end;
  Result:=R;
end;

{ ---------------------------------------------------------------------------- }

function GetGridStoragePath(const aGridName: String): String;
begin
  if aGridName = EmptyStr then
    raise Exception.Create( '無法儲存資料方格設定值。'#13#10'原因:儲存名稱為空值。' );
  Result := IncludeTrailingPathDelimiter( IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + GRID_FOLDER ) + aGridName + '.ini';
end;

{ ---------------------------------------------------------------------------- }

function TranslateError(const aErrMsg: String): Boolean;
var
  aPos1, aPos2: Integer;
begin
  Result := False;
  aPos1 := AnsiPos( '-', aErrMsg );
  aPos2 := AnsiPos( ':', aErrMsg );
  if aPos1 < 0 then Exit;
  OracleErros.ErrorType := Copy( aErrMsg, 1, aPos1 - 1 );
  if aPos2 < 0 then Exit;
  OracleErros.ErrorCode := StrToInt( Copy( aErrMsg, aPos1 + 1, aPos2 - ( aPos1 + 1 ) ) );
  OracleErros.ErrorMsg := Copy( aErrMsg, aPos2 + 1, Length( aErrMsg ) - aPos2 );
end;

{ ---------------------------------------------------------------------------- }

function MappingChineseNumber(const aNumber: Integer): String;
const
  aMapping: array [0..10] of String =
    ( '零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十' );
var
  aNumText: String;
begin
  aNumText := IntToStr( aNumber );
  if Length( aNumText ) <= 1 then
  begin
    Result := aMapping[aNumber];
  end else
  if Length( aNumText ) = 2 then
  begin
    if aNumber = 10 then
    begin
      Result := aMapping[10];
    end else
    if ( Copy( aNumText, 2, 1 ) = '0' ) then
    begin
      Result := aMapping[StrToInt( Copy( aNumText, 1, 1 ) )] + '十';
    end else
    begin
      Result := aMapping[StrToInt( Copy( aNumText, 1, 1 ) )] + '十' +
        aMapping[StrToInt( Copy( aNumText, 2, 1 ) )];
    end;    
  end else
  if Length( aNumText ) = 3 then
  begin
    if ( Copy( aNumText, 2, 2 ) = '00' ) then
    begin
      Result := aMapping[StrToInt( Copy( aNumText, 1, 1 ) )] + '百';
    end else
    if ( Copy( aNumText, 2, 1 ) = '0' ) then
    begin
      Result := aMapping[StrToInt( Copy( aNumText, 1, 1 ) )] + '百' + aMapping[0] +
        aMapping[StrToInt( Copy( aNumText, 3, 1 ) )];
    end else
    if ( Copy( aNumText, 3, 1 ) = '0' ) then
    begin
      Result := aMapping[StrToInt( Copy( aNumText, 1, 1 ) )] + '百' +
        aMapping[StrToInt( Copy( aNumText, 2, 1 ) )] + '十';
    end else
    begin
      Result := aMapping[StrToInt( Copy( aNumText, 1, 1 ) )] + '百' +
        aMapping[StrToInt( Copy( aNumText, 2, 1 ) )] + '十' +
        aMapping[StrToInt( Copy( aNumText, 3, 1 ) )];
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function FacisNoAddMask(aText: String): String;
var
  aHeader, aFooter: String;
begin
  Result := EmptyStr;
  if Trim( aText ) = EmptyStr then Exit;
  { 6碼含以內, 遮罩前面, 留後 2 碼 }
  if Length( aText ) in [1..6] then
  begin
    Result := Copy( aText, Length( aText ) - 1, 2 );
    Result := Lpad( Result, Length( aText ), 'X' );
  end else
  { 超過6碼, 前面取4碼, 後面取2碼, 中間遮罩 }
  begin
    aHeader := Copy( aText, 1, 4 );
    aFooter := Copy( aText, Length( aText ) - 1, 2 );
    Result := aHeader + Lpad( EmptyStr, Length( aText ) - 6, 'X' ) + aFooter;
  end;
end;

{ ---------------------------------------------------------------------------- }

function MapingInvoiceYearMonth(const aDate: TDate): String;
var
  aYear, aMonth, aDay: Word;
begin
  DecodeDate( aDate, aYear, aMonth, aDay );
  if ( aMonth mod 2 ) = 0 then Dec( aMonth );
  Result := Format( '%4d%.2d', [aYear, aMonth] );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMain.GetSoInfo(const aConnSeq: Integer): PSMSDb;
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

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.ReleaseSoInfo;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSMSDbList.Count - 1 do
  begin
    if Assigned( FSMSDbList[aIndex] ) then
    begin
      Dispose( PSMSDb( FSMSDbList[aIndex] ) );
      FSMSDbList[aIndex] := nil;
    end;
  end;
  FSMSDbList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain.GetSMSDb(const aFileName: String);
var
  aIndex: Integer;
  aIniFile: TIniFile;
  aStrList: TStringList;
  aSMSDbId: String;
  aPrizeDbId:String;
  aPrizeObj:PPRIZEDb;
  aStrPrizeList: TStringList;
  aObj: PSMSDb;
  aIniSid, aIniUserId, aIniPassword,aIniCompTitle :String;

begin
  aStrList := TStringList.Create;
  aStrPrizeList := TStringList.Create;
  try
    aIniFile := TIniFile.Create( aFileName );
    try

      aIniFile.ReadSection( 'DATAAREA', aStrList );
      FSMSDbList.Clear;
      FPrizeDbList.Clear;
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
        aObj.EInvDb := aIniFile.ReadString( 'PRIZECOM' ,'SID',EmptyStr);
        aObj.EInvUser := aIniFile.ReadString( 'PRIZECOM','DB_USER', EmptyStr);
        aObj.EInvPassword := aIniFile.ReadString( 'PRIZECOM', 'DB_USER_PASSWORD',EmptyStr);
        if ( aObj.EInvDb <> EmptyStr ) and ( aObj.EInvUser <> EmptyStr )
          and ( aObj.EInvPassword <> EmptyStr ) then
        begin
          aObj.USEPrizeCom := True;
        end else
        begin
          aObj.USEPrizeCom := False;
        end;
        FSMSDbList.Add( aObj );
        {#5668 多加檢核對獎發票的類別 By Kin 2010/07/07}
        aPrizeDbId := 'E' + aStrList.Strings[ aIndex ];
        aIniSid := aIniFile.ReadString( aPrizeDbId, 'SID', EmptyStr );
        aIniUserId :=  aIniFile.ReadString( aPrizeDbId, 'DB_USER', EmptyStr );
        aIniPassword := aIniFile.ReadString( aPrizeDbId, 'DB_USER_PASSWORD', EmptyStr );
        aIniCompTitle := aIniFile.ReadString( aPrizeDbId ,'COMPTITLE', EmptyStr );
        if ( aIniSid <> EmptyStr ) and ( aIniUserId <> EmptyStr )
          and ( aIniPassword <> EmptyStr ) then
        begin
          New( aPrizeObj );
          aPrizeObj.ConnSeq := StrToInt( aSMSDbId );
          aPrizeObj.DBSid := aIniSid;
          aPrizeObj.UserId := aIniUserId;
          aPrizeObj.Password := aIniPassword;
          aPrizeObj.CompTitle := aIniCompTitle;
          FPrizeDbList.Add( aPrizeObj );
        end;
      end;

    finally
      aIniFile.Free;
    end;
  finally
    aStrList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function CreateCxComboBoxItem(aComboBox: TcxComboBox;
  aParam:TComboBoxCreateParam): Integer;
var
  aText: String;
begin
  Result := 0;
  if not Assigned( aComboBox ) then Exit;
  if ( aParam.Sql = EmptyStr ) or ( aParam.KeyField = EmptyStr ) then Exit;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aParam.Sql;
  dtmMain.adoComm.Open;
  if ( dtmMain.adoComm.IsEmpty ) then
  begin
    dtmMain.adoComm.Close;
    Exit;
  end;
  aComboBox.Properties.Items.Clear;
  dtmMain.adoComm.First;
  while not dtmMain.adoComm.Eof do
  begin
    aText := EmptyStr;
    if Assigned( dtmMain.adoComm.FindField( aParam.KeyField ) ) then
      aText := dtmMain.adoComm.FieldByName( aParam.KeyField ).AsString;
    if ( aParam.DescField <> EmptyStr ) then
    begin
      if Assigned( dtmMain.adoComm.FindField( aParam.DescField ) ) then
      begin
        aText := ( aText + Nvl( aParam.DelimiterText, '.' ) +
          dtmMain.adoComm.FieldByName( aParam.DescField ).AsString );
      end;
    end;
    if ( aText <> EmptyStr ) then
    begin
      aComboBox.Properties.Items.Add( aText );
      Inc( Result );
    end;
    dtmMain.adoComm.Next;
  end;
  dtmMain.adoComm.Close;
  if ( aParam.AddAllText ) and ( aComboBox.Properties.Items.Count > 0 ) then
    aComboBox.Properties.Items.Insert( 0, '全部' );
end;

{ ---------------------------------------------------------------------------- }

function GetCxComboBoxItemValue(aComboBox: TcxComboBox;
  var aParam: TComboBoxItemParam): Boolean;
var
  aText: String;
  aPos: Integer;
begin
  Result := False;
  aParam.KeyValue := EmptyStr;
  aParam.DesValue := EmptyStr;
  if not Assigned( aComboBox ) then Exit;
  if ( aComboBox.ItemIndex < 0 ) then Exit;
  aText := aComboBox.Properties.Items[aComboBox.ItemIndex];
  aPos := Pos( Nvl( aParam.DelimiterText, '.' ), aText );
  if ( aPos <= 0 ) then Exit;
  aParam.KeyValue := Copy( aText, 1, aPos - 1 );
  aParam.DesValue := Copy( aText, aPos + 1, Length( aText ) - aPos );
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure SetCxComboBoxItemValue(aComboBox: TcxComboBox;
  aValue: String);
var
  aIndex, aPos: Integer;
  aText: String;
begin
  if not Assigned( aComboBox ) then Exit;
  for aIndex := 0 to aComboBox.Properties.Items.Count - 1 do
  begin
    aText := aComboBox.Properties.Items[aIndex];
    aPos := Pos( '.', aText );
    if ( aPos < 0 ) then Continue;
    aText := Copy( aText, 1, aPos - 1 );
    if SameText( aValue, aText ) then
    begin
      aComboBox.ItemIndex := aIndex;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }
  
procedure TdtmMain.GetSMSDb;
 var aObj: PSMSDb;
begin
  FSMSDbList.Clear;

  New( aObj );
  aObj.ConnSeq := StrToInt( ParamStr(2) );
  aObj.OracleSid := ParamStr(3) ;
  aObj.Description := ParamStr(4);
  aObj.UserId := ParamStr( 5 );
  aObj.Password := ParamStr( 6 );
  aObj.ReportPasth := ParamStr( 7 );
  aObj.MisDbOwner := ParamStr( 8 );
  aObj.MisDbPassword := ParamStr( 9 );
  aObj.MisDbCompCode := ParamStr( 10 );
  aObj.MisDbCompName := ParamStr( 11 );
  aObj.MisDbSid := ParamStr( 12 );
  aObj.MisDbLink := ParamStr( 13 );
  FSMSDbList.Add( aObj );
end;

procedure TdtmMain.ShellApplication(AppName, Caption: String;
  Param: array of String);
  var execinfo : TShellExecuteInfo;
    ApplicationName, ParamString : String;
    i : integer;
begin
  try
    Screen.Cursor := crSQLWait;
    for i := Low(PARAM) to High(Param) do
      ParamString := ' ' + ParamString + ' ' + Param[i];
      FillChar(execinfo,SizeOf(execinfo),0);
      execinfo.cbSize:=sizeof(execinfo);
      execinfo.lpVerb:='Open';
      execinfo.lpFile:=Pchar(AppName); //所要執行的外部

      execinfo.lpParameters:=Pchar(ParamString);    //參數
      execinfo.fMask:=SEE_MASK_NOCLOSEPROCESS;
      execinfo.nShow:=SW_SHOW;

      if (not ShellExecuteEx(@execinfo)) then
        WarningMsg( '無法執行' + AppName );
    //Application.Minimize;
    //WaitForSingleObject(execinfo.hProcess,INFINITE); //會一直等待直到執
        // 行的外部程式結束，才會把控制權交給你的應用程式。
    //Application.Restore;
  finally
    Screen.Cursor := crDefault;
  end;
    
end;

function TdtmMain.GetEInvDb: String;
begin
  Result := sG_EInvDb;
end;

function TdtmMain.GetEInvUser: String;
begin
  Result := sG_EInvUser;
end;

function TdtmMain.GetEInvPassword: String;
begin
  Result := sG_EInvPassword;
end;




function TdtmMain.ConnectToEINV(aObj: PRIZECOMDb): Boolean;
begin
  Result := False;
  try
    if not Assigned( aObj ) then
     raise Exception.Create( '對應不到發票資料區。' );
    EInvConnection.Connected := False;
    EInvConnection.ConnectionString := Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [aObj.Password, aObj.UserId, aObj.DBSid] );
   EInvConnection.Connected := True;
   Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '【發票】資料區連結失敗, 訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;
end;

function TdtmMain.GetStarTVmail: Boolean;
begin
  Result := sG_StarTVmail;
end;

function TdtmMain.GetStarCMSend: Boolean;
begin
  Result := sG_StarCMSend;
end;

procedure TdtmMain.GetCOMDb(const aFileName: String);
var
  aIniFile: TIniFile;
  aObj : PRIZECOMDb;
  aSection: String;
begin
  if aFileName = EmptyStr then
  begin
    sG_UseCOM := False;

    Exit;
  end;
  aIniFile := TIniFile.Create( aFileName );
  try
    aSection := 'COMMON';
    New( aObj );
    aObj.DBSid := aIniFile.ReadString( aSection , 'SID', EmptyStr );
    aObj.UserId := aIniFile.ReadString( aSection, 'DB_USER', EmptyStr );
    aObj.Password := aIniFile.ReadString( aSection, 'DB_USER_PASSWORD', EmptyStr );
    if ( aobj.DBSid <> EmptyStr ) and ( aObj.UserId <> EmptyStr ) and
      ( aObj.Password <> EmptyStr ) then
    begin
      sG_UseCOM := True;
      sG_EInvDb := aobj.DBSid;
      sG_EInvUser := aobj.UserId;
      sG_EInvPassword := aobj.Password;
      FCOMDbList.Add( aObj );
    end;
  finally
    aIniFile.Free;
  end;

end;

function TdtmMain.GetUseCOM: Boolean;
begin
  Result := sG_UseCOM;
end;

procedure TdtmMain.ReleasePrizeInfo;
var
  aIndex : Integer;
begin
  for aIndex := 0 to FPrizeDbList.Count -1  do
  begin
    if Assigned( FPrizeDbList[aIndex] ) then
    begin

      Dispose( PPRIZEDb( FPrizeDbList[aIndex] ) );
      FPrizeDbList[aIndex] := nil;
    end;
  end;
  FPrizeDbList.Clear;
end;

procedure TdtmMain.ReleaseComInfo;
var
  aIndex : Integer;
begin
  for aIndex := 0 to FCOMDbList.Count -1  do
  begin
    if Assigned( FCOMDbList[aIndex] ) then
    begin

      Dispose( PRIZECOMDb( FCOMDbList[aIndex] ) );
      FCOMDbList[aIndex] := nil;
    end;
  end;
  FCOMDbList.Clear;
end;

function TdtmMain.ConnectToEINV(aSid, aUserId, aPassword: String): Boolean;
begin
  Result := False;
  try
    EInvConnection.Connected := False;
    EInvConnection.ConnectionString := Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [aPassword, aUserId, aSid] );
   EInvConnection.Connected := True;
   Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '【發票】資料區連結失敗, 訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;
end;

function TdtmMain.ConnectToEINVComm(aSid, aUserId,
  aPassword: String): Boolean;
begin
  Result := False;
  try
    EInvConnectComm.Connected := False;
    EInvConnectComm.ConnectionString := Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [aPassword, aUserId, aSid] );
   EInvConnectComm.Connected := True;
   Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '檢核中獎發票資料區連結失敗, 訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;
end;

end.
