unit cbCM2CT;

{$WARN SYMBOL_PLATFORM OFF}

interface          

uses
  Windows, SysUtils, ComObj, ActiveX, Classes, DB, ADODB, DBClient,
  LibCM2CT_TLB, xmldom, msxmldom, XMLIntf, XMLDoc,
  { Oracle Access Component }
  MemDS, VirtualTable, DBAccess, Ora;

const

  {}

  XML_NO_DATA = 'EMPTY';
  XML_WAREHOUSEID = '0';

  { 資料長度檢查 }

  XML_BATCHNO_LEN = 15;
  XML_MATERIALNO_LEN = 15;
  XML_DATE_LEN = 7;
  XML_TIME_LEN = 4;
  XML_VENDORCODE_LEN = 5;
  XML_VENDORNAME_LEN = 20;
  XML_DESCRIPTION_LEN = 30;
  XML_SPEC_LEN = 50;
  XML_HFCMAC_LEN = 20;
  XML_ETHERNETMAC_LEN = 20;
  XML_MTAMAC_LEN = 20;
  XML_CHASSISSN_LEN = 30;
  XML_CUSTOMERSN_LEN = 30;
  XML_MODELNO_LEN = 30;
  XML_MATERIALTYPE_LEN = 50;
  XML_HWVER_LEN = 10;

type

  { 系統台資訊 }

  PSoInfo = ^TSoInfo;
  TSoInfo = record
    CompCode: String;                  { 系統台代碼 }
    CompName: String;                  { 系統台名稱 }
    LoginUser: String;                 { 資料庫帳號 }
    LoginPass: String;                 { 資料庫密碼 }
    DbAliase: String;                  { 資料庫別名 }
    Connection: TADOConnection;
    Reader: TADOQuery;
    Writer: TADOQuery;
    OraSession: TOraSession;
    LogWriter: TOraSQL;
  end;


  { CM 入庫設備檔 }
  
  PSo306 = ^TSo306;
  
  TSo306 = record
    CompCode: String;
    OldCompCode: String;
    BatchNo: String;
    MaterialNo: String;
    CartonnuNo: Integer;
    ModelNo: String;
    ChassisSN: String;
    CustomerSN: String;
    HfcMac: String;
    EthernetMac: String;
    MtaMac: String;
    Flag: Integer;
    CMIssue: String;
    ModeId: Integer;
    ModelCode: Integer;
    ActStatus: String;
    UseFlag: Integer;
    LimitDate: String;
    Belong: Integer;
    VendorCode: String;
    VendorName: String;
    Spec: String;
    Description: String;
    UpdEn: String;
    UpdTime: TDateTime;
    BackHFCMac: String;
    BackMTAMac: String;
    MaterialType: String;
    HwVer: String;
  end;

  { 驗退檔 }

  PSo310 = ^TSo310;
  
  TSo310 = record
    CompCode: String;
    BatchNo: String;
    MaterialNo: String;
    OldBatchNo: String;
    CartonnuNo: Integer;
    ModelNo: String;
    ChassisSN: String;
    CustomerSN: String;
    HfcMac: String;
    EthernetMac: String;
    MtaMac: String;
    Flag: Integer;
    CMIssue: String;
    ModeId: Integer;
    ModelCode: Integer;
    ActStatus: String;
    LimitDate: String;
    Belong: Integer;
    VendorCode: String;
    VendorName: String;
    Spec: String;
    UpdEn: String;
    UpdTime: TDateTime;
    MaterinalType: String;
    HwVer: String;
  end;



  { TCM2CT }

  TCM2CT = class(TAutoObject, ICM2CT)
  private
    FPassPhrase: String;
    FSoList: TStringList;
    FXml: TXMLDocument;
    FDataBuffer: TClientDataSet;
    FSo306: PSo306;
    FSo310: PSo310;
    FType: Integer;
    FLogBuffer: TVirtualTable;
    FCallTime: TDateTime;
    FActiveXmlData: String;
    procedure PrepareSoInfo;
    procedure PrepareSoInfoForTest;
    procedure UnPrepareSoInfo;
    procedure PrepareDataBuffer;
    procedure UnPrepareDataBuffer;
    procedure PrepareLogBuffer;
    procedure UnPrepareLogBuffer;
    function PrepareSoDBConnection(const aCompCode: String): String; overload;
    function PrepareSoDBConnection: String; overload;
    procedure UnPrepareSoDBConnection(const aCompCode: String); overload;
    procedure UnPrepareSoDBConnection; overload;
    function GetActiveSo(const aCompCode: String): PSoInfo;
    function VdCompCode(aNode: IDOMNode; aCompare: String): String;
    function VdBatchNo(aNode: IDOMNode): String; overload;
    function VdBatchNo(aNodes: IDOMNodeList): String; overload;
    function VdMaterialNo(aNode: IDOMNode): String; overload;
    function VdMaterialNo(aNodes: IDOMNodeList): String; overload;
    function VdDate(aNode: IDOMNode): String;
    function VdTime(aNode: IDOMNode): String;
    function VdVendorCode(aNode: IDOMNode): String;
    function VdVendorName(aNode: IDOMNode): String;
    function VdItemCount(aNode, aOwner: IDOMNode; aOwnerHasValue: Boolean = False): String;
    function VdDescription(aNode: IDOMNode): String;
    function VdDescription2(aNode: IDOMNode): String;
    function VdSpec(aNode: IDOMNode): String;
    function VdBelongTo(aNode: IDOMNode): String;
    function VdEqum(aNodes: IDOMNodeList): String;
    function VdPrimaryMAC(aNode: IDOMNode; aId: String): String;
    function VdHFCMac(aNode: IDOMNode): String; overload;
    function VdHFCMac(aNodes: IDOMNodeList): String; overload;
    function VdEtherNetMac(aNode: IDOMNode; aId: String): String;
    function VdMTAMac(aNode: IDOMNode; aId: String): String;
    function VdChassisSN(aNode: IDOMNode; aId: String): String;
    function VdCustomerSN(aNode: IDOMNode; aId: String): String;
    function VdModelId(aNode: IDOMNode; aId: String): String;
    function VdModelNo(aNode: IDOMNode; aId: String): String;
    function VdLimitDate(aNode: IDOMNode; aId: String): String;
    function VdGetOrBack(aNode: IDOMNode): String;
    function VdMasterialType(aNode: IDOMNode; aId: String): String;
    function VdHwVer(aNode: IDOMNode; aId: String): String;
    {}
    function VdBackList(aNodes: IDOMNodeList): String;
    function VdTranList(aNodes: IDOMNodeList): String;
    function VdGetBackList(aNodes: IDOMNodeList): String;
    {}
    function VdChangeList(aNodes: IDOMNodeList): String;
    function VdHfcList(aNodes: IDOMNodeList): String;
    function VdPariEqum(aNode1, aNode2: IDOMNode): String;
    {}
    function VdAlreadyInOtherSo(aCompCode, aId: String): String;
    function VdNotExists(aCompCode, aId: String): String;
    function VdNotUse(aCompCode, aId: String): String;
    function VdBlackList(aCompCode, aId:String): String;
    function VdCD043(aCompCode, aModelNo, aId: String; out aModeCode: String): String;
    function VdCD039(aCompCode: String): String;
    function VdCD022(aCompCode, aDesc, aId: String): String;
    {}
    function VdSpecialRule1(aSo306: PSo306): String;
    function VdExchangeNode(var aSend, aBack: IDOMNode): String;
    function VdBackEqumExists(aId: String): String;
    function VdEqumMtaMac(aHfc, aDbMtaMac, aXmlMtaMac: String; aCompare: Boolean = True): String;
    {}
    function MakeErrorText: String;
    function GetActionText(const CallType: Integer = 0): String;
    procedure AppendLog(aCompCode, aAction, aMsg: String;
      aIsErr: Integer);
    procedure InitCallTime;
    function WriteLog: String;
    function GetChangeCompanyCode(aNode: IDOMNode): String;
  protected
    { 入庫檢驗 }
    function ValidateXML1(aXmlData: String): Boolean;
    { 驗退檢驗 }
    function ValidateXML2(aXmlData: String): Boolean;
    { 調撥檢驗 }
    function ValidateXML3(aXmlData: String): Boolean;
    { 領退料檢驗 }
    function ValidateXML4(aXmlData: String): Boolean;
    { 維修送回檢驗 }
    function ValidateXML5(aXmlData: String): Boolean;
    { 入庫 }
    function SaveXML1: Boolean;
    { 驗退 }
    function SaveXML2: Boolean;
    { 調撥 }
    function SaveXML3: Boolean;
    { 領退料 }
    function SaveXML4: Boolean;
    { 維修送回 }
    function SaveXML5: Boolean;
  public
    procedure Initialize; override;
    destructor Destroy; override;
    { CM 入庫/驗退 }
    function CM2CT_DLL1(aData1, aData2: OleVariant): WideString; safecall;
    { CM 調撥 }
    function CM2CT_DLL2(aData1: OleVariant): WideString; safecall;
    { CM 領退料 }
    function CM2CT_DLL3(aData1: OleVariant): WideString; safecall;
    { CM 維修送回 }
    function CM2CT_DLL5(aData1: OleVariant): WideString; safecall;
  end;


implementation


{ LbCipher, LbString --> Turbopower 的加解密 Library ( LockBox ) }

uses ComServ, LbCipher, LbString, Variants, MaskUtils, StrUtils;


{ ---------------------------------------------------------------------------- }

function ConvertCDate(aCDate: String): String;
var
  aYear, aMonth, aDay: String;
begin
  aYear := Copy( aCDate, 1, 3 );
  aMonth := Copy( aCDate, 4, 2 );
  aDay := Copy( aCDate, 6, 2 );
  aYear := IntToStr( StrToInt( aYear ) + 1911 );
  Result := FormatMaskText( '####/##/##;0;_' , aYear + aMonth + aDay );
end;

{ ---------------------------------------------------------------------------- }

function ConvertTime(aTime: String): String;
var
  aHour, aMin: String;
begin
  aHour := Copy( aTime, 1, 2 );
  aMin := Copy( aTime, 3, 2 );
  Result := FormatMaskText( '##:##:##;0;_' , aHour + aMin + '00' );
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

function Nvl(InValue: String; NullValue: String): String; overload;
begin
  Result := NullValue;
  if InValue <> '' then Result := InValue;
end;

{ ---------------------------------------------------------------------------- }

{ TCM2CT }

procedure TCM2CT.Initialize;
begin
  inherited;
  FSoList := TStringList.Create;
  FXml := TXMLDocument.Create( nil );
  FDataBuffer := TClientDataSet.Create( nil );
  FLogBuffer := TVirtualTable.Create( nil );
  FLogBuffer.Options := [];
  New( FSo306 );
  New( FSo310 );
  //PrepareSoInfo;
  PrepareSoInfoForTest;
  FActiveXmlData := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

destructor TCM2CT.Destroy;
begin
  FXml.Free;
  Dispose( FSo306 );
  Dispose( FSo310 );
  UnPrepareDataBuffer;
  UnPrepareLogBuffer;
  UnPrepareSoInfo;
  FDataBuffer.Free;
  FLogBuffer.Free;
  FSoList.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.PrepareSoInfo;
var
  aIndex: Integer;
  aSoInfo: PSoInfo;
  aKey: TKey64;
begin
  { DES 加解密用的 Key --> CS (大寫字母) }
  FPassPhrase := Chr( 67 ) + Chr( 83 );
  { 用加解密FPassPhrase值產生一組解密用的 array 結構 }
  GenerateLMDKey( aKey, SizeOf( aKey ), FPassPhrase );
  for aIndex := 0 to 20 do
  begin
    New( aSoInfo );
    aSoInfo.CompCode := IntToStr( aIndex );
    { 使用 DES 解密回原本字串, 避免被震江讀出 }
    { 以下是 EMC 正式資料區的資料庫連線字元, 若連線字元有改 }
    { 則必須改過後, 重新 Buid 一次再給震江使用 }
    case aIndex of
      { 總倉 com/com@emc }
      0: begin
           aSoInfo.CompName := DESEncryptStringEx( 'GK6rPSGTehY=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
      { 觀昇 emcks/emc5601ks@emc6}
      1: begin
           aSoInfo.CompName := DESEncryptStringEx( 'rGpA+1maUhg=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'qKvF5yaDktg=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'oa/TVQdZIlHk3oz3EzOLkw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'fkO5zCYBdJI=', aKey, False );
         end;
      { 屏南 emcpn/emc6202pn@emc6}
      2: begin
           aSoInfo.CompName := DESEncryptStringEx( 'akhHiBFhS9I=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'f2CQAghzZ5s=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '44F1CiX5YaflkqSw2nKDxg==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'fkO5zCYBdJI=', aKey, False );
         end;
      { 南天 emcnt/emc6403nt@emc5 }
      3: begin
           aSoInfo.CompName := DESEncryptStringEx( 'qfEoE8BXGhQ=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'eR5muOd5ncQ=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'wfb7rptxl5j9onZBwB6okQ==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MUL6BtuWyt4=', aKey, False );
         end;
      { 新頻道 emcncc/emc6105ncc@emc4}
      5: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Zr1PBwa8Byw=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'vfd+I/JKQwQ=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '0k3E8Wv+LtKFbBsej4DGrw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'jixfiCHuC34=', aKey, False );
         end;
      { 豐盟 emcfm/emc1906fm@emc4 }
      6: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Ah8M/NBq8F8=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'SiQSz4YRrEA=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '7Wf24kRroFQ/ZoOJ2rU25A==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'jixfiCHuC34=', aKey, False );
         end;
      { 振道 emcct/emc5707ct@emc3 }
      7: begin
           aSoInfo.CompName := DESEncryptStringEx( '3Vysqp7t/kg=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'n86eGJ+19io=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'Ui+Zp/vzI7H9onZBwB6okQ==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'lqH+cXxXoVM=', aKey, False );
         end;
      { 全聯 emcuc/emc5508uc@emc1}
      8: begin
           aSoInfo.CompName := DESEncryptStringEx( 'xuL0QY75BFU=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'fxaKJE3mXRU=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'O+x1I/fitrq0xgnUHu0CqA==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'mf5f/EFTrNE=', aKey, False );
         end;
      { 陽明山 emcyms/emc5909yms@emc }
      9: begin
           aSoInfo.CompName := DESEncryptStringEx( 'awStqsVwxTA=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( '6Dtr2sVE7eU=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '1wW0aTDGGC3T4qhwC249LA==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { 新台北 emcntp/emc6010ntp@emc }
     10: begin
           aSoInfo.CompName := DESEncryptStringEx( 'mzIxRa4Iq2s=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'fZPrFuqCIU4=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'bLKoJ8ryZnSpUcfLMZxZLw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { 金頻道 emckp/emc2511kp@emc }
     11: begin
           aSoInfo.CompName := DESEncryptStringEx( '6tW9Hl1M7R4=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'dbhSjDnsLqI=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '2osHY0C6C+p0A5aiCm2nOw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { 大安文山 emcdw/emc5812dw@emc }
     12: begin
           aSoInfo.CompName := DESEncryptStringEx( '8uFrustgM9ZrZ2Kx7hdylw==', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'dMHHX3Ucjbs=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'vjKNxtn+QScOsJK+bY9iIw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { 新唐城 emctc/emctc@emc }
     13: begin
           aSoInfo.CompName := DESEncryptStringEx( '2CT3OdYNWHM=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'lY3LUuztcSE=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'lY3LUuztcSE=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { 大新店 emcst/emcst@emc2 }
     14: begin
           aSoInfo.CompName := DESEncryptStringEx( 'nKNwWhHJnJQ=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'w9XKCpLq4iE=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'w9XKCpLq4iE=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'Kw1slo9ivtU=', aKey, False );
         end;
     { 北桃園 emcnty/emc2916nty@emc7 }
     16: begin
           aSoInfo.CompName := DESEncryptStringEx( 'BEsOx4FNwcI=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'HmFB29lZPTg=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '2uubyluDZQ2pIcIsdVqyIA==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'bO/fSqEvXk0=', aKey, False );
         end;
     { 唐城 emctc_cm/emctc_cm@emc7, 唐城的 CM 業務獨一立一間系統台做, 系統台代碼為 17, 舊的代碼 13 不會用到 }
     { 2007/12/25 新唐城資料庫回歸北平台, 改回用原本的系統台代碼 }
     17: begin
           aSoInfo.CompName := DESEncryptStringEx( '2CT3OdYNWHM=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'xc2y/MAu33BrZ2Kx7hdylw==', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'xc2y/MAu33BrZ2Kx7hdylw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'bO/fSqEvXk0=', aKey, False );
         end;
     { 倉管測試資料庫 kptest/kptest@emc }
     20: begin
           aSoInfo.CompName := DESEncryptStringEx( 'W2DfZE0KwHB4C/bFD91dLA==', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'XP0zgPgTEls=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'XP0zgPgTEls=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
    else begin
          Dispose( aSoInfo );
          Continue;
        end;
    end;
    aSoInfo.Connection := TADOConnection.Create( nil );
    aSoInfo.Connection.LoginPrompt := False;
    aSoInfo.Connection.ConnectionString := Format(
      'Provider=MSDAORA.1;Persist Security Info=True;' +
      'Password=%s;User ID=%s;Data Source=%s;',
      [aSoInfo.LoginPass, aSoInfo.LoginUser, aSoInfo.DbAliase] );
    {}
    aSoInfo.Reader := TADOQuery.Create( nil );
    aSoInfo.Reader.Connection := aSoInfo.Connection;
    aSoInfo.Reader.CacheSize := 100;
    aSoInfo.Reader.Prepared := True;
    {}
    aSoInfo.Writer := TADOQuery.Create( nil );
    aSoInfo.Writer.Connection := aSoInfo.Connection;
    aSoInfo.Writer.CacheSize := 100;
    aSoInfo.Writer.Prepared := True;
    { 寫 Log 用, 使用 Oracle Access Component ODAC 來寫 CLOB Column }
    { Log 一律寫到總倉 }
    aSoInfo.OraSession := nil;
    aSoInfo.LogWriter := nil;
    if ( aIndex = 0 ) then
    begin
      aSoInfo.OraSession := TOraSession.Create( nil );
      aSoInfo.OraSession.ConnectPrompt := False;
      aSoInfo.OraSession.AutoCommit := True;
      aSoInfo.OraSession.ConnectString := Format( '%s/%s@%s', [
        aSoInfo.LoginUser, aSoInfo.LoginPass, aSoInfo.DbAliase] );
      {}
      aSoInfo.LogWriter := TOraSQL.Create( nil );
      aSoInfo.LogWriter.Session := aSoInfo.OraSession;
      aSoInfo.LogWriter.AutoCommit := True;
    end;
    FSoList.AddObject( aSoInfo.CompCode, TObject( aSoInfo ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.PrepareSoInfoForTest;
var
  aIndex: Integer;
  aSoInfo: PSoInfo;
  aKey: TKey64;
begin
  { DES 加解密用的 Key --> CS (大寫字母) }
  FPassPhrase := Chr( 67 ) + Chr( 83 );
  { 用加解密FPassPhrase值產生一組解密用的 array 結構 }
  GenerateLMDKey( aKey, SizeOf( aKey ), FPassPhrase );
  for aIndex := 0 to 20 do
  begin
    New( aSoInfo );
    aSoInfo.CompCode := IntToStr( aIndex );
    { 使用 DES 解密回原本字串, 避免被震江讀出 }
    { 以下是 EMC 正式資料區的資料庫連線字元, 若連線字元有改 }
    { 則必須改過後, 重新 Buid 一次再給震江使用 }
    case aIndex of
      { 總倉 COM/COM@RDKNET }
      0: begin
           aSoInfo.CompName := DESEncryptStringEx( 'GK6rPSGTehY=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MnXBjY2nyV4=', aKey, False );
         end;
      { 新頻道 NCCTEST/NCCTEST@CATVO }
      5: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Zr1PBwa8Byw=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'QqLvWdgA/ek=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'QqLvWdgA/ek=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'HnhysRH/TXc=', aKey, False );
         end;
      { 豐盟 EMCFM/EMCFM@RDKNET }
      6: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Ah8M/NBq8F8=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'bGiEDA+FVOc=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'bGiEDA+FVOc=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MnXBjY2nyV4=', aKey, False );
         end;   
      { 北桃園 EMCNTY/EMCNTY@MIS }
     16: begin
           aSoInfo.CompName := DESEncryptStringEx( 'BEsOx4FNwcI=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'ANlgkPXM+Pw=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'ANlgkPXM+Pw=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MP/Q8McZXhA=', aKey, False );
         end;
    else begin
          Dispose( aSoInfo );
          Continue;
        end;
    end;
    {}
    aSoInfo.Connection := TADOConnection.Create( nil );
    aSoInfo.Connection.LoginPrompt := False;
    aSoInfo.Connection.ConnectionString := Format(
      'Provider=MSDAORA.1;Persist Security Info=True;' +
      'Password=%s;User ID=%s;Data Source=%s;',
      [aSoInfo.LoginPass, aSoInfo.LoginUser, aSoInfo.DbAliase] );
    {}
    aSoInfo.Reader := TADOQuery.Create( nil );
    aSoInfo.Reader.Connection := aSoInfo.Connection;
    aSoInfo.Reader.CacheSize := 100;
    aSoInfo.Reader.Prepared := True;
    {}
    aSoInfo.Writer := TADOQuery.Create( nil );
    aSoInfo.Writer.Connection := aSoInfo.Connection;
    aSoInfo.Writer.CacheSize := 100;
    aSoInfo.Writer.Prepared := True;
    { 寫 Log 用, 使用 Oracle Access Component ODAC 來寫 CLOB Column }
    { Log 一律寫到總倉 }
    aSoInfo.OraSession := nil;
    aSoInfo.LogWriter := nil;
    if ( aIndex = 0 ) then
    begin
      aSoInfo.OraSession := TOraSession.Create( nil );
      aSoInfo.OraSession.ConnectPrompt := False;
      aSoInfo.OraSession.AutoCommit := True;
      aSoInfo.OraSession.ConnectString := Format( '%s/%s@%s', [
        aSoInfo.LoginUser, aSoInfo.LoginPass, aSoInfo.DbAliase] );
      {}
      aSoInfo.LogWriter := TOraSQL.Create( nil );
      aSoInfo.LogWriter.Session := aSoInfo.OraSession;
      aSoInfo.LogWriter.AutoCommit := True;
    end;
    FSoList.AddObject( aSoInfo.CompCode, TObject( aSoInfo ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.UnPrepareSoInfo;
var
  aIndex: Integer;
  aSoInfo: PSoInfo;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if not Assigned( FSoList.Objects[aIndex] ) then Continue;
    aSoInfo := PSoInfo( FSoList.Objects[aIndex] );
    if aSoInfo.Connection.InTransaction then aSoInfo.Connection.RollbackTrans;
    if aSoInfo.Reader.Active then aSoInfo.Reader.Close;
    if aSoInfo.Writer.Active then aSoInfo.Writer.Close;
    if aSoInfo.Connection.Connected then aSoInfo.Connection.Connected := False;
    aSoInfo.Reader.Free;
    aSoInfo.Writer.Free;
    aSoInfo.Connection.Free;
    {}
    if Assigned( aSoInfo.LogWriter ) then
      aSoInfo.LogWriter.Free;
    if Assigned( aSoInfo.OraSession ) then
    begin
      if aSoInfo.OraSession.Connected then aSoInfo.OraSession.Connected := False;
      aSoInfo.OraSession.Free;
    end;
    Dispose( aSoInfo );
    FSoList.Objects[aIndex] := nil;
  end;
  FSoList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.PrepareDataBuffer;

  { --------------------------------------------------- }

  procedure InitBufferField1;
  begin
    { 入庫寫入 SO306 用 }
    FDataBuffer.FieldDefs.Add( 'COMPCODE', ftInteger );       { 公司別 ( 調入 ) }
    FDataBuffer.FieldDefs.Add( 'OLDCOMPCODE', ftInteger );    { 公司別 ( 調出 ) }    
    FDataBuffer.FieldDefs.Add( 'BATCHNO', ftString, 15 );     { 批號 }
    FDataBuffer.FieldDefs.Add( 'MATERIALNO', ftString, 15 );  { 料號 }
    FDataBuffer.FieldDefs.Add( 'CARTONNUNO', ftInteger );     { 紙箱序號 }
    FDataBuffer.FieldDefs.Add( 'MODELNO', ftString, 30 );     { 機種/品名 }
    FDataBuffer.FieldDefs.Add( 'CHASSISSN', ftString, 30 );   { 基版序號 }
    FDataBuffer.FieldDefs.Add( 'CUSTOMERSN', ftString, 30 );  { 機器序號 }
    FDataBuffer.FieldDefs.Add( 'HFCMAC', ftString, 20 );      { HFCMAC 序號 }
    FDataBuffer.FieldDefs.Add( 'ETHERNETMAC', ftString, 20 ); { Ethernet MAC 序號 }
    FDataBuffer.FieldDefs.Add( 'MTAMAC', ftString, 20 );      { MTAMAC 序號 }
    FDataBuffer.FieldDefs.Add( 'FLAG', ftInteger );           { 以那一個 MAC 序號為主, 預設 1 }
    FDataBuffer.FieldDefs.Add( 'CMISSUE', ftString, 20 );     { CM 匯入日期, 執行此 DLL 的日期, 年月日時分秒 }
    FDataBuffer.FieldDefs.Add( 'MODEID', ftInteger );         { 設備型號, 0=EMTA, 1=CM }
    FDataBuffer.FieldDefs.Add( 'MODELCODE', ftInteger );      { 設備型號代碼 }
    FDataBuffer.FieldDefs.Add( 'ACTSTATUS', ftString, 6 );    { SPM 設定狀態 }
    FDataBuffer.FieldDefs.Add( 'USEFLAG', ftInteger );        { 可用狀態, 0=可用, 1=不可用 }
    FDataBuffer.FieldDefs.Add( 'LIMITDATE', ftString, 10 );   { 保固期限 , 年月日 }
    FDataBuffer.FieldDefs.Add( 'BELONG', ftInteger );         { 財產規屬單位, 0=EMC, 1=APTG }
    FDataBuffer.FieldDefs.Add( 'VENDORCODE', ftString, 5 );   { 廠商代碼 }
    FDataBuffer.FieldDefs.Add( 'VENDORNAME', ftString, 20 );  { 廠商名稱 }
    FDataBuffer.FieldDefs.Add( 'SPEC', ftString, 50 );        { 規格 }
    FDataBuffer.FieldDefs.Add( 'DESCRIPTION', ftString, 30 ); { 客服品名 }
    FDataBuffer.FieldDefs.Add( 'UPDEN', ftString, 20 );       { 異動人員 }
    FDataBuffer.FieldDefs.Add( 'UPDTIME', ftDateTime );       { 異動時間 }
    FDataBuffer.FieldDefs.Add( 'ALREADYUSE', ftDateTime );    { 首次使用 }
    FDataBuffer.FieldDefs.Add( 'MATERIALTYPE', ftString, 50 ); { 物料型號 }
    FDataBuffer.FieldDefs.Add( 'HWVER', ftString, 10 );       { 硬體版本 }
    {}
    FDataBuffer.FieldDefs.Add( 'SENDMTAMAC', ftString, 20 );  { 維修 EMTAMAC序號 }
    FDataBuffer.FieldDefs.Add( 'BACKHFCMAC', ftString, 20 );  { 送回 HFCMAC序號 }
    FDataBuffer.FieldDefs.Add( 'BACKMTAMAC', ftString, 20 );  { 送回 EMTAMAC序號 }
    {}
  end;

  { --------------------------------------------------- }

  procedure InitBufferField2;
  begin
    { 驗退寫入 SO310 用 }
    FDataBuffer.FieldDefs.Add( 'COMPCODE', ftInteger );       { 公司別 }
    FDataBuffer.FieldDefs.Add( 'BATCHNO', ftString, 15 );     { 驗退單號 }
    FDataBuffer.FieldDefs.Add( 'MATERIALNO', ftString, 15 );  { 料號 }
    FDataBuffer.FieldDefs.Add( 'OLDBATCHNO', ftString, 15 );  { 原入庫批號 }
    FDataBuffer.FieldDefs.Add( 'CARTONNUNO', ftInteger );     { 紙箱序號 }
    FDataBuffer.FieldDefs.Add( 'MODELNO', ftString, 30 );     { 機種/品名 }
    FDataBuffer.FieldDefs.Add( 'CHASSISSN', ftString, 30 );   { 基版序號 }
    FDataBuffer.FieldDefs.Add( 'CUSTOMERSN', ftString, 30 );  { 機器序號 }
    FDataBuffer.FieldDefs.Add( 'HFCMAC', ftString, 20 );      { HFCMAC 序號 }
    FDataBuffer.FieldDefs.Add( 'ETHERNETMAC', ftString, 20 ); { Ethernet MAC 序號 }
    FDataBuffer.FieldDefs.Add( 'MTAMAC', ftString, 20 );      { MTAMAC 序號 }
    FDataBuffer.FieldDefs.Add( 'FLAG', ftInteger  );          { 以那一個 MAC 序號為主, 預設 1 }
    FDataBuffer.FieldDefs.Add( 'CMISSUE', ftString, 20 );     { CM 匯入日期, 執行此 DLL 的日期, 年月日時分秒 }
    FDataBuffer.FieldDefs.Add( 'MODEID', ftInteger );         { 設備型號, 0=EMTA, 1=CM }
    FDataBuffer.FieldDefs.Add( 'MODELCODE', ftInteger );      { 設備型號代碼 }
    FDataBuffer.FieldDefs.Add( 'ACTSTATUS', ftString, 6 );    { SPM 設定狀態 }
    FDataBuffer.FieldDefs.Add( 'LIMITDATE', ftString, 10 );   { 保固期限 , 年月日 }
    FDataBuffer.FieldDefs.Add( 'BELONG', ftInteger );         { 財產規屬單位, 0=EMC, 1=APTG }
    FDataBuffer.FieldDefs.Add( 'VENDORCODE', ftString, 5 );   { 廠商代碼 }
    FDataBuffer.FieldDefs.Add( 'VENDORNAME', ftString, 20 );  { 廠商名稱 }
    FDataBuffer.FieldDefs.Add( 'SPEC', ftString, 50 );        { 規格 }
    FDataBuffer.FieldDefs.Add( 'UPDEN', ftString, 20 );       { 異動人員 }
    FDataBuffer.FieldDefs.Add( 'UPDTIME', ftDateTime );       { 異動時間 }
    FDataBuffer.FieldDefs.Add( 'ALREADYUSE', ftDateTime );    { 首次使用 }
    FDataBuffer.FieldDefs.Add( 'MATERIALTYPE', ftString, 50 ); { 物料型號 }
    FDataBuffer.FieldDefs.Add( 'HWVER', ftString, 10 );       { 硬體版本 }
  end;

  { --------------------------------------------------- }

begin
  FDataBuffer.FieldDefs.Clear;
  case FType of
    1, 3, 4, 5: InitBufferField1;
    2: InitBufferField2;
  end;
  FDataBuffer.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.PrepareLogBuffer;
begin
  FLogBuffer.FieldDefs.Clear;
  FLogBuffer.FieldDefs.Add( 'CALLTIME', ftDateTime );
  FLogBuffer.FieldDefs.Add( 'ACTION', ftString, 50 );
  FLogBuffer.FieldDefs.Add( 'COMPCODE', ftString, 10 );
  FLogBuffer.FieldDefs.Add( 'COMPNAME', ftString, 50 );
  FLogBuffer.FieldDefs.Add( 'ISERR', ftInteger );
  FLogBuffer.FieldDefs.Add( 'MSG', ftMemo );
  FLogBuffer.FieldDefs.Add( 'XMLDATA', ftMemo );
  FLogBuffer.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.UnPrepareDataBuffer;
begin
  if not ( VarIsNull( FDataBuffer.Data ) ) then
    FDataBuffer.EmptyDataSet;
  FDataBuffer.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.UnPrepareLogBuffer;
begin
  if FLogBuffer.Active then
    FLogBuffer.Close;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.PrepareSoDBConnection(const aCompCode: String): String;
var
  aSoInfo: PSoInfo;
  aIndex: Integer;
begin
  Result := EmptyStr;
  try
    aIndex := FSoList.IndexOf( aCompCode );
    if aIndex < 0 then
      raise Exception.CreateFmt( '無此公司別%s。', [aCompCode] )
    else
    begin
      aSoInfo := PSoInfo( FSoList.Objects[aIndex] );
      aSoInfo.Connection.Connected := True;
      if Assigned( aSoInfo.OraSession ) then
        aSoInfo.OraSession.Connected := True;
    end;
  except
    on E: Exception do Result := E.Message;
  end;
  if Result <> EmptyStr then
    Result := Format( '-1: 與CC&B資料庫連線失敗,錯誤訊息:<%s>', [Result] );
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.PrepareSoDBConnection: String;
var
  aIndex: Integer;
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := PSoInfo( FSoList.Objects[aIndex] );
    Result := PrepareSoDBConnection( aSoInfo.CompCode );
    if Result <> EmptyStr then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.UnPrepareSoDBConnection(const aCompCode: String);
var
  aIndex: Integer;
  aSoInfo: PSoInfo;
begin
  try
    aIndex := FSoList.IndexOf( aCompCode );
    aSoInfo := PSoInfo( FSoList.Objects[aIndex] );
    if aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.RollbackTrans;
    aSoInfo.Reader.Close;
    aSoInfo.Writer.Close;  
    aSoInfo.Connection.Connected := False;
    if Assigned( aSoInfo.OraSession ) then
      aSoInfo.OraSession.Connected := False;
  except
    { }
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.UnPrepareSoDBConnection;
var
  aIndex: Integer;
  aSoInfo: PSoInfo;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    aSoInfo := PSoInfo( FSoList.Objects[aIndex] );
    UnPrepareSoDBConnection( aSoInfo.CompCode );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.GetActiveSo(const aCompCode: String): PSoInfo;
var
  aIndex: Integer;
begin
  Result := nil;
  aIndex := FSoList.IndexOf( aCompCode );
  if aIndex >= 0 then
    Result := PSoInfo( FSoList.Objects[aIndex] );
end;

{ ---------------------------------------------------------------------------- }

{ 批號檢查 }

function TCM2CT.VdBatchNo(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫批號格式錯誤, 批號=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退單號格式錯誤, 單號<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫批號長度錯誤, 批號=<%s>,入庫批號的長度最多<%d碼>。',
          [aNode.nodeValue, XML_BATCHNO_LEN] );
      2: Result := Format( '-2:驗退單號長度錯誤, 批號=<%s>,驗退單號的長度最多<%d碼>。',
          [aNode.nodeValue, XML_BATCHNO_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Trim( aNode.nodeValue ) = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( Trim( aNode.nodeValue ) ) > XML_BATCHNO_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdBatchNo(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫批號長度錯誤, 批號=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退批號長度錯誤, 批號=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdCompCode(aNode: IDOMNode; aCompare: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫公司代碼應為總倉代碼<%s>,錯誤公司代碼=<%s>。',
          [XML_WAREHOUSEID, '空值'] );
      2: Result := Format( '-2:驗退公司代碼應為總倉代碼<%s>,錯誤公司代碼=<%s>。',
          [XML_WAREHOUSEID, '空值'] );
      3: Result := Format( '-2:調撥公司代碼必須有值,錯誤公司代碼=<%s>。',
          ['空值'] );
      4: Result := Format( '-2:領退料公司代碼必須有值,錯誤公司代碼=<%s>。',
          ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫公司代碼應為總倉代碼<%s>,錯誤公司代碼=<%s>。',
          [XML_WAREHOUSEID, aNode.nodeValue] );
      2: Result := Format( '-2:驗退公司代碼應為總倉代碼<%s>,錯誤公司代碼=<%s>。',
          [XML_WAREHOUSEID, aNode.nodeValue] );
      3: Result := Format( '-2:調撥公司代碼應為有效的公司代碼,錯誤公司代碼=<%s>。',
          [aNode.nodeValue] );
      4: Result := Format( '-2:領退料公司代碼不可為總倉代碼,錯誤公司代碼=<%s>。',
          [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription3: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:無此入庫公司代碼=<%s>。',
          [aNode.nodeValue] );
      2: Result := Format( '-2:無此驗退公司代碼=<%s>。',
          [aNode.nodeValue] );
      3: Result := Format( '-2:無此調撥公司代碼=<%s>。',
          [aNode.nodeValue] );
      4: Result := Format( '-2:無此領退料公司代碼=<%s>。',
          [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }

var
  aIndex: Integer;
begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  { 公司代碼空值 }
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  { 須要比對公司別, Ex: aCompare = 1 }
  if ( aCompare <> EmptyStr ) then
  begin
    { 假如 aCompare <> aNode.nodeValue, 入庫/驗退 }
    if ( aNode.nodeValue <> aCompare ) and ( FType in [1,2] ) then
    begin
      Result := GetErrorDescription2;
      Exit;
    end;
    { 假如 aCompare = aNode.nodeValue, 領/退料 }
    if ( aNode.nodeValue = aCompare ) and ( FType in [4] ) then
    begin
      Result := GetErrorDescription2;
      Exit;
    end;
  end;
  { 檢驗公司別是否合法 }
  aIndex := FSoList.IndexOf( aNode.nodeValue );
  if ( aIndex < 0 ) then
  begin
    Result := GetErrorDescription3;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdMaterialNo(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫料號格式錯誤,料號=<%s>,。', ['空值'] );
      2: Result := Format( '-2:驗退料號格式錯誤,料號=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫料號長度錯誤,料號=<%s>,入庫料號的長度最多<%d碼>。',
          [aNode.nodeValue, XML_MATERIALNO_LEN] );
      2: Result := Format( '-2:驗退料號長度錯誤,料號=<%s>,入庫料號的長度最多<%d碼>。',
          [aNode.nodeValue, XML_MATERIALNO_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( Trim( aNode.nodeValue ) ) > XML_MATERIALNO_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdMaterialNo(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫料號格式錯誤,料號=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退料號格式錯誤,料號=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdDate(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫日期格式錯誤,日期=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退日期格式錯誤,日期=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫日期格式錯誤,日期=<%s>。', [aNode.nodeValue] );
      2: Result := Format( '-2:驗退日期格式錯誤,日期=<%s>。', [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }

var
  aYear, aMonth, aDay: String;
begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  aYear := Copy( aNode.nodeValue, 1, 3 );
  aMonth := Copy( aNode.nodeValue, 4, 2 );
  aDay := Copy( aNode.nodeValue, 6, 2 );
  if ( Length( aYear ) <> 3 ) or
     ( Length( aMonth ) <> 2 ) or
     ( Length( aDay ) <> 2 ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
  try
    aYear := IntToStr( StrToInt( aYear ) + 1911 );
    StrToDate( FormatMaskText( '####/##/##;0;_', aYear + aMonth + aDay ) );
  except
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdTime(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫時間格式錯誤,時間=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退時間格式錯誤,時間=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫時間格式錯誤,時間=<%s>。', [aNode.nodeValue] );
      2: Result := Format( '-2:驗退時間格式錯誤,時間=<%s>。', [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }

var
  aHour, aMin: String;
begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  aHour := Copy( aNode.nodeValue, 1, 2 );
  aMin := Copy( aNode.nodeValue, 3, 2 );
  if ( Length( aHour ) <> 2 ) or
     ( Length( aMin ) <> 2 ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
  try
    StrToTime( FormatMaskText( '##:##:##;0;_', aHour + aMin + '00' ) );
  except
    Result := GetErrorDescription2;
    Exit;
  end;
end;
                                    
{ ---------------------------------------------------------------------------- }

function TCM2CT.VdVendorCode(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫廠商代碼格式錯誤,廠商代碼=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退廠商代碼格式錯誤,廠商代碼=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫廠商代碼長度錯誤,廠商代碼=<%s>,廠商代碼的長度最多<%d碼>。',
          [aNode.nodeValue, XML_VENDORCODE_LEN] );
      2: Result := Format( '-2:驗退廠商代碼長度錯誤,廠商代碼=<%s>,廠商代碼的長度最多<%d碼>。',
          [aNode.nodeValue, XML_VENDORCODE_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( aNode.nodeValue ) > XML_VENDORCODE_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdVendorName(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫廠商名稱格式錯誤,廠商名稱=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退廠商名稱格式錯誤,廠商名稱=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫廠商名稱長度錯誤,廠商名稱=<%s>,廠商名稱的長度最多<%d碼>。',
          [aNode.nodeValue, XML_VENDORNAME_LEN] );
      2: Result := Format( '-2:驗退廠商名稱長度錯誤,廠商名稱=<%s>,廠商名稱的長度最多<%d碼>。',
          [aNode.nodeValue, XML_VENDORNAME_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( aNode.nodeValue ) > XML_VENDORNAME_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdItemCount(aNode, aOwner: IDOMNode; aOwnerHasValue: Boolean = False): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫數量格式錯誤,數量=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退數量格式錯誤,數量=<%s>。', ['空值'] );
      3: Result := Format( '-2:調撥數量格式錯誤,數量=<%s>。', ['空值'] );
      4: Result := Format( '-2:領退料數量格式錯誤,數量=<%s>。', ['空值'] );
      5: Result := Format( '-2:維修送回數量格式錯誤,數量=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫數量格式錯誤,數量=<%s>。', [aNode.nodeValue] );
      2: Result := Format( '-2:驗退數量格式錯誤,數量=<%s>。', [aNode.nodeValue] );
      3: Result := Format( '-2:調撥數量格式錯誤,數量=<%s>。', [aNode.nodeValue] );
      4: Result := Format( '-2:領退料數量格式錯誤,數量=<%s>。', [aNode.nodeValue] );
      5: Result := Format( '-2:維修送回數量格式錯誤,數量=<%s>。', [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }
  
  function GetErrorDescription3: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫數量錯誤,入庫數量=<%s>,實際明細筆數=<%d>。',
          [aNode.nodeValue, aOwner.childNodes.length] );
      2: Result := Format( '-2:驗退數量錯誤,驗退數量=<%s>,實際明細筆數=<%d>。',
          [aNode.nodeValue, aOwner.childNodes.length] );
      3: Result := Format( '-2:調撥數量錯誤,調撥數量=<%s>,實際明細筆數=<%d>。',
          [aNode.nodeValue, aOwner.childNodes.length] );
      4: Result := Format( '-2:領退料數量錯誤,領退料數量=<%s>,實際明細筆數=<%d>。',
          [aNode.nodeValue, aOwner.childNodes.length] );
      5: Result := Format( '-2:維修送回數量錯誤,維修送回數量=<%s>,實際明細筆數=<%d>。',
          [aNode.nodeValue, aOwner.childNodes.length] );
    end;
  end;

  { ----------------------------------------------- }

var
  aCount, aCompare: Integer;
begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if not TryStrToInt( aNode.nodeValue, aCount ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
  aCompare := aOwner.childNodes.length;
  if aOwnerHasValue then Dec( aCompare );
  if ( aCount <> aCompare ) then
  begin
    Result := GetErrorDescription3;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdDescription(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫品名格式錯誤,品名=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退品名格式錯誤,品名=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫品名長度錯誤,品名=<%s>,最多品名長度=<%d>。',
          [aNode.nodeValue, XML_DESCRIPTION_LEN] );
      2: Result := Format( '-2:驗退品名長度錯誤,品名=<%s>,最多品名長度=<%d>。',
          [aNode.nodeValue,XML_DESCRIPTION_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( aNode.nodeValue ) > XML_DESCRIPTION_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdDescription2(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫客服品名格式錯誤,客服品名=<%s>。', ['空值'] );
      3: Result := Format( '-2:調撥客服品名格式錯誤,客服品名=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫客服品名長度錯誤,客服品名=<%s>,最多客服品名長度=<%d>。',
          [aNode.nodeValue, XML_DESCRIPTION_LEN] );
      3: Result := Format( '-2:調撥客服品名長度錯誤,客服品名=<%s>,最多客服品名長度=<%d>。',
          [aNode.nodeValue, XML_DESCRIPTION_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( aNode.nodeValue ) > XML_DESCRIPTION_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdSpec(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫規格長度錯誤,規格=<%s>,最多規格長度=<%d碼>。',
          [aNode.nodeValue, XML_SPEC_LEN] );
      2: Result := Format( '-2:驗退規格長度錯誤,規格=<%s>,最多長度=<%d碼>。',
          [aNode.nodeValue,XML_SPEC_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then Exit;
  if ( Length( aNode.nodeValue ) > XML_SPEC_LEN ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdBelongTo(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫列帳單位錯誤,列帳單位=<%s>,列帳單位必須符合下列值<%s>。',
          [aNode.nodeValue, ( '0,1' )] );
      2: Result := Format( '-2:驗退列帳單位錯誤,列帳單位=<%s>,列帳單位必須符合下列值<%s>。',
          [aNode.nodeValue, ( '0,1' )] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then  Exit;
  if ( aNode.nodeValue = EmptyWideStr ) then Exit;
  if ( aNode.nodeValue <> '0' ) and
     ( aNode.nodeValue <> '1' ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdEqum(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備格式錯誤,設備格式內容=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退設備格式錯誤,設備格式內容=<%s>。', ['空值'] );
      3: Result := Format( '-2:調撥設備格式錯誤,設備格式內容=<%s>。', ['空值'] );
      5: Result := Format( '-2:維修送回設備格式錯誤,設備格式內容=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdPrimaryMAC(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備主序號別格式錯誤,主序號別=<%s>, HFCMAC序號=<%s>。',
        ['空值', aId] );
      2: Result := Format( '-2:驗退設備主序號別格式錯誤,主序號別=<%s>, HFCMAC序號=<%s>。',
        ['空值', aId] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備主序號別格式錯誤,主序號別=<%s>,主序號別必須符合下列值<%s>, HFCMAC序號=<%s>。',
          [aNode.nodeValue, '1', aId] );
      2: Result := Format( '-2:驗退設備主序號別格式錯誤,主序號別=<%s>,主序號別必須符合下列值<%s>, HFCMAC序號=<%s>。',
          [aNode.nodeValue, '1', aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue <> '1' ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdHFCMac(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
      3: Result := Format( '-2:調撥HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
      4: Result := Format( '-2:領/退料HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
      5: Result := Format( '-2:維修送回HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫HFCMAC序號格式錯誤,HFCMAC序號=<%s>,HFCMAC序號最大長度<%d碼>。',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      2: Result := Format( '-2:驗退HFCMAC序號格式錯誤,HFCMAC序號=<%s>,HFCMAC序號最大長度<%d碼>。',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      3: Result := Format( '-2:調撥HFCMAC序號格式錯誤,HFCMAC序號=<%s>,HFCMAC序號最大長度<%d碼>。',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      4: Result := Format( '-2:領/退料HFCMAC序號格式錯誤,HFCMAC序號=<%s>,HFCMAC序號最大長度<%d碼>。',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      5: Result := Format( '-2:維修送回HFCMAC序號格式錯誤,HFCMAC序號=<%s>,HFCMAC序號最大長度<%d碼>。',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( aNode.nodeValue ) > XML_HFCMAC_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdHFCMac(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
      2: Result := Format( '-2:驗退HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
      3: Result := Format( '-2:調撥HFCMAC序號格式錯誤,HFCMAC序號=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdEtherNetMac(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫EthernetMAC序號格式錯誤,EthernetMAC序號=<%s>,EthernetMAC序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_ETHERNETMAC_LEN, aId] );
      2: Result := Format( '-2:驗退EthernetMAC序號格式錯誤,EthernetMAC序號=<%s>,EthernetMAC序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_ETHERNETMAC_LEN, aId] );
      3: Result := Format( '-2:調撥EthernetMAC序號格式錯誤,EthernetMAC序號=<%s>,EthernetMAC序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_ETHERNETMAC_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if ( Length( aNode.nodeValue ) > XML_ETHERNETMAC_LEN ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdMTAMac(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫MTAMAC序號格式錯誤,MTAMAC序號=<%s>,MTAMAC序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_MTAMAC_LEN, aId] );
      2: Result := Format( '-2:驗退MTAMAC序號格式錯誤,MTAMAC序號=<%s>,MTAMAC序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_MTAMAC_LEN, aId] );
      3: Result := Format( '-2:調撥MTAMAC序號格式錯誤,MTAMAC序號=<%s>,MTAMAC序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_MTAMAC_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if ( Length( aNode.nodeValue ) > XML_MTAMAC_LEN ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdChassisSN(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫基版序號格式錯誤,基版序號=<%s>,基版序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_CHASSISSN_LEN, aId] );
      2: Result := Format( '-2:驗退基版序號格式錯誤,基版序號=<%s>,基版序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_CHASSISSN_LEN, aId] );
      3: Result := Format( '-2:調撥基版序號格式錯誤,基版序號=<%s>,基版序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_CHASSISSN_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if Length( aNode.nodeValue ) > XML_CHASSISSN_LEN then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdCustomerSN(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫機器序號格式錯誤,機器序號=<%s>,機器序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_CUSTOMERSN_LEN, aId] );
      2: Result := Format( '-2:驗退機器序號格式錯誤,機器序號=<%s>,機器序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_CUSTOMERSN_LEN, aId] );
      3: Result := Format( '-2:調撥機器序號格式錯誤,機器序號=<%s>,機器序號最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_CUSTOMERSN_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if Length( aNode.nodeValue ) > XML_CUSTOMERSN_LEN then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdModelId(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備型式格式錯誤,設備型式=<%s>,HFCMAC序號=<%s>。',
        ['空值', aId] );
      2: Result := Format( '-2:驗退設備型式格式錯誤,設備型式=<%s>,HFCMAC序號=<%s>。',
        ['空值', aId] );
      3: Result := Format( '-2:調撥設備型式格式錯誤,設備型式=<%s>,HFCMAC序號=<%s>。',
        ['空值', aId] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備型式格式錯誤,設備型式=<%s>,設備型式必須符合下列值<%s>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, '0,1,2', aId] );
      2: Result := Format( '-2:驗退設備型式格式錯誤,設備型式=<%s>,設備型式必須符合下列值<%s>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, '0,1,2', aId] );
      3: Result := Format( '-2:調撥設備型式格式錯誤,設備型式=<%s>,設備型式必須符合下列值<%s>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, '0,1,2', aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue <> '0' ) and
     ( aNode.nodeValue <> '1' ) and
     ( aNode.nodeValue <> '2' ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdModelNo(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備機種格式錯誤,設備機種=<%s>,HFCMAC序號=<%s>。',
        ['空值', aId] );
      2: Result := Format( '-2:驗退設備機種格式錯誤,設備機種=<%s>,HFCMAC序號=<%s>。',
        ['空值', aId] );
      3: Result := Format( '-2:調撥設備機種格式錯誤,設備機種=<%s>,HFCMAC序號=<%s>。',
        ['空值', aId] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備機種格式錯誤,機種=<%s>,機種最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_MODELNO_LEN, aId] );
      2: Result := Format( '-2:驗退設備機種格式錯誤,機種=<%s>,機種最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_MODELNO_LEN, aId] );
      3: Result := Format( '-2:調撥設備機種格式錯誤,機種=<%s>,機種最大長度<%d碼>,HFCMAC序號=<%s>。',
          [aNode.nodeValue, XML_MODELNO_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( Length( aNode.nodeValue ) > XML_MODELNO_LEN ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdLimitDate(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:入庫設備保固期限格式錯誤,保固期限=<%s>,HFCMAC序號=<%s>。',
        [aNode.nodeValue, aId] );
      2: Result := Format( '-2:驗退設備保固期限格式錯誤,保固期限=<%s>,HFCMAC序號=<%s>。',
        [aNode.nodeValue, aId] );
      3: Result := Format( '-2:調撥設備保固期限格式錯誤,保固期限=<%s>,HFCMAC序號=<%s>。',
        [aNode.nodeValue, aId] );
    end;
  end;

  { ----------------------------------------------- }

var
  aYear, aMonth, aDay: String;
begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if ( aNode.nodeValue = EmptyWideStr ) then Exit;
  aYear := Copy( aNode.nodeValue, 1, 3 );
  aMonth := Copy( aNode.nodeValue, 4, 2 );
  aDay := Copy( aNode.nodeValue, 6, 2 );
  if ( Length( aYear ) <> 3 ) or
     ( Length( aMonth ) <> 2 ) or
     ( Length( aDay ) <> 2 ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
  try
    aYear := IntToStr( StrToInt( aYear ) + 1911 );
    StrToDate( FormatMaskText( '####/##/##;0;_', aYear + aMonth + aDay ) );
  except
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdGetOrBack(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      4: Result := Format( '-2:領退料識別格式錯誤,識別格式=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      4: Result := Format( '-2:領退料識別格式錯誤,識別格式=<%s>,識別格式必須符合下列值<%s>。',
          [aNode.nodeValue, '1,2'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if not Assigned( aNode ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue = EmptyStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  if ( aNode.nodeValue <> '1' ) and
     ( aNode.nodeValue <> '2' ) then
  begin
    Result := GetErrorDescription2;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdMasterialType(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format(
        '-2:入庫設備物料型號格式錯誤,物料型號=<%s>,物料型號格式最大長度<%d碼>,HFCMAC序號=<%s>。',
        [aNode.nodeValue, XML_MATERIALTYPE_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if ( aNode.nodeValue = EmptyWideStr ) then Exit;
  if ( Length( aNode.nodeValue ) > XML_MATERIALTYPE_LEN ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdHwVer(aNode: IDOMNode; aId: String): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format(
        '-2:入庫設備硬體版本格式錯誤,硬體版本=<%s>,硬體版本格式最大長度<%d碼>,HFCMAC序號=<%s>。',
        [aNode.nodeValue, XML_HWVER_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '空值';
  if not Assigned( aNode ) then Exit;
  if ( aNode.nodeValue = EmptyWideStr ) then Exit;
  if ( Length( aNode.nodeValue ) > XML_HWVER_LEN ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdBackList(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      2: Result := Format( '-2:驗退資料格式錯誤,驗退格式=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdTranList(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      3: Result := Format( '-2:調撥資料格式錯誤,調撥格式=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdGetBackList(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      3: Result := Format( '-2:領退料資料格式錯誤,領退料格式=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }
  
begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdChangeList(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      5: Result := Format( '-2:維修送回資料格式錯誤,維修送回格式=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if aNodes.length <= 0 then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdHfcList(aNodes: IDOMNodeList): String;

  { ----------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      5: Result := Format( '-2:維修送回資料格式錯誤,序號內容=<%s>。', ['空值'] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aNodes.length <= 0 ) then
  begin
    Result := GetErrorDescription;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdPariEqum(aNode1, aNode2: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      5: Result := '-2:維修送回資料格式錯誤,序號內容必須有=<送修>及<領回>。';
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2(aName: String): String;
  begin
    Result := EmptyStr;
    if ( aName = '送修' ) then
      aName := '領回'
    else
      aName := '送修';
    case FType of
      5: Result := Format( '-2:維修送回資料格式錯誤,序號內容無<%s>資料。', [aName] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( not Assigned( aNode1 ) ) and ( not Assigned( aNode2 ) ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end else
  if ( not Assigned( aNode1 ) ) and ( Assigned( aNode2 ) ) then
  begin
    Result := GetErrorDescription2( aNode2.nodeName );
    Exit;
  end else
  if ( Assigned( aNode1 ) ) and ( not Assigned( aNode2 ) ) then
  begin
    Result := GetErrorDescription2( aNode1.nodeName );
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdAlreadyInOtherSo(aCompCode, aId: String): String;
var
  aSoInfo, aOtherSo: PSoInfo;
  aOtherSoId, aCompName: String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:已調撥至公司別=<%s %s>,HFCMAC序號=<%s>不可再次入庫。',
          [aOtherSoId, aCompName, aId] );
      2: Result := Format( '-2:已調撥至公司別=<%s %s>,HFCMAC序號=<%s>不可驗退。',
          [aOtherSoId, aCompName, aId] );
      3: Result := Format( '-2:已調撥至公司別=<%s %s>,HFCMAC序號=<%s>不可再次調撥。',
          [aOtherSoId, aCompName, aId] );
    end;
  end;

  { ----------------------------------------------------- }

begin
  Result := EmptyStr;
  aSoInfo := GetActiveSo( aCompCode );
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    ' SELECT HFCMAC, COMPCODE FROM SO306 ' +
    '  WHERE HFCMAC = :1                 ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aId;
  aSoInfo.Reader.Open;
  if not aSoInfo.Reader.IsEmpty then
  begin
    aOtherSoId := aSoInfo.Reader.FieldByName( 'COMPCODE' ).AsString;
    { 已調撥至其它系統台 }
    if ( aOtherSoId <> aCompCode ) then
    begin
      aOtherSo := GetActiveSo( aOtherSoId );
      aCompName := EmptyStr;
      if Assigned( aOtherSo ) then
        aCompName := aOtherSo.CompName;
      Result := GetErrorDescription;
    end;
  end;
  aSoInfo.Reader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdNotExists(aCompCode, aId: String): String;
var
  aSoInfo: PSoInfo;
  aCompName: String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:公司別<%s %s>此設備不存在,無法入庫,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      2: Result := Format( '-2:公司別<%s %s>此設備不存在,無法註銷/驗退,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      3: Result := Format( '-2:公司別<%s %s>此設備不存在,無法調撥,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      4: Result := Format( '-2:公司別<%s %s>此設備不存在,無法領/退料,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      5: Result := Format( '-2:公司別<%s %s>此設備不存在,無法維修/領回,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
    end;
  end;

  { ----------------------------------------------------- }

begin
  Result := EmptyStr;
  aSoInfo := GetActiveSo( aCompCode );
  aCompName := aSoInfo.CompName;
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    ' SELECT HFCMAC, COMPCODE FROM SO306 ' +
    '  WHERE HFCMAC = :1                 ' +
    '    AND COMPCODE = :2               ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aId;
  aSoInfo.Reader.Parameters.ParamByName( '2' ).Value := aCompCode;
  aSoInfo.Reader.Open;
  if aSoInfo.Reader.IsEmpty then
    Result := GetErrorDescription;
  aSoInfo.Reader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdNotUse(aCompCode, aId: String): String;
var
  aCompName: String;

  { ----------------------------------------------------- }
  
  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:公司別<%s %s>此設備已使用中無法入庫,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      2: Result := Format( '-2:公司別<%s %s>此設備已使用中無法驗退,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      3: Result := Format( '-2:公司別<%s %s>此設備已使用中無法調撥,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aId] );
      4: Result := Format( '-2:公司別<%s %s>此設備已使用中無法領/退料,HFCMAC序號=<%s> 。',
          [aCompCode, aCompName, aId] );
      5: Result := Format( '-2:公司別<%s %s>此設備已使用中無法維修/冷回,HFCMAC序號=<%s> 。',
          [aCompCode, aCompName, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  { 檢核是否 "使用中" ? ( 若公司別為總倉則不須此檢核 )  }
  if aCompCode = XML_WAREHOUSEID then Exit;
  aSoInfo := GetActiveSo( aCompCode );
  aCompName := aSoInfo.CompName;
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    '  SELECT PRDATE, INSTDATE FROM SO004 WHERE FACISNO = :1 ' +
    '   ORDER BY INSTDATE DESC                               ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aId;
  aSoInfo.Reader.Open;
  try
    { 退料檢核 "使用中" 特殊規則 }
    { 存在於 SO004 的客戶設備檔中, 有安裝日期, 且無拆除日期, 視為 "使用中" }
    if ( not aSoInfo.Reader.IsEmpty ) then
    begin
      if ( aSoInfo.Reader.FieldByName( 'INSTDATE' ).AsString <> EmptyStr ) and
         ( aSoInfo.Reader.FieldByName( 'PRDATE' ).AsString = EmptyStr ) then
      begin
        Result := GetErrorDescription;
        Exit;
      end;
      (*
      else
      { 不然的話, 有裝機日, 拆除日也有, 可是裝機日大於拆除日, 也視為 "使用中" }
      if ( aSoInfo.Reader.FieldByName( 'INSTDATE' ).AsString <> EmptyStr ) and
         ( aSoInfo.Reader.FieldByName( 'PRDATE' ).AsString <> EmptyStr ) then
      begin
        if ( aSoInfo.Reader.FieldByName( 'INSTDATE' ).AsDateTime >
          aSoInfo.Reader.FieldByName( 'PRDATE' ).AsDateTime ) then
        begin
          Result := GetErrorDescription;
          Exit;
        end;
      end;  *)
    end;
  finally
    aSoInfo.Reader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdBlackList(aCompCode, aId: String): String;
var
  aDesc, aCompName: String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:公司別<%s %s>此設備已遺失,原因:<%s>,無法入庫,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aDesc, aId] );
      2: Result := Format( '-2:公司別<%s %s>此設備已遺失,原因:<%s>,無法驗退,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aDesc, aId] );
      3: Result := Format( '-2:公司別<%s %s>此設備已遺失,原因:<%s>,無法調撥,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aDesc, aId] );
      4: Result := Format( '-2:公司別<%s %s>此設備已遺失,原因:<%s>,無法領/退料,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aDesc, aId] );
      5: Result := Format( '-2:公司別<%s %s>此設備已遺失,原因:<%s>,無法維修/送回,HFCMAC序號=<%s>。',
          [aCompCode, aCompName, aDesc, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
  aUseFlag: Integer;
begin
  Result := EmptyStr;
  aDesc := EmptyStr;
  { 檢核是否 "已被註銷" ? ( 若公司別為總倉則不須此檢核 )  }
  if aCompCode = XML_WAREHOUSEID then Exit;
  aSoInfo := GetActiveSo( aCompCode );
  aCompName := aSoInfo.CompName;
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    '  SELECT USEFLAG FROM SO306 WHERE HFCMAC = :1  ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aId;
  aSoInfo.Reader.Open;
  try
    { 存在於 SO004 的客戶設備檔中, UseFlag = 2,  視為 "註銷" }
    if ( not aSoInfo.Reader.IsEmpty ) then
    begin
      aUseFlag := aSoInfo.Reader.FieldByName( 'USEFLAG' ).AsInteger;
      aSoInfo.Reader.Close;
      if ( aUseFlag = 2 ) then
      begin
        aSoInfo.Reader.SQL.Clear;
        aSoInfo.Reader.SQL.Text :=
          ' SELECT ( NVL( TO_CHAR( DEREASONCODE ), ''空值'' ) || '','' || NVL( DEREASONNAME, ''空值'' ) ) DEREASON ' +
          '   FROM SO004 WHERE FACISNO = :1 ORDER BY INSTDATE DESC               ';
        aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aId;
        aSoInfo.Reader.Open;
        aDesc := aSoInfo.Reader.FieldByName( 'DEREASON' ).AsString;
        aSoInfo.Reader.Close;
        Result := GetErrorDescription;
        Exit;
      end;
    end;
  finally
    aSoInfo.Reader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdCD043(aCompCode, aModelNo, aId: String; out aModeCode: String): String;
var
  aCompName: String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      3: Result := Format( '-2:公司別<%s %s>此設備查無對應機種型號,機種=<%s>, HFCMAC序號=<%s> 無法調撥。',
          [aCompCode, aCompName, aModelNo, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  aModeCode := EmptyStr;
  Result := EmptyStr;
  { 調入的是總倉, 不檢核 }
  if aCompCode = XML_WAREHOUSEID then Exit;
  aSoInfo := GetActiveSo( aCompCode );
  aCompName := aSoInfo.CompName;
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    ' SELECT CODENO FROM CD043   ' +
    '  WHERE DESCRIPTION = :1    ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aModelNo;
  aSoInfo.Reader.Open;
  try
    if ( aSoInfo.Reader.IsEmpty ) or
       ( aSoInfo.Reader.RecordCount > 1 ) then
    begin
      if ( aModelNo = EmptyStr ) then aModelNo := '空值';
      Result := GetErrorDescription;
      Exit;
    end;
    aModeCode := aSoInfo.Reader.FieldByName( 'CODENO' ).AsString;
  finally
    aSoInfo.Reader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ 調撥時呼叫此函式, 驗證連接的 DB 是否是該系統台的 DB, 預防程式有誤
  寫錯到別家系統台 }

function TCM2CT.VdCD039(aCompCode: String): String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      3: Result := Format( '-1: CC&B調撥資料庫存檔失敗,錯誤訊息:<%s>。',
          ['調撥公司代碼與實際資料庫不符'] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  { 總倉, 不檢核 }
  if aCompCode = XML_WAREHOUSEID then Exit;
  aSoInfo := GetActiveSo( aCompCode );
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    ' SELECT CODENO FROM CD039   ' +
    '  WHERE CODENO = :1         ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := StrToInt( aCompCode );
  aSoInfo.Reader.Open;
  try
    if ( aSoInfo.Reader.IsEmpty ) or
       ( aSoInfo.Reader.RecordCount > 1 ) then
    begin
      Result := GetErrorDescription;
      Exit;
    end;
  finally
    aSoInfo.Reader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdCD022(aCompCode, aDesc, aId: String): String;
var
  aCompName: String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      3: Result := Format( '-2:公司別<%s %s>此設備查無對應客服品名,客服品名=<%s>, HFCMAC序號=<%s> 無法調撥。',
          [aCompCode, aCompName, aDesc, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  { 調入的是總倉, 不檢核 }
  if aCompCode = XML_WAREHOUSEID then Exit;
  aSoInfo := GetActiveSo( aCompCode );
  aCompName := aSoInfo.CompName;
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text :=
    ' SELECT CODENO FROM CD022   ' +
    '  WHERE DESCRIPTION = :1    ';
  aSoInfo.Reader.Parameters.ParamByName( '1' ).Value := aDesc;
  aSoInfo.Reader.Open;
  try
    if ( aSoInfo.Reader.IsEmpty ) or
       ( aSoInfo.Reader.RecordCount > 1 ) then
    begin
      if ( aDesc = EmptyStr ) then aDesc := '空值';
      Result := GetErrorDescription;
      Exit;
    end;
  finally
    aSoInfo.Reader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdSpecialRule1(aSo306: PSo306): String;
begin
  Result := EmptyStr;
  if ( aSo306.ModeId = 0 ) and ( Trim( aSo306.MtaMac ) = EmptyStr )  then
  begin
    Result := Format( '-2:入庫設備MTAMAC序號不可為空值,該設備型式為<EMTA>, HFCMAC序號=<%s>。',
      [aSo306.HfcMac] );
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdExchangeNode(var aSend, aBack: IDOMNode): String;
var
  aTmp: IDOMNode;
begin
  Result := EmptyStr;
  if ( aSend.nodeName = aBack.nodeName ) then
  begin
    Result := '-2:維修送回序號格式錯誤, <維修>及<領回>必須成對出現。';
    Exit;
  end;
  if ( aSend.nodeName = '領回' ) then
  begin
    aTmp := aSend;
    aSend := aBack;
    aBack := aTmp;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdBackEqumExists(aId: String): String;
var
  aSoInfo, aBackSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  aSoInfo := GetActiveSo( XML_WAREHOUSEID );
  aSoInfo.Reader.Close;
  aSoInfo.Reader.SQL.Clear;
  aSoInfo.Reader.SQL.Text := Format(
    ' SELECT COMPCODE, HFCMAC FROM SO306 ' +
    '  WHERE HFCMAC = ''%s''             ', [aId] );
  aSoInfo.Reader.Open;
  if not aSoInfo.Reader.IsEmpty then
  begin
    aBackSoInfo := GetActiveSo( aSoInfo.Reader.FieldByName( 'compcode' ).AsString );
    Result := Format( '-2:公司別<%s %s>維修領回之設備序號已存在, HFCMAC序號=<%s> 不可維修領回。',
      [aBackSoInfo.CompCode, aBackSoInfo.CompName, aId] );
  end;
  aSoInfo.Reader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.VdEqumMtaMac(aHfc, aDbMtaMac, aXmlMtaMac: String; aCompare: Boolean = True): String;
var
  aText: String;
begin
  Result := EmptyStr;
  if ( aDbMtaMac <> EmptyStr ) and ( aXmlMtaMac = EmptyStr ) then
  begin
    Result := Format( '-2:維修領回之設備型式不符, MTAMAC序號不可為<空值>, HFCMAC=<%s>。',
      [aHfc] );
    Exit;
  end;
  if ( aDbMtaMac = EmptyStr ) and ( aXmlMtaMac <> EmptyStr ) then
  begin
    Result := Format( '-2:維修領回之設備型式不符, MTAMAC序號必須為<空值>, HFCMAC=<%s>。',
      [aXmlMtaMac, aHfc] );
    Exit;
  end;
  if ( aCompare ) then
  begin
    if ( Nvl( aDbMtaMac, 'x' ) <> Nvl( aXmlMtaMac, 'x' ) ) then
    begin
      aText := aDbMtaMac;
      if ( aXmlMtaMac <> EmptyStr ) then aText := aXmlMtaMac;
      Result := Format( '-2:維修領回之設備型式不符, MTAMAC序號有誤=<%s>, HFCMAC=<%s>。',
        [aText, aHfc] );
      Exit;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.CM2CT_DLL1(aData1, aData2: OleVariant): WideString;
var
  aXmlData1, aXmlData2: String;
  aExcResult: Boolean;
begin
  Result := EmptyWideStr;
  aXmlData1 := VarToStrDef( aData1, XML_NO_DATA );
  aXmlData2 := VarToStrDef( aData2, XML_NO_DATA );
  { 入庫 }
  if UpperCase( Nvl( Trim( aXmlData1 ), XML_NO_DATA ) ) <> XML_NO_DATA then
  begin
    FType := 1;
    FActiveXmlData := aXmlData1;
    UnPrepareDataBuffer;
    UnPrepareLogBuffer;
    PrepareDataBuffer;
    PrepareLogBuffer;
    InitCallTime;
    try
      { 先驗證入庫格式或資料 }
      aExcResult := ValidateXML1( aXmlData1 );
      if not aExcResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { 驗證OK, 寫入 Table }
      SaveXML1;
      WriteLog;
      Result := MakeErrorText;
    finally
      UnPrepareDataBuffer;
      UnPrepareLogBuffer;
      UnPrepareSoDBConnection;
    end;
    if ( Result <> EmptyWideStr ) then Exit;
  end;
  { 驗退 }
  if UpperCase( Nvl( Trim( aXmlData2 ), XML_NO_DATA ) ) <> XML_NO_DATA then
  begin
    FType := 2;
    FActiveXmlData := aXmlData2;
    UnPrepareDataBuffer;
    UnPrepareLogBuffer;
    PrepareDataBuffer;
    PrepareLogBuffer;
    InitCallTime;
    try
      { 先驗證驗退格式或資料 }
      aExcResult := ValidateXML2( aXmlData2 );
      if not aExcResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { 驗證OK, 寫入 Table }
      SaveXML2;
      WriteLog;
      Result := MakeErrorText;
    finally
      UnPrepareDataBuffer;
      UnPrepareLogBuffer;
      UnPrepareSoDBConnection;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.CM2CT_DLL2(aData1: OleVariant): WideString;
var
  aXmlData1: String;
  aExecResult: Boolean;
begin
  Result := EmptyWideStr;
  aXmlData1 := VarToStrDef( aData1, XML_NO_DATA );
  { 調撥 }
  if UpperCase( Nvl( Trim( aXmlData1 ), XML_NO_DATA ) ) <> XML_NO_DATA then
  begin
    FType := 3;
    FActiveXmlData := aXmlData1;
    UnPrepareDataBuffer;
    UnPrepareLogBuffer;
    PrepareDataBuffer;
    PrepareLogBuffer;
    InitCallTime;
    try
      { 先驗證調撥格式或資料 }
      aExecResult := ValidateXML3( aXmlData1 );
      if ( not aExecResult ) then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { 驗證OK, 寫入 Table }
      SaveXML3;
      WriteLog;
      Result := MakeErrorText;
    finally
      UnPrepareDataBuffer;
      UnPrepareLogBuffer;
      UnPrepareSoDBConnection;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.CM2CT_DLL3(aData1: OleVariant): WideString;
var
  aXmlData1: String;
  aExecResult: Boolean;
begin
  Result := EmptyWideStr;
  aXmlData1 := VarToStrDef( aData1, XML_NO_DATA );
  { 領/退料 }
  if UpperCase( Nvl( Trim( aXmlData1 ), XML_NO_DATA ) ) <> XML_NO_DATA then
  begin
    FType := 4;
    FActiveXmlData := aXmlData1;
    UnPrepareDataBuffer;
    UnPrepareLogBuffer;
    PrepareDataBuffer;
    PrepareLogBuffer;
    InitCallTime;
    try
      { 先驗證領/退料格式或資料 }
      aExecResult := ValidateXML4( aXmlData1 );
      if not aExecResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { 驗證OK, 寫入 Table }
      SaveXML4;
      WriteLog;
      Result := MakeErrorText;
    finally
      UnPrepareDataBuffer;
      UnPrepareLogBuffer;
      UnPrepareSoDBConnection;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.CM2CT_DLL5(aData1: OleVariant): WideString;
var
  aXmlData1: String;
  aExecResult: Boolean;
begin
  Result := EmptyWideStr;
  aXmlData1 := VarToStrDef( aData1, XML_NO_DATA );
  { 領/退料 }
  if UpperCase( Nvl( Trim( aXmlData1 ), XML_NO_DATA ) ) <> XML_NO_DATA then
  begin
    FType := 5;
    FActiveXmlData := aXmlData1;
    UnPrepareDataBuffer;
    UnPrepareLogBuffer;
    PrepareDataBuffer;
    PrepareLogBuffer;
    InitCallTime;
    try
      { 先驗證領/退料格式或資料 }
      aExecResult := ValidateXML5( aXmlData1 );
      if not aExecResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { 驗證OK, 寫入 Table }
      SaveXML5;
      WriteLog;
      Result := MakeErrorText;
    finally
      UnPrepareDataBuffer;
      UnPrepareLogBuffer;
      UnPrepareSoDBConnection;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.ValidateXML1(aXmlData: String): Boolean;
var
  aBatchList, aMaterialList, aEqumList: IDOMNodeList;
  aNode: IDOMNode;
  aIndex, aIndex2, aIndex3: Integer;
  aMsg, aActionText: String;
  aDetailError: Boolean;
begin
  Result := False;
  aDetailError := False;
  aMsg := EmptyStr;
  aActionText := GetActionText;
  try
    FXml.LoadFromXML( aXmlData );
  except
    on E: Exception do aMsg := E.Message;
  end;
  if ( aMsg <> EmptyStr ) then
  begin
    aMsg := Format( '-1:傳入之入庫 XML 格式有誤, 錯誤原因:%s。', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { 批號檢查 }
  aBatchList := FXml.DOMDocument.getElementsByTagName( '批號' );
  aMsg := VdBatchNo( aBatchList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { 逐一取出該批號下的資料 }
  for aIndex := 0 to aBatchList.length - 1 do
  begin
    { 初始 FSo306 填入值結構 }
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    {}
    aMsg := VdBatchNo( aBatchList.item[aIndex].firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.BatchNo := Trim( aBatchList.item[aIndex].firstChild.nodeValue );
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '日期' );
    aMsg := VdDate( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.CMIssue := ConvertCDate( aNode.firstChild.nodeValue );
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '時間' );
    aMsg := VdTime( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.CMIssue := ( FSo306.CMIssue + #32 + ConvertTime( aNode.firstChild.nodeValue ) );
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '廠商代碼' );
    aMsg := VdVendorCode( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.VendorCode := aNode.firstChild.nodeValue;
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '廠商名稱' );
    aMsg := VdVendorName( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;  
    FSo306.VendorName := aNode.firstChild.nodeValue;
    {}
    { 入庫一定是總倉, 總倉代碼為 0 }
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '公司代碼' );
    aMsg := VdCompCode( aNode.firstChild, '0' );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;  
    FSo306.CompCode := aNode.firstChild.nodeValue;
    {}
    aMaterialList := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNodes( '料號' );
    aMsg := VdMaterialNo( aMaterialList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { 逐一取出該批號下所有料號資料 }
    for aIndex2 := 0 to aMaterialList.length - 1 do
    begin
      aMsg := VdMaterialNo( aMaterialList.item[aIndex2].firstChild );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;  
      FSo306.MaterialNo := Trim( aMaterialList.item[aIndex2].firstChild.nodeValue );
      {}
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '數量' );
      aMsg := VdItemCount( aNode, aMaterialList.item[aIndex2], True );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;  
      { 品名暫時沒用到, 有給值先照抓, 但是不寫 Table }
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '品名' );
      aMsg := VdDescription( aNode );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      FSo306.Spec := EmptyStr;
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '規格' );
      if Assigned( aNode ) then
      begin
        aMsg := VdSpec( aNode );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          Exit;
        end;
        FSo306.Spec := aNode.nodeValue;
      end;
      {}
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '列帳單位' );
      if Assigned( aNode ) then
      begin
        aMsg := VdBelongTo( aNode );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          Exit;
        end;
        FSo306.Belong := StrToIntDef( aNode.nodeValue, 0 );
      end;
      {}
      aEqumList := ( aMaterialList.item[aIndex2] as IDOMNodeSelect ).selectNodes( '設備' );
      aMsg := VdEqum( aEqumList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      { 取出所有設備 }
      for aIndex3 := 0 to aEqumList.length - 1 do
      begin
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( 'HFCMAC序號' );
        FSo306.HfcMac :=  EmptyStr;
        aMsg := VdHFCMac( aNode.firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.HfcMac := aNode.firstChild.nodeValue;
        {}
        aNode := aEqumList.item[aIndex3].attributes.getNamedItem( '主序號' );
        aMsg := VdPrimaryMAC( aNode, FSo306.HfcMac );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.Flag := StrToInt( aNode.nodeValue );
        {}
        FSo306.EthernetMac := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( 'ETHERNETMAC序號' );
        if Assigned( aNode ) then
        begin
          aMsg := VdEtherNetMac( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
            FSo306.EthernetMac := aNode.firstChild.nodeValue;
        end;
        {}
        FSo306.MtaMac := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( 'MTAMAC序號' );
        if Assigned( aNode ) then
        begin
          aMsg := VdMTAMac( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
            FSo306.MtaMac := aNode.firstChild.nodeValue;
        end;
        {}
        FSo306.ChassisSN := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '基版序號' );
        if Assigned( aNode ) then
        begin
          aMsg := VdChassisSN( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
            FSo306.ChassisSN := aNode.firstChild.nodeValue;
        end;
        {}
        FSo306.CustomerSN := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '機器序號' );
        if Assigned( aNode ) then
        begin
          aMsg := VdCustomerSN( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
            FSo306.CustomerSN := aNode.firstChild.nodeValue;
        end;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '機種' );
        aMsg := VdModelNo( aNode.firstChild, FSo306.HfcMac );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.ModelNo := aNode.firstChild.nodeValue;
        {}
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '型式' );
        aMsg := VdModelId( aNode.firstChild, FSo306.HfcMac );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.ModeId := StrToInt( aNode.firstChild.nodeValue );
        { 保固期限 }
        FSo306.LimitDate := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '保固期限' );
        if Assigned( aNode ) then
        begin
          aMsg := VdLimitDate( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
          begin
            if Length( aNode.firstChild.nodeValue ) > 0 then
              FSo306.LimitDate := ConvertCDate( aNode.firstChild.nodeValue );
          end;
        end;
        { 客服品名 }
        FSo306.Description := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '客服品名' );
        aMsg := VdDescription2( aNode.firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.Description := aNode.firstChild.nodeValue;
        { 物料型號 }
        FSo306.MaterialType := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '物料型號' );
        if Assigned( aNode ) then
        begin
          aMsg := VdMasterialType( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
          begin
            if Assigned( aNode.firstChild ) then
              FSo306.MaterialType := aNode.firstChild.nodeValue;
          end;
        end;
        { 硬體版本 }
        FSo306.HwVer := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '硬體版本' );
        if Assigned( aNode ) then
        begin
          aMsg := VdHwVer( aNode.firstChild, FSo306.HfcMac );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end else
          begin
            if Assigned( aNode.firstChild )  then
              FSo306.HwVer := aNode.firstChild.nodeValue;
          end;
        end;
        { 前面都沒發生錯誤 }
        if not aDetailError then
        begin
          aMsg := VdSpecialRule1( FSo306 );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end;
        end;
        { 只要有任何錯誤, 一律不存檔 }
        if not aDetailError then
        begin
          FDataBuffer.Append;
          FDataBuffer.FieldByName( 'COMPCODE' ).Value := FSo306.CompCode ;      { 公司別 }
          FDataBuffer.FieldByName( 'BATCHNO' ).Value := FSo306.BatchNo;         { 批號 }
          FDataBuffer.FieldByName( 'MATERIALNO' ).Value := FSo306.MaterialNo;   { 料號 }
          FDataBuffer.FieldByName( 'CARTONNUNO' ).Value := Null;                { 紙箱序號 }
          FDataBuffer.FieldByName( 'MODELNO' ).Value := FSo306.ModelNo;         { 機種 }
          FDataBuffer.FieldByName( 'CHASSISSN' ).Value := FSo306.ChassisSN;     { 基版序號 }
          FDataBuffer.FieldByName( 'CUSTOMERSN' ).Value := FSo306.CustomerSN;   { 機器序號 }
          FDataBuffer.FieldByName( 'HFCMAC' ).Value := FSo306.HfcMac;           { HFCMAC 序號 }
          FDataBuffer.FieldByName( 'ETHERNETMAC' ).Value := FSo306.EthernetMac; { Ethernet MAC 序號 }
          FDataBuffer.FieldByName( 'MTAMAC' ).Value := FSo306.MtaMac;           { MTAMAC 序號 }
          FDataBuffer.FieldByName( 'FLAG' ).Value := FSo306.Flag;               { 以那一個 MAC 序號為主, 預設 1 }
          FDataBuffer.FieldByName( 'CMISSUE' ).Value := FSo306.CMIssue;         { CM 匯入日期, 執行此 DLL 的日期, 年月日時分秒 }
          FDataBuffer.FieldByName( 'MODEID' ).Value := FSo306.ModeId;           { 設備型號, 0=EMTA, 1=CM, 2=WIFLY }
          FDataBuffer.FieldByName( 'MODELCODE' ).Value := Null;                 { 設備型號代碼 }
          FDataBuffer.FieldByName( 'ACTSTATUS' ).Value := EmptyStr;             { SPM 設定狀態 }
          FDataBuffer.FieldByName( 'USEFLAG' ).Value := Null;                   { 可用狀態, 0=可用, 1=不可用 }
          FDataBuffer.FieldByName( 'LIMITDATE' ).Value := FSo306.LimitDate;     { 保固期限 , 年月日 }
          FDataBuffer.FieldByName( 'BELONG' ).Value := FSo306.Belong;           { 財產規屬單位, 0=EMC, 1=APTG }
          FDataBuffer.FieldByName( 'VENDORCODE' ).Value := FSo306.VendorCode;   { 廠商代碼 }
          FDataBuffer.FieldByName( 'VENDORNAME' ).Value := FSo306.VendorName;   { 廠商名稱 }
          FDataBuffer.FieldByName( 'SPEC' ).Value := FSo306.Spec;               { 規格 }
          FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString := FSo306.Description; { 客服品名 }
          FDataBuffer.FieldByName( 'UPDEN' ).Value := '物料系統';               { 異動人員 }
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;              { 異動時間 }
          FDataBuffer.FieldByName( 'MATERIALTYPE' ).AsString := FSo306.MaterialType; { 物料型號 }
          FDataBuffer.FieldByName( 'HWVER' ).AsString := FSo306.HwVer;          { 硬體版本 }
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC序號=<%s>,正常。', [FSo306.HfcMac] );
          AppendLog( '0', aActionText, aMsg, 0 );
        end;
      end;
    end;
  end;
  Result := not aDetailError;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.SaveXML1: Boolean;
var
  aSoInfo: PSoInfo;

  { ------------------------------------------------------- }

  procedure UpdateRecord;
  begin
    aSoInfo.Writer.Close;
    aSoInfo.Writer.SQL.Clear;
    aSoInfo.Writer.Parameters.Clear;
    aSoInfo.Writer.SQL.Text :=
      ' UPDATE SO306                  ' +
      '    SET COMPCODE = :1,         ' +
      '        BATCHNO = :2,          ' +
      '        MATERIALNO = :3,       ' +
      '        MODELNO = :4,          ' +
      '        CHASSISSN = :5,        ' +
      '        CUSTOMERSN = :6,       ' +
      '        ETHERNETMAC = :7,      ' +
      '        MTAMAC = :8,           ' +
      '        FLAG = :9,             ' +
      '        CMISSUE = TO_DATE( :10, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '        MODEID = :11,          ' +
      '        MODELCODE = :12,       ' +
      '        ACTSTATUS = :13,       ' +
      '        USEFLAG = :14,         ' +
      '        LIMITDATE = TO_DATE( :15, ''YYYY/MM/DD'' ), ' +
      '        BELONG = :16,          ' +
      '        VENDORCODE = :17,      ' +
      '        VENDORNAME = :18,      ' +
      '        SPEC = :19,            ' +
      '        DESCRIPTION = :20,     ' +
      '        UPDEN = :21,           ' +
      '        UPDTIME = TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '        MATERIALMODE = :22,    ' +
      '        HARDWAREVER = :23      ' +
      '  WHERE HFCMAC = :24           ';
    aSoInfo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '2' ).Value :=
     FDataBuffer.FieldByName( 'BATCHNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '3' ).Value :=
     FDataBuffer.FieldByName( 'MATERIALNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '4' ).Value :=
     FDataBuffer.FieldByName( 'MODELNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '5' ).Value :=
     FDataBuffer.FieldByName( 'CHASSISSN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '6' ).Value :=
     FDataBuffer.FieldByName( 'CUSTOMERSN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '7' ).Value :=
     FDataBuffer.FieldByName( 'ETHERNETMAC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '8' ).Value :=
     FDataBuffer.FieldByName( 'MTAMAC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '9' ).Value :=
     FDataBuffer.FieldByName( 'FLAG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '10' ).Value :=
     FDataBuffer.FieldByName( 'CMISSUE' ).Value;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '11' ).Value :=
     FDataBuffer.FieldByName( 'MODEID' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '12' ).Value :=
     FDataBuffer.FieldByName( 'MODELCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '13' ).Value :=
     FDataBuffer.FieldByName( 'ACTSTATUS' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '14' ).Value :=
     FDataBuffer.FieldByName( 'USEFLAG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '15' ).Value :=
     FDataBuffer.FieldByName( 'LIMITDATE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '16' ).Value :=
     FDataBuffer.FieldByName( 'BELONG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '17' ).Value :=
     FDataBuffer.FieldByName( 'VENDORCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '18' ).Value :=
     FDataBuffer.FieldByName( 'VENDORNAME' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '19' ).Value :=
      FDataBuffer.FieldByName( 'SPEC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '20' ).Value :=
      FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '21' ).Value :=
      FDataBuffer.FieldByName( 'UPDEN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '22' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALTYPE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '23' ).Value :=
      FDataBuffer.FieldByName( 'HWVER' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '24' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    {}
    aSoInfo.Writer.ExecSQL;
  end;

  { ------------------------------------------------------- }

  procedure InsertRecord;
  begin
    aSoInfo.Writer.Close;
    aSoInfo.Writer.SQL.Clear;
    aSoInfo.Writer.Parameters.Clear;
    aSoInfo.Writer.SQL.Text :=
      ' INSERT INTO SO306 (   ' +
      '    COMPCODE,          ' +
      '    BATCHNO,           ' +
      '    MATERIALNO,        ' +
      '    CARTONNUNO,        ' +
      '    MODELNO,           ' +
      '    CHASSISSN,         ' +
      '    CUSTOMERSN,        ' +
      '    HFCMAC,            ' +
      '    ETHERNETMAC,       ' +
      '    MTAMAC,            ' +
      '    FLAG,              ' +
      '    CMISSUE,           ' +
      '    MODEID,            ' +
      '    MODELCODE,         ' +
      '    ACTSTATUS,         ' +
      '    USEFLAG,           ' +
      '    LIMITDATE,         ' +
      '    BELONG,            ' +
      '    VENDORCODE,        ' +
      '    VENDORNAME,        ' +
      '    SPEC,              ' +
      '    DESCRIPTION,       ' +
      '    UPDEN,             ' +
      '    UPDTIME,           ' +
      '    MATERIALMODE,      ' +
      '    HARDWAREVER    )   ' +
      ' VALUES (              ' +
      '    :1,                ' +
      '    :2,                ' +
      '    :3,                ' +
      '    :4,                ' +
      '    :5,                ' +
      '    :6,                ' +
      '    :7,                ' +
      '    :8,                ' +
      '    :9,                ' +
      '    :10,               ' +
      '    :11,               ' +
      '    TO_DATE( :12, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '    :13,               ' +
      '    :14,               ' +
      '    :15,               ' +
      '    :16,               ' +
      '    TO_DATE( :17, ''YYYY/MM/DD'' ), ' +
      '    :18,               ' +
      '    :19,               ' +
      '    :20,               ' +
      '    :21,               ' +
      '    :22,               ' +
      '    :23,               ' +
      '    TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ),' +
      '    :24,               ' +
      '    :25 )              ';
    aSoInfo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '2' ).Value :=
      FDataBuffer.FieldByName( 'BATCHNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '3' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '4' ).Value :=
      FDataBuffer.FieldByName( 'CARTONNUNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '5' ).Value :=
      FDataBuffer.FieldByName( 'MODELNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '6' ).Value :=
      FDataBuffer.FieldByName( 'CHASSISSN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '7' ).Value :=
      FDataBuffer.FieldByName( 'CUSTOMERSN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '8' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '9' ).Value :=
      FDataBuffer.FieldByName( 'ETHERNETMAC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '10' ).Value :=
      FDataBuffer.FieldByName( 'MTAMAC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '11' ).Value :=
      FDataBuffer.FieldByName( 'FLAG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '12' ).Value :=
      FDataBuffer.FieldByName( 'CMISSUE' ).Value;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '13' ).Value :=
      FDataBuffer.FieldByName( 'MODEID' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '14' ).Value :=
      FDataBuffer.FieldByName( 'MODELCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '15' ).Value :=
      FDataBuffer.FieldByName( 'ACTSTATUS' ).Value;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '16' ).Value :=
      FDataBuffer.FieldByName( 'USEFLAG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '17' ).Value :=
      FDataBuffer.FieldByName( 'LIMITDATE' ).Value;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '18' ).Value :=
      FDataBuffer.FieldByName( 'BELONG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '19' ).Value :=
      FDataBuffer.FieldByName( 'VENDORCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '20' ).Value :=
      FDataBuffer.FieldByName( 'VENDORNAME' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '21' ).Value :=
      FDataBuffer.FieldByName( 'SPEC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '22' ).Value :=
      FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '23' ).Value :=
      FDataBuffer.FieldByName( 'UPDEN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '24' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALTYPE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '25' ).Value :=
      FDataBuffer.FieldByName( 'HWVER' ).AsString;
    {}
    aSoInfo.Writer.ExecSQL;
  end;


  { ------------------------------------------------------- }

var
  aMsg, aActionText: String;
  aIsErr: Boolean;
begin
  Result := False;
  aActionText := GetActionText( 1 );
  { 資料庫連結至總倉 }
  aMsg := PrepareSoDBConnection( XML_WAREHOUSEID );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( XML_WAREHOUSEID, aActionText, aMsg, 1 );
    Exit;
  end;
  aMsg := EmptyStr;
  aSoInfo := GetActiveSo( '0' );
  aSoInfo.Connection.BeginTrans;
  try
    aIsErr := False;
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      { 該筆 CM 是否已調撥至其它系統台, 只檢查調撥即可,
        不須檢核 SO004 是否使用中 }
      aMsg := VdAlreadyInOtherSo( '0', FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        aIsErr := True;
        AppendLog( '0', aActionText, aMsg, 1 );
      end else
      begin
        aMsg := Format( 'HFCMAC序號<%s>,正常。', [
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
        AppendLog( '0', aActionText, aMsg, 0 );
      end;
      if ( not aIsErr ) then
      begin
        { 先更新總倉 Table, 更新不到用新增 }
        UpdateRecord;
        if aSoInfo.Writer.RowsAffected <= 0 then
        begin
          InsertRecord;
        end;
      end;
      FDataBuffer.Next;
    end;
    if aIsErr then
      aSoInfo.Connection.RollbackTrans
    else begin
      aSoInfo.Connection.CommitTrans;
      Result := True;
    end;  
  except
    on E: Exception do
    begin
      aSoInfo.Connection.RollbackTrans;
      aMsg := Format( '-1: CC&B入庫資料庫存檔失敗,錯誤訊息:<%s>', [E.Message] );
      AppendLog( '0', aActionText, aMsg, 1 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.ValidateXML2(aXmlData: String): Boolean;
var
  aBackList, aEqumList, aSeqList: IDOMNodeList;
  aNode: IDOMNode;
  aIndex, aIndex2, aIndex3: Integer;
  aMsg, aActionText: String;
  aDetailError: Boolean;
begin
  Result := False;
  aMsg := EmptyStr;
  aDetailError := False;
  aActionText := GetActionText;
  try
    FXml.LoadFromXML( aXmlData );
  except
    on E: Exception do aMsg := E.Message;
  end;
  if ( aMsg <> EmptyStr ) then
  begin
    aMsg := Format( '-1:傳入之驗退 XML 格式有誤, 錯誤原因:%s。', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  aBackList := FXml.DOMDocument.getElementsByTagName( '驗退' );
  aMsg := VdBackList( aBackList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { 逐一取出驗退單號 }
  for aIndex := 0 to aBackList.length - 1 do
  begin
    { 初始 FSo310 填入值結構 }
    ZeroMemory( FSo310, SizeOf( FSo310^ ) );
    {}
    aNode := aBackList.item[aIndex].attributes.getNamedItem( '單號' );
    aMsg := VdBatchNo( aNode );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;  
    FSo310.BatchNo := aNode.nodeValue;
    {}
    aNode := aBackList.item[aIndex].attributes.getNamedItem( '公司代碼' );
    aMsg := VdCompCode( aNode, '0' );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo310.CompCode := aNode.nodeValue;
    {}
    aEqumList := ( aBackList.item[aIndex] as IDOMNodeSelect ).selectNodes( '設備' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { 逐一取出設備 }
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '數量' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      {} 
      aSeqList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( 'HFCMAC序號' );
      aMsg := VdHFCMac( aSeqList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      { 逐一取出序號 }
      for aIndex3 := 0 to aSeqList.length - 1 do
      begin
        aMsg := VdHFCMac( aSeqList.item[aIndex3].firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo310.HfcMac := aSeqList.item[aIndex3].firstChild.nodeValue;
        if not aDetailError then
        begin
          FDataBuffer.Append;
          FDataBuffer.FieldByName( 'COMPCODE' ).AsString := FSo310.CompCode;
          FDataBuffer.FieldByName( 'BATCHNO' ).AsString := FSo310.BatchNo;
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString := FSo310.HfcMac;
          FDataBuffer.FieldByName( 'UPDEN' ).AsString := '物料系統';
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC序號=<%s>,正常。', [FSo310.HfcMac] );
          AppendLog( '0', aActionText, aMsg, 0 );
        end;
      end;
    end;
  end;
  Result := not aDetailError;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.SaveXML2: Boolean;
var
  aSoInfo: PSoInfo;

  { ------------------------------------------------------- }

  procedure UpdateBufferRecord;
  begin
    aSoInfo.Reader.Close;
    aSoInfo.Reader.SQL.Clear;
    aSoInfo.Reader.SQL.Text :=
      ' SELECT COMPCODE, BATCHNO, MATERIALNO,                     ' +
      '        CARTONNUNO, CHASSISSN, CUSTOMERSN,                 ' +
      '        ETHERNETMAC, MTAMAC, FLAG,                         ' +
      '        TO_CHAR( CMISSUE, ''YYYY/MM/DD HH24:MI:SS'' ) AS CMISSUE,  ' +
      '        MODELNO, MODEID, MODELCODE, ACTSTATUS,             ' +
      '        TO_CHAR( LIMITDATE, ''YYYY/MM/DD'' ) AS LIMITDATE, ' +
      '        BELONG, VENDORCODE, VENDORNAME, SPEC,              ' +
      '        ALREADYUSE, MATERIALMODE, HARDWAREVER              ' +
      '   FROM SO306                                              ' +
      '  WHERE HFCMAC = :1                                        ';
     aSoInfo.Reader.Parameters.ParamByName( '1' ).Value :=
       FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
     aSoInfo.Reader.Open;
     FDataBuffer.Edit;
     FDataBuffer.FieldByName( 'MATERIALNO' ).Value :=
       aSoInfo.Reader.FieldByName( 'MATERIALNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'OLDBATCHNO' ).Value :=
       aSoInfo.Reader.FieldByName( 'BATCHNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'CARTONNUNO' ).Value :=
       aSoInfo.Reader.FieldByName( 'CARTONNUNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'CHASSISSN' ).Value :=
       aSoInfo.Reader.FieldByName( 'CHASSISSN' ).Value;
     {}
     FDataBuffer.FieldByName( 'CUSTOMERSN' ).Value :=
       aSoInfo.Reader.FieldByName( 'CUSTOMERSN' ).Value;
     {}
     FDataBuffer.FieldByName( 'ETHERNETMAC' ).Value :=
       aSoInfo.Reader.FieldByName( 'ETHERNETMAC' ).Value;
     {}
     FDataBuffer.FieldByName( 'MTAMAC' ).Value :=
       aSoInfo.Reader.FieldByName( 'MTAMAC' ).Value;
     {}
     FDataBuffer.FieldByName( 'FLAG' ).Value :=
       aSoInfo.Reader.FieldByName( 'FLAG' ).Value;
     {}
     FDataBuffer.FieldByName( 'CMISSUE' ).Value :=
       aSoInfo.Reader.FieldByName( 'CMISSUE' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODELNO' ).Value :=
       aSoInfo.Reader.FieldByName( 'MODELNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODEID' ).Value :=
       aSoInfo.Reader.FieldByName( 'MODEID' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODELCODE' ).Value :=
       aSoInfo.Reader.FieldByName( 'MODELCODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'ACTSTATUS' ).Value :=
       aSoInfo.Reader.FieldByName( 'ACTSTATUS' ).Value;
     {}
     FDataBuffer.FieldByName( 'LIMITDATE' ).Value :=
       aSoInfo.Reader.FieldByName( 'LIMITDATE' ).Value;
     {}
     FDataBuffer.FieldByName( 'BELONG' ).Value :=
       aSoInfo.Reader.FieldByName( 'BELONG' ).Value;
     {}
     FDataBuffer.FieldByName( 'VENDORCODE' ).Value :=
       aSoInfo.Reader.FieldByName( 'VENDORCODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'VENDORNAME' ).Value :=
       aSoInfo.Reader.FieldByName( 'VENDORNAME' ).Value;
     {}
     FDataBuffer.FieldByName( 'SPEC' ).Value :=
       aSoInfo.Reader.FieldByName( 'SPEC' ).Value;
     {}
     FDataBuffer.FieldByName( 'MATERIALTYPE' ).Value :=
       aSoInfo.Reader.FieldByName( 'MATERIALMODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'HWVER' ).Value :=
       aSoInfo.Reader.FieldByName( 'HARDWAREVER' ).Value;
     {}
     FDataBuffer.FieldByName( 'ALREADYUSE' ).Value :=
       aSoInfo.Reader.FieldByName( 'ALREADYUSE' ).Value;
     {}
     FDataBuffer.Post;
     aSoInfo.Reader.Close;
  end;

  { ------------------------------------------------------- }

  procedure InsertRecord;
  begin
    aSoInfo.Writer.Close;
    aSoInfo.Writer.SQL.Clear;
    aSoInfo.Writer.SQL.Text :=
      '  INSERT INTO SO310 (    ' +
      '     COMPCODE,           ' +
      '     BATCHNO,            ' +
      '     MATERIALNO,         ' +
      '     OLDBATCHNO,         ' +
      '     CARTONNUNO,         ' +
      '     CHASSISSN,          ' +
      '     CUSTOMERSN,         ' +
      '     HFCMAC,             ' +
      '     ETHERNETMAC,        ' +
      '     MTAMAC,             ' +
      '     FLAG,               ' +
      '     CMISSUE,            ' +
      '     MODELNO,            ' +
      '     MODEID,             ' +
      '     MODELCODE,          ' +
      '     ACTSTATUS,          ' +
      '     LIMITDATE,          ' +
      '     BELONG,             ' +
      '     VENDORCODE,         ' +
      '     VENDORNAME,         ' +
      '     SPEC,               ' +
      '     UPDEN,              ' +
      '     UPDTIME,            ' +
      '     ALREADYUSE,         ' +
      '     MATERIALMODE,       ' +
      '     HARDWAREVER   )     ' +
      ' VALUES (                ' +
      '     :1,                 ' +
      '     :2,                 ' +
      '     :3,                 ' +
      '     :4,                 ' +
      '     :5,                 ' +
      '     :6,                 ' +
      '     :7,                 ' +
      '     :8,                 ' +
      '     :9,                 ' +
      '     :10,                ' +
      '     :11,                ' +
      '     TO_DATE( :12, ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '     :13,                ' +
      '     :14,                ' +
      '     :15,                ' +
      '     :16,                ' +
      '     TO_DATE( :17, ''YYYY/MM/DD'' ), ' +
      '     :18,                ' +
      '     :19,                ' +
      '     :20,                ' +
      '     :21,                ' +
      '     :22,                ' +
      '     TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '     TO_DATE( :23, ''YYYY/MM/DD'' ),                ' +
      '     :24,                ' +
      '     :25   )             ';
    aSoInfo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '2' ).Value :=
      FDataBuffer.FieldByName( 'BATCHNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '3' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '4' ).Value :=
      FDataBuffer.FieldByName( 'OLDBATCHNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '5' ).Value :=
      FDataBuffer.FieldByName( 'CARTONNUNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '6' ).Value :=
      FDataBuffer.FieldByName( 'CHASSISSN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '7' ).Value :=
      FDataBuffer.FieldByName( 'CUSTOMERSN' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '8' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '9' ).Value :=
      FDataBuffer.FieldByName( 'ETHERNETMAC' ).AsString;
    {}  
    aSoInfo.Writer.Parameters.ParamByName( '10' ).Value :=
      FDataBuffer.FieldByName( 'MTAMAC' ).AsString;
    {}  
    aSoInfo.Writer.Parameters.ParamByName( '11' ).Value :=
      FDataBuffer.FieldByName( 'FLAG' ).AsString;
    {}  
    aSoInfo.Writer.Parameters.ParamByName( '12' ).Value :=
      FDataBuffer.FieldByName( 'CMISSUE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '13' ).Value :=
      FDataBuffer.FieldByName( 'MODELNO' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '14' ).Value :=
      FDataBuffer.FieldByName( 'MODEID' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '15' ).Value :=
      FDataBuffer.FieldByName( 'MODELCODE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '16' ).Value :=
      FDataBuffer.FieldByName( 'ACTSTATUS' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '17' ).Value :=
      FDataBuffer.FieldByName( 'LIMITDATE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '18' ).Value :=
      FDataBuffer.FieldByName( 'BELONG' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '19' ).Value :=
      FDataBuffer.FieldByName( 'VENDORCODE' ).AsString;
    {}  
    aSoInfo.Writer.Parameters.ParamByName( '20' ).Value :=
      FDataBuffer.FieldByName( 'VENDORNAME' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '21' ).Value :=
      FDataBuffer.FieldByName( 'SPEC' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '22' ).Value :=
      FDataBuffer.FieldByName( 'UPDEN' ).AsString;
    {}  
    aSoInfo.Writer.Parameters.ParamByName( '23' ).Value :=
      FDataBuffer.FieldByName( 'ALREADYUSE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '24' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALTYPE' ).AsString;
    {}
    aSoInfo.Writer.Parameters.ParamByName( '25' ).Value :=
      FDataBuffer.FieldByName( 'HWVER' ).AsString;
    {}  
    aSoInfo.Writer.ExecSQL;
  end;

  { ------------------------------------------------------- }

  procedure DeleteRecord;
  begin
    aSoInfo.Writer.Close;
    aSoInfo.Writer.SQL.Clear;
    aSoInfo.Writer.SQL.Text :=
      ' DELETE FROM SO306 WHERE HFCMAC = :1 ';
    aSoInfo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    aSoInfo.Writer.ExecSQL;  
  end;

  { ------------------------------------------------------- }

  procedure InsertLog(aValue: String);
  begin
    aSoInfo.Writer.Close;
    aSoInfo.Writer.SQL.Clear;
    aSoInfo.Writer.SQL.Text :=
      ' INSERT INTO SOAC0204 (   ' +
      '    CALLTIME,             ' +
      '    COMPCODE,             ' +
      '    PROGNAME,             ' +
      '    PARA     )            ' +
      ' VALUES (                 ' +
      '    TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '    :1,                   ' +
      '    ''CM2CT_DLL1'',       ' +
      '    :2       )            ';
    aSoInfo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    aSoInfo.Writer.Parameters.ParamByName( '2' ).Value := aValue;
    aSoInfo.Writer.ExecSQL;
  end;

  { ------------------------------------------------------- }

var
  aMsg, aActionText, aBatchNo: String;
  aCount: Integer;
  aIsErr, aTotalErr: Boolean;
begin
  Result := False;
  aActionText := GetActionText( 1 );
  { 資料庫連結至總倉 }
  aMsg := PrepareSoDBConnection( XML_WAREHOUSEID );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( XML_WAREHOUSEID, aActionText, aMsg, 1 );
    Exit;
  end;
  aTotalErr := False;
  aMsg := EmptyStr;
  aSoInfo := GetActiveSo( XML_WAREHOUSEID );
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    aIsErr := False;
    { 該筆 CM 是否已調撥至其它系統台, 只檢查調撥即可,
      不須檢核 SO004 是否使用中 }
    aMsg := VdAlreadyInOtherSo( XML_WAREHOUSEID,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end else
    begin
      { 該筆 CM 是否不在總倉中 }
      aMsg := VdNotExists( XML_WAREHOUSEID,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
    end;
    { 將原本該設備資料讀出來, 只要發生過錯誤就不做 }
    if not aTotalErr then
    begin
      UpdateBufferRecord;
    end;
    { 該筆資料沒問題 }
    if not aIsErr then
    begin
      aMsg := Format( 'HFCMAC序號<%s>,正常。', [
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      AppendLog( '0', aActionText, aMsg, 0 );
    end;
    FDataBuffer.Next;
  end;
  if not aTotalErr then
  begin
    aSoInfo.Connection.BeginTrans;
    try
      FDataBuffer.First;
      while not FDataBuffer.Eof do
      begin
        { 寫入 SO310, CM 設備驗退檔 }
        InsertRecord;
        { 刪除原本的  CM 設備檔 }
        DeleteRecord;
        FDataBuffer.Next;
      end;
      { 寫 Log }
      FDataBuffer.First;
      aCount := 0;
      aBatchNo := FDataBuffer.FieldByName( 'BATCHNO' ).AsString;
      while not FDataBuffer.Eof do
      begin
        Inc( aCount );
        FDataBuffer.Next;
        if ( FDataBuffer.FieldByName( 'BATCHNO' ).AsString <> aBatchNo ) or
           ( FDataBuffer.Eof ) then
        begin
          InsertLog( Format( '驗退單號:%s,設備別:CM,公司代碼:%s,數量:%d。',
            [FDataBuffer.FieldByName( 'BATCHNO' ).AsString,
             FDataBuffer.FieldByName( 'COMPCODE' ).AsString, aCount] ) );
          aBatchNo := FDataBuffer.FieldByName( 'BATCHNO' ).AsString;
          aCount := 0;
        end;
      end;
      aSoInfo.Connection.CommitTrans;
      Result := True;
    except
      on E: Exception do
      begin
        aSoInfo.Connection.RollbackTrans;
        aMsg := Format( '-1: CC&B驗退資料庫存檔失敗,錯誤訊息:<%s>', [E.Message] );
        AppendLog( '0', aActionText, aMsg, 1 );
      end;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.ValidateXML3(aXmlData: String): Boolean;
var
  aTransList, aEqumList, aSeqList: IDOMNodeList;
  aNode: IDOMNode;
  aIndex, aIndex2, aIndex3: Integer;
  aMsg, aActionText, aCompCodeText: String;
  aDetailErr: Boolean;
begin
  Result := False;
  aMsg := EmptyStr;
  aDetailErr := False;
  aActionText := GetActionText;
  try
    FXml.LoadFromXML( aXmlData );
  except
    on E: Exception do aMsg := E.Message;
  end;
  if ( aMsg <> EmptyStr ) then
  begin
    aMsg := Format( '-1:傳入之調撥 XML 格式有誤, 錯誤原因:%s。', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  {}
  aTransList := FXml.DOMDocument.getElementsByTagName( '調撥' );
  aMsg := VdTranList( aTransList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( EmptyStr, aActionText, aMsg, 1 );
    Exit;
  end;  
  { 逐一取出調撥資料 }
  for aIndex := 0 to aTransList.length - 1 do
  begin
    { 初始 FSo306 填入值結構 }
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    aEqumList := ( aTransList.item[aIndex] as IDOMNodeSelect ).selectNodes( '設備' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { 逐一取出設備 }
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aCompCodeText := GetChangeCompanyCode( aEqumList.item[aIndex2] );
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '數量' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '調入公司' );
      aMsg := VdCompCode( aNode, EmptyStr );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      FSo306.CompCode := aNode.nodeValue;
      {}
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '調出公司' );
      aMsg := VdCompCode( aNode, EmptyStr );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      FSo306.OldCompCode := aNode.nodeValue;
      {}
      aSeqList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( 'HFCMAC序號' );
      aMsg := VdHFCMac( aSeqList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      { 逐一取出序號 }
      for aIndex3 := 0 to aSeqList.length - 1 do
      begin
        aMsg := VdHFCMac( aSeqList.item[aIndex3].firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDetailErr := True;
        end else
          FSo306.HfcMac := aSeqList.item[aIndex3].firstChild.nodeValue;
        if not aDetailErr then
        begin
          FDataBuffer.Append;
          FDataBuffer.FieldByName( 'COMPCODE' ).AsString := FSo306.CompCode;
          FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString := FSo306.OldCompCode;
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString := FSo306.HfcMac;
          FDataBuffer.FieldByName( 'UPDEN' ).AsString := '物料系統';
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC序號=<%s>,正常。', [FSo306.HfcMac] );
          AppendLog( aCompCodeText, aActionText, aMsg, 0 );
          aMsg := EmptyStr;
        end;
      end;
    end;
  end;
  Result := not aDetailErr;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.SaveXML3: Boolean;

  { ---------------------------------------------------- }

  function GetAllCompanyString: String;
  begin
    Result := EmptyStr;
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      { 調入公司 }
      if Pos( FDataBuffer.FieldByName( 'COMPCODE' ).AsString, Result ) <= 0 then
        Result := ( Result + FDataBuffer.FieldByName( 'COMPCODE' ).AsString + ',' );
      { 調出公司 }
      if Pos( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString, Result ) <= 0 then
        Result := ( Result + FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString + ',' );
      FDataBuffer.Next;
    end;
    if Length( Result ) > 0 then Delete( Result, Length( Result ), 1 );
  end;

  { ---------------------------------------------------- }

  procedure UpdateBufferRecord(aCompCode: String);
  var
    aSo: PSoInfo;
  begin
    aSo := GetActiveSo( aCompCode );
    aSo.Reader.Close;
    aSo.Reader.SQL.Clear;
    aSo.Reader.SQL.Text :=
      ' SELECT BATCHNO, MATERIALNO,                                 ' +
      '        CARTONNUNO, CHASSISSN, CUSTOMERSN,                   ' +
      '        ETHERNETMAC, MTAMAC, FLAG,                           ' +
      '        TO_CHAR( CMISSUE, ''YYYY/MM/DD HH24:MI:SS'' ) AS CMISSUE,  ' +
      '        MODELNO, MODEID, MODELCODE, ACTSTATUS,               ' +
      '        TO_CHAR( LIMITDATE, ''YYYY/MM/DD'' ) AS LIMITDATE,   ' +
      '        BELONG, VENDORCODE, VENDORNAME, SPEC, DESCRIPTION,   ' +
      '        ALREADYUSE, MATERIALMODE, HARDWAREVER                ' +
      '   FROM SO306                                                ' +
      '  WHERE HFCMAC = :1                                          ';
     aSo.Reader.Parameters.ParamByName( '1' ).Value :=
       FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
     aSo.Reader.Open;
     FDataBuffer.Edit;
     {}
     FDataBuffer.FieldByName( 'BATCHNO' ).Value :=
       aSo.Reader.FieldByName( 'BATCHNO' ).AsString;
     {}
     FDataBuffer.FieldByName( 'MATERIALNO' ).Value :=
       aSo.Reader.FieldByName( 'MATERIALNO' ).AsString;
     {}
     FDataBuffer.FieldByName( 'CARTONNUNO' ).Value :=
       aSo.Reader.FieldByName( 'CARTONNUNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'CHASSISSN' ).Value :=
       aSo.Reader.FieldByName( 'CHASSISSN' ).Value;
     {}
     FDataBuffer.FieldByName( 'CUSTOMERSN' ).Value :=
       aSo.Reader.FieldByName( 'CUSTOMERSN' ).Value;
     {}
     FDataBuffer.FieldByName( 'ETHERNETMAC' ).Value :=
       aSo.Reader.FieldByName( 'ETHERNETMAC' ).Value;
     {}
     FDataBuffer.FieldByName( 'MTAMAC' ).Value :=
       aSo.Reader.FieldByName( 'MTAMAC' ).Value;
     {}
     FDataBuffer.FieldByName( 'FLAG' ).Value :=
       aSo.Reader.FieldByName( 'FLAG' ).Value;
     {}
     FDataBuffer.FieldByName( 'CMISSUE' ).Value :=
       aSo.Reader.FieldByName( 'CMISSUE' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODELNO' ).Value :=
       aSo.Reader.FieldByName( 'MODELNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODEID' ).Value :=
       aSo.Reader.FieldByName( 'MODEID' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODELCODE' ).Value :=
       aSo.Reader.FieldByName( 'MODELCODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'ACTSTATUS' ).Value :=
       aSo.Reader.FieldByName( 'ACTSTATUS' ).Value;
     {}
     FDataBuffer.FieldByName( 'USEFLAG' ).Value := '0';
     {}
     FDataBuffer.FieldByName( 'LIMITDATE' ).Value :=
       aSo.Reader.FieldByName( 'LIMITDATE' ).Value;
     {}
     FDataBuffer.FieldByName( 'BELONG' ).Value :=
       aSo.Reader.FieldByName( 'BELONG' ).Value;
     {}
     FDataBuffer.FieldByName( 'VENDORCODE' ).Value :=
       aSo.Reader.FieldByName( 'VENDORCODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'VENDORNAME' ).Value :=
       aSo.Reader.FieldByName( 'VENDORNAME' ).Value;
     {}
     FDataBuffer.FieldByName( 'SPEC' ).Value :=
       aSo.Reader.FieldByName( 'SPEC' ).Value;
     {}
     FDataBuffer.FieldByName( 'DESCRIPTION' ).Value :=
       aSo.Reader.FieldByName( 'DESCRIPTION' ).Value;
     {}
     FDataBuffer.FieldByName( 'ALREADYUSE' ).Value :=
       aSo.Reader.FieldByName( 'ALREADYUSE' ).Value;
     {}
     FDataBuffer.FieldByName( 'MATERIALTYPE' ).Value :=
       aSo.Reader.FieldByName( 'MATERIALMODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'HWVER' ).Value :=
       aSo.Reader.FieldByName( 'HARDWAREVER' ).Value;
     {}
     FDataBuffer.Post;
     aSo.Reader.Close;
  end;

  { ---------------------------------------------------- }

  procedure DeleteRecord(aCompCode: String);
  var
    aSo: PSoInfo;
  begin
    if aCompCode = XML_WAREHOUSEID then Exit;
    aSo := GetActiveSo( aCompCode );
    aSo.Writer.Close;
    aSo.Writer.SQL.Clear;
    aSo.Writer.SQL.Text :=
      ' DELETE FROM SO306 WHERE HFCMAC = :1 ';
    aSo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    aSo.Writer.ExecSQL;
  end;

  { ---------------------------------------------------- }

  procedure InsertRecord(aCompCode: String);
  var
    aSo: PSoInfo;
    aIsFound: Boolean;
  begin
    aSo := GetActiveSo( aCompCode );
    aSo.Reader.Close;
    aSo.Writer.SQL.Clear;
    aSo.Reader.SQL.Text :=
      ' SELECT HFCMAC FROM SO306 WHERE HFCMAC = :1 ';
    aSo.Reader.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    aSo.Reader.Open;
    aIsFound := ( not aSo.Reader.IsEmpty );
    aSo.Reader.Close;
    {}
    if ( aIsFound ) then Exit;
    {}
    aSo.Writer.Close;
    aSo.Writer.SQL.Clear;
    aSo.Writer.SQL.Text :=
      ' INSERT INTO SO306 (   ' +
      '    COMPCODE,          ' +
      '    BATCHNO,           ' +
      '    MATERIALNO,        ' +
      '    CARTONNUNO,        ' +
      '    MODELNO,           ' +
      '    CHASSISSN,         ' +
      '    CUSTOMERSN,        ' +
      '    HFCMAC,            ' +
      '    ETHERNETMAC,       ' +
      '    MTAMAC,            ' +
      '    FLAG,              ' +
      '    CMISSUE,           ' +
      '    MODEID,            ' +
      '    MODELCODE,         ' +
      '    ACTSTATUS,         ' +
      '    USEFLAG,           ' +
      '    LIMITDATE,         ' +
      '    BELONG,            ' +
      '    VENDORCODE,        ' +
      '    VENDORNAME,        ' +
      '    SPEC,              ' +
      '    DESCRIPTION,       ' +
      '    UPDEN,             ' +
      '    UPDTIME,           ' +
      '    ALREADYUSE,        ' +
      '    MATERIALMODE,      ' +
      '    HARDWAREVER     )  ' +
      ' VALUES (              ' +
      '    :1,                ' +
      '    :2,                ' +
      '    :3,                ' +
      '    :4,                ' +
      '    :5,                ' +
      '    :6,                ' +
      '    :7,                ' +
      '    :8,                ' +
      '    :9,                ' +
      '    :10,               ' +
      '    :11,               ' +
      '    TO_DATE( :12, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '    :13,               ' +
      '    :14,               ' +
      '    :15,               ' +
      '    :16,               ' +
      '    TO_DATE( :17, ''YYYY/MM/DD'' ),  ' +
      '    :18,               ' +
      '    :19,               ' +
      '    :20,               ' +
      '    :21,               ' +
      '    :22,               ' +
      '    :23,               ' +
      '    TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '    TO_DATE( :24, ''YYYY/MM/DD'' ),                ' +
      '    :25,               ' +
      '    :26    )           ';
    aSo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '2' ).Value :=
      FDataBuffer.FieldByName( 'BATCHNO' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '3' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALNO' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '4' ).Value :=
      FDataBuffer.FieldByName( 'CARTONNUNO' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '5' ).Value :=
      FDataBuffer.FieldByName( 'MODELNO' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '6' ).Value :=
      FDataBuffer.FieldByName( 'CHASSISSN' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '7' ).Value :=
      FDataBuffer.FieldByName( 'CUSTOMERSN' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '8' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '9' ).Value :=
      FDataBuffer.FieldByName( 'ETHERNETMAC' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '10' ).Value :=
      FDataBuffer.FieldByName( 'MTAMAC' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '11' ).Value :=
      FDataBuffer.FieldByName( 'FLAG' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '12' ).Value :=
      FDataBuffer.FieldByName( 'CMISSUE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '13' ).Value :=
      FDataBuffer.FieldByName( 'MODEID' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '14' ).Value :=
      FDataBuffer.FieldByName( 'MODELCODE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '15' ).Value :=
      FDataBuffer.FieldByName( 'ACTSTATUS' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '16' ).Value :=
      FDataBuffer.FieldByName( 'USEFLAG' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '17' ).Value :=
      FDataBuffer.FieldByName( 'LIMITDATE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '18' ).Value :=
      FDataBuffer.FieldByName( 'BELONG' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '19' ).Value :=
      FDataBuffer.FieldByName( 'VENDORCODE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '20' ).Value :=
      FDataBuffer.FieldByName( 'VENDORNAME' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '21' ).Value :=
      FDataBuffer.FieldByName( 'SPEC' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '22' ).Value :=
      FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '23' ).Value :=
      FDataBuffer.FieldByName( 'UPDEN' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '24' ).Value :=
      FDataBuffer.FieldByName( 'ALREADYUSE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '25' ).Value :=
      FDataBuffer.FieldByName( 'MATERIALTYPE' ).AsString;
    {}
    aSo.Writer.Parameters.ParamByName( '26' ).Value :=
      FDataBuffer.FieldByName( 'HWVER' ).AsString;
    {}
    aSo.Writer.ExecSQL;
  end;

  { ---------------------------------------------------- }

  procedure UpdateRecord(aCompCode: String);
  var
    aSo: PSoInfo;
  begin
    aSo := GetActiveSo( '0' );
    aSo.Writer.Close;
    aSo.Writer.SQL.Clear;
    aSo.Writer.SQL.Text :=
      ' UPDATE SO306             ' +
      '    SET COMPCODE = :1,    ' +
      '        MODELCODE = :2,   ' +
      '        UPDEN = :3,       ' +
      '        UPDTIME = TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ) ' +
      '  WHERE HFCMAC = :4       ';
     aSo.Writer.Parameters.ParamByName( '1' ).Value := aCompCode;
     aSo.Writer.Parameters.ParamByName( '2' ).Value :=
       FDataBuffer.FieldByName( 'MODELCODE' ).AsString;
     aSo.Writer.Parameters.ParamByName( '3' ).Value :=
       FDataBuffer.FieldByName( 'UPDEN' ).AsString;
     aSo.Writer.Parameters.ParamByName( '4' ).Value :=
       FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
     aSo.Writer.ExecSQL;
  end;

  { ---------------------------------------------------- }

  procedure InsertLog(aValue: String);
  var
    aSo: PSoInfo;
  begin
    aSo := GetActiveSo( XML_WAREHOUSEID );
    aSo.Writer.Close;
    aSo.Writer.SQL.Clear;
    aSo.Writer.SQL.Text :=
      ' INSERT INTO SOAC0204 (   ' +
      '    CALLTIME,             ' +
      '    COMPCODE,             ' +
      '    PROGNAME,             ' +
      '    PARA     )            ' +
      ' VALUES (                 ' +
      '    TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '    :1,                   ' +
      '    ''CM2CT_DLL2'',       ' +
      '    :2       )            ';
    aSo.Writer.Parameters.ParamByName( '1' ).Value := EmptyStr;
    aSo.Writer.Parameters.ParamByName( '2' ).Value := aValue;
    aSo.Writer.ExecSQL;
  end;

  { ------------------------------------------------------- }

var
  aSoInfo: PSoInfo;
  aComps, aTempComps, aModelCode: String;
  aLogComp, aLogOldComp: String;
  aCount: Integer;
  aIgnore: Boolean;
  aMsg, aActionText, aDbComp: String;
  aDbErr, aTotalErr, aIsErr: Boolean;
begin
  Result := False;
  aActionText := GetActionText( 1 );
  { 取出所有公司, 資料庫連線用 }
  aComps := GetAllCompanyString;
  { 下面 Transaction 要用到 }
  aTempComps := aComps;
  aDbErr := False;
  while ( aComps <> EmptyStr ) do
  begin
    aDbComp := ExtractValue( aComps );
    aMsg := PrepareSoDBConnection( aDbComp );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      Exit;
    end;
  end;
  if ( aDbErr ) then Exit;
  { 檢查資料 }
  aTotalErr := False;
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    aIsErr := False;
    { 1. 檢查調出的系統台 DB 是否正確, 預防連錯資料庫 }
    aDbComp :=
      FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString + ',' +
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    aMsg := VdCD039( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 2. 檢查調入的系統台 DB 是否正確, 預防連錯資料庫 }
    aMsg := VdCD039( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 3. 調出此設備的系統台, 是否有此設備?  }
    aMsg := VdNotExists( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 4. 調出此設備的系統台, 是否使用中 ? }
    aMsg := VdNotUse( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 5. 調出此設備的系統台, 該設備是否已註銷(黑名單?) }
    aMsg := VdBlackList( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 6. 從調出設備的系統台讀取原本的設備資料 }
    UpdateBufferRecord( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString );
    { 7. 檢查調入設備系統台是否有 該設備的 ModelCode }
    aMsg := VdCD043(
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'MODELNO' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString, aModelCode  );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 8. 更新 ModelCode }
    if not aTotalErr then
    begin
      FDataBuffer.Edit;
      FDataBuffer.FieldByName( 'MODELCODE' ).AsString := aModelCode;
      FDataBuffer.Post;
    end;
    { 8. 檢核調入的系統台是否有此客服品名 }
    aMsg := VdCD022(
      FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    {}
    if not aIsErr then
    begin
      aMsg := Format( 'HFCMAC序號<%s>,正常。', [
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      AppendLog( aDbComp, aActionText, aMsg, 0 );
    end;
    FDataBuffer.Next;
  end;
  {}
  if aTotalErr then Exit;
  {}
  { 啟動交易 }
  aComps := aTempComps;
  while ( aComps <> EmptyStr ) do
  begin
    aSoInfo := GetActiveSo( ExtractValue( aComps ) );
    if not aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.BeginTrans;
  end;
  { 存檔 }
  try
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      aDbComp :=
        FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString + ',' +
        FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
      { 調出, 調入系統台別必須不同, 才異動資料 }
      if ( FDataBuffer.FieldByName( 'COMPCODE' ).AsString <>
           FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString ) then
      begin
        { 處理調出此設備系統台 }
        DeleteRecord( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString );
        { 處理調入此設備系統台 }
        InsertRecord( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
        { 更新總倉, 傳入的倉別為調入的倉別 }
        UpdateRecord( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
      end;
      FDataBuffer.Next;
    end;
    { 最後, 寫 Log }
    FDataBuffer.First;
    aLogComp := EmptyStr;
    aLogOldComp := EmptyStr;
    aCount := 0;
    while not FDataBuffer.Eof do
    begin
      aDbComp :=
        FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString + ',' +
        FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
      {}
      aIgnore := ( FDataBuffer.FieldByName( 'COMPCODE' ).AsString =
        FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString );
      if ( aIgnore ) then
      begin
        if ( aLogComp <> EmptyStr ) or ( aLogOldComp <> EmptyStr ) then
        begin
          InsertLog( Format( '設備別:CM,調出公司:%s,調入公司:%s,數量:%d',
            [FDataBuffer.FieldByName( 'OLDCOMP' ).AsString,
             FDataBuffer.FieldByName( 'COMPCODE' ).AsString, aCount] ) );
        end;
        aLogComp := EmptyStr;
        aLogOldComp := EmptyStr;
        aCount := 0;
      end else
      begin
        if ( aLogComp = EmptyStr ) then
          aLogComp := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
        if ( aLogOldComp = EmptyStr ) then
          aLogOldComp := FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString;
        if ( aLogComp <> FDataBuffer.FieldByName( 'COMPCODE' ).AsString ) or
           ( aLogOldComp <> FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString ) then
        begin
          InsertLog( Format( '設備別:CM,調出公司:%s,調入公司:%s,數量:%d',
            [FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
             FDataBuffer.FieldByName( 'COMPCODE' ).AsString, aCount] ) );
          aLogComp := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
          aLogOldComp := FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString;
          aCount := 0;
        end;
        Inc( aCount );
      end;
      FDataBuffer.Next;
      if ( FDataBuffer.Eof and ( aCount > 0 ) ) then
      begin
        InsertLog( Format( '設備別:CM,調出公司:%s,調入公司:%s,數量:%d',
          [FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
           FDataBuffer.FieldByName( 'COMPCODE' ).AsString, aCount] ) );
      end;
    end;
    { Commit }
    aComps := aTempComps;
    while ( aComps <> EmptyStr ) do
    begin
      aSoInfo := GetActiveSo( ExtractValue( aComps ) );
      if aSoInfo.Connection.InTransaction then
        aSoInfo.Connection.CommitTrans;
    end;
    Result := True;
  except
    on E: Exception do
    begin
      { Rollback }
      aComps := aTempComps;
      while ( aComps <> EmptyStr ) do
      begin
        aSoInfo := GetActiveSo( ExtractValue( aComps ) );
        if aSoInfo.Connection.InTransaction then
          aSoInfo.Connection.RollbackTrans;
      end;
      aMsg := Format( '-1: CC&B調撥資料庫存檔失敗,錯誤訊息:<%s>', [E.Message] );
      AppendLog( aDbComp, aActionText, aMsg, 1 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.ValidateXML4(aXmlData: String): Boolean;
var
  aGetBackList, aEqumList, aSeqList: IDOMNodeList;
  aNode: IDOMNode;
  aIndex, aIndex2, aIndex3: Integer;
  aDetailErr: Boolean;
  aMsg, aActionText, aCompCodeText: String;
begin
  Result := False;
  aDetailErr := False;
  aActionText := GetActionText;
  try
    FXml.LoadFromXML( aXmlData );
  except
    on E: Exception do aMsg := E.Message;
  end;
  if ( aMsg <> EmptyStr ) then
  begin
    aMsg := Format( '-1:傳入之領退料 XML 格式有誤, 錯誤原因:%s。', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  aGetBackList := FXml.DOMDocument.getElementsByTagName( '領退料' );
  aMsg := VdGetBackList( aGetBackList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { 取出領退料節點 }
  for aIndex := 0 to aGetBackList.length - 1 do
  begin
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    aEqumList := ( aGetBackList.item[aIndex] as IDOMNodeSelect ).selectNodes( '設備' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '公司代碼' );
      aMsg := VdCompCode( aNode, '0' );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end else
      begin
        FSo306.CompCode := aNode.nodeValue;
        aCompCodeText := FSo306.CompCode;
      end;
      {}
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '數量' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '作業' );
      aMsg := VdGetOrBack( aNode );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      { 領料, useflag = 1 --> 可用 }
      if ( aNode.nodeValue = '1' ) then
        FSo306.UseFlag := 1
      else
      { 退料, useflag = 0 --> 不可用 }
        FSo306.UseFlag := 0;
      {}
      aSeqList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( 'HFCMAC序號' );
      aMsg := VdHFCMac( aSeqList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      for aIndex3 := 0 to aSeqList.length - 1 do
      begin
        aMsg := VdHFCMac( aSeqList.item[aIndex3].firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDetailErr := True;
        end else
          FSo306.HfcMac := aSeqList.item[aIndex3].firstChild.nodeValue;
        {}
        if ( not aDetailErr ) then
        begin
          FDataBuffer.Append;
          FDataBuffer.FieldByName( 'COMPCODE' ).AsString := FSo306.CompCode;
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString := FSo306.HfcMac;
          FDataBuffer.FieldByName( 'USEFLAG' ).AsInteger := FSo306.UseFlag;
          FDataBuffer.FieldByName( 'UPDEN' ).AsString := '物料系統';
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC序號=<%s>,正常。', [FSo306.HfcMac] );
          AppendLog( aCompCodeText, aActionText, aMsg, 0 );
          aMsg := EmptyStr;
        end;
      end;
    end;
  end;
  Result := not aDetailErr;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.SaveXML4: Boolean;

  { ---------------------------------------------------- }

  function GetAllCompanyString: String;
  begin
    Result := EmptyStr;
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      if Pos( FDataBuffer.FieldByName( 'COMPCODE' ).AsString, Result ) <= 0 then
        Result := ( Result + FDataBuffer.FieldByName( 'COMPCODE' ).AsString + ',' );
      FDataBuffer.Next;
    end;
    if Length( Result ) > 0 then Delete( Result, Length( Result ), 1 );
  end;

  { ---------------------------------------------------- }

  procedure UpdateRecords(aCompCode: String);
  var
    aSo: PSoInfo;
  begin
    aSo := GetActiveSo( aCompCode );
    aSo.Writer.Close;
    aSo.Writer.SQL.Clear;
    aSo.Writer.SQL.Text :=
     ' UPDATE SO306 SET USEFLAG = :1 ' +
     '  WHERE HFCMAC = :2            ';
    aSo.Writer.Parameters.ParamByName( '1' ).Value :=
      FDataBuffer.FieldByName( 'USEFLAG' ).AsInteger;
    aSo.Writer.Parameters.ParamByName( '2' ).Value :=
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
    aSO.Writer.ExecSQL;
  end;

  { ---------------------------------------------------- }

  procedure InsertLog(aCompCode, aValue: String);
  var
    aSo: PSoInfo;
  begin
    aSo := GetActiveSo( aCompCode );
    aSo.Writer.Close;
    aSo.Writer.SQL.Clear;
    aSo.Writer.SQL.Text :=
      ' INSERT INTO SOAC0204 (   ' +
      '    CALLTIME,             ' +
      '    COMPCODE,             ' +
      '    PROGNAME,             ' +
      '    PARA     )            ' +
      ' VALUES (                 ' +
      '    TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '    :1,                   ' +
      '    ''CM2CT_DLL3'',       ' +
      '    :2       )            ';
    aSo.Writer.Parameters.ParamByName( '1' ).Value := aCompCode;
    aSo.Writer.Parameters.ParamByName( '2' ).Value := aValue;
    aSo.Writer.ExecSQL;
  end;

  { ---------------------------------------------------- }

var
  aComps, aTempComps: String;
  aSoInfo: PSoInfo;
  aAction: String;
  aDbErr, aTotalErr, aIsErr: Boolean;
  aMsg, aActionText, aCompCodeText: String;
begin
  Result := False;
  aActionText := GetActionText( 1 );
  aComps := GetAllCompanyString;
  aTempComps := aComps;
  aDbErr := False;
  while ( aComps <> EmptyStr ) do
  begin
    aMsg := PrepareSoDBConnection( ExtractValue( aComps ) );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( XML_WAREHOUSEID, aActionText, aMsg, 1 );
      aDbErr := True;
    end;
  end;
  if aDbErr then Exit;
  { 啟動交易 }
  aComps := aTempComps;
  while ( aComps <> EmptyStr ) do
  begin
    aSoInfo := GetActiveSo( ExtractValue( aComps ) );
    if not aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.BeginTrans;
  end;
  { 存檔 }
  aTotalErr := False;
  try
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      aIsErr := False;
      aCompCodeText := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
      { 1.該系統台是否有此設備 }
      aMsg := VdNotExists( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      { 2.是否使用中, 退料特殊規則 }
      aMsg := VdNotUse( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      { 3.是否被註銷 }
      aMsg := VdBlackList( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      {}
      if not aTotalErr then
        UpdateRecords( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
      if not aIsErr then
      begin
        aMsg := Format( 'HFCMAC序號=<%s>,正常。', [
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
        AppendLog( aCompCodeText, aActionText, aMsg, 0 );
        aMsg := EmptyStr;
      end;
      FDataBuffer.Next;
    end;
    { 寫 Log }
    { 不須要寫到 SOAC0204A, 因為已經有 materiel_log 這個 Table }
//    if not aTotalErr then
//    begin
//      FDataBuffer.First;
//      aLogComp := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
//      aLogUseFlag := FDataBuffer.FieldByName( 'USEFLAG' ).AsString;
//      aLogHfcMac := EmptyStr;
//      while not FDataBuffer.Eof do
//      begin
//        if ( aLogComp <> FDataBuffer.FieldByName( 'COMPCODE' ).AsString ) or
//           ( aLogUseFlag <> FDataBuffer.FieldByName( 'USEFLAG' ).AsString ) then
//        begin
//          if ( Length( aLogHfcMac ) > 0 ) then
//          begin
//            if aLogHfcMac[Length( aLogHfcMac )] = ',' then
//              Delete( aLogHfcMac, Length( aLogHfcMac ), 1 );
//            if Length( aLogHfcMac ) >= 4000 then
//              aLogHfcMac := Copy( aLogHfcMac, 1, 4000 );
//          end;
//          aAction := '領料';
//          if aLogUseFlag = '0' then aAction := '退料';
//          InsertLog( aLogComp, Format( '設備別:CM,動作:%s,序號:%s',
//            [aAction, aLogHfcMac] ) );
//          aLogComp := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
//          aLogUseFlag := FDataBuffer.FieldByName( 'USEFLAG' ).AsString;
//          aLogHfcMac := EmptyStr;
//        end;
//        aLogHfcMac := ( aLogHfcMac + FDataBuffer.FieldByName( 'HFCMAC' ).AsString + ',' );
//        FDataBuffer.Next;
//        if ( FDataBuffer.Eof and ( aLogHfcMac <> EmptyStr ) ) then
//        begin
//          if ( Length( aLogHfcMac ) > 0 ) then
//          begin
//            if aLogHfcMac[Length( aLogHfcMac )] = ',' then
//              Delete( aLogHfcMac, Length( aLogHfcMac ), 1 );
//            if Length( aLogHfcMac ) >= 4000 then
//              aLogHfcMac := Copy( aLogHfcMac, 1, 4000 );
//          end;
//          aAction := '領料';
//          if aLogUseFlag = '0' then aAction := '退料';
//          InsertLog( aLogComp, Format( '設備別:CM,動作:%s,序號:%s',
//            [aAction, aLogHfcMac] ) );
//        end;
//      end;
//    end;
    { Commit }
    aComps := aTempComps;
    while ( aComps <> EmptyStr ) do
    begin
      aSoInfo := GetActiveSo( ExtractValue( aComps ) );
      if aSoInfo.Connection.InTransaction then
      begin
        if aTotalErr then
          aSoInfo.Connection.RollbackTrans
        else begin
          aSoInfo.Connection.CommitTrans;
          Result := True;
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      { Rollback }
      aComps := aTempComps;
      while ( aComps <> EmptyStr ) do
      begin
        aSoInfo := GetActiveSo( ExtractValue( aComps ) );
        if aSoInfo.Connection.InTransaction then
          aSoInfo.Connection.RollbackTrans;
      end;
      Result := False;
      aMsg := Format( '-1: CC&B領/退料資料庫存檔失敗,錯誤訊息:<%s>', [E.Message] );
      AppendLog( aCompCodeText, aAction, aMsg, 1 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.MakeErrorText: String;
var
  aList: TStringList;
  aIndex: Integer;
begin
  Result := EmptyStr;
  aList := TStringList.Create;
  try
    FLogBuffer.First;
    while not FLogBuffer.Eof do
    begin
      if FLogBuffer.FieldByName( 'ISERR' ).AsInteger = 1 then
      begin
        aList.Text := FLogBuffer.FieldByName( 'MSG' ).AsString;
        for aIndex := 0 to aList.Count - 1 do
        begin
          if ( AnsiPos( '正常', aList.Strings[aIndex] ) <= 0 ) then 
          Result := ( Result + aList.Strings[aIndex] + #13#10 );
        end;  
      end;
      FLogBuffer.Next;
    end;
  finally
    aList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.AppendLog(aCompCode, aAction, aMsg: String;
  aIsErr: Integer);
var
  aSo: PSoInfo;
  aTempText, aTempCode, aCompName: String;
begin
  aTempText := aCompCode;
  aCompName := EmptyStr;
  repeat
    aTempCode := ExtractValue( aTempText );
    aSo := GetActiveSo( aTempCode );
    if Assigned( aSo ) then
      aCompName := ( aCompName + aSo.CompName );
    if ( aTempText <> EmptyStr ) and ( aCompName <> EmptyStr ) then
      aCompName := ( aCompName + ',' );
  until ( aTempText = EmptyStr );
  {}
  FLogBuffer.Last;
  if ( FLogBuffer.IsEmpty ) or
     ( FLogBuffer.FieldByName( 'ACTION' ).AsString <> aAction ) or
     ( FLogBuffer.FieldByName( 'COMPCODE' ).AsString <> aCompCode ) then
    FLogBuffer.Append
  else
    FLogBuffer.Edit;
  if FLogBuffer.State in [dsInsert] then
  begin
    FLogBuffer.FieldByName( 'CALLTIME' ).AsDateTime := FCallTime;
    FLogBuffer.FieldByName( 'ACTION' ).AsString := aAction;
    FLogBuffer.FieldByName( 'COMPCODE' ).AsString := aCompCode;
    FLogBuffer.FieldByName( 'COMPNAME' ).AsString := aCompName;
    FLogBuffer.FieldByName( 'XMLDATA' ).AsString := FActiveXmlData;
    FLogBuffer.FieldByName( 'ISERR').AsInteger := aIsErr;
  end else
  begin
    { 只要 IsErr 的 Flag 變成 1 後, 就不可再改回來 }
    if ( FLogBuffer.FieldByName( 'ISERR' ).AsInteger = 0 ) and ( aIsErr = 1 ) then
      FLogBuffer.FieldByName( 'ISERR' ).AsInteger := aIsErr;
  end;    
  FLogBuffer.FieldByName( 'MSG' ).AsString := (
    FLogBuffer.FieldByName( 'MSG' ).AsString + aMsg + #13#10 );
  FLogBuffer.Post;   
end;

{ ---------------------------------------------------------------------------- }

procedure TCM2CT.InitCallTime;
var
  aSo: PSoInfo;
  aMsg: String;
begin
  FCallTime := Now;
  aMsg := PrepareSoDBConnection( XML_WAREHOUSEID );
  if ( aMsg = EmptyStr ) then
  begin
    aSo := GetActiveSo( XML_WAREHOUSEID );
    if Assigned( aSo ) then
    begin
      aSo.Reader.Close;
      aSo.Reader.SQL.Text := ' select sysdate from dual ';
      try
        aSo.Reader.Open;
        FCallTime := aSo.Reader.Fields[0].AsDateTime;
      except
        {}
      end;
      aSo.Reader.Close;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.GetActionText(const CallType: Integer = 0): String;
begin
  // CallType = 0 --> 檢核, 從 validate 函式開頭呼叫進來
  // CallType = 1 --> 存檔, 從 save 函式開頭呼叫進來
  Result := '檢核';
  if ( CallType <> 0 ) then Result := '存檔';
  case FType of
    1: Result := ( '入庫,' + Result );
    2: Result := ( '驗退,' + Result );
    3: Result := ( '調撥,' + Result );
    4: Result := ( '領/退料,' + Result );
    5: Result := ( '維修送回,' + Result );
  end;    
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.WriteLog: String;
var
  aSo: PSoInfo;
begin
  Result := EmptyStr;
  { 資料庫連結至總倉 }
  Result := PrepareSoDBConnection( XML_WAREHOUSEID );
  if ( Result <> EmptyStr ) then Exit;
  { Log 資料一律寫總倉資料區 COM 區 }
  aSo := GetActiveSo( '0' );
  if not Assigned( aSo ) then
  begin
    Result := '-2:CC&B資料庫無法對應到總倉資料區。';
    Exit;
  end;
  aSo.LogWriter.SQL.Clear;
  aSo.LogWriter.SQL.Text :=
    ' insert into materiel_log ( calltime, action,       ' +
    '   compcode, compname, iserr, msg, xmldata )        ' +
    ' values ( :calltime, :action, :compcode,            ' +
    '   :compname, :iserr, empty_clob(), empty_clob() )  ' +
    ' returning msg, xmldata into :msg, :xmldata         ';
  {}
  aSo.LogWriter.ParamByName( 'msg' ).DataType := ftOraClob;
  aSo.LogWriter.ParamByName( 'msg' ).ParamType := ptInput;
  aSo.LogWriter.ParamByName( 'xmldata' ).DataType := ftOraClob;
  aSo.LogWriter.ParamByName( 'xmldata' ).ParamType := ptInput;
  {}
  try
    FLogBuffer.First;
    while not FLogBuffer.Eof do
    begin
      aSo.LogWriter.ParamByName( 'calltime' ).AsDateTime :=
        FLogBuffer.FieldByName( 'calltime' ).AsDateTime;
      {}
      aSo.LogWriter.ParamByName( 'action' ).AsString :=
        FLogBuffer.FieldByName( 'action' ).AsString;
      {}
      aSo.LogWriter.ParamByName( 'compcode' ).AsString :=
        FLogBuffer.FieldByName( 'compcode' ).AsString;
      {}
      aSo.LogWriter.ParamByName( 'compname' ).AsString :=
        FLogBuffer.FieldByName( 'compname' ).AsString;
      {}
      aSo.LogWriter.ParamByName( 'iserr' ).AsInteger :=
        FLogBuffer.FieldByName( 'iserr' ).AsInteger;
      {}
      aSo.LogWriter.ParamByName( 'msg' ).AsString :=
        FLogBuffer.FieldByName( 'msg' ).AsString;
      {}
      aSo.LogWriter.ParamByName( 'xmldata' ).AsString :=
        FLogBuffer.FieldByName( 'xmldata' ).AsString;
      {}
      aSo.LogWriter.Execute;
      FLogBuffer.Next;
    end;
  except
    on E: Exception do
    begin
      Result := Format('-2:CC&B資料庫寫入Log檔失敗, 原因:<%s>。', [E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.GetChangeCompanyCode(aNode: IDOMNode): String;
var
  aOutNode, aInNode: IDOMNode;
begin
  Result := EmptyStr;
  if not Assigned( aNode ) then Exit;
  aOutNode := aNode.attributes.getNamedItem( '調出公司' );
  aInNode := aNode.attributes.getNamedItem( '調入公司' );
  if Assigned( aOutNode ) and Assigned( aInNode ) then
    Result := ( aOutNode.nodeValue + ',' + aInNode.nodeValue );
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.ValidateXML5(aXmlData: String): Boolean;
var
  aMsg, aActionText, aCompCodeText: String;
  aDeatilErr: Boolean;
  aChangeList, aEqumList, aHFCList: IDOMNodeList;
  aIndex, aIndex2, aIndex3: Integer;
  aNode, aSendNode, aBackNode: IDOMNode;
begin
  Result := False;
  aMsg := EmptyStr;
  aDeatilErr := False;
  aActionText := GetActionText;
  try
    FXml.LoadFromXML( aXmlData );
  except
    on E: Exception do aMsg := E.Message;
  end;
  if ( aMsg <> EmptyStr ) then
  begin
    aMsg := Format( '-1:傳入之維修送回 XML 格式有誤, 錯誤原因:%s。', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  aChangeList := FXml.DOMDocument.getElementsByTagName( '維修送回' );
  aMsg := VdChangeList( aChangeList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { 取出維修送回節點 }
  for aIndex := 0 to aChangeList.length - 1 do
  begin
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    aEqumList := ( aChangeList.item[aIndex] as IDOMNodeSelect ).selectNodes( '設備' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { 逐一取出設備 }
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '公司代碼' );
      aMsg := VdCompCode( aNode, EmptyStr );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end else
      begin
        FSo306.CompCode := aNode.nodeValue;
        aCompCodeText := FSo306.CompCode;
      end;
      {}
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '數量' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      aHFCList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( '序號' );
      aMsg := VdHfcList( aHFCList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      for aIndex3 := 0 to aHFCList.length - 1 do
      begin
        aSendNode := aHFCList.item[aIndex3].childNodes.item[0];
        aBackNode := aHFCList.item[aIndex3].childNodes.item[1];
        aMsg := VdPariEqum( aSendNode, aBackNode );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDeatilErr := True;
          Continue;
        end;
        { 先決定那一個才是 送修 的 node }
        aMsg := VdExchangeNode( aSendNode, aBackNode );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDeatilErr := True;
          Continue;
        end;
        { 送修 HFCMAC }
        aMsg := VdHFCMac( aSendNode.attributes.getNamedItem( 'HFCMAC' ) );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDeatilErr := True;
          Continue;
        end;
        { 領回 HFCMAC }
        aMsg := VdHFCMac( aBackNode.attributes.getNamedItem( 'HFCMAC' ) );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDeatilErr := True;
          Continue;
        end;
        {}
        FSo306.HfcMac := EmptyStr;
        FSo306.MtaMac := EmptyStr;
        {}
        FSo306.BackHFCMac := EmptyStr;
        FSo306.BackMTAMac := EmptyStr;
        {}
        FSo306.HfcMac := aSendNode.attributes.getNamedItem( 'HFCMAC' ).nodeValue;
        if Assigned( aSendNode.attributes.getNamedItem( 'MTAMAC' ) ) then
          FSo306.MtaMac := aSendNode.attributes.getNamedItem( 'MTAMAC' ).nodeValue;
        {}
        FSo306.BackHFCMac := aBackNode.attributes.getNamedItem( 'HFCMAC' ).nodeValue;
        if Assigned( aBackNode.attributes.getNamedItem( 'MTAMAC' ) ) then
          FSo306.BackMtaMac := aBackNode.attributes.getNamedItem( 'MTAMAC' ).nodeValue;
        {}
        FDataBuffer.Append;
        FDataBuffer.FieldByName( 'COMPCODE' ).AsString := FSo306.CompCode;
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString := FSo306.HfcMac;
        FDataBuffer.FieldByName( 'SENDMTAMAC' ).AsString := FSo306.MtaMac;
        FDataBuffer.FieldByName( 'BACKHFCMAC' ).AsString := FSo306.BackHFCMac;
        FDataBuffer.FieldByName( 'BACKMTAMAC' ).AsString := FSo306.BackMTAMac;
        FDataBuffer.FieldByName( 'UPDEN' ).AsString := '物料系統';
        FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
        FDataBuffer.Post;
        aMsg := Format( '送修HFCMAC序號=<%s>, 領回HFCMAC序號=<%s>, 正常。',
          [FSo306.HfcMac, FSo306.BackHFCMac] );
        AppendLog( aCompCodeText, aActionText, aMsg, 0 );
        aMsg := EmptyStr;
      end;
    end;
  end;
  Result := not aDeatilErr;
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.SaveXML5: Boolean;

  { ---------------------------------------------------- }

  function GetAllCompanyString: String;
  begin
    Result := EmptyStr;
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      if Pos( FDataBuffer.FieldByName( 'COMPCODE' ).AsString, Result ) <= 0 then
        Result := ( Result + FDataBuffer.FieldByName( 'COMPCODE' ).AsString + ',' );
      FDataBuffer.Next;
    end;
    { 總倉加進來 }
    if Pos( '0', Result ) <= 0 then
      Result := ( Result + '0' + ',' );
    if Length( Result ) > 0 then Delete( Result, Length( Result ), 1 );
  end;

  { ---------------------------------------------------- }

  procedure UpdateBufferRecord(aCompCode: String);
  var
    aSo: PSoInfo;
  begin
    aSo := GetActiveSo( aCompCode );
    aSo.Reader.Close;
    aSo.Reader.SQL.Clear;
    aSo.Reader.SQL.Text :=
      ' SELECT BATCHNO, MATERIALNO,                                 ' +
      '        CARTONNUNO, CHASSISSN, CUSTOMERSN,                   ' +
      '        ETHERNETMAC, MTAMAC, FLAG,                           ' +
      '        TO_CHAR( CMISSUE, ''YYYY/MM/DD HH24:MI:SS'' ) AS CMISSUE,  ' +
      '        MODELNO, MODEID, MODELCODE, ACTSTATUS,               ' +
      '        TO_CHAR( LIMITDATE, ''YYYY/MM/DD'' ) AS LIMITDATE,   ' +
      '        BELONG, VENDORCODE, VENDORNAME, SPEC, DESCRIPTION    ' +
      '   FROM SO306                                                ' +
      '  WHERE HFCMAC = :1                                          ';
     aSo.Reader.Parameters.ParamByName( '1' ).Value :=
       FDataBuffer.FieldByName( 'HFCMAC' ).AsString;
     aSo.Reader.Open;
     FDataBuffer.Edit;
     {}
     FDataBuffer.FieldByName( 'BATCHNO' ).Value :=
       aSo.Reader.FieldByName( 'BATCHNO' ).AsString;
     {}
     FDataBuffer.FieldByName( 'MATERIALNO' ).Value :=
       aSo.Reader.FieldByName( 'MATERIALNO' ).AsString;
     {}
     FDataBuffer.FieldByName( 'CARTONNUNO' ).Value :=
       aSo.Reader.FieldByName( 'CARTONNUNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'CHASSISSN' ).Value :=
       aSo.Reader.FieldByName( 'CHASSISSN' ).Value;
     {}
     FDataBuffer.FieldByName( 'CUSTOMERSN' ).Value :=
       aSo.Reader.FieldByName( 'CUSTOMERSN' ).Value;
     {}
     FDataBuffer.FieldByName( 'ETHERNETMAC' ).Value :=
       aSo.Reader.FieldByName( 'ETHERNETMAC' ).Value;
     {}
     FDataBuffer.FieldByName( 'MTAMAC' ).Value :=
       aSo.Reader.FieldByName( 'MTAMAC' ).Value;
     {}
     FDataBuffer.FieldByName( 'FLAG' ).Value :=
       aSo.Reader.FieldByName( 'FLAG' ).Value;
     {}
     FDataBuffer.FieldByName( 'CMISSUE' ).Value :=
       aSo.Reader.FieldByName( 'CMISSUE' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODELNO' ).Value :=
       aSo.Reader.FieldByName( 'MODELNO' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODEID' ).Value :=
       aSo.Reader.FieldByName( 'MODEID' ).Value;
     {}
     FDataBuffer.FieldByName( 'MODELCODE' ).Value :=
       aSo.Reader.FieldByName( 'MODELCODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'ACTSTATUS' ).Value :=
       aSo.Reader.FieldByName( 'ACTSTATUS' ).Value;
     {}
     FDataBuffer.FieldByName( 'USEFLAG' ).Value := '0';
     {}
     FDataBuffer.FieldByName( 'LIMITDATE' ).Value :=
       aSo.Reader.FieldByName( 'LIMITDATE' ).Value;
     {}
     FDataBuffer.FieldByName( 'BELONG' ).Value :=
       aSo.Reader.FieldByName( 'BELONG' ).Value;
     {}
     FDataBuffer.FieldByName( 'VENDORCODE' ).Value :=
       aSo.Reader.FieldByName( 'VENDORCODE' ).Value;
     {}
     FDataBuffer.FieldByName( 'VENDORNAME' ).Value :=
       aSo.Reader.FieldByName( 'VENDORNAME' ).Value;
     {}
     FDataBuffer.FieldByName( 'SPEC' ).Value :=
       aSo.Reader.FieldByName( 'SPEC' ).Value;
     {}
     FDataBuffer.FieldByName( 'DESCRIPTION' ).Value :=
       aSo.Reader.FieldByName( 'DESCRIPTION' ).Value;
     {}
     FDataBuffer.Post;
     aSo.Reader.Close;
  end;

  { ---------------------------------------------------- }

  procedure UpdateSoRecord(aCompCode: String);
  var
    aIndex: Integer;
    aSo: PSoInfo;
  begin
    for aIndex := 1 to 2 do
    begin
      aSo := GetActiveSo( aCompCode );
      aSo.Writer.Close;
      aSo.Writer.SQL.Clear;
      aSo.Writer.SQL.Text := Format(
        ' UPDATE SO306                 ' +
        '    SET HFCMAC = ''%s'',      ' +
        '        MTAMAC = ''%s'',      ' +
        '        UPDTIME = TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
        '        UPDEN = ''物料系統''  ' +
        '  WHERE HFCMAC = ''%s''       ',
        [FDataBuffer.FieldByName( 'BACKHFCMAC' ).AsString,
         FDataBuffer.FieldByName( 'BACKMTAMAC' ).AsString,
         FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      aSo.Writer.ExecSQL;
      { 若本身就是總倉, 則不用再做第2次 }
      if ( aCompCode = XML_WAREHOUSEID ) then Break;
      { 若不是總倉, 則還要將總倉更新 }
      aCompCode := XML_WAREHOUSEID;
    end;
  end;

  { ---------------------------------------------------- }

var
  aActionText, aComps, aTempComps, aDbComp: String;
  aMsg, aCompCodeText: String;
  aDbErr, aTotalErr, aIsErr: Boolean;
  aSoInfo: PSoInfo;
begin
  Result := False;
  aActionText := GetActionText( 1 );
  aComps := GetAllCompanyString;
  aTempComps := aComps;
  aDbErr := False;
  while ( aComps <> EmptyStr ) do
  begin
    aDbComp := ExtractValue( aComps );
    aMsg := PrepareSoDBConnection( aDbComp );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aDbErr := True;
    end;
  end;
  if aDbErr then Exit;
  {}
  aTotalErr := False;
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    aIsErr := False;
    aCompCodeText := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
    { 1.送修的設備, 該系統台是否有此設備 }
    aMsg := VdNotExists( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 2.送修的設備是否使用中, 維修送回特殊規則 }
    aMsg := VdNotUse( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 3.送修的設備是否被註銷 }
    aMsg := VdBlackList( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    {}
    UpdateBufferRecord( FDataBuffer.FieldByName( 'COMPCODE' ).AsString  );
    { 維修送回的機器序號, 若是有變更, 必須檢核不可以存在於庫存內 }
    if ( FDataBuffer.FieldByName( 'HFCMAC' ).AsString <>
         FDataBuffer.FieldByName( 'BACKHFCMAC' ).AsString ) then
    begin
      aMsg := VdBackEqumExists( FDataBuffer.FieldByName( 'BACKHFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      { 4.檢核送修的 MTAMAC 是否等於 DB 裏原有的 MTAMAC }
      aMsg := VdEqumMtaMac( FDataBuffer.FieldByName( 'HFCMAC' ).AsString,
        FDataBuffer.FieldByName( 'MTAMAC' ).AsString,
        FDataBuffer.FieldByName( 'SENDMTAMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      {}
      { 5.檢核領回的型式 是否等於 DB 裏原有的型式 (用MTAMAC有值或無值判斷 ) }
      aMsg := VdEqumMtaMac( FDataBuffer.FieldByName( 'HFCMAC' ).AsString,
        FDataBuffer.FieldByName( 'MTAMAC' ).AsString,
        FDataBuffer.FieldByName( 'BACKMTAMAC' ).AsString, False );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
    end else
    { 若是送修前後的 HFCMAC 是一樣的, 則看 MTAMAC 送修前後是否一樣 }
    begin
      aMsg := VdEqumMtaMac( FDataBuffer.FieldByName( 'HFCMAC' ).AsString,
        FDataBuffer.FieldByName( 'MTAMAC' ).AsString,
        FDataBuffer.FieldByName( 'BACKMTAMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
    end;
    {}
    if not aIsErr then
    begin
      aMsg := Format( 'HFCMAC序號<%s>,正常。', [
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      AppendLog( aCompCodeText, aActionText, aMsg, 0 );
    end;
    FDataBuffer.Next;
  end;
  {}
  if aTotalErr then Exit;
  {}
  { 啟動交易 }
  aComps := aTempComps;
  while ( aComps <> EmptyStr ) do
  begin
    aSoInfo := GetActiveSo( ExtractValue( aComps ) );
    if not aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.BeginTrans;
  end;
  { 存檔 }
  try
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      if ( FDataBuffer.FieldByName( 'HFCMAC' ).AsString <>
           FDataBuffer.FieldByName( 'BACKHFCMAC' ).AsString ) then
      begin
        UpdateSoRecord( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
      end;
      FDataBuffer.Next;
    end;
    { Commit }
    aComps := aTempComps;
    while ( aComps <> EmptyStr ) do
    begin
      aSoInfo := GetActiveSo( ExtractValue( aComps ) );
      if aSoInfo.Connection.InTransaction then
        aSoInfo.Connection.CommitTrans;
    end;
    Result := True;
  except
    on E: Exception do
    begin
      { Rollback }
      aComps := aTempComps;
      while ( aComps <> EmptyStr ) do
      begin
        aSoInfo := GetActiveSo( ExtractValue( aComps ) );
        if aSoInfo.Connection.InTransaction then
          aSoInfo.Connection.RollbackTrans;
      end;
      Result := False;
      aMsg := Format( '-1: CC&維修送會資料庫存檔失敗,錯誤訊息:<%s>', [E.Message] );
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

initialization
  TAutoObjectFactory.Create(ComServer, TCM2CT, Class_CM2CT,
    ciMultiInstance, tmApartment);
end.
