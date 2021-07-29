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

  { ��ƪ����ˬd }

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

  { �t�Υx��T }

  PSoInfo = ^TSoInfo;
  TSoInfo = record
    CompCode: String;                  { �t�Υx�N�X }
    CompName: String;                  { �t�Υx�W�� }
    LoginUser: String;                 { ��Ʈw�b�� }
    LoginPass: String;                 { ��Ʈw�K�X }
    DbAliase: String;                  { ��Ʈw�O�W }
    Connection: TADOConnection;
    Reader: TADOQuery;
    Writer: TADOQuery;
    OraSession: TOraSession;
    LogWriter: TOraSQL;
  end;


  { CM �J�w�]���� }
  
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

  { ��h�� }

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
    { �J�w���� }
    function ValidateXML1(aXmlData: String): Boolean;
    { ��h���� }
    function ValidateXML2(aXmlData: String): Boolean;
    { �ռ����� }
    function ValidateXML3(aXmlData: String): Boolean;
    { ��h������ }
    function ValidateXML4(aXmlData: String): Boolean;
    { ���װe�^���� }
    function ValidateXML5(aXmlData: String): Boolean;
    { �J�w }
    function SaveXML1: Boolean;
    { ��h }
    function SaveXML2: Boolean;
    { �ռ� }
    function SaveXML3: Boolean;
    { ��h�� }
    function SaveXML4: Boolean;
    { ���װe�^ }
    function SaveXML5: Boolean;
  public
    procedure Initialize; override;
    destructor Destroy; override;
    { CM �J�w/��h }
    function CM2CT_DLL1(aData1, aData2: OleVariant): WideString; safecall;
    { CM �ռ� }
    function CM2CT_DLL2(aData1: OleVariant): WideString; safecall;
    { CM ��h�� }
    function CM2CT_DLL3(aData1: OleVariant): WideString; safecall;
    { CM ���װe�^ }
    function CM2CT_DLL5(aData1: OleVariant): WideString; safecall;
  end;


implementation


{ LbCipher, LbString --> Turbopower ���[�ѱK Library ( LockBox ) }

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
  { DES �[�ѱK�Ϊ� Key --> CS (�j�g�r��) }
  FPassPhrase := Chr( 67 ) + Chr( 83 );
  { �Υ[�ѱKFPassPhrase�Ȳ��ͤ@�ոѱK�Ϊ� array ���c }
  GenerateLMDKey( aKey, SizeOf( aKey ), FPassPhrase );
  for aIndex := 0 to 20 do
  begin
    New( aSoInfo );
    aSoInfo.CompCode := IntToStr( aIndex );
    { �ϥ� DES �ѱK�^�쥻�r��, �קK�Q�_��Ū�X }
    { �H�U�O EMC ������ưϪ���Ʈw�s�u�r��, �Y�s�u�r������ }
    { �h������L��, ���s Buid �@���A���_���ϥ� }
    case aIndex of
      { �`�� com/com@emc }
      0: begin
           aSoInfo.CompName := DESEncryptStringEx( 'GK6rPSGTehY=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
      { �[�@ emcks/emc5601ks@emc6}
      1: begin
           aSoInfo.CompName := DESEncryptStringEx( 'rGpA+1maUhg=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'qKvF5yaDktg=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'oa/TVQdZIlHk3oz3EzOLkw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'fkO5zCYBdJI=', aKey, False );
         end;
      { �̫n emcpn/emc6202pn@emc6}
      2: begin
           aSoInfo.CompName := DESEncryptStringEx( 'akhHiBFhS9I=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'f2CQAghzZ5s=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '44F1CiX5YaflkqSw2nKDxg==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'fkO5zCYBdJI=', aKey, False );
         end;
      { �n�� emcnt/emc6403nt@emc5 }
      3: begin
           aSoInfo.CompName := DESEncryptStringEx( 'qfEoE8BXGhQ=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'eR5muOd5ncQ=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'wfb7rptxl5j9onZBwB6okQ==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MUL6BtuWyt4=', aKey, False );
         end;
      { �s�W�D emcncc/emc6105ncc@emc4}
      5: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Zr1PBwa8Byw=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'vfd+I/JKQwQ=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '0k3E8Wv+LtKFbBsej4DGrw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'jixfiCHuC34=', aKey, False );
         end;
      { �׷� emcfm/emc1906fm@emc4 }
      6: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Ah8M/NBq8F8=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'SiQSz4YRrEA=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '7Wf24kRroFQ/ZoOJ2rU25A==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'jixfiCHuC34=', aKey, False );
         end;
      { ���D emcct/emc5707ct@emc3 }
      7: begin
           aSoInfo.CompName := DESEncryptStringEx( '3Vysqp7t/kg=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'n86eGJ+19io=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'Ui+Zp/vzI7H9onZBwB6okQ==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'lqH+cXxXoVM=', aKey, False );
         end;
      { ���p emcuc/emc5508uc@emc1}
      8: begin
           aSoInfo.CompName := DESEncryptStringEx( 'xuL0QY75BFU=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'fxaKJE3mXRU=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'O+x1I/fitrq0xgnUHu0CqA==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'mf5f/EFTrNE=', aKey, False );
         end;
      { �����s emcyms/emc5909yms@emc }
      9: begin
           aSoInfo.CompName := DESEncryptStringEx( 'awStqsVwxTA=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( '6Dtr2sVE7eU=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '1wW0aTDGGC3T4qhwC249LA==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { �s�x�_ emcntp/emc6010ntp@emc }
     10: begin
           aSoInfo.CompName := DESEncryptStringEx( 'mzIxRa4Iq2s=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'fZPrFuqCIU4=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'bLKoJ8ryZnSpUcfLMZxZLw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { ���W�D emckp/emc2511kp@emc }
     11: begin
           aSoInfo.CompName := DESEncryptStringEx( '6tW9Hl1M7R4=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'dbhSjDnsLqI=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '2osHY0C6C+p0A5aiCm2nOw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { �j�w��s emcdw/emc5812dw@emc }
     12: begin
           aSoInfo.CompName := DESEncryptStringEx( '8uFrustgM9ZrZ2Kx7hdylw==', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'dMHHX3Ucjbs=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'vjKNxtn+QScOsJK+bY9iIw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { �s�� emctc/emctc@emc }
     13: begin
           aSoInfo.CompName := DESEncryptStringEx( '2CT3OdYNWHM=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'lY3LUuztcSE=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'lY3LUuztcSE=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'oIdxyXylQI0=', aKey, False );
         end;
     { �j�s�� emcst/emcst@emc2 }
     14: begin
           aSoInfo.CompName := DESEncryptStringEx( 'nKNwWhHJnJQ=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'w9XKCpLq4iE=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'w9XKCpLq4iE=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'Kw1slo9ivtU=', aKey, False );
         end;
     { �_��� emcnty/emc2916nty@emc7 }
     16: begin
           aSoInfo.CompName := DESEncryptStringEx( 'BEsOx4FNwcI=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'HmFB29lZPTg=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( '2uubyluDZQ2pIcIsdVqyIA==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'bO/fSqEvXk0=', aKey, False );
         end;
     { �� emctc_cm/emctc_cm@emc7, �𫰪� CM �~�ȿW�@�ߤ@���t�Υx��, �t�Υx�N�X�� 17, �ª��N�X 13 ���|�Ψ� }
     { 2007/12/25 �s�𫰸�Ʈw�^�k�_���x, ��^�έ쥻���t�Υx�N�X }
     17: begin
           aSoInfo.CompName := DESEncryptStringEx( '2CT3OdYNWHM=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'xc2y/MAu33BrZ2Kx7hdylw==', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'xc2y/MAu33BrZ2Kx7hdylw==', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'bO/fSqEvXk0=', aKey, False );
         end;
     { �ܺ޴��ո�Ʈw kptest/kptest@emc }
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
    { �g Log ��, �ϥ� Oracle Access Component ODAC �Ӽg CLOB Column }
    { Log �@�߼g���`�� }
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
  { DES �[�ѱK�Ϊ� Key --> CS (�j�g�r��) }
  FPassPhrase := Chr( 67 ) + Chr( 83 );
  { �Υ[�ѱKFPassPhrase�Ȳ��ͤ@�ոѱK�Ϊ� array ���c }
  GenerateLMDKey( aKey, SizeOf( aKey ), FPassPhrase );
  for aIndex := 0 to 20 do
  begin
    New( aSoInfo );
    aSoInfo.CompCode := IntToStr( aIndex );
    { �ϥ� DES �ѱK�^�쥻�r��, �קK�Q�_��Ū�X }
    { �H�U�O EMC ������ưϪ���Ʈw�s�u�r��, �Y�s�u�r������ }
    { �h������L��, ���s Buid �@���A���_���ϥ� }
    case aIndex of
      { �`�� COM/COM@RDKNET }
      0: begin
           aSoInfo.CompName := DESEncryptStringEx( 'GK6rPSGTehY=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'Rtg7wUpjct4=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MnXBjY2nyV4=', aKey, False );
         end;
      { �s�W�D NCCTEST/NCCTEST@CATVO }
      5: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Zr1PBwa8Byw=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'QqLvWdgA/ek=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'QqLvWdgA/ek=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'HnhysRH/TXc=', aKey, False );
         end;
      { �׷� EMCFM/EMCFM@RDKNET }
      6: begin
           aSoInfo.CompName := DESEncryptStringEx( 'Ah8M/NBq8F8=', aKey, False );
           aSoInfo.LoginUser := DESEncryptStringEx( 'bGiEDA+FVOc=', aKey, False );
           aSoInfo.LoginPass := DESEncryptStringEx( 'bGiEDA+FVOc=', aKey, False );
           aSoInfo.DbAliase := DESEncryptStringEx( 'MnXBjY2nyV4=', aKey, False );
         end;   
      { �_��� EMCNTY/EMCNTY@MIS }
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
    { �g Log ��, �ϥ� Oracle Access Component ODAC �Ӽg CLOB Column }
    { Log �@�߼g���`�� }
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
    { �J�w�g�J SO306 �� }
    FDataBuffer.FieldDefs.Add( 'COMPCODE', ftInteger );       { ���q�O ( �դJ ) }
    FDataBuffer.FieldDefs.Add( 'OLDCOMPCODE', ftInteger );    { ���q�O ( �եX ) }    
    FDataBuffer.FieldDefs.Add( 'BATCHNO', ftString, 15 );     { �帹 }
    FDataBuffer.FieldDefs.Add( 'MATERIALNO', ftString, 15 );  { �Ƹ� }
    FDataBuffer.FieldDefs.Add( 'CARTONNUNO', ftInteger );     { �Ƚc�Ǹ� }
    FDataBuffer.FieldDefs.Add( 'MODELNO', ftString, 30 );     { ����/�~�W }
    FDataBuffer.FieldDefs.Add( 'CHASSISSN', ftString, 30 );   { �򪩧Ǹ� }
    FDataBuffer.FieldDefs.Add( 'CUSTOMERSN', ftString, 30 );  { �����Ǹ� }
    FDataBuffer.FieldDefs.Add( 'HFCMAC', ftString, 20 );      { HFCMAC �Ǹ� }
    FDataBuffer.FieldDefs.Add( 'ETHERNETMAC', ftString, 20 ); { Ethernet MAC �Ǹ� }
    FDataBuffer.FieldDefs.Add( 'MTAMAC', ftString, 20 );      { MTAMAC �Ǹ� }
    FDataBuffer.FieldDefs.Add( 'FLAG', ftInteger );           { �H���@�� MAC �Ǹ����D, �w�] 1 }
    FDataBuffer.FieldDefs.Add( 'CMISSUE', ftString, 20 );     { CM �פJ���, ���榹 DLL �����, �~���ɤ��� }
    FDataBuffer.FieldDefs.Add( 'MODEID', ftInteger );         { �]�ƫ���, 0=EMTA, 1=CM }
    FDataBuffer.FieldDefs.Add( 'MODELCODE', ftInteger );      { �]�ƫ����N�X }
    FDataBuffer.FieldDefs.Add( 'ACTSTATUS', ftString, 6 );    { SPM �]�w���A }
    FDataBuffer.FieldDefs.Add( 'USEFLAG', ftInteger );        { �i�Ϊ��A, 0=�i��, 1=���i�� }
    FDataBuffer.FieldDefs.Add( 'LIMITDATE', ftString, 10 );   { �O�T���� , �~��� }
    FDataBuffer.FieldDefs.Add( 'BELONG', ftInteger );         { �]���W�ݳ��, 0=EMC, 1=APTG }
    FDataBuffer.FieldDefs.Add( 'VENDORCODE', ftString, 5 );   { �t�ӥN�X }
    FDataBuffer.FieldDefs.Add( 'VENDORNAME', ftString, 20 );  { �t�ӦW�� }
    FDataBuffer.FieldDefs.Add( 'SPEC', ftString, 50 );        { �W�� }
    FDataBuffer.FieldDefs.Add( 'DESCRIPTION', ftString, 30 ); { �ȪA�~�W }
    FDataBuffer.FieldDefs.Add( 'UPDEN', ftString, 20 );       { ���ʤH�� }
    FDataBuffer.FieldDefs.Add( 'UPDTIME', ftDateTime );       { ���ʮɶ� }
    FDataBuffer.FieldDefs.Add( 'ALREADYUSE', ftDateTime );    { �����ϥ� }
    FDataBuffer.FieldDefs.Add( 'MATERIALTYPE', ftString, 50 ); { ���ƫ��� }
    FDataBuffer.FieldDefs.Add( 'HWVER', ftString, 10 );       { �w�骩�� }
    {}
    FDataBuffer.FieldDefs.Add( 'SENDMTAMAC', ftString, 20 );  { ���� EMTAMAC�Ǹ� }
    FDataBuffer.FieldDefs.Add( 'BACKHFCMAC', ftString, 20 );  { �e�^ HFCMAC�Ǹ� }
    FDataBuffer.FieldDefs.Add( 'BACKMTAMAC', ftString, 20 );  { �e�^ EMTAMAC�Ǹ� }
    {}
  end;

  { --------------------------------------------------- }

  procedure InitBufferField2;
  begin
    { ��h�g�J SO310 �� }
    FDataBuffer.FieldDefs.Add( 'COMPCODE', ftInteger );       { ���q�O }
    FDataBuffer.FieldDefs.Add( 'BATCHNO', ftString, 15 );     { ��h�渹 }
    FDataBuffer.FieldDefs.Add( 'MATERIALNO', ftString, 15 );  { �Ƹ� }
    FDataBuffer.FieldDefs.Add( 'OLDBATCHNO', ftString, 15 );  { ��J�w�帹 }
    FDataBuffer.FieldDefs.Add( 'CARTONNUNO', ftInteger );     { �Ƚc�Ǹ� }
    FDataBuffer.FieldDefs.Add( 'MODELNO', ftString, 30 );     { ����/�~�W }
    FDataBuffer.FieldDefs.Add( 'CHASSISSN', ftString, 30 );   { �򪩧Ǹ� }
    FDataBuffer.FieldDefs.Add( 'CUSTOMERSN', ftString, 30 );  { �����Ǹ� }
    FDataBuffer.FieldDefs.Add( 'HFCMAC', ftString, 20 );      { HFCMAC �Ǹ� }
    FDataBuffer.FieldDefs.Add( 'ETHERNETMAC', ftString, 20 ); { Ethernet MAC �Ǹ� }
    FDataBuffer.FieldDefs.Add( 'MTAMAC', ftString, 20 );      { MTAMAC �Ǹ� }
    FDataBuffer.FieldDefs.Add( 'FLAG', ftInteger  );          { �H���@�� MAC �Ǹ����D, �w�] 1 }
    FDataBuffer.FieldDefs.Add( 'CMISSUE', ftString, 20 );     { CM �פJ���, ���榹 DLL �����, �~���ɤ��� }
    FDataBuffer.FieldDefs.Add( 'MODEID', ftInteger );         { �]�ƫ���, 0=EMTA, 1=CM }
    FDataBuffer.FieldDefs.Add( 'MODELCODE', ftInteger );      { �]�ƫ����N�X }
    FDataBuffer.FieldDefs.Add( 'ACTSTATUS', ftString, 6 );    { SPM �]�w���A }
    FDataBuffer.FieldDefs.Add( 'LIMITDATE', ftString, 10 );   { �O�T���� , �~��� }
    FDataBuffer.FieldDefs.Add( 'BELONG', ftInteger );         { �]���W�ݳ��, 0=EMC, 1=APTG }
    FDataBuffer.FieldDefs.Add( 'VENDORCODE', ftString, 5 );   { �t�ӥN�X }
    FDataBuffer.FieldDefs.Add( 'VENDORNAME', ftString, 20 );  { �t�ӦW�� }
    FDataBuffer.FieldDefs.Add( 'SPEC', ftString, 50 );        { �W�� }
    FDataBuffer.FieldDefs.Add( 'UPDEN', ftString, 20 );       { ���ʤH�� }
    FDataBuffer.FieldDefs.Add( 'UPDTIME', ftDateTime );       { ���ʮɶ� }
    FDataBuffer.FieldDefs.Add( 'ALREADYUSE', ftDateTime );    { �����ϥ� }
    FDataBuffer.FieldDefs.Add( 'MATERIALTYPE', ftString, 50 ); { ���ƫ��� }
    FDataBuffer.FieldDefs.Add( 'HWVER', ftString, 10 );       { �w�骩�� }
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
      raise Exception.CreateFmt( '�L�����q�O%s�C', [aCompCode] )
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
    Result := Format( '-1: �PCC&B��Ʈw�s�u����,���~�T��:<%s>', [Result] );
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

{ �帹�ˬd }

function TCM2CT.VdBatchNo(aNode: IDOMNode): String;

  { ----------------------------------------------- }

  function GetErrorDescription1: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�帹�榡���~, �帹=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�渹�榡���~, �渹<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�帹���׿��~, �帹=<%s>,�J�w�帹�����׳̦h<%d�X>�C',
          [aNode.nodeValue, XML_BATCHNO_LEN] );
      2: Result := Format( '-2:��h�渹���׿��~, �帹=<%s>,��h�渹�����׳̦h<%d�X>�C',
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
      1: Result := Format( '-2:�J�w�帹���׿��~, �帹=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�帹���׿��~, �帹=<%s>�C', ['�ŭ�'] );
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
      1: Result := Format( '-2:�J�w���q�N�X�����`�ܥN�X<%s>,���~���q�N�X=<%s>�C',
          [XML_WAREHOUSEID, '�ŭ�'] );
      2: Result := Format( '-2:��h���q�N�X�����`�ܥN�X<%s>,���~���q�N�X=<%s>�C',
          [XML_WAREHOUSEID, '�ŭ�'] );
      3: Result := Format( '-2:�ռ����q�N�X��������,���~���q�N�X=<%s>�C',
          ['�ŭ�'] );
      4: Result := Format( '-2:��h�Ƥ��q�N�X��������,���~���q�N�X=<%s>�C',
          ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w���q�N�X�����`�ܥN�X<%s>,���~���q�N�X=<%s>�C',
          [XML_WAREHOUSEID, aNode.nodeValue] );
      2: Result := Format( '-2:��h���q�N�X�����`�ܥN�X<%s>,���~���q�N�X=<%s>�C',
          [XML_WAREHOUSEID, aNode.nodeValue] );
      3: Result := Format( '-2:�ռ����q�N�X�������Ī����q�N�X,���~���q�N�X=<%s>�C',
          [aNode.nodeValue] );
      4: Result := Format( '-2:��h�Ƥ��q�N�X���i���`�ܥN�X,���~���q�N�X=<%s>�C',
          [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription3: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�L���J�w���q�N�X=<%s>�C',
          [aNode.nodeValue] );
      2: Result := Format( '-2:�L����h���q�N�X=<%s>�C',
          [aNode.nodeValue] );
      3: Result := Format( '-2:�L���ռ����q�N�X=<%s>�C',
          [aNode.nodeValue] );
      4: Result := Format( '-2:�L����h�Ƥ��q�N�X=<%s>�C',
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
  { ���q�N�X�ŭ� }
  if ( aNode.nodeValue = EmptyWideStr ) then
  begin
    Result := GetErrorDescription1;
    Exit;
  end;
  { ���n��綠�q�O, Ex: aCompare = 1 }
  if ( aCompare <> EmptyStr ) then
  begin
    { ���p aCompare <> aNode.nodeValue, �J�w/��h }
    if ( aNode.nodeValue <> aCompare ) and ( FType in [1,2] ) then
    begin
      Result := GetErrorDescription2;
      Exit;
    end;
    { ���p aCompare = aNode.nodeValue, ��/�h�� }
    if ( aNode.nodeValue = aCompare ) and ( FType in [4] ) then
    begin
      Result := GetErrorDescription2;
      Exit;
    end;
  end;
  { ���礽�q�O�O�_�X�k }
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
      1: Result := Format( '-2:�J�w�Ƹ��榡���~,�Ƹ�=<%s>,�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�Ƹ��榡���~,�Ƹ�=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�Ƹ����׿��~,�Ƹ�=<%s>,�J�w�Ƹ������׳̦h<%d�X>�C',
          [aNode.nodeValue, XML_MATERIALNO_LEN] );
      2: Result := Format( '-2:��h�Ƹ����׿��~,�Ƹ�=<%s>,�J�w�Ƹ������׳̦h<%d�X>�C',
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
      1: Result := Format( '-2:�J�w�Ƹ��榡���~,�Ƹ�=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�Ƹ��榡���~,�Ƹ�=<%s>�C', ['�ŭ�'] );
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
      1: Result := Format( '-2:�J�w����榡���~,���=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h����榡���~,���=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w����榡���~,���=<%s>�C', [aNode.nodeValue] );
      2: Result := Format( '-2:��h����榡���~,���=<%s>�C', [aNode.nodeValue] );
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
      1: Result := Format( '-2:�J�w�ɶ��榡���~,�ɶ�=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�ɶ��榡���~,�ɶ�=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�ɶ��榡���~,�ɶ�=<%s>�C', [aNode.nodeValue] );
      2: Result := Format( '-2:��h�ɶ��榡���~,�ɶ�=<%s>�C', [aNode.nodeValue] );
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
      1: Result := Format( '-2:�J�w�t�ӥN�X�榡���~,�t�ӥN�X=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�t�ӥN�X�榡���~,�t�ӥN�X=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�t�ӥN�X���׿��~,�t�ӥN�X=<%s>,�t�ӥN�X�����׳̦h<%d�X>�C',
          [aNode.nodeValue, XML_VENDORCODE_LEN] );
      2: Result := Format( '-2:��h�t�ӥN�X���׿��~,�t�ӥN�X=<%s>,�t�ӥN�X�����׳̦h<%d�X>�C',
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
      1: Result := Format( '-2:�J�w�t�ӦW�ٮ榡���~,�t�ӦW��=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�t�ӦW�ٮ榡���~,�t�ӦW��=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�t�ӦW�٪��׿��~,�t�ӦW��=<%s>,�t�ӦW�٪����׳̦h<%d�X>�C',
          [aNode.nodeValue, XML_VENDORNAME_LEN] );
      2: Result := Format( '-2:��h�t�ӦW�٪��׿��~,�t�ӦW��=<%s>,�t�ӦW�٪����׳̦h<%d�X>�C',
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
      1: Result := Format( '-2:�J�w�ƶq�榡���~,�ƶq=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�ƶq�榡���~,�ƶq=<%s>�C', ['�ŭ�'] );
      3: Result := Format( '-2:�ռ��ƶq�榡���~,�ƶq=<%s>�C', ['�ŭ�'] );
      4: Result := Format( '-2:��h�Ƽƶq�榡���~,�ƶq=<%s>�C', ['�ŭ�'] );
      5: Result := Format( '-2:���װe�^�ƶq�榡���~,�ƶq=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�ƶq�榡���~,�ƶq=<%s>�C', [aNode.nodeValue] );
      2: Result := Format( '-2:��h�ƶq�榡���~,�ƶq=<%s>�C', [aNode.nodeValue] );
      3: Result := Format( '-2:�ռ��ƶq�榡���~,�ƶq=<%s>�C', [aNode.nodeValue] );
      4: Result := Format( '-2:��h�Ƽƶq�榡���~,�ƶq=<%s>�C', [aNode.nodeValue] );
      5: Result := Format( '-2:���װe�^�ƶq�榡���~,�ƶq=<%s>�C', [aNode.nodeValue] );
    end;
  end;

  { ----------------------------------------------- }
  
  function GetErrorDescription3: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�ƶq���~,�J�w�ƶq=<%s>,��ک��ӵ���=<%d>�C',
          [aNode.nodeValue, aOwner.childNodes.length] );
      2: Result := Format( '-2:��h�ƶq���~,��h�ƶq=<%s>,��ک��ӵ���=<%d>�C',
          [aNode.nodeValue, aOwner.childNodes.length] );
      3: Result := Format( '-2:�ռ��ƶq���~,�ռ��ƶq=<%s>,��ک��ӵ���=<%d>�C',
          [aNode.nodeValue, aOwner.childNodes.length] );
      4: Result := Format( '-2:��h�Ƽƶq���~,��h�Ƽƶq=<%s>,��ک��ӵ���=<%d>�C',
          [aNode.nodeValue, aOwner.childNodes.length] );
      5: Result := Format( '-2:���װe�^�ƶq���~,���װe�^�ƶq=<%s>,��ک��ӵ���=<%d>�C',
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
      1: Result := Format( '-2:�J�w�~�W�榡���~,�~�W=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�~�W�榡���~,�~�W=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�~�W���׿��~,�~�W=<%s>,�̦h�~�W����=<%d>�C',
          [aNode.nodeValue, XML_DESCRIPTION_LEN] );
      2: Result := Format( '-2:��h�~�W���׿��~,�~�W=<%s>,�̦h�~�W����=<%d>�C',
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
      1: Result := Format( '-2:�J�w�ȪA�~�W�榡���~,�ȪA�~�W=<%s>�C', ['�ŭ�'] );
      3: Result := Format( '-2:�ռ��ȪA�~�W�榡���~,�ȪA�~�W=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�ȪA�~�W���׿��~,�ȪA�~�W=<%s>,�̦h�ȪA�~�W����=<%d>�C',
          [aNode.nodeValue, XML_DESCRIPTION_LEN] );
      3: Result := Format( '-2:�ռ��ȪA�~�W���׿��~,�ȪA�~�W=<%s>,�̦h�ȪA�~�W����=<%d>�C',
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
      1: Result := Format( '-2:�J�w�W����׿��~,�W��=<%s>,�̦h�W�����=<%d�X>�C',
          [aNode.nodeValue, XML_SPEC_LEN] );
      2: Result := Format( '-2:��h�W����׿��~,�W��=<%s>,�̦h����=<%d�X>�C',
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
      1: Result := Format( '-2:�J�w�C�b�����~,�C�b���=<%s>,�C�b��쥲���ŦX�U�C��<%s>�C',
          [aNode.nodeValue, ( '0,1' )] );
      2: Result := Format( '-2:��h�C�b�����~,�C�b���=<%s>,�C�b��쥲���ŦX�U�C��<%s>�C',
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
      1: Result := Format( '-2:�J�w�]�Ʈ榡���~,�]�Ʈ榡���e=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��h�]�Ʈ榡���~,�]�Ʈ榡���e=<%s>�C', ['�ŭ�'] );
      3: Result := Format( '-2:�ռ��]�Ʈ榡���~,�]�Ʈ榡���e=<%s>�C', ['�ŭ�'] );
      5: Result := Format( '-2:���װe�^�]�Ʈ榡���~,�]�Ʈ榡���e=<%s>�C', ['�ŭ�'] );
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
      1: Result := Format( '-2:�J�w�]�ƥD�Ǹ��O�榡���~,�D�Ǹ��O=<%s>, HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
      2: Result := Format( '-2:��h�]�ƥD�Ǹ��O�榡���~,�D�Ǹ��O=<%s>, HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�]�ƥD�Ǹ��O�榡���~,�D�Ǹ��O=<%s>,�D�Ǹ��O�����ŦX�U�C��<%s>, HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, '1', aId] );
      2: Result := Format( '-2:��h�]�ƥD�Ǹ��O�榡���~,�D�Ǹ��O=<%s>,�D�Ǹ��O�����ŦX�U�C��<%s>, HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, '1', aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�wHFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��hHFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
      3: Result := Format( '-2:�ռ�HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
      4: Result := Format( '-2:��/�h��HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
      5: Result := Format( '-2:���װe�^HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�wHFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>,HFCMAC�Ǹ��̤j����<%d�X>�C',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      2: Result := Format( '-2:��hHFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>,HFCMAC�Ǹ��̤j����<%d�X>�C',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      3: Result := Format( '-2:�ռ�HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>,HFCMAC�Ǹ��̤j����<%d�X>�C',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      4: Result := Format( '-2:��/�h��HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>,HFCMAC�Ǹ��̤j����<%d�X>�C',
          [aNode.nodeValue, XML_HFCMAC_LEN] );
      5: Result := Format( '-2:���װe�^HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>,HFCMAC�Ǹ��̤j����<%d�X>�C',
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
      1: Result := Format( '-2:�J�wHFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
      2: Result := Format( '-2:��hHFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
      3: Result := Format( '-2:�ռ�HFCMAC�Ǹ��榡���~,HFCMAC�Ǹ�=<%s>�C', ['�ŭ�'] );
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
      1: Result := Format( '-2:�J�wEthernetMAC�Ǹ��榡���~,EthernetMAC�Ǹ�=<%s>,EthernetMAC�Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_ETHERNETMAC_LEN, aId] );
      2: Result := Format( '-2:��hEthernetMAC�Ǹ��榡���~,EthernetMAC�Ǹ�=<%s>,EthernetMAC�Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_ETHERNETMAC_LEN, aId] );
      3: Result := Format( '-2:�ռ�EthernetMAC�Ǹ��榡���~,EthernetMAC�Ǹ�=<%s>,EthernetMAC�Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_ETHERNETMAC_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�wMTAMAC�Ǹ��榡���~,MTAMAC�Ǹ�=<%s>,MTAMAC�Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_MTAMAC_LEN, aId] );
      2: Result := Format( '-2:��hMTAMAC�Ǹ��榡���~,MTAMAC�Ǹ�=<%s>,MTAMAC�Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_MTAMAC_LEN, aId] );
      3: Result := Format( '-2:�ռ�MTAMAC�Ǹ��榡���~,MTAMAC�Ǹ�=<%s>,MTAMAC�Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_MTAMAC_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�w�򪩧Ǹ��榡���~,�򪩧Ǹ�=<%s>,�򪩧Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_CHASSISSN_LEN, aId] );
      2: Result := Format( '-2:��h�򪩧Ǹ��榡���~,�򪩧Ǹ�=<%s>,�򪩧Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_CHASSISSN_LEN, aId] );
      3: Result := Format( '-2:�ռ��򪩧Ǹ��榡���~,�򪩧Ǹ�=<%s>,�򪩧Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_CHASSISSN_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�w�����Ǹ��榡���~,�����Ǹ�=<%s>,�����Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_CUSTOMERSN_LEN, aId] );
      2: Result := Format( '-2:��h�����Ǹ��榡���~,�����Ǹ�=<%s>,�����Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_CUSTOMERSN_LEN, aId] );
      3: Result := Format( '-2:�ռ������Ǹ��榡���~,�����Ǹ�=<%s>,�����Ǹ��̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_CUSTOMERSN_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�w�]�ƫ����榡���~,�]�ƫ���=<%s>,HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
      2: Result := Format( '-2:��h�]�ƫ����榡���~,�]�ƫ���=<%s>,HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
      3: Result := Format( '-2:�ռ��]�ƫ����榡���~,�]�ƫ���=<%s>,HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�]�ƫ����榡���~,�]�ƫ���=<%s>,�]�ƫ��������ŦX�U�C��<%s>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, '0,1,2', aId] );
      2: Result := Format( '-2:��h�]�ƫ����榡���~,�]�ƫ���=<%s>,�]�ƫ��������ŦX�U�C��<%s>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, '0,1,2', aId] );
      3: Result := Format( '-2:�ռ��]�ƫ����榡���~,�]�ƫ���=<%s>,�]�ƫ��������ŦX�U�C��<%s>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, '0,1,2', aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�w�]�ƾ��خ榡���~,�]�ƾ���=<%s>,HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
      2: Result := Format( '-2:��h�]�ƾ��خ榡���~,�]�ƾ���=<%s>,HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
      3: Result := Format( '-2:�ռ��]�ƾ��خ榡���~,�]�ƾ���=<%s>,HFCMAC�Ǹ�=<%s>�C',
        ['�ŭ�', aId] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      1: Result := Format( '-2:�J�w�]�ƾ��خ榡���~,����=<%s>,���س̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_MODELNO_LEN, aId] );
      2: Result := Format( '-2:��h�]�ƾ��خ榡���~,����=<%s>,���س̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_MODELNO_LEN, aId] );
      3: Result := Format( '-2:�ռ��]�ƾ��خ榡���~,����=<%s>,���س̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
          [aNode.nodeValue, XML_MODELNO_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      1: Result := Format( '-2:�J�w�]�ƫO�T�����榡���~,�O�T����=<%s>,HFCMAC�Ǹ�=<%s>�C',
        [aNode.nodeValue, aId] );
      2: Result := Format( '-2:��h�]�ƫO�T�����榡���~,�O�T����=<%s>,HFCMAC�Ǹ�=<%s>�C',
        [aNode.nodeValue, aId] );
      3: Result := Format( '-2:�ռ��]�ƫO�T�����榡���~,�O�T����=<%s>,HFCMAC�Ǹ�=<%s>�C',
        [aNode.nodeValue, aId] );
    end;
  end;

  { ----------------------------------------------- }

var
  aYear, aMonth, aDay: String;
begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      4: Result := Format( '-2:��h���ѧO�榡���~,�ѧO�榡=<%s>�C', ['�ŭ�'] );
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2: String;
  begin
    Result := EmptyStr;
    case FType of
      4: Result := Format( '-2:��h���ѧO�榡���~,�ѧO�榡=<%s>,�ѧO�榡�����ŦX�U�C��<%s>�C',
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
        '-2:�J�w�]�ƪ��ƫ����榡���~,���ƫ���=<%s>,���ƫ����榡�̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
        [aNode.nodeValue, XML_MATERIALTYPE_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
        '-2:�J�w�]�Ƶw�骩���榡���~,�w�骩��=<%s>,�w�骩���榡�̤j����<%d�X>,HFCMAC�Ǹ�=<%s>�C',
        [aNode.nodeValue, XML_HWVER_LEN, aId] );
    end;
  end;

  { ----------------------------------------------- }

begin
  Result := EmptyStr;
  if ( aId = EmptyStr ) then aId := '�ŭ�';
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
      2: Result := Format( '-2:��h��Ʈ榡���~,��h�榡=<%s>�C', ['�ŭ�'] );
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
      3: Result := Format( '-2:�ռ���Ʈ榡���~,�ռ��榡=<%s>�C', ['�ŭ�'] );
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
      3: Result := Format( '-2:��h�Ƹ�Ʈ榡���~,��h�Ʈ榡=<%s>�C', ['�ŭ�'] );
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
      5: Result := Format( '-2:���װe�^��Ʈ榡���~,���װe�^�榡=<%s>�C', ['�ŭ�'] );
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
      5: Result := Format( '-2:���װe�^��Ʈ榡���~,�Ǹ����e=<%s>�C', ['�ŭ�'] );
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
      5: Result := '-2:���װe�^��Ʈ榡���~,�Ǹ����e������=<�e��>��<��^>�C';
    end;
  end;

  { ----------------------------------------------- }

  function GetErrorDescription2(aName: String): String;
  begin
    Result := EmptyStr;
    if ( aName = '�e��' ) then
      aName := '��^'
    else
      aName := '�e��';
    case FType of
      5: Result := Format( '-2:���װe�^��Ʈ榡���~,�Ǹ����e�L<%s>��ơC', [aName] );
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
      1: Result := Format( '-2:�w�ռ��ܤ��q�O=<%s %s>,HFCMAC�Ǹ�=<%s>���i�A���J�w�C',
          [aOtherSoId, aCompName, aId] );
      2: Result := Format( '-2:�w�ռ��ܤ��q�O=<%s %s>,HFCMAC�Ǹ�=<%s>���i��h�C',
          [aOtherSoId, aCompName, aId] );
      3: Result := Format( '-2:�w�ռ��ܤ��q�O=<%s %s>,HFCMAC�Ǹ�=<%s>���i�A���ռ��C',
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
    { �w�ռ��ܨ䥦�t�Υx }
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
      1: Result := Format( '-2:���q�O<%s %s>���]�Ƥ��s�b,�L�k�J�w,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      2: Result := Format( '-2:���q�O<%s %s>���]�Ƥ��s�b,�L�k���P/��h,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      3: Result := Format( '-2:���q�O<%s %s>���]�Ƥ��s�b,�L�k�ռ�,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      4: Result := Format( '-2:���q�O<%s %s>���]�Ƥ��s�b,�L�k��/�h��,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      5: Result := Format( '-2:���q�O<%s %s>���]�Ƥ��s�b,�L�k����/��^,HFCMAC�Ǹ�=<%s>�C',
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
      1: Result := Format( '-2:���q�O<%s %s>���]�Ƥw�ϥΤ��L�k�J�w,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      2: Result := Format( '-2:���q�O<%s %s>���]�Ƥw�ϥΤ��L�k��h,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      3: Result := Format( '-2:���q�O<%s %s>���]�Ƥw�ϥΤ��L�k�ռ�,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aId] );
      4: Result := Format( '-2:���q�O<%s %s>���]�Ƥw�ϥΤ��L�k��/�h��,HFCMAC�Ǹ�=<%s> �C',
          [aCompCode, aCompName, aId] );
      5: Result := Format( '-2:���q�O<%s %s>���]�Ƥw�ϥΤ��L�k����/�N�^,HFCMAC�Ǹ�=<%s> �C',
          [aCompCode, aCompName, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  { �ˮ֬O�_ "�ϥΤ�" ? ( �Y���q�O���`�ܫh�������ˮ� )  }
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
    { �h���ˮ� "�ϥΤ�" �S��W�h }
    { �s�b�� SO004 ���Ȥ�]���ɤ�, ���w�ˤ��, �B�L����, ���� "�ϥΤ�" }
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
      { ���M����, ���˾���, ���]��, �i�O�˾���j����, �]���� "�ϥΤ�" }
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
      1: Result := Format( '-2:���q�O<%s %s>���]�Ƥw��,��]:<%s>,�L�k�J�w,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aDesc, aId] );
      2: Result := Format( '-2:���q�O<%s %s>���]�Ƥw��,��]:<%s>,�L�k��h,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aDesc, aId] );
      3: Result := Format( '-2:���q�O<%s %s>���]�Ƥw��,��]:<%s>,�L�k�ռ�,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aDesc, aId] );
      4: Result := Format( '-2:���q�O<%s %s>���]�Ƥw��,��]:<%s>,�L�k��/�h��,HFCMAC�Ǹ�=<%s>�C',
          [aCompCode, aCompName, aDesc, aId] );
      5: Result := Format( '-2:���q�O<%s %s>���]�Ƥw��,��]:<%s>,�L�k����/�e�^,HFCMAC�Ǹ�=<%s>�C',
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
  { �ˮ֬O�_ "�w�Q���P" ? ( �Y���q�O���`�ܫh�������ˮ� )  }
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
    { �s�b�� SO004 ���Ȥ�]���ɤ�, UseFlag = 2,  ���� "���P" }
    if ( not aSoInfo.Reader.IsEmpty ) then
    begin
      aUseFlag := aSoInfo.Reader.FieldByName( 'USEFLAG' ).AsInteger;
      aSoInfo.Reader.Close;
      if ( aUseFlag = 2 ) then
      begin
        aSoInfo.Reader.SQL.Clear;
        aSoInfo.Reader.SQL.Text :=
          ' SELECT ( NVL( TO_CHAR( DEREASONCODE ), ''�ŭ�'' ) || '','' || NVL( DEREASONNAME, ''�ŭ�'' ) ) DEREASON ' +
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
      3: Result := Format( '-2:���q�O<%s %s>���]�Ƭd�L�������ث���,����=<%s>, HFCMAC�Ǹ�=<%s> �L�k�ռ��C',
          [aCompCode, aCompName, aModelNo, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  aModeCode := EmptyStr;
  Result := EmptyStr;
  { �դJ���O�`��, ���ˮ� }
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
      if ( aModelNo = EmptyStr ) then aModelNo := '�ŭ�';
      Result := GetErrorDescription;
      Exit;
    end;
    aModeCode := aSoInfo.Reader.FieldByName( 'CODENO' ).AsString;
  finally
    aSoInfo.Reader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ �ռ��ɩI�s���禡, ���ҳs���� DB �O�_�O�Өt�Υx�� DB, �w���{�����~
  �g����O�a�t�Υx }

function TCM2CT.VdCD039(aCompCode: String): String;

  { ----------------------------------------------------- }

  function GetErrorDescription: String;
  begin
    Result := EmptyStr;
    case FType of
      3: Result := Format( '-1: CC&B�ռ���Ʈw�s�ɥ���,���~�T��:<%s>�C',
          ['�ռ����q�N�X�P��ڸ�Ʈw����'] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  { �`��, ���ˮ� }
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
      3: Result := Format( '-2:���q�O<%s %s>���]�Ƭd�L�����ȪA�~�W,�ȪA�~�W=<%s>, HFCMAC�Ǹ�=<%s> �L�k�ռ��C',
          [aCompCode, aCompName, aDesc, aId] );
    end;
  end;

  { ----------------------------------------------------- }

var
  aSoInfo: PSoInfo;
begin
  Result := EmptyStr;
  { �դJ���O�`��, ���ˮ� }
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
      if ( aDesc = EmptyStr ) then aDesc := '�ŭ�';
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
    Result := Format( '-2:�J�w�]��MTAMAC�Ǹ����i���ŭ�,�ӳ]�ƫ�����<EMTA>, HFCMAC�Ǹ�=<%s>�C',
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
    Result := '-2:���װe�^�Ǹ��榡���~, <����>��<��^>��������X�{�C';
    Exit;
  end;
  if ( aSend.nodeName = '��^' ) then
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
    Result := Format( '-2:���q�O<%s %s>���׻�^���]�ƧǸ��w�s�b, HFCMAC�Ǹ�=<%s> ���i���׻�^�C',
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
    Result := Format( '-2:���׻�^���]�ƫ�������, MTAMAC�Ǹ����i��<�ŭ�>, HFCMAC=<%s>�C',
      [aHfc] );
    Exit;
  end;
  if ( aDbMtaMac = EmptyStr ) and ( aXmlMtaMac <> EmptyStr ) then
  begin
    Result := Format( '-2:���׻�^���]�ƫ�������, MTAMAC�Ǹ�������<�ŭ�>, HFCMAC=<%s>�C',
      [aXmlMtaMac, aHfc] );
    Exit;
  end;
  if ( aCompare ) then
  begin
    if ( Nvl( aDbMtaMac, 'x' ) <> Nvl( aXmlMtaMac, 'x' ) ) then
    begin
      aText := aDbMtaMac;
      if ( aXmlMtaMac <> EmptyStr ) then aText := aXmlMtaMac;
      Result := Format( '-2:���׻�^���]�ƫ�������, MTAMAC�Ǹ����~=<%s>, HFCMAC=<%s>�C',
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
  { �J�w }
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
      { �����ҤJ�w�榡�θ�� }
      aExcResult := ValidateXML1( aXmlData1 );
      if not aExcResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { ����OK, �g�J Table }
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
  { ��h }
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
      { ��������h�榡�θ�� }
      aExcResult := ValidateXML2( aXmlData2 );
      if not aExcResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { ����OK, �g�J Table }
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
  { �ռ� }
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
      { �����ҽռ��榡�θ�� }
      aExecResult := ValidateXML3( aXmlData1 );
      if ( not aExecResult ) then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { ����OK, �g�J Table }
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
  { ��/�h�� }
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
      { �����һ�/�h�Ʈ榡�θ�� }
      aExecResult := ValidateXML4( aXmlData1 );
      if not aExecResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { ����OK, �g�J Table }
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
  { ��/�h�� }
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
      { �����һ�/�h�Ʈ榡�θ�� }
      aExecResult := ValidateXML5( aXmlData1 );
      if not aExecResult then
      begin
        Result := WriteLog;
        Result := ( Result + MakeErrorText );
        Exit;
      end;
      { ����OK, �g�J Table }
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
    aMsg := Format( '-1:�ǤJ���J�w XML �榡���~, ���~��]:%s�C', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { �帹�ˬd }
  aBatchList := FXml.DOMDocument.getElementsByTagName( '�帹' );
  aMsg := VdBatchNo( aBatchList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { �v�@���X�ӧ帹�U����� }
  for aIndex := 0 to aBatchList.length - 1 do
  begin
    { ��l FSo306 ��J�ȵ��c }
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
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '���' );
    aMsg := VdDate( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.CMIssue := ConvertCDate( aNode.firstChild.nodeValue );
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '�ɶ�' );
    aMsg := VdTime( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.CMIssue := ( FSo306.CMIssue + #32 + ConvertTime( aNode.firstChild.nodeValue ) );
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '�t�ӥN�X' );
    aMsg := VdVendorCode( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo306.VendorCode := aNode.firstChild.nodeValue;
    {}
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '�t�ӦW��' );
    aMsg := VdVendorName( aNode.firstChild );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;  
    FSo306.VendorName := aNode.firstChild.nodeValue;
    {}
    { �J�w�@�w�O�`��, �`�ܥN�X�� 0 }
    aNode := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNode( '���q�N�X' );
    aMsg := VdCompCode( aNode.firstChild, '0' );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;  
    FSo306.CompCode := aNode.firstChild.nodeValue;
    {}
    aMaterialList := ( aBatchList.item[aIndex] as IDOMNodeSelect ).selectNodes( '�Ƹ�' );
    aMsg := VdMaterialNo( aMaterialList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { �v�@���X�ӧ帹�U�Ҧ��Ƹ���� }
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
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '�ƶq' );
      aMsg := VdItemCount( aNode, aMaterialList.item[aIndex2], True );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;  
      { �~�W�ȮɨS�Ψ�, �����ȥ��ӧ�, ���O���g Table }
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '�~�W' );
      aMsg := VdDescription( aNode );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      FSo306.Spec := EmptyStr;
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '�W��' );
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
      aNode := aMaterialList.item[aIndex2].attributes.getNamedItem( '�C�b���' );
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
      aEqumList := ( aMaterialList.item[aIndex2] as IDOMNodeSelect ).selectNodes( '�]��' );
      aMsg := VdEqum( aEqumList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      { ���X�Ҧ��]�� }
      for aIndex3 := 0 to aEqumList.length - 1 do
      begin
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( 'HFCMAC�Ǹ�' );
        FSo306.HfcMac :=  EmptyStr;
        aMsg := VdHFCMac( aNode.firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.HfcMac := aNode.firstChild.nodeValue;
        {}
        aNode := aEqumList.item[aIndex3].attributes.getNamedItem( '�D�Ǹ�' );
        aMsg := VdPrimaryMAC( aNode, FSo306.HfcMac );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.Flag := StrToInt( aNode.nodeValue );
        {}
        FSo306.EthernetMac := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( 'ETHERNETMAC�Ǹ�' );
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
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( 'MTAMAC�Ǹ�' );
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
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '�򪩧Ǹ�' );
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
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '�����Ǹ�' );
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
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '����' );
        aMsg := VdModelNo( aNode.firstChild, FSo306.HfcMac );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.ModelNo := aNode.firstChild.nodeValue;
        {}
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '����' );
        aMsg := VdModelId( aNode.firstChild, FSo306.HfcMac );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.ModeId := StrToInt( aNode.firstChild.nodeValue );
        { �O�T���� }
        FSo306.LimitDate := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '�O�T����' );
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
        { �ȪA�~�W }
        FSo306.Description := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '�ȪA�~�W' );
        aMsg := VdDescription2( aNode.firstChild );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( '0', aActionText, aMsg, 1 );
          aDetailError := True;
        end else
          FSo306.Description := aNode.firstChild.nodeValue;
        { ���ƫ��� }
        FSo306.MaterialType := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '���ƫ���' );
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
        { �w�骩�� }
        FSo306.HwVer := EmptyStr;
        aNode := ( aEqumList.item[aIndex3] as IDOMNodeSelect ).selectNode( '�w�骩��' );
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
        { �e�����S�o�Ϳ��~ }
        if not aDetailError then
        begin
          aMsg := VdSpecialRule1( FSo306 );
          if ( aMsg <> EmptyStr ) then
          begin
            AppendLog( '0', aActionText, aMsg, 1 );
            aDetailError := True;
          end;
        end;
        { �u�n��������~, �@�ߤ��s�� }
        if not aDetailError then
        begin
          FDataBuffer.Append;
          FDataBuffer.FieldByName( 'COMPCODE' ).Value := FSo306.CompCode ;      { ���q�O }
          FDataBuffer.FieldByName( 'BATCHNO' ).Value := FSo306.BatchNo;         { �帹 }
          FDataBuffer.FieldByName( 'MATERIALNO' ).Value := FSo306.MaterialNo;   { �Ƹ� }
          FDataBuffer.FieldByName( 'CARTONNUNO' ).Value := Null;                { �Ƚc�Ǹ� }
          FDataBuffer.FieldByName( 'MODELNO' ).Value := FSo306.ModelNo;         { ���� }
          FDataBuffer.FieldByName( 'CHASSISSN' ).Value := FSo306.ChassisSN;     { �򪩧Ǹ� }
          FDataBuffer.FieldByName( 'CUSTOMERSN' ).Value := FSo306.CustomerSN;   { �����Ǹ� }
          FDataBuffer.FieldByName( 'HFCMAC' ).Value := FSo306.HfcMac;           { HFCMAC �Ǹ� }
          FDataBuffer.FieldByName( 'ETHERNETMAC' ).Value := FSo306.EthernetMac; { Ethernet MAC �Ǹ� }
          FDataBuffer.FieldByName( 'MTAMAC' ).Value := FSo306.MtaMac;           { MTAMAC �Ǹ� }
          FDataBuffer.FieldByName( 'FLAG' ).Value := FSo306.Flag;               { �H���@�� MAC �Ǹ����D, �w�] 1 }
          FDataBuffer.FieldByName( 'CMISSUE' ).Value := FSo306.CMIssue;         { CM �פJ���, ���榹 DLL �����, �~���ɤ��� }
          FDataBuffer.FieldByName( 'MODEID' ).Value := FSo306.ModeId;           { �]�ƫ���, 0=EMTA, 1=CM, 2=WIFLY }
          FDataBuffer.FieldByName( 'MODELCODE' ).Value := Null;                 { �]�ƫ����N�X }
          FDataBuffer.FieldByName( 'ACTSTATUS' ).Value := EmptyStr;             { SPM �]�w���A }
          FDataBuffer.FieldByName( 'USEFLAG' ).Value := Null;                   { �i�Ϊ��A, 0=�i��, 1=���i�� }
          FDataBuffer.FieldByName( 'LIMITDATE' ).Value := FSo306.LimitDate;     { �O�T���� , �~��� }
          FDataBuffer.FieldByName( 'BELONG' ).Value := FSo306.Belong;           { �]���W�ݳ��, 0=EMC, 1=APTG }
          FDataBuffer.FieldByName( 'VENDORCODE' ).Value := FSo306.VendorCode;   { �t�ӥN�X }
          FDataBuffer.FieldByName( 'VENDORNAME' ).Value := FSo306.VendorName;   { �t�ӦW�� }
          FDataBuffer.FieldByName( 'SPEC' ).Value := FSo306.Spec;               { �W�� }
          FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString := FSo306.Description; { �ȪA�~�W }
          FDataBuffer.FieldByName( 'UPDEN' ).Value := '���ƨt��';               { ���ʤH�� }
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;              { ���ʮɶ� }
          FDataBuffer.FieldByName( 'MATERIALTYPE' ).AsString := FSo306.MaterialType; { ���ƫ��� }
          FDataBuffer.FieldByName( 'HWVER' ).AsString := FSo306.HwVer;          { �w�骩�� }
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC�Ǹ�=<%s>,���`�C', [FSo306.HfcMac] );
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
  { ��Ʈw�s�����`�� }
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
      { �ӵ� CM �O�_�w�ռ��ܨ䥦�t�Υx, �u�ˬd�ռ��Y�i,
        �����ˮ� SO004 �O�_�ϥΤ� }
      aMsg := VdAlreadyInOtherSo( '0', FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        aIsErr := True;
        AppendLog( '0', aActionText, aMsg, 1 );
      end else
      begin
        aMsg := Format( 'HFCMAC�Ǹ�<%s>,���`�C', [
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
        AppendLog( '0', aActionText, aMsg, 0 );
      end;
      if ( not aIsErr ) then
      begin
        { ����s�`�� Table, ��s����ηs�W }
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
      aMsg := Format( '-1: CC&B�J�w��Ʈw�s�ɥ���,���~�T��:<%s>', [E.Message] );
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
    aMsg := Format( '-1:�ǤJ����h XML �榡���~, ���~��]:%s�C', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  aBackList := FXml.DOMDocument.getElementsByTagName( '��h' );
  aMsg := VdBackList( aBackList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { �v�@���X��h�渹 }
  for aIndex := 0 to aBackList.length - 1 do
  begin
    { ��l FSo310 ��J�ȵ��c }
    ZeroMemory( FSo310, SizeOf( FSo310^ ) );
    {}
    aNode := aBackList.item[aIndex].attributes.getNamedItem( '�渹' );
    aMsg := VdBatchNo( aNode );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;  
    FSo310.BatchNo := aNode.nodeValue;
    {}
    aNode := aBackList.item[aIndex].attributes.getNamedItem( '���q�N�X' );
    aMsg := VdCompCode( aNode, '0' );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    FSo310.CompCode := aNode.nodeValue;
    {}
    aEqumList := ( aBackList.item[aIndex] as IDOMNodeSelect ).selectNodes( '�]��' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { �v�@���X�]�� }
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�ƶq' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      {} 
      aSeqList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( 'HFCMAC�Ǹ�' );
      aMsg := VdHFCMac( aSeqList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        Exit;
      end;
      { �v�@���X�Ǹ� }
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
          FDataBuffer.FieldByName( 'UPDEN' ).AsString := '���ƨt��';
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC�Ǹ�=<%s>,���`�C', [FSo310.HfcMac] );
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
  { ��Ʈw�s�����`�� }
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
    { �ӵ� CM �O�_�w�ռ��ܨ䥦�t�Υx, �u�ˬd�ռ��Y�i,
      �����ˮ� SO004 �O�_�ϥΤ� }
    aMsg := VdAlreadyInOtherSo( XML_WAREHOUSEID,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end else
    begin
      { �ӵ� CM �O�_���b�`�ܤ� }
      aMsg := VdNotExists( XML_WAREHOUSEID,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( '0', aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
    end;
    { �N�쥻�ӳ]�Ƹ��Ū�X��, �u�n�o�͹L���~�N���� }
    if not aTotalErr then
    begin
      UpdateBufferRecord;
    end;
    { �ӵ���ƨS���D }
    if not aIsErr then
    begin
      aMsg := Format( 'HFCMAC�Ǹ�<%s>,���`�C', [
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
        { �g�J SO310, CM �]����h�� }
        InsertRecord;
        { �R���쥻��  CM �]���� }
        DeleteRecord;
        FDataBuffer.Next;
      end;
      { �g Log }
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
          InsertLog( Format( '��h�渹:%s,�]�ƧO:CM,���q�N�X:%s,�ƶq:%d�C',
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
        aMsg := Format( '-1: CC&B��h��Ʈw�s�ɥ���,���~�T��:<%s>', [E.Message] );
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
    aMsg := Format( '-1:�ǤJ���ռ� XML �榡���~, ���~��]:%s�C', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  {}
  aTransList := FXml.DOMDocument.getElementsByTagName( '�ռ�' );
  aMsg := VdTranList( aTransList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( EmptyStr, aActionText, aMsg, 1 );
    Exit;
  end;  
  { �v�@���X�ռ���� }
  for aIndex := 0 to aTransList.length - 1 do
  begin
    { ��l FSo306 ��J�ȵ��c }
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    aEqumList := ( aTransList.item[aIndex] as IDOMNodeSelect ).selectNodes( '�]��' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { �v�@���X�]�� }
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aCompCodeText := GetChangeCompanyCode( aEqumList.item[aIndex2] );
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�ƶq' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�դJ���q' );
      aMsg := VdCompCode( aNode, EmptyStr );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      FSo306.CompCode := aNode.nodeValue;
      {}
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�եX���q' );
      aMsg := VdCompCode( aNode, EmptyStr );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      FSo306.OldCompCode := aNode.nodeValue;
      {}
      aSeqList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( 'HFCMAC�Ǹ�' );
      aMsg := VdHFCMac( aSeqList );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      { �v�@���X�Ǹ� }
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
          FDataBuffer.FieldByName( 'UPDEN' ).AsString := '���ƨt��';
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC�Ǹ�=<%s>,���`�C', [FSo306.HfcMac] );
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
      { �դJ���q }
      if Pos( FDataBuffer.FieldByName( 'COMPCODE' ).AsString, Result ) <= 0 then
        Result := ( Result + FDataBuffer.FieldByName( 'COMPCODE' ).AsString + ',' );
      { �եX���q }
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
  { ���X�Ҧ����q, ��Ʈw�s�u�� }
  aComps := GetAllCompanyString;
  { �U�� Transaction �n�Ψ� }
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
  { �ˬd��� }
  aTotalErr := False;
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    aIsErr := False;
    { 1. �ˬd�եX���t�Υx DB �O�_���T, �w���s����Ʈw }
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
    { 2. �ˬd�դJ���t�Υx DB �O�_���T, �w���s����Ʈw }
    aMsg := VdCD039( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 3. �եX���]�ƪ��t�Υx, �O�_�����]��?  }
    aMsg := VdNotExists( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 4. �եX���]�ƪ��t�Υx, �O�_�ϥΤ� ? }
    aMsg := VdNotUse( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 5. �եX���]�ƪ��t�Υx, �ӳ]�ƬO�_�w���P(�¦W��?) }
    aMsg := VdBlackList( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aDbComp, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 6. �q�եX�]�ƪ��t�ΥxŪ���쥻���]�Ƹ�� }
    UpdateBufferRecord( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString );
    { 7. �ˬd�դJ�]�ƨt�Υx�O�_�� �ӳ]�ƪ� ModelCode }
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
    { 8. ��s ModelCode }
    if not aTotalErr then
    begin
      FDataBuffer.Edit;
      FDataBuffer.FieldByName( 'MODELCODE' ).AsString := aModelCode;
      FDataBuffer.Post;
    end;
    { 8. �ˮֽդJ���t�Υx�O�_�����ȪA�~�W }
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
      aMsg := Format( 'HFCMAC�Ǹ�<%s>,���`�C', [
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      AppendLog( aDbComp, aActionText, aMsg, 0 );
    end;
    FDataBuffer.Next;
  end;
  {}
  if aTotalErr then Exit;
  {}
  { �Ұʥ�� }
  aComps := aTempComps;
  while ( aComps <> EmptyStr ) do
  begin
    aSoInfo := GetActiveSo( ExtractValue( aComps ) );
    if not aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.BeginTrans;
  end;
  { �s�� }
  try
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      aDbComp :=
        FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString + ',' +
        FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
      { �եX, �դJ�t�Υx�O�������P, �~���ʸ�� }
      if ( FDataBuffer.FieldByName( 'COMPCODE' ).AsString <>
           FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString ) then
      begin
        { �B�z�եX���]�ƨt�Υx }
        DeleteRecord( FDataBuffer.FieldByName( 'OLDCOMPCODE' ).AsString );
        { �B�z�դJ���]�ƨt�Υx }
        InsertRecord( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
        { ��s�`��, �ǤJ���ܧO���դJ���ܧO }
        UpdateRecord( FDataBuffer.FieldByName( 'COMPCODE' ).AsString );
      end;
      FDataBuffer.Next;
    end;
    { �̫�, �g Log }
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
          InsertLog( Format( '�]�ƧO:CM,�եX���q:%s,�դJ���q:%s,�ƶq:%d',
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
          InsertLog( Format( '�]�ƧO:CM,�եX���q:%s,�դJ���q:%s,�ƶq:%d',
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
        InsertLog( Format( '�]�ƧO:CM,�եX���q:%s,�դJ���q:%s,�ƶq:%d',
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
      aMsg := Format( '-1: CC&B�ռ���Ʈw�s�ɥ���,���~�T��:<%s>', [E.Message] );
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
    aMsg := Format( '-1:�ǤJ����h�� XML �榡���~, ���~��]:%s�C', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  aGetBackList := FXml.DOMDocument.getElementsByTagName( '��h��' );
  aMsg := VdGetBackList( aGetBackList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { ���X��h�Ƹ`�I }
  for aIndex := 0 to aGetBackList.length - 1 do
  begin
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    aEqumList := ( aGetBackList.item[aIndex] as IDOMNodeSelect ).selectNodes( '�]��' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '���q�N�X' );
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
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�ƶq' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�@�~' );
      aMsg := VdGetOrBack( aNode );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      { ���, useflag = 1 --> �i�� }
      if ( aNode.nodeValue = '1' ) then
        FSo306.UseFlag := 1
      else
      { �h��, useflag = 0 --> ���i�� }
        FSo306.UseFlag := 0;
      {}
      aSeqList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( 'HFCMAC�Ǹ�' );
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
          FDataBuffer.FieldByName( 'UPDEN' ).AsString := '���ƨt��';
          FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
          FDataBuffer.Post;
          aMsg := Format( 'HFCMAC�Ǹ�=<%s>,���`�C', [FSo306.HfcMac] );
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
  { �Ұʥ�� }
  aComps := aTempComps;
  while ( aComps <> EmptyStr ) do
  begin
    aSoInfo := GetActiveSo( ExtractValue( aComps ) );
    if not aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.BeginTrans;
  end;
  { �s�� }
  aTotalErr := False;
  try
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      aIsErr := False;
      aCompCodeText := FDataBuffer.FieldByName( 'COMPCODE' ).AsString;
      { 1.�Өt�Υx�O�_�����]�� }
      aMsg := VdNotExists( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      { 2.�O�_�ϥΤ�, �h�ƯS��W�h }
      aMsg := VdNotUse( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        aTotalErr := True;
        aIsErr := True;
      end;
      { 3.�O�_�Q���P }
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
        aMsg := Format( 'HFCMAC�Ǹ�=<%s>,���`�C', [
          FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
        AppendLog( aCompCodeText, aActionText, aMsg, 0 );
        aMsg := EmptyStr;
      end;
      FDataBuffer.Next;
    end;
    { �g Log }
    { �����n�g�� SOAC0204A, �]���w�g�� materiel_log �o�� Table }
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
//          aAction := '���';
//          if aLogUseFlag = '0' then aAction := '�h��';
//          InsertLog( aLogComp, Format( '�]�ƧO:CM,�ʧ@:%s,�Ǹ�:%s',
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
//          aAction := '���';
//          if aLogUseFlag = '0' then aAction := '�h��';
//          InsertLog( aLogComp, Format( '�]�ƧO:CM,�ʧ@:%s,�Ǹ�:%s',
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
      aMsg := Format( '-1: CC&B��/�h�Ƹ�Ʈw�s�ɥ���,���~�T��:<%s>', [E.Message] );
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
          if ( AnsiPos( '���`', aList.Strings[aIndex] ) <= 0 ) then 
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
    { �u�n IsErr �� Flag �ܦ� 1 ��, �N���i�A��^�� }
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
  // CallType = 0 --> �ˮ�, �q validate �禡�}�Y�I�s�i��
  // CallType = 1 --> �s��, �q save �禡�}�Y�I�s�i��
  Result := '�ˮ�';
  if ( CallType <> 0 ) then Result := '�s��';
  case FType of
    1: Result := ( '�J�w,' + Result );
    2: Result := ( '��h,' + Result );
    3: Result := ( '�ռ�,' + Result );
    4: Result := ( '��/�h��,' + Result );
    5: Result := ( '���װe�^,' + Result );
  end;    
end;

{ ---------------------------------------------------------------------------- }

function TCM2CT.WriteLog: String;
var
  aSo: PSoInfo;
begin
  Result := EmptyStr;
  { ��Ʈw�s�����`�� }
  Result := PrepareSoDBConnection( XML_WAREHOUSEID );
  if ( Result <> EmptyStr ) then Exit;
  { Log ��Ƥ@�߼g�`�ܸ�ư� COM �� }
  aSo := GetActiveSo( '0' );
  if not Assigned( aSo ) then
  begin
    Result := '-2:CC&B��Ʈw�L�k�������`�ܸ�ưϡC';
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
      Result := Format('-2:CC&B��Ʈw�g�JLog�ɥ���, ��]:<%s>�C', [E.Message] );
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
  aOutNode := aNode.attributes.getNamedItem( '�եX���q' );
  aInNode := aNode.attributes.getNamedItem( '�դJ���q' );
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
    aMsg := Format( '-1:�ǤJ�����װe�^ XML �榡���~, ���~��]:%s�C', [aMsg] );
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  aChangeList := FXml.DOMDocument.getElementsByTagName( '���װe�^' );
  aMsg := VdChangeList( aChangeList );
  if ( aMsg <> EmptyStr ) then
  begin
    AppendLog( '0', aActionText, aMsg, 1 );
    Exit;
  end;
  { ���X���װe�^�`�I }
  for aIndex := 0 to aChangeList.length - 1 do
  begin
    ZeroMemory( FSo306, SizeOf( FSo306^ ) );
    aEqumList := ( aChangeList.item[aIndex] as IDOMNodeSelect ).selectNodes( '�]��' );
    aMsg := VdEqum( aEqumList );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( '0', aActionText, aMsg, 1 );
      Exit;
    end;
    { �v�@���X�]�� }
    for aIndex2 := 0 to aEqumList.length - 1 do
    begin
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '���q�N�X' );
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
      aNode := aEqumList.item[aIndex2].attributes.getNamedItem( '�ƶq' );
      aMsg := VdItemCount( aNode, aEqumList.item[aIndex2] );
      if ( aMsg <> EmptyStr ) then
      begin
        AppendLog( aCompCodeText, aActionText, aMsg, 1 );
        Exit;
      end;
      {}
      aHFCList := ( aEqumList.item[aIndex2] as IDOMNodeSelect ).selectNodes( '�Ǹ�' );
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
        { ���M�w���@�Ӥ~�O �e�� �� node }
        aMsg := VdExchangeNode( aSendNode, aBackNode );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDeatilErr := True;
          Continue;
        end;
        { �e�� HFCMAC }
        aMsg := VdHFCMac( aSendNode.attributes.getNamedItem( 'HFCMAC' ) );
        if ( aMsg <> EmptyStr ) then
        begin
          AppendLog( aCompCodeText, aActionText, aMsg, 1 );
          aDeatilErr := True;
          Continue;
        end;
        { ��^ HFCMAC }
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
        FDataBuffer.FieldByName( 'UPDEN' ).AsString := '���ƨt��';
        FDataBuffer.FieldByName( 'UPDTIME' ).Value := FCallTime;
        FDataBuffer.Post;
        aMsg := Format( '�e��HFCMAC�Ǹ�=<%s>, ��^HFCMAC�Ǹ�=<%s>, ���`�C',
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
    { �`�ܥ[�i�� }
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
        '        UPDEN = ''���ƨt��''  ' +
        '  WHERE HFCMAC = ''%s''       ',
        [FDataBuffer.FieldByName( 'BACKHFCMAC' ).AsString,
         FDataBuffer.FieldByName( 'BACKMTAMAC' ).AsString,
         FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      aSo.Writer.ExecSQL;
      { �Y�����N�O�`��, �h���ΦA����2�� }
      if ( aCompCode = XML_WAREHOUSEID ) then Break;
      { �Y���O�`��, �h�٭n�N�`�ܧ�s }
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
    { 1.�e�ת��]��, �Өt�Υx�O�_�����]�� }
    aMsg := VdNotExists( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 2.�e�ת��]�ƬO�_�ϥΤ�, ���װe�^�S��W�h }
    aMsg := VdNotUse( FDataBuffer.FieldByName( 'COMPCODE' ).AsString,
      FDataBuffer.FieldByName( 'HFCMAC' ).AsString );
    if ( aMsg <> EmptyStr ) then
    begin
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
      aTotalErr := True;
      aIsErr := True;
    end;
    { 3.�e�ת��]�ƬO�_�Q���P }
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
    { ���װe�^�������Ǹ�, �Y�O���ܧ�, �����ˮ֤��i�H�s�b��w�s�� }
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
      { 4.�ˮְe�ת� MTAMAC �O�_���� DB �ح즳�� MTAMAC }
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
      { 5.�ˮֻ�^������ �O�_���� DB �ح즳������ (��MTAMAC���ȩεL�ȧP�_ ) }
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
    { �Y�O�e�׫e�᪺ HFCMAC �O�@�˪�, �h�� MTAMAC �e�׫e��O�_�@�� }
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
      aMsg := Format( 'HFCMAC�Ǹ�<%s>,���`�C', [
        FDataBuffer.FieldByName( 'HFCMAC' ).AsString] );
      AppendLog( aCompCodeText, aActionText, aMsg, 0 );
    end;
    FDataBuffer.Next;
  end;
  {}
  if aTotalErr then Exit;
  {}
  { �Ұʥ�� }
  aComps := aTempComps;
  while ( aComps <> EmptyStr ) do
  begin
    aSoInfo := GetActiveSo( ExtractValue( aComps ) );
    if not aSoInfo.Connection.InTransaction then
      aSoInfo.Connection.BeginTrans;
  end;
  { �s�� }
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
      aMsg := Format( '-1: CC&���װe�|��Ʈw�s�ɥ���,���~�T��:<%s>', [E.Message] );
      AppendLog( aCompCodeText, aActionText, aMsg, 1 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

initialization
  TAutoObjectFactory.Create(ComServer, TCM2CT, Class_CM2CT,
    ciMultiInstance, tmApartment);
end.
