unit cbSoInfo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, IniFiles, Contnrs,
  Encryption_TLB, cxPC, cxTL, xmldom, XMLIntf, msxmldom, XMLDoc;

const
  DBSECTION = 'DBINFO';
  CONFIGSECTION = 'GENERAL';

type

  TDbConnectStatus = ( dbNoSelect, dbNone, dbOK, dbError, dbActive );
  
  TLoader = class
  private
    FIniObj: TIniFile;
    FEncrypObj: _Password;
    FEncrypKey: WideString;
    function ProduceDecryptFile(aSourceFileName: String): String;
  protected
    procedure LoadFromFile(aFullFileName: String); virtual; abstract;
  public
    constructor Create;
    destructor Destroy; override;
  end;


  TSoLoader = class(TLoader)
  private
    FSoIndex: Integer;
    FSoList: TObjectList;
    function GetCompCode: String;
    function GetCompName: String;
    function GetSid: String;
    function GetOwner: String;
    function GetPassword: String;
    function GetEnabled: Boolean;
    property SoIndex: Integer read FSoIndex write FSoIndex;
  public
    procedure LoadFromFile(aFullFileName: String); override;
    property LoaderResult: TObjectList read FSoList;
    constructor Create;
    destructor Destroy; override;
  end;

  TSoInfo = class
  private
    FCompCode: String;
    FPassword: String;
    FCompName: String;
    FSid: String;
    FOwner: String;
    FEnable: Boolean;
    FDisplaySheet: TcxTabSheet;
    FDisplayTreeList: TcxTreeList;
    FExecThread: TThread;
    FDbStatus: TDbConnectStatus;
  public
    property CompCode: String read FCompCode;
    property CompName: String read FCompName;
    property Sid: String read FSid;
    property Owner: String read FOwner;
    property Password: String read FPassword;
    property Enable: Boolean read FEnable;
    property DisplaySheet: TcxTabSheet read FDisplaySheet write FDisplaySheet;
    property DisplayTreeList: TcxTreeList read FDisplayTreeList write FDisplayTreeList;
    property ExecThread: TThread read FExecThread write FExecThread;
    property DbStatus: TDbConnectStatus read FDbStatus write FDbStatus;
    constructor Create(aLoader: TLoader; const aSoIndex: Integer);
    destructor Destroy; override;
  end;


  TConfigInfo = class
  private
    FDbScanInterval: Integer;
    FXmlParseInterval: Integer;
    FXmlParseRecords: Integer;
    FSOAPUrl: String;
    FSOAPAction: String;
  public
    property DbScanInterval: Integer read FDbScanInterval;
    property XmlParseInterval: Integer read FXmlParseInterval;
    property XmlParseRecords: Integer read FXmlParseRecords;
    property SOAPUrl: String read FSOAPUrl;
    property SOAPAction: String read FSOAPAction;
    constructor Create(aLoader: TLoader);
  end;


  TConfigLoader = class(TLoader)
  private
    FConfigInfo: TConfigInfo;
    function GetDbScanInterval: Integer;
    function GetXmlParseInterval: Integer;
    function GetXmlParseRecords: Integer;
    function GetSOAPUrl: String;
    function GetSOAPAction: String;
  public
    procedure LoadFromFile(aFullFileName: String); override;
    property LoaderResult: TConfigInfo read FConfigInfo;
  end;


  PMsg = ^TMsg;
  TMsg = record
    Kind: Integer;
    Msg: String;
  end;

  PSO311 = ^TSO311;
  TSO311 = record
    CmdSeqNO: String;
    MsgCode: String;
    MsgName: String;
    WorkId: String;
    CusSo: String;
    CusId: String;
    CMMac: String;
    EMTAMac: String;
    EMTAPort: String;
    CPNo: String;
    ClassId: String;
    PCIPNo: String;
    PCIP: String;
    OperType: String;
    StartDateTime: String;
    EndDateTime: String;
    Zone: String;
    Lin: String;
    Section: String;
    Lane: String;
    Alley: String;
    SubAlley: String;
    NO1: String;
    NO2: String;
    SOAPRequest: String;
    RequestTime: String;
    SOAPResponse: String;
    ReponseTime: String;
    DisplayNode: Pointer;
    ErrMsg: String;
    ExtendedInfo: String;
    IsBatch: Boolean;
    {}
    GroupId: String;
    CMStatus: String;
    RCMIp: String;
    RCMUpFreq: String;
    RCMDownFreq: String;
    RCMRecPw: String;
    RCMTransPw: String;
    RCMSnr: String;
    RCMOnlineTimes: String;
    RCMInoctets: String;
    RCMOutOctets: String;
    RCMInErrors: String;
    RCMOutErrors: String;
    RPCIp: String;
    RCMTsrecPw: String;
    RCMTsrXsNr: String;
    RCMTsupMod: String;
    RHfcType: String;
    RPCOfflineTime: String;
    RPCOnlineTime: String;
    ROfflineTime: String;
    ROnlineTime: String;
    RDetctTime: String;
    RCMMac: String;
    RCTIp: String;
    RNodeId: String;
    {}
    QueryResult: String;
    OperResult: String;
    FaultReason: String;
  end;

  TCMWebServiceManager = class
  private
    FXml: TXMLDocument;
    FOwner: String;
    {}
    function GetDOMNodeValue(aNode:IDOMNode): String;
    {}
    function CMReg(aSO311: PSO311): String;
    function CPIPReg(aSO311: PSO311): String;
    function ResetCM(aSO311: PSO311): String;
    function PingCM(aSO311: PSO311): String;
    function ClearCPE(aSO311: PSO311): String;
    function CMStatusQuery(aSO311: PSO311): String;
    function HFCTypeQuery(aSO311: PSO311): String;
    function CMConnLog(aSO311: PSO311): String;
    function CMQALog(aSO311: PSO311): String;
    function CMPCLog(aSO311: PSO311): String;
    {}
    function CMRegErr(aErrMsg: String): String;
    function CPIPRegErr(aErrMsg: String): String;
    function ResetCMErr(aErrMsg: String): String;
    function PingCMErr(aErrMsg: String): String;
    function ClearCPEErr(aErrMsg: String): String;
    function CMStatusQueryErr(aErrMsg: String): String;
    function HFCTypeQueryErr(aErrMsg: String): String;
    function CMConnLogErr(aErrMsg: String): String;
    function CMQALogErr(aErrMsg: String): String;
    function CMPCLogErr(aErrMsg: String): String;
    {}
    procedure ParseCMReg(aSO311: PSO311);
    procedure ParseResetCM(aSO311: PSO311);
    procedure ParsePingCM(aSO311: PSO311);
    procedure ParseClearCPE(aSO311: PSO311);
    procedure ParseCMStatusQuery(aSO311: PSO311);
    procedure ParseHFCTypeQuery(aSO311: PSO311);
    procedure ParseCMConnLog(aSO311: PSO311);
    procedure ParseCMQALog(aSO311: PSO311);
    procedure ParseCMPCLog(aSO311: PSO311);
    procedure ParseCPIPReg(aSO311: PSO311);
    {}
    function SqlCMReg(aSO311: PSO311): String;
    function SqlResetCM(aSO311: PSO311): String;
    function SqlPingCM(aSO311: PSO311): String;
    function SqlClearCPE(aSO311: PSO311): String;
    function SqlCMStatusQuery(aSO311: PSO311): String;
    function SqlHFCTypeQuery(aSO311: PSO311): String;
    function SqlCMConnLog(aSO311: PSO311): String;
    function SqlCMQALog(aSO311: PSO311): String;
    function SqlCMPCLog(aSO311: PSO311): String;
    function SqlCPIPReg(aSO311: PSO311): String;
  public
    constructor Create(aXml: TXMLDocument; aOwner: String);
    destructor Destroy; override;
    function GetSOAPRequestText(aSO311: PSO311; aXmlns: String = ''): String;
    function GetSOAPActionEndText(aSO311: PSO311): String;
    function GetErrorXml(aSO311: PSO311): String;
    function GetUpdateXmlFieldSql(aSO311: PSO311): String;
    procedure ParseXml(aSO311: PSO311);
  end;


implementation

{ ---------------------------------------------------------------------------- }

{ TSoInfo }

constructor TSoInfo.Create(aLoader: TLoader; const aSoIndex: Integer);
begin
  TSoLoader( aLoader ).SoIndex := aSoIndex;
  FCompCode := TSoLoader( aLoader ).GetCompCode;
  FCompName := TSoLoader( aLoader ).GetCompName;
  FSid := TSoLoader( aLoader ).GetSid;
  FOwner := TSoLoader( aLoader ).GetOwner;
  FPassword := TSoLoader( aLoader ).GetPassword;
  FEnable := TSoLoader( aLoader ).GetEnabled;
  FDbStatus := dbNone;
end;

{ ---------------------------------------------------------------------------- }

destructor TSoInfo.Destroy;
begin
  FDisplaySheet := nil;
  FDisplayTreeList := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

{ TLoader }

constructor TLoader.Create;
begin
  FEncrypObj := CoPassword.Create;
  FEncrypKey := 'CS';
end;

{ ---------------------------------------------------------------------------- }

destructor TLoader.Destroy;
begin
  FEncrypObj := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TLoader.ProduceDecryptFile(aSourceFileName: String): String;

     { ---------------------------------------------------- }

     function GetTempDir: String;
     var
       aBuffer: array [0..MAX_PATH] of Char;
     begin
       ZeroMemory( @aBuffer[0], SizeOf( aBuffer ) );
       GetTempPath( SizeOf( aBuffer ), @aBuffer[0] );
       Result := aBuffer;
     end;

     { ---------------------------------------------------- }

var
  aIndex: Integer;
  aStrList: TStringList;
  aPrefix, aLineStr, aTempDir: String;
begin
  Result := EmptyStr;
  if not FileExists( aSourceFileName ) then Exit;
  aStrList := TStringList.Create;
  try
    aStrList.LoadFromFile( aSourceFileName );
    for aIndex := 0 to aStrList.Count - 1 do
    begin
      aLineStr := aStrList[aIndex];
      aPrefix := Copy( aLineStr, 1, 2 );
      if ( aPrefix = '//' ) or ( aPrefix = EmptyStr ) then Continue;
      if Pos( 'COMPCODE', UpperCase( aStrList[aIndex] ) ) <= 0 then
        aStrList[aIndex] := FEncrypObj.Decrypt( aLineStr, FEncrypKey )
      else
        aStrList[aIndex] := aLineStr;
    end;
    aTempDir := GetTempDir;
    Result := IncludeTrailingPathDelimiter( aTempDir ) + ExtractFileName(
      aSourceFileName );
    aStrList.SaveToFile( Result );  
  finally
     aStrList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSoLoader }

constructor TSoLoader.Create;
begin
  inherited Create;
  FSoList := nil;
  FSoIndex := -1;
end;

{ ---------------------------------------------------------------------------- }

destructor TSoLoader.Destroy;
begin
  FSoList := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetCompCode: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'COMPCODE_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetCompName: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'COMPNAME_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetOwner: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'USERID_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetPassword: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'PASSWORD_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetSid: String;
begin
  Result := FIniObj.ReadString( DBSECTION, Format( 'ALIAS_%d', [FSoIndex+1] ),
    EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TSoLoader.GetEnabled: Boolean;
var
  aStr: String;
begin
  aStr := FIniObj.ReadString( DBSECTION, Format( 'ENABLE_%d', [FSoIndex+1] ),
    'N' );
  Result := ( aStr = 'Y' );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoLoader.LoadFromFile(aFullFileName: String);
var
  aIndex, aCount: Integer;
begin
  FIniObj := TIniFile.Create( ProduceDecryptFile( aFullFileName ) );
  try
    FSoList := nil;
    aCount := FIniObj.ReadInteger( DBSECTION, 'DB_COUNT', 0 );
    if aCount <= 0 then Exit;
    FSoList := TObjectList.Create( True );
    for aIndex := 0 to aCount - 1 do
      FSoList.Add( TSoInfo.Create( Self, aIndex ) );
    DeleteFile( FIniObj.FileName );
  finally
    FIniObj.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TConfigLoader }

function TConfigLoader.GetDbScanInterval: Integer;
begin
  Result := FIniObj.ReadInteger( CONFIGSECTION, 'DB_SCAN_INTERVAL', 3 );
end;

{ ---------------------------------------------------------------------------- }

function TConfigLoader.GetXmlParseInterval: Integer;
begin
  Result := FIniObj.ReadInteger( CONFIGSECTION, 'XML_PARSE_INTERVAL', 30 );
end;

{ ---------------------------------------------------------------------------- }

function TConfigLoader.GetXmlParseRecords: Integer;
begin
  Result := FIniObj.ReadInteger( CONFIGSECTION, 'XML_PARSE_RECORDS', 50 );
end;

{ ---------------------------------------------------------------------------- }

function TConfigLoader.GetSOAPAction: String;
begin
  Result := FIniObj.ReadString( CONFIGSECTION, 'SOAP_ACTION', EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TConfigLoader.GetSOAPUrl: String;
begin
  Result := FIniObj.ReadString( CONFIGSECTION, 'SOAP_URL', EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigLoader.LoadFromFile(aFullFileName: String);
begin
  FIniObj := TIniFile.Create( ProduceDecryptFile( aFullFileName ) );
  try
    FConfigInfo := TConfigInfo.Create( Self );
    DeleteFile( FIniObj.FileName );
  finally
    FIniObj.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TConfigInfo }

constructor TConfigInfo.Create(aLoader: TLoader);
begin
  FDbScanInterval := TConfigLoader( aLoader ).GetDbScanInterval;
  FXmlParseInterval := TConfigLoader( aLoader ).GetXmlParseInterval;
  FXmlParseRecords := TConfigLoader( aLoader ).GetXmlParseRecords;
  FSOAPUrl := TConfigLoader( aLoader ).GetSOAPUrl;
  FSOAPAction := TConfigLoader( aLoader ).GetSOAPAction;
end;

{ ---------------------------------------------------------------------------- }

{ TCMWebServiceManager }

constructor TCMWebServiceManager.Create(aXml: TXMLDocument; aOwner: String);
begin
  inherited Create;
  FXml := aXml;
  FOwner := aOwner;
end;

{ ---------------------------------------------------------------------------- }

destructor TCMWebServiceManager.Destroy;
begin
  FXml := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.GetSOAPRequestText(aSO311: PSO311;
  aXmlns: String = '' ): String;
var
  aCode: Integer;
begin
  aCode := StrToInt( aSO311.MsgCode );
  case aCode of
    2: Result := CMReg( aSO311 );
    12: Result := ResetCM( aSO311 );
    20: Result := PingCM( aSO311 );
    25: Result := ClearCPE( aSO311 );
    13: Result := CMStatusQuery( aSO311 );
    21: Result := HFCTypeQuery( aSO311 );
    22: Result := CMConnLog( aSO311 );
    23: Result := CMQALog( aSO311 );
    26: Result := CMPCLog( aSO311 );
    27: Result := CPIPReg( aSO311 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.ClearCPE(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s',
    [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMConnLog(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s&StartDateTime=%s&EndDateTime=%s',
   [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac,
    aSO311.StartDateTime, aSO311.EndDateTime] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMPCLog(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&PCIP=%s&StartDateTime=%s&EndDateTime=%s',
   [aSO311.WorkId, '10', aSO311.CusId, aSO311.PCIP,
    aSO311.StartDateTime, aSO311.EndDateTime] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMQALog(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s&StartDateTime=%s&EndDateTime=%s',
   [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac,
    aSO311.StartDateTime, aSO311.EndDateTime] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMReg(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s&EMTAMAC=%s&CLASSID=%s&PCIPNO=%s&OPERTYPE=%s',
   [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac, aSO311.EMTAMac,
    aSO311.ClassId, aSO311.PCIPNo, aSO311.OperType] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMStatusQuery(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s',
    [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CPIPReg(aSO311: PSO311): String;
var
  aZone: String;
begin
  aZone := aSO311.Zone;
  aSO311.ExtendedInfo := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s&EMTAMAC=%s&EMTAPort=%s&CPNO=%s&Zone=%s&Lin=%s&Section=%s&Lane=%s&Alley=%s&SubAlley=%s&NO1=%s&NO2=%s&OPERTYPE=%s',
   [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac, aSO311.EMTAMac,
    aSO311.EMTAPort, aSO311.CPNo, aZone, aSO311.Lin, aSO311.Section,
    aSO311.Lane, aSO311.Alley, aSO311.SubAlley, aSO311.NO1, aSO311.NO2, aSO311.OperType] );
  if ( aZone <> EmptyStr ) then
    aZone := AnsiToUtf8( aZone );
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s&EMTAMAC=%s&EMTAPort=%s&CPNO=%s&Zone=%s&Lin=%s&Section=%s&Lane=%s&Alley=%s&SubAlley=%s&NO1=%s&NO2=%s&OPERTYPE=%s',
   [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac, aSO311.EMTAMac,
    aSO311.EMTAPort, aSO311.CPNo, aZone, aSO311.Lin, aSO311.Section,
    aSO311.Lane, aSO311.Alley, aSO311.SubAlley, aSO311.NO1, aSO311.NO2, aSO311.OperType] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.HFCTypeQuery(aSO311: PSO311): String;
var
  aZone: String;
begin
  aZone := aSO311.Zone;
  aSo311.ExtendedInfo := Format( 'WORKID=%s&CUSSO=%s&Zone=%s&Lin=%s&Section=%s&Lane=%s&Alley=%s&SubAlley=%s&NO1=%s&NO2=%s',
   [aSO311.WorkId, '10', aZone, aSO311.Lin, aSO311.Section,
    aSO311.Lane, aSO311.Alley, aSO311.SubAlley, aSO311.NO1, aSO311.NO1] );
  if ( aZone <> EmptyStr ) then
    aZone := AnsiToUtf8( aZone );
  Result := Format( 'WORKID=%s&CUSSO=%s&Zone=%s&Lin=%s&Section=%s&Lane=%s&Alley=%s&SubAlley=%s&NO1=%s&NO2=%s',
   [aSO311.WorkId, '10', aZone, aSO311.Lin, aSO311.Section,
    aSO311.Lane, aSO311.Alley, aSO311.SubAlley, aSO311.NO1, aSO311.NO1] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.PingCM(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s',
    [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.ResetCM(aSO311: PSO311): String;
begin
  Result := Format( 'WORKID=%s&CUSSO=%s&CUSID=%s&CMMAC=%s',
    [aSO311.WorkId, '10', aSO311.CusId, aSO311.CMMac] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.GetSOAPActionEndText(aSO311: PSO311): String;
var
  aCode: Integer;
begin
  aCode := StrToInt( aSo311.MsgCode );
  case aCode of
    2: Result := 'CMReg';
    12: Result := 'ResetCM';
    20: Result := 'PingCM';
    25: Result := 'ClearCPE';
    13: Result := 'CMStatusQuery';
    21: Result := 'HFCTypeQuery';
    22: Result := 'CMConnLog';
    23: Result := 'CMQALog';
    26: Result := 'CMPCLog';
    27: Result := 'CPIPReg';
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.GetErrorXml(aSO311: PSO311): String;
var
  aCode: Integer;
begin
  aCode := StrToInt( aSO311.MsgCode );
  case aCode of
    2: Result := CMRegErr( aSO311.ErrMsg );
    12: Result := ResetCMErr( aSO311.ErrMsg );
    20: Result := PingCMErr( aSO311.ErrMsg );
    25: Result := ClearCPEErr( aSO311.ErrMsg );
    13: Result := CMStatusQueryErr( aSO311.ErrMsg );
    21: Result := HFCTypeQueryErr( aSO311.ErrMsg );
    22: Result := CMConnLogErr( aSO311.ErrMsg );
    23: Result := CMQALogErr( aSO311.ErrMsg );
    26: Result := CMPCLogErr( aSO311.ErrMsg );
    27: Result := CPIPRegErr( aSO311.ErrMsg );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMRegErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="CMRegDS" targetNamespace="http://tempuri.org/CMRegDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/CMRegDS.xsd" xmlns="http://tempuri.org/CMRegDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="CMRegDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="CMRegDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="OperResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="GROUPID" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CTIP" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <CMRegDS xmlns="http://tempuri.org/CMRegDS.xsd"> ' +
  '      <CMRegDT diffgr:id="CMRegDT1" msdata:rowOrder="0"> ' +
  '        <OperResult>2</OperResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <GROUPID /> ' +
  '        <CTIP /> ' +
  '      </CMRegDT> ' +
  '    </CMRegDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet>', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.ClearCPEErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="ClearCPEDS" targetNamespace="http://tempuri.org/ClearCPEDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/ClearCPEDS.xsd" xmlns="http://tempuri.org/ClearCPEDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" '+
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="ClearCPEDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="ClearCPEDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="OperResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <ClearCPEDS xmlns="http://tempuri.org/ClearCPEDS.xsd"> ' +
  '      <ClearCPEDT diffgr:id="ClearCPEDT1" msdata:rowOrder="0"> ' +
  '        <OperResult>2</OperResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '      </ClearCPEDT> ' +
  '    </ClearCPEDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMConnLogErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="CMConnLogDS" targetNamespace="http://tempuri.org/CMConnLogDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/CMConnLogDS.xsd" xmlns="http://tempuri.org/CMConnLogDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="CMConnLogDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="CMConnLogDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="QueryResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="OnlineTime" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="OfflineTime" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMIP" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="PCIP" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <CMConnLogDS xmlns="http://tempuri.org/CMConnLogDS.xsd"> ' +
  '      <CMConnLogDT diffgr:id="CMConnLogDT1" msdata:rowOrder="0"> ' +
  '        <QueryResult>2</QueryResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <OnlineTime /> ' +
  '        <OfflineTime /> ' +
  '        <CMIP /> ' +
  '        <PCIP /> ' +
  '      </CMConnLogDT> ' +
  '    </CMConnLogDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMPCLogErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="CMPCLogDS" targetNamespace="http://tempuri.org/CMPCLogDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/CMPCLogDS.xsd" xmlns="http://tempuri.org/CMPCLogDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="CMPCLogDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="CMPCLogDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="QueryResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="PCOnlineTime" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="PCOfflineTime" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMMAC" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <CMPCLogDS xmlns="http://tempuri.org/CMPCLogDS.xsd"> ' +
  '      <CMPCLogDT diffgr:id="CMPCLogDT1" msdata:rowOrder="0"> ' +
  '        <QueryResult>2</QueryResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <PCOnlineTime /> ' +
  '        <PCOfflineTime /> ' +
  '        <CMMAC /> ' +
  '      </CMPCLogDT> ' +
  '    </CMPCLogDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMQALogErr(aErrMsg: String): String;
begin
  Result := Format( 
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="CMQALogDS" targetNamespace="http://tempuri.org/CMQALogDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/CMQALogDS.xsd" xmlns="http://tempuri.org/CMQALogDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="CMQALogDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="CMQALogDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="QueryResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="DetectTime" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTSRECPW" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTSRXSNR" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMRECPW" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTRANSPW" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMSNR" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMInOctets" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMOutOctets" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMInErrors" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMOutErrors" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <CMQALogDS xmlns="http://tempuri.org/CMQALogDS.xsd"> ' +
  '      <CMQALogDT diffgr:id="CMQALogDT1" msdata:rowOrder="0"> ' +
  '        <QueryResult>2</QueryResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <DetectTime /> ' +
  '        <CMTSRECPW /> ' +
  '        <CMTSRXSNR /> ' +
  '        <CMRECPW /> ' +
  '        <CMTRANSPW /> ' +
  '        <CMSNR /> ' +
  '        <CMInOctets /> ' +
  '        <CMOutOctets /> ' +
  '        <CMInErrors /> ' +
  '        <CMOutErrors /> ' +
  '      </CMQALogDT> ' +
  '    </CMQALogDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CMStatusQueryErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="CMStatusQueryDS" targetNamespace="http://tempuri.org/Dataset1.xsd" ' +
  'xmlns:mstns="http://tempuri.org/Dataset1.xsd" xmlns="http://tempuri.org/Dataset1.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="CMStatusQueryDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="CMSTDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="QueryResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="GROUPID" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMStatus" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMIP" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMUPFREQ" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMDOWNFREQ" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMRECPW" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTRANSPW" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMSNR" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMOnlineTimes" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMInOctets" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMOutOctets" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMInErrors" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMOutErrors" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="PCIP" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTSRECPW" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTSRXSNR" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CMTSUPMOD" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <CMStatusQueryDS xmlns="http://tempuri.org/Dataset1.xsd"> ' +
  '      <CMSTDT diffgr:id="CMSTDT1" msdata:rowOrder="0"> ' +
  '        <QueryResult>2</QueryResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <GROUPID /> ' +
  '        <CMStatus /> ' +
  '        <CMIP /> ' +
  '        <CMUPFREQ /> ' +
  '        <CMDOWNFREQ /> ' +
  '        <CMRECPW /> ' +
  '        <CMTRANSPW /> ' +
  '        <CMSNR /> ' +
  '        <CMOnlineTimes /> ' +
  '        <CMInOctets /> ' +
  '        <CMOutOctets /> ' +
  '        <CMInErrors /> ' +
  '        <CMOutErrors /> ' +
  '        <PCIP /> ' +
  '        <CMTSRECPW /> ' +
  '        <CMTSRXSNR /> ' +
  '        <CMTSUPMOD /> ' +
  '      </CMSTDT> ' +
  '    </CMStatusQueryDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.CPIPRegErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="CPIPRegDS" targetNamespace="http://tempuri.org/CPIPRegDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/CPIPRegDS.xsd" xmlns="http://tempuri.org/CPIPRegDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="CPIPRegDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="CPIPRegDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="OperResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="CTIP" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <CPIPRegDS xmlns="http://tempuri.org/CPIPRegDS.xsd"> ' +
  '      <CPIPRegDT diffgr:id="CPIPRegDT1" msdata:rowOrder="0"> ' +
  '        <OperResult>2</OperResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <CTIP /> ' +
  '      </CPIPRegDT> ' +
  '    </CPIPRegDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.HFCTypeQueryErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="HFCTypeQueryDS" targetNamespace="http://tempuri.org/HFCTypeQueryDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/HFCTypeQueryDS.xsd" xmlns="http://tempuri.org/HFCTypeQueryDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="HFCTypeQueryDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="HFCTypeQueryDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="OperResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="HFCType" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="NODEID" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <HFCTypeQueryDS xmlns="http://tempuri.org/HFCTypeQueryDS.xsd"> ' +
  '      <HFCTypeQueryDT diffgr:id="HFCTypeQueryDT1" msdata:rowOrder="0"> ' +
  '        <OperResult>2</OperResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '        <HFCType /> ' +
  '        <NODEID /> ' +
  '      </HFCTypeQueryDT> ' +
  '    </HFCTypeQueryDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.PingCMErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="PingCMDS" targetNamespace="http://tempuri.org/PingCMDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/PingCMDS.xsd" xmlns="http://tempuri.org/PingCMDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="PingCMDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="PingCMDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="OperResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <PingCMDS xmlns="http://tempuri.org/PingCMDS.xsd"> ' +
  '      <PingCMDT diffgr:id="PingCMDT1" msdata:rowOrder="0"> ' +
  '        <OperResult>2</OperResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '      </PingCMDT> ' +
  '    </PingCMDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.ResetCMErr(aErrMsg: String): String;
begin
  Result := Format(
  '<?xml version="1.0" encoding="utf-8"?> ' +
  '<DataSet xmlns="http://tempuri.org/CMWebServiceGW/EMCWebServiceGW"> ' +
  '  <xs:schema id="ResetCMDS" targetNamespace="http://tempuri.org/ResetCMDS.xsd" ' +
  'xmlns:mstns="http://tempuri.org/ResetCMDS.xsd" xmlns="http://tempuri.org/ResetCMDS.xsd" ' +
  'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'attributeFormDefault="qualified" elementFormDefault="qualified"> ' +
  '    <xs:element name="ResetCMDS" msdata:IsDataSet="true"> ' +
  '      <xs:complexType> ' +
  '        <xs:choice maxOccurs="unbounded"> ' +
  '          <xs:element name="ResetCMDT"> ' +
  '            <xs:complexType> ' +
  '              <xs:sequence> ' +
  '                <xs:element name="OperResult" type="xs:string" minOccurs="0" /> ' +
  '                <xs:element name="FaultReason" type="xs:string" minOccurs="0" /> ' +
  '              </xs:sequence> ' +
  '            </xs:complexType> ' +
  '          </xs:element> ' +
  '        </xs:choice> ' +
  '      </xs:complexType> ' +
  '    </xs:element> ' +
  '  </xs:schema> ' +
  '  <diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" ' +
  'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"> ' +
  '    <ResetCMDS xmlns="http://tempuri.org/ResetCMDS.xsd"> ' +
  '      <ResetCMDT diffgr:id="ResetCMDT1" msdata:rowOrder="0"> ' +
  '        <OperResult>2</OperResult> ' +
  '        <FaultReason>無法正常傳送指令，可能為網管介接失敗，請洽網管人員, 訊息:%s</FaultReason> ' +
  '      </ResetCMDT> ' +
  '    </ResetCMDS> ' +
  '  </diffgr:diffgram> ' +
  '</DataSet> ', [aErrMsg] );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseXml(aSO311: PSO311);
var
  aCode: Integer;
begin
  aCode := StrToInt( aSO311.MsgCode );
  case aCode of
    2: ParseCMReg( aSO311 );
    12: ParseResetCM( aSO311 );
    20: ParsePingCM( aSO311 );
    25: ParseClearCPE( aSO311 );
    13: ParseCMStatusQuery( aSO311 );
    21: ParseHFCTypeQuery( aSO311 );
    22: ParseCMConnLog( aSO311 );
    23: ParseCMQALog( aSO311 );
    26: ParseCMPCLog( aSO311 );
    27: ParseCPIPReg( aSO311 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.GetDOMNodeValue(aNode:IDOMNode): String;
begin
  Result := EmptyStr;
  if Assigned( aNode ) then
    if Assigned( aNode.firstChild ) then
      Result := aNode.firstChild.nodeValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseClearCPE(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''OperResult'']' );
  aSO311.OperResult := aValueNode.firstChild.nodeValue;
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseCMConnLog(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''QueryResult'']' );
  aSO311.QueryResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''OnlineTime'']' );
  aSO311.ROnlineTime := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''OfflineTime'']' );
  aSO311.ROfflineTime := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMIP'']' );
  aSO311.RCMIp := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''PCIP'']' );
  aSO311.RPCIp := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseCMPCLog(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''QueryResult'']' );
  aSO311.QueryResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''PCOnlineTime'']' );
  aSO311.RPCOnlineTime := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''PCOfflineTime'']' );
  aSO311.RPCOfflineTime := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMMAC'']' );
  aSO311.RCMMac := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseCMQALog(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''QueryResult'']' );
  aSO311.QueryResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''DetectTime'']' );
  aSO311.RDetctTime := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTSRECPW'']' );
  aSO311.RCMTsrecPw := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTSRXSNR'']' );
  aSO311.RCMTsrXsNr := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMRECPW'']' );
  aSO311.RCMRecPw := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTRANSPW'']' );
  aSO311.RCMTransPw := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMSNR'']' );
  aSO311.RCMSnr := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMInOctets'']' );
  aSO311.RCMInoctets := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMOutOctets'']' );
  aSO311.RCMOutOctets := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMInErrors'']' );
  aSO311.RCMInErrors := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMOutErrors'']' );
  aSO311.RCMOutErrors := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseCMReg(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''OperResult'']' );
  aSO311.OperResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''GROUPID'']' );
  aSO311.GroupId := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CTIP'']' );
  aSO311.RCTIp := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseCMStatusQuery(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''QueryResult'']' );
  aSO311.QueryResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''GROUPID'']' );
  aSO311.GroupId := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMStatus'']' );
  aSO311.CMStatus := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMIP'']' );
  aSO311.RCMIp := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMUPFREQ'']' );
  aSO311.RCMUpFreq := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMDOWNFREQ'']' );
  aSO311.RCMDownFreq := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMRECPW'']' );
  aSO311.RCMRecPw := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTRANSPW'']' );
  aSO311.RCMTransPw := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMSNR'']' );
  aSO311.RCMSnr := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMOnlineTimes'']' );
  aSO311.RCMOnlineTimes := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMInOctets'']' );
  aSO311.RCMInoctets := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMOutOctets'']' );
  aSO311.RCMOutOctets := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMInErrors'']' );
  aSO311.RCMInErrors := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMOutErrors'']' );
  aSO311.RCMOutErrors := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''PCIP'']' );
  aSO311.RPCIp := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTSRECPW'']' );
  aSO311.RCMTsrecPw := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTSRXSNR'']' );
  aSO311.RCMTsrXsNr := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CMTSUPMOD'']' );
  aSO311.RCMTsupMod := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseCPIPReg(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''OperResult'']' );
  aSO311.OperResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''CTIP'']' );
  aSO311.RCTIp := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseHFCTypeQuery(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''OperResult'']' );
  aSO311.OperResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''HFCType'']' );
  aSO311.RHfcType := GetDOMNodeValue( aValueNode );
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''NODEID'']' );
  aSO311.RNodeId := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParseResetCM(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''OperResult'']' );
  aSO311.OperResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMWebServiceManager.ParsePingCM(
  aSO311: PSO311);
var
  aRootNode, aValueNode: IDOMNode;
  aSelectNode: IDOMNodeSelect;
begin
  FXml.LoadFromXML( aSO311.SOAPResponse );
  aRootNode := FXml.Node.ChildNodes.Nodes['DataSet'].DOMNode;
  aSelectNode := ( aRootNode as IDOMNodeSelect );
  aValueNode := aSelectNode.selectNode( '//*[name()=''OperResult'']' );
  aSO311.OperResult := aValueNode.firstChild.nodeValue;
  {}
  aValueNode := aSelectNode.selectNode( '//*[name()=''FaultReason'']' );
  aSO311.FaultReason := GetDOMNodeValue( aValueNode );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.GetUpdateXmlFieldSql(aSO311: PSO311): String;
var
  aCode: Integer;
begin
  aCode := StrToInt( aSO311.MsgCode );
  case aCode of
    2: Result := SqlCMReg( aSO311 );
    12: Result := SqlResetCM( aSO311 );
    20: Result := SqlPingCM( aSO311 );
    25: Result := SqlClearCPE( aSO311 );
    13: Result := SqlCMStatusQuery( aSO311 );
    21: Result := SqlHFCTypeQuery( aSO311 );
    22: Result := SqlCMConnLog( aSO311 );
    23: Result := SqlCMQALog( aSO311 );
    26: Result := SqlCMPCLog( aSO311 );
    27: Result := SqlCPIPReg( aSO311 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlClearCPE(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set      ' +
    '   operresult = ''%s'',   ' +
    '   faultreason = ''%s''   ' +
    ' where cmdseqno = ''%s''  ',
    [FOwner, aSO311.OperResult, aSO311.FaultReason, aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlCMConnLog(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set       ' +
    '   queryresult = ''%s'',   ' +
    '   faultreason = ''%s'',   ' +
    '   ronlinetime = ''%s'',   ' +
    '   rofflinetime = ''%s'',  ' +
    '   rcmip = ''%s'',         ' +
    '   rpcip = ''%s''          ' +
    ' where cmdseqno = ''%s''   ',
    [FOwner, aSO311.QueryResult, aSO311.FaultReason,
     aSO311.ROnlineTime, aSO311.ROfflineTime,
     aSO311.RCMIp, aSO311.RPCIp,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlCMPCLog(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set         ' +
    '   queryresult = ''%s'',     ' +
    '   faultreason = ''%s'',     ' +
    '   rpconlinetime = ''%s'',   ' +
    '   rpcofflinetime = ''%s'',  ' +
    '   rcmmac = ''%s''           ' +
    ' where cmdseqno = ''%s''     ',
    [FOwner, aSO311.QueryResult, aSO311.FaultReason,
     aSO311.RPCOnlineTime, aSO311.RPCOfflineTime,
     aSO311.RCMMac,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlCMQALog(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set        ' +
    '   queryresult = ''%s'',    ' +
    '   faultreason = ''%s'',    ' +
    '   rdetcttime = ''%s'',     ' +
    '   rcmtsrecpw = ''%s'',     ' +
    '   rcmtranspw = ''%s'',     ' +
    '   rcmsnr = ''%s'',         ' +
    '   rcminoctets = ''%s'',    ' +
    '   rcmoutoctets = ''%s'',   ' +
    '   rcminerrors = ''%s'',    ' +
    '   rcmouterrors = ''%s''    ' +
    ' where cmdseqno = ''%s''    ',
    [FOwner, aSO311.QueryResult, aSO311.FaultReason,
     aSO311.RDetctTime, aSO311.RCMTsrecPw, aSO311.RCMTransPw,
     aSO311.RCMSnr, aSO311.RCMInoctets, aSO311.RCMOutOctets,
     aSO311.RCMInErrors, aSO311.RCMOutErrors,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlCMReg(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set        ' +
    '   operresult = ''%s'',     ' +
    '   faultreason = ''%s'',    ' +
    '   groupid = ''%s'',        ' +
    '   rctip = ''%s''           ' +
    ' where cmdseqno = ''%s''    ',
    [FOwner, aSO311.OperResult, aSO311.FaultReason,
     aSO311.GroupId, aSO311.RCTIp,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlCMStatusQuery(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set             ' +
    '   queryresult = ''%s'',         ' +
    '   faultreason = ''%s'',         ' +
    '   groupid = ''%s'',             ' +
    '   cmstatus = ''%s'',            ' +
    '   rcmip = ''%s'',               ' +
    '   rcmupfreq = ''%s'',           ' +
    '   rcmdownfreq = ''%s'',         ' +
    '   rcmrecpw = ''%s'',            ' +
    '   rcmtranspw = ''%s'',          ' +
    '   rcmsnr = ''%s'',              ' +
    '   rcmonlinetimes = ''%s'',      ' +
    '   rcminoctets = ''%s'',         ' +
    '   rcmoutoctets = ''%s'',        ' +
    '   rcminerrors = ''%s'',         ' +
    '   rcmouterrors = ''%s'',        ' +
    '   rpcip = ''%s'',               ' +
    '   rcmtsrecpw = ''%s'',          ' +
    '   rcmtsrxsnr = ''%s'',          ' +
    '   rcmtsupmod = ''%s''           ' +
    ' where cmdseqno = ''%s''         ',
    [FOwner, aSO311.QueryResult, aSO311.FaultReason,
     aSO311.GroupId, aSO311.CMStatus, aSO311.RCMIp,
     aSO311.RCMUpFreq, aSO311.RCMDownFreq,
     aSO311.RCMRecPw, aSO311.RCMTransPw,
     aSO311.RCMSnr, aSO311.RCMOnlineTimes,
     aSO311.RCMInoctets, aSO311.RCMOutOctets,
     aSO311.RCMInErrors, aSO311.RCMOutErrors,
     aSO311.RPCIp, aSO311.RCMTsrecPw,
     aSO311.RCMTsrXsNr, aSO311.RCMTsupMod,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlCPIPReg(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set             ' +
    '   operresult = ''%s'',          ' +
    '   faultreason = ''%s'',         ' +
    '   rctip = ''%s''                ' +
    ' where cmdseqno = ''%s''         ',
    [FOwner, aSO311.OperResult, aSO311.FaultReason,
     aSO311.RCTIp,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlHFCTypeQuery(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set             ' +
    '   operresult = ''%s'',          ' +
    '   faultreason = ''%s'',         ' +
    '   rhfctype = ''%s'',            ' +
    '   rnodeid = ''%s''              ' +
    ' where cmdseqno = ''%s''         ',
    [FOwner, aSO311.OperResult, aSO311.FaultReason,
     aSO311.RHfcType, aSO311.RNodeId,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlPingCM(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set             ' +
    '   operresult = ''%s'',          ' +
    '   faultreason = ''%s''          ' +
    ' where cmdseqno = ''%s''         ',
    [FOwner, aSO311.OperResult, aSO311.FaultReason,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

function TCMWebServiceManager.SqlResetCM(aSO311: PSO311): String;
begin
  Result := Format(
    ' update %s.so311 set             ' +
    '   operresult = ''%s'',          ' +
    '   faultreason = ''%s''          ' +
    ' where cmdseqno = ''%s''         ',
    [FOwner, aSO311.OperResult, aSO311.FaultReason,
     aSO311.CmdSeqNO] );
end;

{ ---------------------------------------------------------------------------- }

end.
