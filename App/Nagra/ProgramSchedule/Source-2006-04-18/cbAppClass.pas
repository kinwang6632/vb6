unit cbAppClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DateUtils, xmldom, XMLIntf,
  msxmldom, XMLDoc, Contnrs, DB, DBClient, ADODB, cbClass;

const

  COMM_DELAY = 150;
  SMS_PROGRAM_SCHEDULAR = 'SMS Schedular';

type
{ Application Runing State }
  TAppRunSate = ( rsStop, rsPlay );

{ TApplication User Define Option, Load from Access }
  TAppParam = class(TPersistent)
  private
    FDexSrc: String;
    FDexDest: String;
    FDexErr: String;
    FDexBackup: String;
    FDexLastProcFile: String;
    FWrapperSrc: String;
    FWrapperDest: String;
    FWrapperErr: String;
    FWrapperBackup: String;
    FWrapperLastProcFile: String;
    FAsRunSrc: String;
    FAsRunDest: String;
    FAsRunErr: String;
    FAsRunBackup: String;
    FAsRunLastProcFile: String;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    property DexSrc: String read FDexSrc write FDexSrc;
    property DexDest: String read FDexDest write FDexDest;
    property DexErr: String read FDexErr write FDexErr;
    property DexBackup: String read FDexBackup write FDexBackup;
    property WrapperSrc: String read FWrapperSrc write FWrapperSrc;
    property WrapperDest: String read FWrapperDest write FWrapperDest;
    property WrapperErr: String read FWrapperErr write FWrapperErr;
    property WrapperBackup: String read FWrapperBackup write FWrapperBackup;
    property AsRunSrc: String read FAsRunSrc write FAsRunSrc;
    property AsRunDest: String read FAsRunDest write FAsRunDest;
    property AsRunErr: String read FAsRunErr write FAsRunErr;
    property AsRunBackup: String read FAsRunBackup write FAsRunBackup;
    property DexLastProcFile: String read FDexLastProcFile write FDexLastProcFile;
    property WrapperLastProcFile: String read FWrapperLastProcFile write FWrapperLastProcFile;
    property AsRunLastProcFile: String read FAsRunLastProcFile write FAsRunLastProcFile;
    procedure AssureDirectoryExist;
  end;

  PClockItem = ^TClockItem;
  TClockItem = record
    IsSet: Boolean;
    Expire: Boolean;
    Execute: Boolean;
  end;

  TAppPeriod = class(TPersistent)
  private
    FList: TList;
    FPeriodType: String;
    FPeriodMinute: Integer;
    FLastExecute: TDateTime;
    function GetClockByIndex(const Index: Integer): PClockItem;
    function GetClockByName(const Name: String): PClockItem;
    function GetClockCount: Integer;
    procedure FreeClockItem;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    procedure CheckExpire;
    procedure ResetExecute;
    property PeriodType: String read FPeriodType write FPeriodType;
    property PeriodMinute: Integer read FPeriodMinute write FPeriodMinute;
    property LastExecute: TDateTime read FLastExecute write FLastExecute;
    property Clocks[const Index: Integer]: PClockItem read GetClockByIndex; default;
    property ClockNames[const Name: String]: PClockItem read GetClockByName;
    property ClockCount: Integer read GetClockCount;
  end;


  TAppDatabase = class(TPersistent)
  private
    FDbAccount: String;
    FDbPassoword: String;
    FDbSid: String;
    FDbRetrySec: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    property DbAccount: String read FDbAccount write FDbAccount;
    property DbPassoword: String read FDbPassoword write FDbPassoword;
    property DbSid: String read FDbSid write FDbSid;
    property DbRetrySec: Integer read FDbRetrySec write FDbRetrySec;
  end;

  TAppLogParam = class(TPersistent)
  private
    FActLoadDays: Integer;
    FMsgLoadDays: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    property ActLoadDays: Integer read FActLoadDays write FActLoadDays;
    property MsgLoadDays: Integer read FMsgLoadDays write FMsgLoadDays;
  end;


  TSMTPSetting = class
  private
    FServer: String;
    FEmail: String;
    FAccount: String;
    FPassword: String;
  public
    property Server: String read FServer write FServer;
    property Email: String read FEmail write FEmail;
    property Account: String read FAccount write FAccount;
    property Password: String read FPassword write FPassword;
  end;


  TAppNotify = class(TPersistent)
  private
    FErrorNotify: Boolean;
    FChangeNotiy: Boolean;
    FErrorEmail: TStringList;
    FChangeEmail: TStringList;
    FSMTPSetting: TSMTPSetting;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    property ErrorNotify: Boolean read FErrorNotify write FErrorNotify;
    property ChangeNotiy: Boolean read FChangeNotiy write FChangeNotiy;
    property ErrorEmail: TStringList read FErrorEmail;
    property ChangeEmail: TStringList read FChangeEmail;
    property SMTPSetting: TSMTPSetting read FSMTPSetting write FSMTPSetting;
  end;


  { Monitor XML Source Directory  }

  TXmlDirectory = ( xdsWrapper = 1, xdsDex = 2, xdsAsRun = 3 );

  { Nagra Program List XML File Type }

  TXmlFileType = ( xfMerge = 0 , xfWrapper = 1, xfPpv = 2, xfSub = 3, xfAsRun = 4 );


{ Xml Parse Record Status }
  PActObject = ^TActObject;
  TActObject = record
    KeyId: String;
    ActDate: String;
    ActTime: String;
    ActFileName: String;
    ActFileSize: String;
    ActFilePath: String;
    ActSource: String;
    ActFileType: TXmlFileType;
    ActStatus: String;
    ActCost: Integer;
    ActProgress: Integer;
    ActErrMsg: String;
  end;

  TParseSubject = class(TSubject)
  private
    FActionObject: Pointer;
    procedure SetActionObject(Value: Pointer);
  public
    property ActionObject: Pointer read FActionObject write SetActionObject;
  end;


  TXmlFile = class(TObject)
  private
    FXml: TXMLDocument;
    FXmlFileName: String;
    FParerList: TObjectList;
    procedure SetXmlFileName(const aValue: String);
  public
    constructor Create; virtual;
    destructor Destroy; override;
    function Parse(var aMsg: String): Boolean; virtual; abstract;
    function Save(var aMsg: String): Boolean; virtual; abstract;
    property XmlFileName: String read FXmlFileName write SetXmlFileName;
  end;


  { S2 Wrapper Xml File }
  TWrapperXmlFile = class(TXmlFile)
  public
    constructor Create; override;
    destructor Destroy; override;
    function Parse(var aMsg: String): Boolean; override;
    function Save(var aMsg: String): Boolean; override;
  end;


  { Dex Ppv Xml File }
  TDexPpvXmlFile = class(TXmlFile)
  public
    constructor Create; override;
    destructor Destroy; override;
    function Parse(var aMsg: String): Boolean; override;
    function Save(var aMsg: String): Boolean; override;
  end;


  { Dex Sub Xml File }
  TDexSubXmlFile = class(TXmlFile)
  public
    constructor Create; override;
    destructor Destroy; override;
    function Parse(var aMsg: String): Boolean; override;
    function Save(var aMsg: String): Boolean; override;
  end;
  

  { AsRun Xml File }
  TAsRunXmlFile = class(TXmlFile)
  public
    constructor Create; override;
    destructor Destroy; override;
    function Parse(var aMsg: String): Boolean; override;
    function Save(var aMsg: String): Boolean; override;
  end;


  TXmlFileClass = class of TXmlFile;

  
  TXmlParser = class(TObject)
  private
    FDataBuffer: TClientDataSet;
    FXml: TXMLDocument;
    FMsg: String;
    FDataConnection: TADOConnection;
    FDataUpdate: TADOQuery;
    FDataInsert: TADOQuery;
    FDataReader: TADOQuery;
    FDataDelete: TADOQuery;
    FName: String;
  protected
    procedure PrepareDataBuffer; virtual;
    procedure UnPrepareDataBuffer; virtual;
    procedure InternalParse; virtual;
    procedure InternalApplyUpdate; virtual;
    property DataConnection: TADOConnection read FDataConnection write FDataConnection;
    property DataUpdate: TADOQuery read FDataUpdate write FDataUpdate;
    property DataInsert: TADOQuery read FDataInsert write FDataInsert;
    property DataReader: TADOQuery read FDataReader write FDataReader;
    property DataDelete: TADOQuery read FDataDelete write FDataDelete;
  public
    constructor Create(aXml: TXMLDocument);
    destructor Destroy; override;
    procedure Parse;
    procedure ApplyUpdate;
    property Msg: String read FMsg;
    property DataBuffer: TClientDataSet read FDataBuffer;
    property Name: String read FName write FName;
  end;


  { S2 Wrapper Xml Parser }
  TParseSO159 = class(TXmlParser)
  protected
    procedure PrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  end;

  { S2 Wrapper Xml Parser }
  TParseSO161 = class(TXmlParser)
  protected
    procedure PrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  end;


  { S2 Wrapper Xml Parser }
  TParseSO162 = class(TXmlParser)
  protected
    procedure PrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  end;

  { Dex Xml Parser }
  TParseSO158Ppv = class(TXmlParser)
  private
    FDeleteProduct: TClientDataSet;
  protected
    procedure PrepareDataBuffer; override;
    procedure UnPrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  public
    constructor Create(aXml: TXMLDocument);
    destructor Destroy; override;
  end;

  { Dex Xml Parser }

  TParseSO158Sub = class(TXmlParser)
  private
    FDeleteProduct: TClientDataSet;
  protected
    procedure PrepareDataBuffer; override;
    procedure UnPrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  public
    constructor Create(aXml: TXMLDocument);
    destructor Destroy; override;
  end;

  { Dex Xml Parser }
  TParseSO163 = class(TXmlParser)
  private
    FDeleteProduct: TClientDataSet;
  protected
    procedure PrepareDataBuffer; override;
    procedure UnPrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  public
    constructor Create(aXml: TXMLDocument);
    destructor Destroy; override;
  end;


  { AsRun Xml Parser }

  TParseAsRun = class(TXmlParser)
  private
    FOrignalEvent: TClientDataSet;
    FChannelTime: TClientDataSet;
  protected
    procedure PrepareDataBuffer; override;
    procedure UnPrepareDataBuffer; override;
    procedure InternalParse; override;
    procedure InternalApplyUpdate; override;
  public
    constructor Create(aXml: TXMLDocument);
    destructor Destroy; override;
  end;


  TStoreAction = ( saBefore, saAftrer );

  TLogType = ( ltProductDelete, ltEventDelete, ltEventTimeChange,
    ltEpgPriceChange );

  TMergeChange = class
  private
    FSo160: TClientDataSet;
    FProductBuffer: TClientDataSet;
    FEventBuffer: TClientDataSet;
    FProductBuffer2: TClientDataSet;
    FEventBuffer2: TClientDataSet;
    FLogText: TClientDataSet;
    FDataConnection: TADOConnection;
    FDataReader: TADOQuery;
    FDataWriter: TADOQuery;
    FDataDelete: TADOQuery;
    FDataInsert: TADOQuery;
    FDataUpdate: TADOQuery;
  protected
    procedure PrepareDataBuffer;
    procedure UnPrepareDataBuffer;
    function StoreRecord(const aAction: TStoreAction; var aMsg: String): Boolean;
    function DoSo160(var aMsg: String): Boolean;
    function DoSo154(var aMsg: String): Boolean;
    function DoSo162(var aMsg: String): Boolean;
    function DoMergin(var aMsg: String): Boolean;
    function DoMerginEx(var aMsg: String): Boolean;
    function MakeLogText(const aType: TLogType): String;
    function DoMakeLog(var aMsg: String): Boolean;
    function DoWriteLog(var aMsg: String): Boolean;
  public
    constructor Create;
    destructor Destroy; override;
    function Merge(var aMsg: String): Boolean;
    function DetectChangeCount(var aMsg: String; var aCount: Integer): Boolean;
    property DataConnection: TADOConnection read FDataConnection write FDataConnection;
    property DataReader: TADOQuery read FDataReader write FDataReader;
    property DataWriter: TADOQuery read FDataWriter write FDataWriter;
    property DataInsert: TADOQuery read FDataInsert write FDataInsert;
    property DataUpdate: TADOQuery read FDataUpdate write FDataUpdate;
    property DataDelete: TADOQuery read FDataDelete write FDataDelete;
  end;


  { Nagra Program List XML File Type Text }

  var XmlFileTypeText: array [TXmlFileType] of String =
    ( 'None', 'Wrapper', 'Ppv', 'Sub', 'AsRun' );


implementation

uses cbParseController, cbUtilis;

{ ---------------------------------------------------------------------------- }

procedure GetParentalratingName(aReader: TADOQuery;
  var aParentalrating, aParentalratingName: String);
begin
  aParentalratingName := EmptyStr;
  if Assigned( aReader ) and ( aParentalrating <> EmptyStr ) then
  begin
    aReader.Close;
    aReader.SQL.Text := Format(
      ' select codeno, description from cd029       ' +
      '  where ''%s'' between age1 and age2 ', [aParentalrating] );
    aReader.Open;
    aParentalrating := aReader.FieldByName( 'codeno' ).AsString;
    aParentalratingName := aReader.FieldByName( 'description' ).AsString;
    aReader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetChannelNo(aReader: TADOQuery; const aChannelId: String): String;
begin
  Result := EmptyStr;
  if Assigned( aReader ) and ( aChannelId <> EmptyStr ) then
  begin
    aReader.Close;
    aReader.SQL.Text := Format(
      ' select channelno from so162       ' +
      '  where channelid = ''%s''         ', [aChannelId] );
    aReader.Open;
    Result := aReader.FieldByName( 'channelno' ).AsString;
    aReader.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function ROCDate: String;
var
  aYear, aMonth, aDay: Word;
begin
  DecodeDate( Date, aYear, aMonth, aDay );
  Result := Format( '%d/%.2d/%.2d', [aYear-1911, aMonth, aDay] ) + #32 +
    FormatDateTime( 'hh:nn:ss', Now );
end;

{ ---------------------------------------------------------------------------- }

{ TAppParam }

procedure TAppParam.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TAppParam ) then
  begin
    TAppParam( Dest ).DexSrc := Self.FDexSrc;
    TAppParam( Dest ).DexDest := Self.FDexDest;
    TAppParam( Dest ).DexErr := Self.FDexErr;
    TAppParam( Dest ).DexBackup := Self.FDexBackup;
    TAppParam( Dest ).DexLastProcFile := Self.DexLastProcFile;
    TAppParam( Dest ).WrapperSrc := Self.WrapperSrc;
    TAppParam( Dest ).WrapperDest := Self.WrapperDest;
    TAppParam( Dest ).WrapperErr := Self.WrapperErr;
    TAppParam( Dest ).WrapperBackup := Self.WrapperBackup;
    TAppParam( Dest ).WrapperLastProcFile := Self.WrapperLastProcFile;
    TAppParam( Dest ).AsRunSrc := Self.AsRunSrc;
    TAppParam( Dest ).AsRunDest := Self.AsRunDest;
    TAppParam( Dest ).AsRunErr := Self.AsRunErr;
    TAppParam( Dest ).AsRunBackup := Self.AsRunBackup;
    TAppParam( Dest ).AsRunLastProcFile := Self.AsRunLastProcFile;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

procedure TAppParam.AssureDirectoryExist;

  { ------------------------------------------------ }

  procedure DoMakeDir(const aDir: String);
  begin
    if ( aDir <> EmptyStr ) then
      ForceDirectories( aDir );
  end;

  { ------------------------------------------------ }

begin
  DoMakeDir( FDexSrc );
  DoMakeDir( FDexDest );
  DoMakeDir( FDexErr );
  DoMakeDir( FDexBackup );
  DoMakeDir( FWrapperSrc );
  DoMakeDir( FWrapperDest );
  DoMakeDir( FWrapperErr );
  DoMakeDir( FWrapperBackup );
  DoMakeDir( FAsRunSrc );
  DoMakeDir( FAsRunDest );
  DoMakeDir( FAsRunErr );
  DoMakeDir( FAsRunBackup );
end;

{ ---------------------------------------------------------------------------- }

{ TAppPeriod }

constructor TAppPeriod.Create;
var
  aIndex: Integer;
  aClokcItem: PClockItem;
begin
  inherited Create;
  FList := TList.Create;
  for aIndex := 0 to 23 do
  begin
    New( aClokcItem );
    aClokcItem.IsSet := False;
    aClokcItem.Expire := False;
    aClokcItem.Execute := False;
    FList.Add( aClokcItem );
  end;
  CheckExpire;
end;

{ ---------------------------------------------------------------------------- }

destructor TAppPeriod.Destroy;
begin
  FreeClockItem;
  FList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppPeriod.AssignTo(Dest: TPersistent);
var
  aIndex: Integer;
begin
  if ( Dest is TAppPeriod ) then
  begin
    TAppPeriod( Dest ).PeriodType := Self.FPeriodType;
    TAppPeriod( Dest ).PeriodMinute := Self.FPeriodMinute;
    TAppPeriod( Dest ).LastExecute := Self.FLastExecute;
    for aIndex := 0 to TAppPeriod( Dest ).ClockCount - 1 do
    begin
      TAppPeriod( Dest ).Clocks[aIndex].IsSet := Self.Clocks[aIndex].IsSet;
      TAppPeriod( Dest ).Clocks[aIndex].Expire := Self.Clocks[aIndex].Expire;
    end;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

function TAppPeriod.GetClockByIndex(const Index: Integer): PClockItem;
begin
  Result := FList[Index];
end;

{ ---------------------------------------------------------------------------- }

function TAppPeriod.GetClockByName(const Name: String): PClockItem;
begin
  if ( Name = '24' ) then
    Result := FList[0]
  else
    Result := FList[StrToInt( Name )];
end;

{ ---------------------------------------------------------------------------- }

function TAppPeriod.GetClockCount: Integer;
begin
  Result := FList.Count;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppPeriod.FreeClockItem;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if Assigned( FList[aIndex] ) then
    begin
      Dispose( PClockItem( FList[aIndex] ) );
      FList[aIndex] := nil;
    end;
  end;
  FList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppPeriod.CheckExpire;
var
  aIndex: Integer;
  aHour: Word;
begin
  aHour := HourOf( Now );
  for aIndex := 0 to FList.Count - 1 do
    PClockItem( FList[aIndex] ).Expire := ( aHour > aIndex );
end;

{ ---------------------------------------------------------------------------- }

procedure TAppPeriod.ResetExecute;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
    PClockItem( FList[aIndex] ).Execute := False;
end;

{ ---------------------------------------------------------------------------- }

{ TAppDatabase }

procedure TAppDatabase.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TAppDatabase ) then
  begin
    TAppDatabase( Dest ).DbAccount := Self.DbAccount;
    TAppDatabase( Dest ).DbPassoword := Self.DbPassoword;
    TAppDatabase( Dest ).DbSid := Self.DbSid;
    TAppDatabase( Dest ).DbRetrySec := Self.DbRetrySec;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TAppLogs }

procedure TAppLogParam.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TAppLogParam ) then
  begin
    TAppLogParam( Dest ).ActLoadDays := Self.ActLoadDays;
    TAppLogParam( Dest ).MsgLoadDays := Self.MsgLoadDays;
  end else
  inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TAppNotify }

constructor TAppNotify.Create;
begin
  inherited Create;
  FErrorNotify := False;
  FChangeNotiy := False;
  FErrorEmail := TStringList.Create;
  FChangeEmail := TStringList.Create;
  FSMTPSetting := TSMTPSetting.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TAppNotify.Destroy;
begin
  FErrorEmail.Free;
  FChangeEmail.Free;
  FSMTPSetting.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TAppNotify.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TAppNotify ) then
  begin
    TAppNotify( Dest ).ErrorNotify := Self.FErrorNotify;
    TAppNotify( Dest ).ChangeNotiy := Self.FChangeNotiy;
    TAppNotify( Dest ).ErrorEmail.Assign( Self.FErrorEmail );
    TAppNotify( Dest ).ChangeEmail.Assign( Self.ChangeEmail );
    TAppNotify( Dest ).SMTPSetting.Server := Self.FSMTPSetting.Server;
    TAppNotify( Dest ).SMTPSetting.Email := Self.FSMTPSetting.FEmail;
    TAppNotify( Dest ).SMTPSetting.Account := Self.FSMTPSetting.FAccount;
    TAppNotify( Dest ).SMTPSetting.Password := Self.FSMTPSetting.FPassword;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TParseSubject }

procedure TParseSubject.SetActionObject(Value: Pointer);
begin
  FActionObject := Value;
  if Assigned( FActionObject ) then NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TXmlFile }

constructor TXmlFile.Create;
begin
  FXml := TXMLDocument.Create( nil );
  FParerList := TObjectList.Create( False );
end;

{ ---------------------------------------------------------------------------- }

destructor TXmlFile.Destroy;
var
  aIndex: Integer;
begin
  FXml.Active := False;
  FXml.Free;
  for aIndex := 0 to FParerList.Count - 1do
    FParerList[aIndex].Free;
  FParerList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlFile.SetXmlFileName(const aValue: String);
begin
  FXmlFileName := aValue;
  FXml.LoadFromFile( FXmlFileName );
end;

{ ---------------------------------------------------------------------------- }

{ TWrapperXmlFile }

constructor TWrapperXmlFile.Create;
var
  aIndex: Integer;
  aParse: TXmlParser;
begin
  inherited Create;
  for aIndex := 1 to 3 do
  begin
    case aIndex of
      1: aParse := TParseSO162.Create( Self.FXml );
      2: aParse := TParseSO159.Create( Self.FXml );
      3: aParse := TParseSO161.Create( Self.FXml );
    else
      aParse := nil;
    end;
    case aIndex of
      1: aParse.Name := 'SO162';
      2: aParse.Name := 'SO159';
      3: aParse.Name := 'SO161';
    end;
    if Assigned( aParse ) then
    begin
      aParse.DataConnection := ParseController.DataConnection;
      aParse.DataUpdate := ParseController.DataUpdate;
      aParse.DataInsert := ParseController.DataInsert;
      aParse.DataDelete := ParseController.DataDelete;
      aParse.DataReader := ParseController.DataReader;
      FParerList.Add( aParse );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

destructor TWrapperXmlFile.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TWrapperXmlFile.Parse(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.Parse;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TWrapperXmlFile.Save(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.ApplyUpdate;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TDexPpvXmlFile }

constructor TDexPpvXmlFile.Create;
var
  aIndex: Integer;
  aParse: TXmlParser;
begin
  inherited Create;
  for aIndex := 1 to 1 do
  begin
    case aIndex of
      1: aParse := TParseSO158Ppv.Create( Self.FXml );
    else
      aParse := nil;
    end;
    case aIndex of
      1: aParse.Name := 'SO158';
    else
      aParse := nil;
    end;
    if Assigned( aParse ) then
    begin
      aParse.DataConnection := ParseController.DataConnection;
      aParse.DataUpdate := ParseController.DataUpdate;
      aParse.DataInsert := ParseController.DataInsert;
      aParse.DataDelete := ParseController.DataDelete;
      FParerList.Add( aParse );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

destructor TDexPpvXmlFile.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TDexPpvXmlFile.Parse(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.Parse;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDexPpvXmlFile.Save(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.ApplyUpdate;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TXmlParser }

constructor TXmlParser.Create(aXml: TXMLDocument);
begin
  FDataBuffer := TClientDataSet.Create( nil );
  FXml := aXml;
  FDataConnection := nil;
  FDataUpdate := nil;
  FDataInsert := nil;
  FDataReader := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TXmlParser.Destroy;
begin
  inherited;
  FDataBuffer.Free;
  FXml := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlParser.InternalApplyUpdate;
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlParser.InternalParse;
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlParser.Parse;
begin
  UnPrepareDataBuffer;
  PrepareDataBuffer;
  FMsg := EmptyStr;
  try
    InternalParse;
  except
    on E: Exception do FMsg := Format( '%s - %s', [FName, E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlParser.ApplyUpdate;
begin
  FDataConnection.BeginTrans;
  try
    InternalApplyUpdate;
    FDataConnection.CommitTrans;
  except
    on E: Exception do
    begin
      FDataConnection.RollbackTrans;
      FMsg := Format( '%s - %s', [FName, E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlParser.PrepareDataBuffer;
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TXmlParser.UnPrepareDataBuffer;
begin
  if not VarIsNull( FDataBuffer.Data ) then
    FDataBuffer.EmptyDataSet;
  FDataBuffer.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

{ TParseSO159 }

procedure TParseSO159.PrepareDataBuffer;
begin
  inherited;
  { SO159 Epg XML資料檔 }
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'ChannelPeriodBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ChannelPeriodEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'FileName', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'CreationDate', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'EventBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'Duration', ftInteger );
    FDataBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FDataBuffer.FieldDefs.Add( 'EventType', ftString, 1 );
    FDataBuffer.FieldDefs.Add( 'Name', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'EpgDesc', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'EngName', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'EngEpgDesc', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'ParentalRating', ftInteger );
    FDataBuffer.FieldDefs.Add( 'ContentNibble1', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'ContentNibble2', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UserNibble1', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UserNibble2', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UpdTime', ftString, 19 );
    FDataBuffer.FieldDefs.Add( 'UpdEn', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'UTCEventBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'Action', ftInteger );
  end;
  FDataBuffer.IndexName := EmptyStr;
  FDataBuffer.CreateDataSet;
  FDataBuffer.AddIndex( 'IDX1', 'ChannelId;EventBeginTime', []  );
  FDataBuffer.IndexName := 'IDX1';
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO159.InternalParse;

type
  TSo159 = record
    ChannelPeriodBeginTime: String;
    ChannelPeriodEndTime: String;
    FileName: String;
    CreationDate: String;
    ChannelId: String;
    EventBeginTime: String;
    Duration: String;
    EventId: String;
    EvnetType: String;
    Name: String;
    EpgDesc: String;
    EngName: String;
    EngEpgDesc: String;
    ParentalRating: String;
    ContentNibble1: String;
    ContentNibble2: String;
    UserNibble1: String;
    UserNibble2: String;
    UpdTime: String;
    UpdEn: String;
    ParentalRatingName: String;
    UTCEventBeginTime: String;
  end;

  { ------------------------------------------------------ }

  function GetEpgDesc(aNode: IDOMNode): String;
  var
    aIndex: Integer;
    aTmpNode: IDOMNode;
    aNodeLis: IDOMNodeList;
  begin
    Result := EmptyStr;
    if not Assigned( aNode ) then Exit;
    aNodeLis := ( aNode as IDOMNodeSelect ).selectNodes( 'ExtendedInfo' );
    for aIndex := 0 to aNodeLis.length - 1 do
    begin
      aTmpNode := aNodeLis.item[aIndex].attributes.getNamedItem( 'name' );
      if Assigned( aTmpNode ) then
        Result := ( Result + aTmpNode.nodeValue + ':' );
      Result := Result + aNodeLis.item[aIndex].firstChild.nodeValue;
      if ( Result <> EmptyStr ) then Result := ( Result + #13#10 );
    end;
    aTmpNode := ( aNode as IDOMNodeSelect ).selectNode( 'Description' );
    if Assigned( aTmpNode ) then
      Result := Result + aTmpNode.firstChild.nodeValue;
  end;

  { ------------------------------------------------------ }

var
  aIDOMList, aIDOMList2, aIDOMList3: IDOMNodeList;
  aEventNode, aLanguageNode, aNode: IDOMNode;
  aIndex, aIndex2, aIndex3: Integer;
  a159: TSo159;
  aCreationDate, aFileName: String;
begin
  aFileName := FileNameWithoutExt( ExtractFileName( FXml.FileName ) );
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'BroadcastData' );
  aCreationDate := aIDOMList.item[0].attributes.getNamedItem( 'creationDate' ).nodeValue;
  aIDOMList := nil;
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'ChannelPeriod' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    aIDOMList2 := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNodes( 'Event' );
    for aIndex2 := 0 to aIDOMList2.length - 1 do
    begin
      { 判斷是否是 PPV Product 的 Event ( 或是可能成為PPV Event ), EventType = 'P' }
      aNode := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode( 'EventType' );
      if ( aNode.firstChild.nodeValue <> 'P' ) then Continue;
      {}
      ZeroMemory( @a159, SizeOf( a159 ) );
      {}
      a159.ChannelPeriodBeginTime := ( aIDOMList.item[aIndex] ).attributes.getNamedItem( 'beginTime' ).nodeValue;
      a159.ChannelPeriodEndTime := ( aIDOMList.item[aIndex] ).attributes.getNamedItem( 'endTime' ).nodeValue;
      {}
      a159.ChannelPeriodBeginTime := UTCDateTimeToLocalTime(
         FormatMaskText( '####/##/## ##:##:##;0;_', a159.ChannelPeriodBeginTime ) );
      a159.ChannelPeriodEndTime := UTCDateTimeToLocalTime(
         FormatMaskText( '####/##/## ##:##:##;0;_', a159.ChannelPeriodEndTime ) );
      {}
      a159.ChannelPeriodBeginTime := TrimChar( a159.ChannelPeriodBeginTime, ['/', ':', ' '] );
      a159.ChannelPeriodEndTime := TrimChar( a159.ChannelPeriodEndTime, ['/', ':', ' '] );
      {}
      a159.FileName := aFileName;
      a159.CreationDate := aCreationDate;
      {}
      a159.ChannelId := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNode(
        'ChannelId' ).firstChild.nodeValue;
      {}
      aEventNode := ( aIDOMList2.item[aIndex2] );
      a159.UTCEventBeginTime := aEventNode.attributes.getNamedItem( 'beginTime' ).nodeValue;
      a159.EventBeginTime := UTCDateTimeToLocalTime(
         FormatMaskText( '####/##/## ##:##:##;0;_', a159.UTCEventBeginTime ) );
      {}
      a159.EventBeginTime := TrimChar( a159.EventBeginTime, ['/', ':', ' '] );
      {}
      a159.Duration := aEventNode.attributes.getNamedItem( 'duration' ).nodeValue;
      {}
      a159.EventId := ( aEventNode as IDOMNodeSelect ).selectNode(
        'EventId' ).firstChild.nodeValue;
      {}
      a159.EvnetType := ( aEventNode as IDOMNodeSelect ).selectNode(
        'EventType' ).firstChild.nodeValue;
      {}
      aIDOMList3 := ( ( aEventNode as IDOMNodeSelect ).selectNode(
        'EpgProduction' ) as IDOMNodeSelect ).selectNodes( 'EpgText' );
      for aIndex3 := 0 to aIDOMList3.length - 1 do
      begin
        aLanguageNode := aIDOMList3.item[aIndex3];
        if ( aLanguageNode.attributes.getNamedItem( 'language' ).nodeValue = 'chi' ) then
        begin
          a159.Name := ( aLanguageNode as IDOMNodeSelect ).selectNode(
            'Name' ).firstChild.nodeValue;
          a159.EpgDesc := GetEpgDesc( aLanguageNode );
        end else
        begin
          a159.EngName := ( aLanguageNode as IDOMNodeSelect ).selectNode(
            'Name' ).firstChild.nodeValue;
          a159.EngEpgDesc := GetEpgDesc( aLanguageNode );
        end;
      end;
      aIDOMList3 := nil;
      {}
      aNode := ( ( aEventNode as IDOMNodeSelect ).selectNode(
        'EpgProduction' ) as IDOMNodeSelect ).selectNode( 'ParentalRating' );
      if Assigned( aNode ) then
        a159.ParentalRating := aNode.firstChild.nodeValue;
      {}
      aNode := ( ( aEventNode as IDOMNodeSelect ).selectNode(
        'EpgProduction' ) as IDOMNodeSelect ).selectNode( 'DvbContent' );
      if Assigned( aNode ) then
      begin
        if Assigned( ( aNode as IDOMNodeSelect ).selectNode( 'Content' ) ) then
        begin
          a159.ContentNibble1 := ( aNode as IDOMNodeSelect ).selectNode(
            'Content' ) .attributes.getNamedItem( 'nibble1' ).nodeValue;
          a159.ContentNibble2 := ( aNode as IDOMNodeSelect ).selectNode(
            'Content' ) .attributes.getNamedItem( 'nibble2' ).nodeValue;
        end;
        {}
        if Assigned( ( aNode as IDOMNodeSelect ).selectNode( 'User' ) ) then
        begin
          a159.UserNibble1 := ( aNode as IDOMNodeSelect ).selectNode(
            'User' ) .attributes.getNamedItem( 'nibble1' ).nodeValue;
          a159.UserNibble2 := ( aNode as IDOMNodeSelect ).selectNode(
            'User' ) .attributes.getNamedItem( 'nibble2' ).nodeValue;
        end;
      end;
      {}
      a159.UpdTime := ROCDate;
      a159.UpdEn := SMS_PROGRAM_SCHEDULAR;
      {}
      FDataBuffer.Append;
      FDataBuffer.FieldByName( 'ChannelPeriodBeginTime' ).AsString := a159.ChannelPeriodBeginTime;
      FDataBuffer.FieldByName( 'ChannelPeriodEndTime' ).AsString := a159.ChannelPeriodEndTime;
      FDataBuffer.FieldByName( 'FileName' ).Value := Copy( a159.FileName, Length( a159.FileName ) - 13, 14 );
      FDataBuffer.FieldByName( 'CreationDate' ).Value := a159.CreationDate;
      FDataBuffer.FieldByName( 'ChannelId' ).Value := a159.ChannelId;
      FDataBuffer.FieldByName( 'EventBeginTime' ).Value := a159.EventBeginTime;
      FDataBuffer.FieldByName( 'Duration' ).Value := a159.Duration;
      FDataBuffer.FieldByName( 'EventId' ).Value := a159.EventId;
      FDataBuffer.FieldByName( 'EventType' ).Value := a159.EvnetType;
      FDataBuffer.FieldByName( 'Name' ).Value := a159.Name;
      FDataBuffer.FieldByName( 'EpgDesc' ).Value := a159.EpgDesc;
      FDataBuffer.FieldByName( 'EngName' ).Value := a159.EngName;
      FDataBuffer.FieldByName( 'EngEpgDesc' ).Value := a159.EngEpgDesc;
      FDataBuffer.FieldByName( 'ParentalRating' ).Value := a159.ParentalRating;
      FDataBuffer.FieldByName( 'ContentNibble1' ).Value := a159.ContentNibble1;
      FDataBuffer.FieldByName( 'ContentNibble2' ).Value := a159.ContentNibble2;
      FDataBuffer.FieldByName( 'UserNibble1' ).Value := a159.UserNibble1;
      FDataBuffer.FieldByName( 'UserNibble2' ).Value := a159.UserNibble2;
      FDataBuffer.FieldByName( 'UpdTime' ).Value := a159.UpdTime;
      FDataBuffer.FieldByName( 'UpdEn' ).Value := a159.UpdEn;
      FDataBuffer.FieldByName( 'UTCEventBeginTime' ).Value := a159.UTCEventBeginTime;
      FDataBuffer.FieldByName( 'Action' ).Value := '2';
      FDataBuffer.Post;
    end;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO159.InternalApplyUpdate;
var
  aRecordEffect: Integer;
  aPreChannelId: String;
begin
  inherited;
  if ( FDataBuffer.IsEmpty ) then Exit;
  FDataDelete.SQL.Clear;
  FDataInsert.SQL.Clear;
  FDataUpdate.SQL.Clear;
  FDataDelete.SQL.Text :=
    '  UPDATE SO159 SET eStopFlag = 1, Action = 2                   ' +
    '   WHERE ChannelId = :ChannelId                                ' +
    '     AND ( EventBeginTime >= :St and EventBeginTime < :Ed )    ';
  FDataBuffer.First;
  aPreChannelId := EmptyStr;
  while not FDataBuffer.Eof do
  begin
    if ( aPreChannelId <> FDataBuffer.FieldByName( 'ChannelId' ).AsString ) then
    begin
      FDataDelete.Parameters.ParamByName( 'St' ).Value :=
        FDataBuffer.FieldByName( 'ChannelPeriodBeginTime' ).AsString;
      FDataDelete.Parameters.ParamByName( 'Ed' ).Value :=
        FDataBuffer.FieldByName( 'ChannelPeriodEndTime' ).AsString;
      FDataDelete.Parameters.ParamByName( 'ChannelId' ).Value :=
        FDataBuffer.FieldByName( 'ChannelId' ).AsString;
      FDataDelete.ExecSQL;
      aPreChannelId := FDataBuffer.FieldByName( 'ChannelId' ).AsString;
    end;
    FDataBuffer.Next;
  end;
  FDataInsert.SQL.Text :=
    ' INSERT INTO SO159 ( FileName, CreationDate, ChannelId, EventBeginTime, ' +
    '   Duration, EventId, EventType, Name, EpgDesc, EngName, EngEpgDesc,    ' +
    '   ParentalRating, ContentNibble1, ContentNibble2, UserNibble1,         ' +
    '   UserNibble2, UpdTime, UpdEn, eStopFlag, UTCEventBeginTime,           ' +
    '   Action )                                                             ' +
    ' VALUES ( :FileName, :CreationDate, :ChannelId, :EventBeginTime,        ' +
    '   :Duration, :EventId, :EventType, :Name, :EpgDesc, :EngName,          ' +
    '   :EngEpgDesc, :ParentalRating, :ContentNibble1, :ContentNibble2,      ' +
    '   :UserNibble1, :UserNibble2, :UpdTime, :UpdEn, 0,                     ' +
    '   :UTCEventBeginTime, 1 )                                              ';
  FDataUpdate.SQL.Text :=
    ' UPDATE SO159 SET FileName = :FileName, CreationDate = :CreationDate,   ' +
    '   EventBeginTime = :EventBeginTime, Duration = :Duration,              ' +
    '   EventType = :EventType, Name = :Name, EpgDesc = :EpgDesc,            ' +
    '   EngName = :EngName, EngEpgDesc = :EngEpgDesc,                        ' +
    '   ParentalRating = :ParentalRating, ContentNibble1 = :ContentNibble1,  ' +
    '   ContentNibble2 = :ContentNibble2, UserNibble1 = :UserNibble1,        ' +
    '   UserNibble2 = :UserNibble2, UpdTime = :UpdTime, UpdEn = :UpdEn,      ' +
    '   eStopFlag = 0, ChannelId = :ChannelId, Action = :Action,             ' +
    '   UTCEventBeginTime = :UTCEventBeginTime                               ' +
    ' WHERE EventId = :EventId                                               ';
  {}
  { SO161 存放的是一般 Event, 這裏 SO159 存放的是 PPV Event, 如果出現在 PPV 裏
    必須把 SO159 的刪掉, 因為 Event 從一般的 evnet 變成 PPV Event }
  FDataDelete.SQL.Text :=
    ' DELETE FROM SO161 WHERE EventId = :EventId  ';
  {}
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    FDataUpdate.Parameters.ParamByName( 'FileName' ).Value :=
      FDataBuffer.FieldByName( 'FileName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'CreationDate' ).Value :=
      FDataBuffer.FieldByName( 'CreationDate' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EventBeginTime' ).Value :=
      FDataBuffer.FieldByName( 'EventBeginTime' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Duration' ).Value :=
      FDataBuffer.FieldByName( 'Duration' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EventType' ).Value :=
      FDataBuffer.FieldByName( 'EventType' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Name' ).Value :=
      FDataBuffer.FieldByName( 'Name' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EpgDesc' ).Value :=
      FDataBuffer.FieldByName( 'EpgDesc' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngName' ).Value :=
      FDataBuffer.FieldByName( 'EngName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngEpgDesc' ).Value :=
      FDataBuffer.FieldByName( 'EngEpgDesc' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ParentalRating' ).Value :=
      FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ContentNibble1' ).Value :=
      FDataBuffer.FieldByName( 'ContentNibble1' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ContentNibble2' ).Value :=
      FDataBuffer.FieldByName( 'ContentNibble2' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UserNibble1' ).Value :=
      FDataBuffer.FieldByName( 'UserNibble1' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UserNibble2' ).Value :=
      FDataBuffer.FieldByName( 'UserNibble2' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UpdTime' ).Value :=
      FDataBuffer.FieldByName( 'UpdTime' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UpdEn' ).Value :=
      FDataBuffer.FieldByName( 'UpdEn' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EventId' ).Value :=
      FDataBuffer.FieldByName( 'EventId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ChannelId' ).Value :=
      FDataBuffer.FieldByName( 'ChannelId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Action' ).Value :=
      FDataBuffer.FieldByName( 'Action' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UTCEventBeginTime' ).Value :=
      FDataBuffer.FieldByName( 'UTCEventBeginTime' ).AsString;
    aRecordEffect := FDataUpdate.ExecSQL;
    if ( aRecordEffect <= 0 ) then
    begin
      FDataInsert.Parameters.ParamByName( 'FileName' ).Value :=
        FDataBuffer.FieldByName( 'FileName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'CreationDate' ).Value :=
        FDataBuffer.FieldByName( 'CreationDate' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ChannelId' ).Value :=
        FDataBuffer.FieldByName( 'ChannelId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EventBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'EventBeginTime' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Duration' ).Value :=
        FDataBuffer.FieldByName( 'Duration' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EventId' ).Value :=
        FDataBuffer.FieldByName( 'EventId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EventType' ).Value :=
        FDataBuffer.FieldByName( 'EventType' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Name' ).Value :=
        FDataBuffer.FieldByName( 'Name' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EpgDesc' ).Value :=
        FDataBuffer.FieldByName( 'EpgDesc' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngName' ).Value :=
        FDataBuffer.FieldByName( 'EngName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngEpgDesc' ).Value :=
        FDataBuffer.FieldByName( 'EngEpgDesc' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ParentalRating' ).Value :=
        FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ContentNibble1' ).Value :=
        FDataBuffer.FieldByName( 'ContentNibble1' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ContentNibble2' ).Value :=
        FDataBuffer.FieldByName( 'ContentNibble2' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UserNibble1' ).Value :=
        FDataBuffer.FieldByName( 'UserNibble1' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UserNibble2' ).Value :=
        FDataBuffer.FieldByName( 'UserNibble2' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UpdTime' ).Value :=
        FDataBuffer.FieldByName( 'UpdTime' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UpdEn' ).Value :=
        FDataBuffer.FieldByName( 'UpdEn' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UTCEventBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'UTCEventBeginTime' ).AsString;
      FDataInsert.ExecSQL;
    end;
    FDataDelete.Parameters.ParamByName( 'EventId' ).Value :=
      FDataBuffer.FieldByName( 'EventId' ).AsString;
    FDataDelete.ExecSQL;
    FDataBuffer.Next;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TParseSO161 }

procedure TParseSO161.PrepareDataBuffer;
begin
  inherited;
  { SO161 Sub時段資料檔 }
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'ChannelPeriodBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ChannelPeriodEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'FileName', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'CreationDate', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'EventBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'Duration', ftInteger );
    FDataBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FDataBuffer.FieldDefs.Add( 'EventType', ftString, 1 );
    FDataBuffer.FieldDefs.Add( 'Name', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'EpgDesc', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'EngName', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'EngEpgDesc', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'ParentalRating', ftInteger );
    FDataBuffer.FieldDefs.Add( 'ContentNibble1', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'ContentNibble2', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UserNibble1', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UserNibble2', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UpdTime', ftString, 19 );
    FDataBuffer.FieldDefs.Add( 'UpdEn', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ParentalRatingName', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'UTCEventBeginTime', ftString, 14 );
  end;
  FDataBuffer.IndexName := EmptyStr;
  FDataBuffer.CreateDataSet;
  FDataBuffer.AddIndex( 'IDX1', 'ChannelId;EventId', [] );
  FDataBuffer.IndexName := 'IDX1';
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO161.InternalParse;

type
  TSo161 = record
    ChannelPeriodBeginTime: String;
    ChannelPeriodEndTime: String;
    FileName: String;
    CreationDate: String;
    ChannelId: String;
    EventBeginTime: String;
    Duration: String;
    EventId: String;
    EvnetType: String;
    Name: String;
    EpgDesc: String;
    EngName: String;
    EngEpgDesc: String;
    ParentalRating: String;
    ContentNibble1: String;
    ContentNibble2: String;
    UserNibble1: String;
    UserNibble2: String;
    UpdTime: String;
    UpdEn: String;
    ParentalRatingName: String;
    UTCEventBeginTime: String;
  end;

  { ------------------------------------------------------ }

  function GetEpgDesc(aNode: IDOMNode): String;
  var
    aIndex: Integer;
    aTmpNode: IDOMNode;
    aNodeLis: IDOMNodeList;
  begin
    Result := EmptyStr;
    if not Assigned( aNode ) then Exit;
    aNodeLis := ( aNode as IDOMNodeSelect ).selectNodes( 'ExtendedInfo' );
    for aIndex := 0 to aNodeLis.length - 1 do
    begin
      aTmpNode := aNodeLis.item[aIndex].attributes.getNamedItem( 'name' );
      if Assigned( aTmpNode ) then
        Result := ( Result + aTmpNode.nodeValue + ':' );
      Result := Result + aNodeLis.item[aIndex].firstChild.nodeValue;
      if ( Result <> EmptyStr ) then Result := ( Result + #13#10 ); 
    end;
    aTmpNode := ( aNode as IDOMNodeSelect ).selectNode( 'Description' );
    if Assigned( aTmpNode ) then
      Result := Result + aTmpNode.firstChild.nodeValue;
  end;

  { ------------------------------------------------------ }

var
  aIDOMList, aIDOMList2, aIDOMList3: IDOMNodeList;
  aEventNode, aLanguageNode, aNode: IDOMNode;
  aIndex, aIndex2, aIndex3: Integer;
  a161: TSo161;
  aCreationDate, aFileName: String;
begin
  aFileName := FileNameWithoutExt( ExtractFileName( FXml.FileName ) );
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'BroadcastData' );
  aCreationDate := aIDOMList.item[0].attributes.getNamedItem( 'creationDate' ).nodeValue;
  aIDOMList := nil;
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'ChannelPeriod' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    aIDOMList2 := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNodes( 'Event' );
    for aIndex2 := 0 to aIDOMList2.length - 1 do
    begin
      { 判斷是否是 PPV Product 的 Event ( 或是可能成為PPV Event ), EventType = 'P' }
      aNode := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode( 'EventType' );
      if ( aNode.firstChild.nodeValue = 'P' ) then Continue;
      {}
      ZeroMemory( @a161, SizeOf( a161 ) );
      {}
      a161.ChannelPeriodBeginTime := aIDOMList.item[aIndex].attributes.getNamedItem( 'beginTime' ).nodeValue;
      a161.ChannelPeriodEndTime := aIDOMList.item[aIndex].attributes.getNamedItem( 'endTime' ).nodeValue;
      {}
      a161.FileName := aFileName;
      a161.CreationDate := aCreationDate;
      {}
      a161.ChannelId := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNode(
        'ChannelId' ).firstChild.nodeValue;
      {}
      aEventNode := ( aIDOMList2.item[aIndex2] );
      a161.UTCEventBeginTime := aEventNode.attributes.getNamedItem( 'beginTime' ).nodeValue;
      a161.EventBeginTime := UTCDateTimeToLocalTime(
         FormatMaskText( '####/##/## ##:##:##;0;_', a161.UTCEventBeginTime ) );
      a161.EventBeginTime := TrimChar( a161.EventBeginTime, ['/', ':', ' '] );
      {}
      a161.Duration := aEventNode.attributes.getNamedItem( 'duration' ).nodeValue;
      {}
      a161.EventId := ( aEventNode as IDOMNodeSelect ).selectNode(
        'EventId' ).firstChild.nodeValue;
      {}
      a161.EvnetType := ( aEventNode as IDOMNodeSelect ).selectNode(
        'EventType' ).firstChild.nodeValue;
      {}
      aIDOMList3 := ( ( aEventNode as IDOMNodeSelect ).selectNode(
        'EpgProduction' ) as IDOMNodeSelect ).selectNodes( 'EpgText' );
      for aIndex3 := 0 to aIDOMList3.length - 1 do
      begin
        aLanguageNode := aIDOMList3.item[aIndex3];
        if ( aLanguageNode.attributes.getNamedItem( 'language' ).nodeValue = 'chi' ) then
        begin
          a161.Name := ( aLanguageNode as IDOMNodeSelect ).selectNode(
            'Name' ).firstChild.nodeValue;
          a161.EpgDesc := GetEpgDesc( aLanguageNode );
        end else
        begin
          a161.EngName := ( aLanguageNode as IDOMNodeSelect ).selectNode(
            'Name' ).firstChild.nodeValue;
          a161.EngEpgDesc := GetEpgDesc( aLanguageNode );
        end;
      end;
      aIDOMList3 := nil;
      {}
      aNode := ( ( aEventNode as IDOMNodeSelect ).selectNode(
        'EpgProduction' ) as IDOMNodeSelect ).selectNode( 'ParentalRating' );
      if Assigned( aNode ) then
        a161.ParentalRating := aNode.firstChild.nodeValue;
      {}
      aNode := ( ( aEventNode as IDOMNodeSelect ).selectNode(
        'EpgProduction' ) as IDOMNodeSelect ).selectNode( 'DvbContent' );
      if Assigned( aNode ) then
      begin
        if Assigned( ( aNode as IDOMNodeSelect ).selectNode( 'Content' ) ) then
        begin
          a161.ContentNibble1 := ( aNode as IDOMNodeSelect ).selectNode(
            'Content' ) .attributes.getNamedItem( 'nibble1' ).nodeValue;
          a161.ContentNibble2 := ( aNode as IDOMNodeSelect ).selectNode(
            'Content' ) .attributes.getNamedItem( 'nibble2' ).nodeValue;
        end;
        {}
        if Assigned( ( aNode as IDOMNodeSelect ).selectNode( 'User' ) ) then
        begin
          a161.UserNibble1 := ( aNode as IDOMNodeSelect ).selectNode(
            'User' ) .attributes.getNamedItem( 'nibble1' ).nodeValue;
          a161.UserNibble2 := ( aNode as IDOMNodeSelect ).selectNode(
            'User' ) .attributes.getNamedItem( 'nibble2' ).nodeValue;
        end;
      end;
      {}
      Sleep( 5 );
      {}
      a161.UpdTime := ROCDate;
      a161.UpdEn := SMS_PROGRAM_SCHEDULAR;
      {}
      FDataBuffer.Append;
      FDataBuffer.FieldByName( 'ChannelPeriodBeginTime' ).Value := a161.ChannelPeriodBeginTime;
      FDataBuffer.FieldByName( 'ChannelPeriodEndTime' ).Value := a161.ChannelPeriodEndTime;
      FDataBuffer.FieldByName( 'FileName' ).Value := Copy( a161.FileName, Length( a161.FileName ) - 13, 14 );
      FDataBuffer.FieldByName( 'CreationDate' ).Value := a161.CreationDate;
      FDataBuffer.FieldByName( 'ChannelId' ).Value := a161.ChannelId;
      FDataBuffer.FieldByName( 'EventBeginTime' ).Value := a161.EventBeginTime;
      FDataBuffer.FieldByName( 'Duration' ).Value := a161.Duration;
      FDataBuffer.FieldByName( 'EventId' ).Value := a161.EventId;
      FDataBuffer.FieldByName( 'EventType' ).Value := a161.EvnetType;
      FDataBuffer.FieldByName( 'Name' ).Value := a161.Name;
      FDataBuffer.FieldByName( 'EpgDesc' ).Value := a161.EpgDesc;
      FDataBuffer.FieldByName( 'EngName' ).Value := a161.EngName;
      FDataBuffer.FieldByName( 'EngEpgDesc' ).Value := a161.EngEpgDesc;
      FDataBuffer.FieldByName( 'ParentalRating' ).Value := a161.ParentalRating;
      FDataBuffer.FieldByName( 'ContentNibble1' ).Value := a161.ContentNibble1;
      FDataBuffer.FieldByName( 'ContentNibble2' ).Value := a161.ContentNibble2;
      FDataBuffer.FieldByName( 'UserNibble1' ).Value := a161.UserNibble1;
      FDataBuffer.FieldByName( 'UserNibble2' ).Value := a161.UserNibble2;
      FDataBuffer.FieldByName( 'UpdTime' ).Value := a161.UpdTime;
      FDataBuffer.FieldByName( 'UpdEn' ).Value := a161.UpdEn;
      FDataBuffer.FieldByName( 'UTCEventBeginTime' ).Value := a161.UTCEventBeginTime;
      FDataBuffer.Post;
    end;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO161.InternalApplyUpdate;
var
  aRecordEffect: Integer;
  aPreChannelId, aPRCodeWrapper, aPRNameWrapper: String;
begin
  inherited;
  if ( FDataBuffer.IsEmpty ) then Exit;
  FDataDelete.SQL.Clear;
  FDataInsert.SQL.Clear;
  FDataUpdate.SQL.Clear;
  FDataDelete.SQL.Text :=
    '  UPDATE SO161 SET eStopFlag = 1                               ' +
    '   WHERE ChannelId = :ChannelId                                ' +
    '     AND ( EventBeginTime >= :St and EventBeginTime < :Ed )    ';
  FDataBuffer.First;
  aPreChannelId := EmptyStr;
  while not  FDataBuffer.Eof do
  begin
    if ( aPreChannelId <> FDataBuffer.FieldByName( 'ChannelId' ).AsString ) then
    begin
      FDataDelete.Parameters.ParamByName( 'St' ).Value :=
        FDataBuffer.FieldByName( 'ChannelPeriodBeginTime' ).AsString;
      FDataDelete.Parameters.ParamByName( 'Ed' ).Value :=
        FDataBuffer.FieldByName( 'ChannelPeriodEndTime' ).AsString;
      FDataDelete.Parameters.ParamByName( 'ChannelId' ).Value :=
        FDataBuffer.FieldByName( 'ChannelId' ).AsString;
      FDataDelete.ExecSQL;
      aPreChannelId := FDataBuffer.FieldByName( 'ChannelId' ).AsString;
    end;
    FDataBuffer.Next;
  end;
  FDataInsert.SQL.Text :=
    ' INSERT INTO SO161 ( FileName, CreationDate, ChannelId, EventBeginTime, ' +
    '   Duration, EventId, EventType, Name, EpgDesc, EngName, EngEpgDesc,    ' +
    '   ParentalRating, ContentNibble1, ContentNibble2, UserNibble1,         ' +
    '   UserNibble2, UpdTime, UpdEn, ParentalRatingName, eStopFlag,          ' +
    '   UTCEventBeginTime )                                                  ' +
    ' VALUES ( :FileName, :CreationDate, :ChannelId, :EventBeginTime,        ' +
    '   :Duration, :EventId, :EventType, :Name, :EpgDesc, :EngName,          ' +
    '   :EngEpgDesc, :ParentalRating, :ContentNibble1, :ContentNibble2,      ' +
    '   :UserNibble1, :UserNibble2, :UpdTime, :UpdEn, :ParentalRatingName,   ' +
    '    0, :UTCEventBeginTime )                                             ';
  FDataUpdate.SQL.Text :=
    ' UPDATE SO161 SET FileName = :FileName, CreationDate = :CreationDate,   ' +
    '   EventBeginTime = :EventBeginTime, Duration = :Duration,              ' +
    '   EventType = :EventType, Name = :Name, EpgDesc = :EpgDesc,            ' +
    '   EngName = :EngName, EngEpgDesc = :EngEpgDesc,                        ' +
    '   ParentalRating = :ParentalRating, ContentNibble1 = :ContentNibble1,  ' +
    '   ContentNibble2 = :ContentNibble2, UserNibble1 = :UserNibble1,        ' +
    '   UserNibble2 = :UserNibble2, UpdTime = :UpdTime, UpdEn = :UpdEn,      ' +
    '   eStopFlag = 0, ChannelId = :ChannelId,                               ' +
    '   ParentalRatingName = :ParentalRatingName,                            ' +
    '   UTCEventBeginTime = :UTCEventBeginTime                               ' +
    ' WHERE EventId = :EventId                                               ';
  {}
  FDataDelete.SQL.Text :=
    ' DELETE FROM SO159 WHERE EventId = :EventId  ';
  {}
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    aPRCodeWrapper := FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
    GetParentalratingName( FDataReader, aPRCodeWrapper, aPRNameWrapper );
    FDataUpdate.Parameters.ParamByName( 'FileName' ).Value :=
      FDataBuffer.FieldByName( 'FileName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'CreationDate' ).Value :=
      FDataBuffer.FieldByName( 'CreationDate' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EventBeginTime' ).Value :=
      FDataBuffer.FieldByName( 'EventBeginTime' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Duration' ).Value :=
      FDataBuffer.FieldByName( 'Duration' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EventType' ).Value :=
      FDataBuffer.FieldByName( 'EventType' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Name' ).Value :=
      FDataBuffer.FieldByName( 'Name' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EpgDesc' ).Value :=
      FDataBuffer.FieldByName( 'EpgDesc' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngName' ).Value :=
      FDataBuffer.FieldByName( 'EngName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngEpgDesc' ).Value :=
      FDataBuffer.FieldByName( 'EngEpgDesc' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ContentNibble1' ).Value :=
      FDataBuffer.FieldByName( 'ContentNibble1' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ContentNibble2' ).Value :=
      FDataBuffer.FieldByName( 'ContentNibble2' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UserNibble1' ).Value :=
      FDataBuffer.FieldByName( 'UserNibble1' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UserNibble2' ).Value :=
      FDataBuffer.FieldByName( 'UserNibble2' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UpdTime' ).Value :=
      FDataBuffer.FieldByName( 'UpdTime' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UpdEn' ).Value :=
      FDataBuffer.FieldByName( 'UpdEn' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ChannelId' ).Value :=
      FDataBuffer.FieldByName( 'ChannelId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EventId' ).Value :=
      FDataBuffer.FieldByName( 'EventId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ParentalRating' ).Value := aPRCodeWrapper;
    FDataUpdate.Parameters.ParamByName( 'ParentalRatingName' ).Value := aPRNameWrapper;
    FDataUpdate.Parameters.ParamByName( 'UTCEventBeginTime' ).Value :=
      FDataBuffer.FieldByName( 'UTCEventBeginTime' ).AsString;
    aRecordEffect := FDataUpdate.ExecSQL;
    if ( aRecordEffect <= 0 ) then
    begin
      FDataInsert.Parameters.ParamByName( 'FileName' ).Value :=
        FDataBuffer.FieldByName( 'FileName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'CreationDate' ).Value :=
        FDataBuffer.FieldByName( 'CreationDate' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ChannelId' ).Value :=
        FDataBuffer.FieldByName( 'ChannelId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EventBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'EventBeginTime' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Duration' ).Value :=
        FDataBuffer.FieldByName( 'Duration' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EventId' ).Value :=
        FDataBuffer.FieldByName( 'EventId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EventType' ).Value :=
        FDataBuffer.FieldByName( 'EventType' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Name' ).Value :=
        FDataBuffer.FieldByName( 'Name' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EpgDesc' ).Value :=
        FDataBuffer.FieldByName( 'EpgDesc' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngName' ).Value :=
        FDataBuffer.FieldByName( 'EngName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngEpgDesc' ).Value :=
        FDataBuffer.FieldByName( 'EngEpgDesc' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ContentNibble1' ).Value :=
        FDataBuffer.FieldByName( 'ContentNibble1' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ContentNibble2' ).Value :=
        FDataBuffer.FieldByName( 'ContentNibble2' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UserNibble1' ).Value :=
        FDataBuffer.FieldByName( 'UserNibble1' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UserNibble2' ).Value :=
        FDataBuffer.FieldByName( 'UserNibble2' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UpdTime' ).Value :=
        FDataBuffer.FieldByName( 'UpdTime' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UpdEn' ).Value :=
        FDataBuffer.FieldByName( 'UpdEn' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ParentalRating' ).Value := aPRCodeWrapper;
      FDataInsert.Parameters.ParamByName( 'ParentalRatingName' ).Value := aPRNameWrapper;
      FDataInsert.Parameters.ParamByName( 'UTCEventBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'UTCEventBeginTime' ).AsString;
      FDataInsert.ExecSQL;
    end;
    FDataDelete.Parameters.ParamByName( 'EventId' ).Value :=
      FDataBuffer.FieldByName( 'EventId' ).AsString;
    FDataDelete.ExecSQL;
    {}
    FDataBuffer.Next;
    if ( FDataBuffer.RecNo mod 10 ) = 0 then Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TParseSO162 }

procedure TParseSO162.PrepareDataBuffer;
begin
  inherited;
  { SO162 Channel資訊 }
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ChannelNo', ftInteger );
    FDataBuffer.FieldDefs.Add( 'Language', ftString, 10 );
    FDataBuffer.FieldDefs.Add( 'Name', ftString, 80 );
    FDataBuffer.FieldDefs.Add( 'ShortName', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ProviderName', ftString, 50 );
    FDataBuffer.FieldDefs.Add( 'Description', ftString, 150 );
    FDataBuffer.FieldDefs.Add( 'EngName', ftString, 80 );
    FDataBuffer.FieldDefs.Add( 'EngShortName', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'EngDescription', ftString, 150 );
    FDataBuffer.FieldDefs.Add( 'ContentNibble1', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'ContentNibble2', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UserNibble1', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UserNibble2', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'ServiceId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'TransportId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'OrignalNetworkId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'Kind', ftInteger );
  end;
  FDataBuffer.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO162.InternalParse;

type
  TSo162 = record
    ChannelId: String;
    ChannelNo: String;
    Language: String;
    Name: String;
    ShortName: String;
    ProviderName: String;
    Description: String;
    EngName: String;
    EngShortName: String;
    EngDescription: String;
    ContentNibble1: String;
    ContentNibble2: String;
    UserNibble1: String;
    UserNibble2: String;
    ServiceId: String;
    TransportId: String;
    OrignalNetworkId: String;
    Kind: String;
  end;
  
var
  aIDOMList, aIDOMList2: IDOMNodeList;
  aNode: IDOMNode;
  aIndex, aIndex2: Integer;
  a162: TSo162;
begin
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'Channel' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    ZeroMemory( @a162, SizeOf( a162 ) );
    a162.ChannelId := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNode(
      'ChannelId' ).firstChild.nodeValue;
    {}
    a162.ChannelNo := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNode(
      'ChannelNumber' ).firstChild.nodeValue;
    {}
    aIDOMList2 := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNodes(
      'ChannelText' );
    for aIndex2 := 0 to aIDOMList2.length - 1 do
    begin
      if ( aIDOMList2.item[aIndex2].attributes.getNamedItem(
        'language' ).nodeValue = 'chi' ) then
      begin
        a162.Language := 'chi';
        a162.Name := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
          'ChannelName' ).firstChild.nodeValue;
        a162.ShortName := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
          'ChannelShortName' ).firstChild.nodeValue;
        aNode := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
         'ChannelProviderName' );
        if Assigned( aNode ) then
          a162.ProviderName := aNode.firstChild.nodeValue;
        aNode := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
         'ChannelDescription' );
        if Assigned( aNode ) then
          a162.Description := aNode.firstChild.nodeValue;
      end else
      begin
        if ( a162.Language = EmptyStr ) then
        begin
          a162.Language := aIDOMList2.item[aIndex2].attributes.getNamedItem(
            'language' ).nodeValue;
        end;
        a162.EngName := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
          'ChannelName' ).firstChild.nodeValue;
        a162.EngShortName := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
          'ChannelShortName' ).firstChild.nodeValue;
        if ( a162.ProviderName = EmptyStr ) then
        begin
          aNode := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
           'ChannelProviderName' );
          if Assigned( aNode ) then
            a162.ProviderName := aNode.firstChild.nodeValue;
        end;    
        aNode := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
         'ChannelDescription' );
        if Assigned( aNode ) then
          a162.EngDescription := aNode.firstChild.nodeValue;
      end;
      if aIndex2 >= 1 then Break;
    end;
    aNode := ( aIDOMList.item[aIndex] as IDOMNodeSelect ).selectNode(
      'PhysicalServiceId' );
    if Assigned( aNode ) then
    begin
      a162.ServiceId := ( aNode as IDOMNodeSelect ).selectNode(
        'DvbServiceId' ).firstChild.nodeValue;
      a162.TransportId := ( aNode as IDOMNodeSelect ).selectNode(
        'TransportId' ).firstChild.nodeValue;
      a162.OrignalNetworkId := ( aNode as IDOMNodeSelect ).selectNode(
        'TransportId' ).attributes.getNamedItem( 'originalNetworkId' ).nodeValue;
    end;
    {}
    FDataBuffer.Append;
    FDataBuffer.FieldByName( 'ChannelId' ).AsString :=
      a162.ChannelId;
    FDataBuffer.FieldByName( 'ChannelNo' ).AsString :=
      a162.ChannelNo;
    FDataBuffer.FieldByName( 'Language' ).AsString :=
      a162.Language;
    FDataBuffer.FieldByName( 'Name' ).AsString :=
      a162.Name;
    FDataBuffer.FieldByName( 'ShortName' ).AsString :=
      a162.ShortName;
    FDataBuffer.FieldByName( 'ProviderName' ).AsString :=
      a162.ProviderName;
    FDataBuffer.FieldByName( 'Description' ).AsString :=
      a162.Description;
    FDataBuffer.FieldByName( 'EngName' ).AsString :=
      a162.EngName;
    FDataBuffer.FieldByName( 'EngShortName' ).AsString :=
      a162.EngShortName;
    FDataBuffer.FieldByName( 'EngDescription' ).AsString :=
      a162.EngDescription;
    FDataBuffer.FieldByName( 'ContentNibble1' ).AsString :=
      a162.ContentNibble1;
    FDataBuffer.FieldByName( 'ContentNibble2' ).AsString :=
      a162.ContentNibble2;
    FDataBuffer.FieldByName( 'UserNibble1' ).AsString :=
      a162.UserNibble1;
    FDataBuffer.FieldByName( 'UserNibble2' ).AsString :=
      a162.UserNibble2;
    FDataBuffer.FieldByName( 'ServiceId' ).AsString :=
      a162.ServiceId;
    FDataBuffer.FieldByName( 'TransportId' ).AsString :=
      a162.TransportId;
    FDataBuffer.FieldByName( 'OrignalNetworkId' ).AsString :=
      a162.OrignalNetworkId;
    FDataBuffer.FieldByName( 'Kind' ).AsString :=
      a162.Kind;
    FDataBuffer.Post;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO162.InternalApplyUpdate;
var
  aRecordEffect: Integer;
begin
  inherited;
  if ( FDataBuffer.IsEmpty ) then Exit;
  FDataInsert.SQL.Clear;
  FDataUpdate.SQL.Clear;
  FDataInsert.SQL.Text :=
    ' INSERT INTO SO162 ( ChannelId, ChannelNo, Language, Name, ShortName,   ' +
    '   ProviderName, Description, EngName, EngShortName, EngDescription,    ' +
    '   ContentNibble1, ContentNibble2, UserNibble1, UserNibble2, ServiceId, ' +
    '   TransportId, OrignalNetworkId, Kind )                                ' +
    ' VALUES ( :ChannelId, :ChannelNo, :Language_, :Name, :ShortName,        ' +
    '   :ProviderName, :Description, :EngName, :EngShortName, :EngDescription,     ' +
    '   :ContentNibble1, :ContentNibble2, :UserNibble1, :UserNibble2, :ServiceId,  ' +
    '   :TransportId, :OrignalNetworkId, :Kind )                                   ';
  FDataUpdate.SQL.Text :=
    ' UPDATE SO162 SET ChannelNo = :ChannelNo, Language = :Language,       ' +
    '  Name = :Name, ShortName = :ShortName, ProviderName = :ProviderName, ' +
    '  Description = :Description, EngName = :EngName,                     ' +
    '  EngShortName = :EngShortName, EngDescription = :EngDescription,     ' +
    '  ContentNibble1 = :ContentNibble1, ContentNibble2 = :ContentNibble2, ' +
    '  UserNibble1 = :UserNibble1, UserNibble2 = :UserNibble2,             ' +
    '  ServiceId = :ServiceId, TransportId = :TransportId,                 ' +
    '  OrignalNetworkId = :OrignalNetworkId, Kind = :Kind                  ' +
    ' WHERE ChannelId = :ChannelId                                         ';
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    FDataUpdate.Parameters.ParamByName( 'ChannelNo' ).Value :=
      FDataBuffer.FieldByName( 'ChannelNo' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Language' ).Value :=
      FDataBuffer.FieldByName( 'Language' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Name' ).Value :=
      FDataBuffer.FieldByName( 'Name' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ShortName' ).Value :=
      FDataBuffer.FieldByName( 'ShortName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ProviderName' ).Value :=
      FDataBuffer.FieldByName( 'ProviderName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Description' ).Value :=
      FDataBuffer.FieldByName( 'Description' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngName' ).Value :=
      FDataBuffer.FieldByName( 'EngName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngShortName' ).Value :=
      FDataBuffer.FieldByName( 'EngShortName' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'EngDescription' ).Value :=
      FDataBuffer.FieldByName( 'EngDescription' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ContentNibble1' ).Value :=
      FDataBuffer.FieldByName( 'ContentNibble1' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ContentNibble2' ).Value :=
      FDataBuffer.FieldByName( 'ContentNibble2' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UserNibble1' ).Value :=
      FDataBuffer.FieldByName( 'UserNibble1' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'UserNibble2' ).Value :=
      FDataBuffer.FieldByName( 'UserNibble2' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ServiceId' ).Value :=
      FDataBuffer.FieldByName( 'ServiceId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'TransportId' ).Value :=
      FDataBuffer.FieldByName( 'TransportId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'OrignalNetworkId' ).Value :=
      FDataBuffer.FieldByName( 'OrignalNetworkId' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'Kind' ).Value :=
      FDataBuffer.FieldByName( 'Kind' ).AsString;
    FDataUpdate.Parameters.ParamByName( 'ChannelId' ).Value :=
      FDataBuffer.FieldByName( 'ChannelId' ).AsString;
    aRecordEffect := FDataUpdate.ExecSQL;
    if ( aRecordEffect <= 0 ) then
    begin
      FDataInsert.Parameters.ParamByName( 'ChannelId' ).Value :=
        FDataBuffer.FieldByName( 'ChannelId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ChannelNo' ).Value :=
        FDataBuffer.FieldByName( 'ChannelNo' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Language_' ).Value :=
        FDataBuffer.FieldByName( 'Language' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Name' ).Value :=
        FDataBuffer.FieldByName( 'Name' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ShortName' ).Value :=
        FDataBuffer.FieldByName( 'ShortName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ProviderName' ).Value :=
        FDataBuffer.FieldByName( 'ProviderName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Description' ).Value :=
        FDataBuffer.FieldByName( 'Description' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngName' ).Value :=
        FDataBuffer.FieldByName( 'EngName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngShortName' ).Value :=
        FDataBuffer.FieldByName( 'EngShortName' ).AsString;
      FDataInsert.Parameters.ParamByName( 'EngDescription' ).Value :=
        FDataBuffer.FieldByName( 'EngDescription' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ContentNibble1' ).Value :=
        FDataBuffer.FieldByName( 'ContentNibble1' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ContentNibble2' ).Value :=
        FDataBuffer.FieldByName( 'ContentNibble2' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UserNibble1' ).Value :=
        FDataBuffer.FieldByName( 'UserNibble1' ).AsString;
      FDataInsert.Parameters.ParamByName( 'UserNibble2' ).Value :=
        FDataBuffer.FieldByName( 'UserNibble2' ).AsString;
      FDataInsert.Parameters.ParamByName( 'ServiceId' ).Value :=
        FDataBuffer.FieldByName( 'ServiceId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'TransportId' ).Value :=
        FDataBuffer.FieldByName( 'TransportId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'OrignalNetworkId' ).Value :=
        FDataBuffer.FieldByName( 'OrignalNetworkId' ).AsString;
      FDataInsert.Parameters.ParamByName( 'Kind' ).Value :=
        FDataBuffer.FieldByName( 'Kind' ).AsString;
      FDataInsert.ExecSQL;
    end;
    FDataBuffer.Next;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TParseSO158Ppv }

constructor TParseSO158Ppv.Create(aXml: TXMLDocument);
begin
  inherited Create( aXml );
  FDeleteProduct := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

destructor TParseSO158Ppv.Destroy;
begin
  FDeleteProduct.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Ppv.PrepareDataBuffer;
begin
  inherited;
  { SO158 Dex XML資料檔 }
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'FileName', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'CreationDate', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ProductId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ImpulsiveFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'ProductKind', ftString, 3 );
    FDataBuffer.FieldDefs.Add( 'SpecialPpv', ftInteger );
    FDataBuffer.FieldDefs.Add( 'PurchasableFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'SoldFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'SuspendedFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'ProductName', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'ProductDescription', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'EngProductName', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'EngProductDescription', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'SaleBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'SaleEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ValidityPeriodBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ValidityPeriodEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ProductCategory', ftInteger );
    FDataBuffer.FieldDefs.Add( 'EpgPrice', ftFloat );
    FDataBuffer.FieldDefs.Add( 'SubAllowPpvEventFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FDataBuffer.FieldDefs.Add( 'EventPreviewTime', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UpdTime', ftString, 19 );
    FDataBuffer.FieldDefs.Add( 'UpdEn', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ParentalRating', ftInteger );
    FDataBuffer.FieldDefs.Add( 'dStopFlag', ftString, 1 );
    FDataBuffer.FieldDefs.Add( 'eStopFlag', ftString, 1 );
    FDataBuffer.FieldDefs.Add( 'UTCSaleBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'UTCSaleEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'UTCValidityPeriodBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'UTCValidityPeriodEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'Action', ftInteger );
  end;
  FDataBuffer.IndexName := EmptyStr;
  FDataBuffer.CreateDataSet;
  FDataBuffer.AddIndex( 'IDX1', 'ProductId;EventId', [] );
  FDataBuffer.IndexName := 'IDX1';
  {}
  if FDeleteProduct.FieldDefs.Count <= 0 then
    FDeleteProduct.FieldDefs.Add( 'ProductId', ftString, 20 );
  FDeleteProduct.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Ppv.UnPrepareDataBuffer;
begin
  inherited;
  if not VarIsNull( FDeleteProduct.Data ) then
    FDeleteProduct.EmptyDataSet;
  FDeleteProduct.Data := Null;  
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Ppv.InternalParse;

type
  TSo158 = record
    FileName: String;
    CreationDate: String;
    ProductId: String;
    ImpulsiveFlag: String;
    ProductKind: String;
    SpecialPpv: String;
    PurchasableFlag: String;
    SoldFlag: String;
    SuspendedFlag: String;
    ProductName: String;
    ProductDescription: String;
    EngProductName: String;
    EngProductDescription: String;
    SaleBeginTime: String;
    SaleEndTime: String;
    ValidityPeriodBeginTime: String;
    ValidityPeriodEndTime: String;
    ProductCategory: String;
    EpgPrice: String;
    SubAllowPpvEventFlag: String;
    EventId: String;
    EventPreviewTime: String;
    UpdTime: String;
    UpdEn: String;
    ParentalRating: String;
    UTCSaleBeginTime: String;
    UTCSaleEndTime: String;
    UTCValidityPeriodBeginTime: String;
    UTCValidityPeriodEndTime: String;
  end;

var
  aIDOMList, aIDOMList2, aEventList: IDOMNodeList;
  aProductNode, aLanguageNode, aNode: IDOMNode;
  aIndex, aIndex2: Integer;
  a158: TSo158;
  aCreationDate, aFileName: String;
begin
  aFileName := FileNameWithoutExt( ExtractFileName( FXml.FileName ) );
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'ProductManagementData' );
  aCreationDate := aIDOMList.item[0].attributes.getNamedItem( 'creationDate' ).nodeValue;
  aIDOMList := nil;
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'Product' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    aProductNode := aIDOMList.item[aIndex];
    { 只處理 PPV }
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductType' );
    if WideUpperCase( aNode.attributes.getNamedItem(
      'productKind' ).nodeValue ) <> 'PPV' then Continue;
    {}
    ZeroMemory( @a158, SizeOf( a158 ) );
    a158.FileName := aFileName;
    a158.CreationDate := aCreationDate;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductId' );
    a158.ProductId := aNode.firstChild.nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductType' );
    a158.ImpulsiveFlag := aNode.attributes.getNamedItem( 'impulsiveFlag' ).nodeValue;
    a158.ProductKind := aNode.attributes.getNamedItem( 'productKind' ).nodeValue;
    a158.SpecialPpv := aNode.attributes.getNamedItem( 'specialPpv' ).nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductStatus' );
    a158.SuspendedFlag := aNode.attributes.getNamedItem( 'suspendedFlag' ).nodeValue;
    a158.SoldFlag := aNode.attributes.getNamedItem( 'soldFlag' ).nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductCategory' );
    if Assigned( aNode ) then
      a158.ProductCategory := aNode.firstChild.nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'EpgPrice' );
    if Assigned( aNode ) then a158.EpgPrice := aNode.firstChild.nodeValue;
    {}
    aIDOMList2 := ( aProductNode as IDOMNodeSelect ).selectNodes( 'ProductText' );
    for aIndex2 := 0 to aIDOMList2.length - 1 do
    begin
      aLanguageNode := aIDOMList2.item[aIndex2];
      if ( aLanguageNode.attributes.getNamedItem( 'language' ).nodeValue = 'chi' ) then
      begin
        a158.ProductName := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductName' ).firstChild.nodeValue;
        a158.ProductDescription := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductDescription' ).firstChild.nodeValue;
      end else
      begin
        a158.EngProductName := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductName' ).firstChild.nodeValue;
        a158.EngProductDescription := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductDescription' ).firstChild.nodeValue;
      end;
    end;
    {}
    if ( a158.ProductName = EmptyStr ) then a158.ProductName := a158.EngProductName;
    if ( a158.ProductDescription = EmptyStr ) then a158.ProductDescription := a158.EngProductDescription;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'SalePeriod' );
    a158.UTCSaleBeginTime := aNode.attributes.getNamedItem( 'beginTime' ).nodeValue;
    a158.UTCSaleEndTime := aNode.attributes.getNamedItem( 'endTime' ).nodeValue;
    {}
    a158.SaleBeginTime := UTCDateTimeToLocalTime(
      FormatMaskText( '####/##/## ##:##:##;0;_', a158.UTCSaleBeginTime ) );
    a158.SaleBeginTime := TrimChar( a158.SaleBeginTime, ['/', ':', ' '] );
    a158.SaleEndTime := UTCDateTimeToLocalTime(
      FormatMaskText( '####/##/## ##:##:##;0;_', a158.UTCSaleEndTime ) );
    a158.SaleEndTime := TrimChar( a158.SaleEndTime, ['/', ':', ' '] );
    {}
    aNode := ( ( aProductNode as IDOMNodeSelect ).selectNode(
      'PpvData' ) as IDOMNodeSelect ).selectNode( 'ValidityPeriod' );
    a158.UTCValidityPeriodBeginTime := aNode.attributes.getNamedItem( 'beginTime' ).nodeValue;
    a158.UTCValidityPeriodEndTime := aNode.attributes.getNamedItem( 'endTime' ).nodeValue;
    {}
    a158.ValidityPeriodBeginTime := UTCDateTimeToLocalTime(
      FormatMaskText( '####/##/## ##:##:##;0;_', a158.UTCValidityPeriodBeginTime ) );
    a158.ValidityPeriodBeginTime := TrimChar( a158.ValidityPeriodBeginTime, ['/', ':', ' '] );
    a158.ValidityPeriodEndTime := UTCDateTimeToLocalTime(
      FormatMaskText( '####/##/## ##:##:##;0;_', a158.UTCValidityPeriodEndTime ) );
    a158.ValidityPeriodEndTime := TrimChar( a158.ValidityPeriodEndTime, ['/', ':', ' '] );
    {}
    aEventList := ( ( aProductNode as IDOMNodeSelect ).selectNode(
      'PpvData' ) as IDOMNodeSelect ).selectNodes( 'Event' );
    for aIndex2 := 0 to aEventList.length - 1 do
    begin
      aNode := aEventList.item[aIndex2];
      a158.EventId := ( aNode as IDOMNodeSelect ).selectNode( 'EventId' ).firstChild.nodeValue;
      a158.EventPreviewTime := ( aNode as IDOMNodeSelect ).selectNode( 'PreviewTime' ).firstChild.nodeValue;
      a158.ParentalRating := EmptyStr;
      if Assigned( ( aNode as IDOMNodeSelect ).selectNode( 'ParentalRating' ) ) then
        a158.ParentalRating := ( aNode as IDOMNodeSelect ).selectNode( 'ParentalRating' ).firstChild.nodeValue;
      {}
      a158.UpdTime := ROCDate;
      a158.UpdEn := SMS_PROGRAM_SCHEDULAR;
      {}
      FDataBuffer.Append;
      FDataBuffer.FieldByName( 'FileName' ).AsString := Copy( a158.FileName, Length( a158.FileName ) - 13, 14 );
      FDataBuffer.FieldByName( 'CreationDate' ).AsString := a158.CreationDate;
      FDataBuffer.FieldByName( 'ProductId' ).AsString := a158.ProductId;
      FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString := a158.ImpulsiveFlag;
      FDataBuffer.FieldByName( 'ProductKind' ).AsString := a158.ProductKind;
      FDataBuffer.FieldByName( 'SpecialPpv' ).AsString := a158.SpecialPpv;
      FDataBuffer.FieldByName( 'PurchasableFlag' ).AsString := a158.PurchasableFlag;
      FDataBuffer.FieldByName( 'SoldFlag' ).AsString := a158.SoldFlag;
      FDataBuffer.FieldByName( 'SuspendedFlag' ).AsString := a158.SuspendedFlag;
      FDataBuffer.FieldByName( 'ProductName' ).AsString := a158.ProductName;
      FDataBuffer.FieldByName( 'ProductDescription' ).AsString := a158.ProductDescription;
      FDataBuffer.FieldByName( 'EngProductName' ).AsString := a158.EngProductName;
      FDataBuffer.FieldByName( 'EngProductDescription' ).AsString := a158.EngProductDescription;
      FDataBuffer.FieldByName( 'SaleBeginTime' ).AsString := a158.SaleBeginTime;
      FDataBuffer.FieldByName( 'SaleEndTime' ).AsString := a158.SaleEndTime;
      FDataBuffer.FieldByName( 'ValidityPeriodBeginTime' ).AsString := a158.ValidityPeriodBeginTime;
      FDataBuffer.FieldByName( 'ValidityPeriodEndTime' ).AsString := a158.ValidityPeriodEndTime;
      FDataBuffer.FieldByName( 'ProductCategory' ).AsString := a158.ProductCategory;
      FDataBuffer.FieldByName( 'EpgPrice' ).AsString := a158.EpgPrice;
      FDataBuffer.FieldByName( 'SubAllowPpvEventFlag' ).AsString := '0';
      FDataBuffer.FieldByName( 'EventId' ).AsString := a158.EventId;
      FDataBuffer.FieldByName( 'EventPreviewTime' ).AsString := a158.EventPreviewTime;
      FDataBuffer.FieldByName( 'UpdTime' ).AsString := a158.UpdTime;
      FDataBuffer.FieldByName( 'UpdEn' ).AsString := a158.UpdEn;
      FDataBuffer.FieldByName( 'ParentalRating' ).AsString := a158.ParentalRating;
      FDataBuffer.FieldByName( 'UTCSaleBeginTime' ).AsString := a158.UTCSaleBeginTime;
      FDataBuffer.FieldByName( 'UTCSaleEndTime' ).AsString := a158.UTCSaleEndTime;
      FDataBuffer.FieldByName( 'UTCValidityPeriodBeginTime' ).AsString := a158.UTCValidityPeriodBeginTime;
      FDataBuffer.FieldByName( 'UTCValidityPeriodEndTime' ).AsString := a158.UTCValidityPeriodEndTime;
      FDataBuffer.FieldByName( 'Action' ).AsString := '2';
      FDataBuffer.Post;
      Sleep( 10 );
    end;  
  end;
  { 標示為 DeleteProduct 的資料 }
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'DeletedProductData' );
  if ( aIDOMList.length > 0 ) then
  begin
    aIDOMList := ( aIDOMList.item[0] as IDOMNodeSelect ).selectNodes( 'ProductId' );
    for aIndex := 0 to aIDOMList.length - 1 do
    begin
      FDeleteProduct.Append;
      FDeleteProduct.FieldByName( 'ProductId' ).AsString :=
        aIDOMList.item[aIndex].firstChild.nodeValue;
      FDeleteProduct.Post;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Ppv.InternalApplyUpdate;
var
  aRecordEffect: Integer;
  aPreProductId: String;
begin
  inherited;
  FDataDelete.SQL.Clear;
  FDataInsert.SQL.Clear;
  FDataUpdate.SQL.Clear;
  if ( not FDataBuffer.IsEmpty ) then
  begin
    FDataInsert.SQL.Text :=
      ' INSERT INTO SO158 ( FileName, CreationDate, ProductId, ImpulsiveFlag,  ' +
      '    ProductKind, SpecialPpv, PurchasableFlag, SoldFlag, SuspendedFlag,  ' +
      '    ProductName, ProductDescription, EngProductName,                    ' +
      '    EngProductDescription, SaleBeginTime, SaleEndTime,                  ' +
      '    ValidityPeriodBeginTime, ValidityPeriodEndTime, ProductCategory,    ' +
      '    EpgPrice, SubAllowPpvEventFlag, EventId,                            ' +
      '    EventPreviewTime, UpdTime, UpdEn, ParentalRating,                   ' +
      '    dStopFlag, eStopFlag, UTCSaleBeginTime, UTCSaleEndTime,             ' +
      '    UTCValidityPeriodBeginTime, UTCValidityPeriodEndTime, Action )      ' +
      ' VALUES ( :FileName, :CreationDate, :ProductId, :ImpulsiveFlag,         ' +
      '    :ProductKind, :SpecialPpv, :PurchasableFlag, :SoldFlag,             ' +
      '    :SuspendedFlag, :ProductName, :ProductDescription, :EngProductName, ' +
      '    :EngProductDescription, :SaleBeginTime, :SaleEndTime,               ' +
      '    :ValidityPeriodBeginTime, :ValidityPeriodEndTime, :ProductCategory, ' +
      '    :EpgPrice, :SubAllowPpvEventFlag, :EventId,                         ' +
      '    :EventPreviewTime, :UpdTime, :UpdEn, :ParentalRating,               ' +
      '     0, 0, :UTCSaleBeginTime, :UTCSaleEndTime,                          ' +
      '     :UTCValidityPeriodBeginTime, :UTCValidityPeriodEndTime, 1 )        ';
    { 先將該 Product 下所有 Event 先 Update 成 Delete }
    FDataDelete.SQL.Text :=
      ' UPDATE SO158 SET eStopFlag = 1, Action = 2                             ' +
      '  WHERE ProductId = :ProductId                                          ';
    {}
    FDataUpdate.SQL.Text :=
      ' UPDATE SO158 SET                                                       ' +
      '    FileName = :FileName, CreationDate = :CreationDate,                 ' +
      '    ImpulsiveFlag = :ImpulsiveFlag, ProductKind = :ProductKind,         ' +
      '    SpecialPpv = :SpecialPpv, PurchasableFlag = :PurchasableFlag,       ' +
      '    SoldFlag = :SoldFlag, SuspendedFlag = :SuspendedFlag,               ' +
      '    ProductName = :ProductName,                                         ' +
      '    ProductDescription = :ProductDescription,                           ' +
      '    EngProductName = :EngProductName,                                   ' +
      '    EngProductDescription = :EngProductDescription,                     ' +
      '    SaleBeginTime = :SaleBeginTime, SaleEndTime = :SaleEndTime,         ' +
      '    ValidityPeriodBeginTime = :ValidityPeriodBeginTime,                 ' +
      '    ValidityPeriodEndTime = :ValidityPeriodEndTime,                     ' +
      '    ProductCategory = :ProductCategory, EpgPrice = :EpgPrice,           ' +
      '    SubAllowPpvEventFlag = :SubAllowPpvEventFlag,                       ' +
      '    EventPreviewTime = :EventPreviewTime, UpdTime = :UpdTime,           ' +
      '    UpdEn = :UpdEn, ParentalRating = :ParentalRating,                   ' +
      '    eStopFlag = 0, UTCSaleBeginTime = :UTCSaleBeginTime,                ' +
      '    UTCSaleEndTime = :UTCSaleEndTime,                                   ' +
      '    UTCValidityPeriodBeginTime = :UTCValidityPeriodBeginTime,           ' +
      '    UTCValidityPeriodEndTime = :UTCValidityPeriodEndTime,               ' +
      '    Action = :Action                                                    ' +
      '  WHERE ProductId = :ProductId                                          ' +
      '    AND EventId = :EventId                                              ';
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      if ( aPreProductId <> FDataBuffer.FieldByName( 'ProductId' ).AsString ) then
      begin
        FDataDelete.Parameters.ParamByName( 'ProductId' ).Value :=
          FDataBuffer.FieldByName( 'ProductId' ).AsString;
        FDataDelete.ExecSQL;  
        aPreProductId := FDataBuffer.FieldByName( 'ProductId' ).AsString;
      end;
      FDataUpdate.Parameters.ParamByName( 'FileName' ).Value :=
        FDataBuffer.FieldByName( 'FileName' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'CreationDate' ).Value :=
        FDataBuffer.FieldByName( 'CreationDate' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ImpulsiveFlag' ).Value :=
        FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductKind' ).Value :=
        FDataBuffer.FieldByName( 'ProductKind' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SpecialPpv' ).Value :=
        FDataBuffer.FieldByName( 'SpecialPpv' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'PurchasableFlag' ).Value :=
        FDataBuffer.FieldByName( 'PurchasableFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SoldFlag' ).Value :=
        FDataBuffer.FieldByName( 'SoldFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SuspendedFlag' ).Value :=
        FDataBuffer.FieldByName( 'SuspendedFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductName' ).Value :=
        FDataBuffer.FieldByName( 'ProductName' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductDescription' ).Value :=
        FDataBuffer.FieldByName( 'ProductDescription' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EngProductName' ).Value :=
        FDataBuffer.FieldByName( 'EngProductName' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EngProductDescription' ).Value :=
        FDataBuffer.FieldByName( 'EngProductDescription' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SaleBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'SaleBeginTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SaleEndTime' ).Value :=
        FDataBuffer.FieldByName( 'SaleEndTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ValidityPeriodBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'ValidityPeriodBeginTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ValidityPeriodEndTime' ).Value :=
        FDataBuffer.FieldByName( 'ValidityPeriodEndTime' ).Value;
      FDataUpdate.Parameters.ParamByName( 'ProductCategory' ).Value :=
        FDataBuffer.FieldByName( 'ProductCategory' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EpgPrice' ).Value :=
        FDataBuffer.FieldByName( 'EpgPrice' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SubAllowPpvEventFlag' ).Value :=
        FDataBuffer.FieldByName( 'SubAllowPpvEventFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EventPreviewTime' ).Value :=
        FDataBuffer.FieldByName( 'EventPreviewTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UpdTime' ).Value :=
        FDataBuffer.FieldByName( 'UpdTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UpdEn' ).Value :=
        FDataBuffer.FieldByName( 'UpdEn' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ParentalRating' ).Value :=
        FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductId' ).Value :=
        FDataBuffer.FieldByName( 'ProductId' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EventId' ).Value :=
        FDataBuffer.FieldByName( 'EventId' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UTCSaleBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'UTCSaleBeginTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UTCSaleEndTime' ).Value :=
        FDataBuffer.FieldByName( 'UTCSaleEndTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UTCValidityPeriodBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'UTCValidityPeriodBeginTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UTCValidityPeriodEndTime' ).Value :=
        FDataBuffer.FieldByName( 'UTCValidityPeriodEndTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'Action' ).Value :=
        FDataBuffer.FieldByName( 'Action' ).AsString;
      aRecordEffect := FDataUpdate.ExecSQL;
      if ( aRecordEffect <= 0 ) then
      begin
        FDataInsert.Parameters.ParamByName( 'FileName' ).Value :=
          FDataBuffer.FieldByName( 'FileName' ).Value;
        FDataInsert.Parameters.ParamByName( 'CreationDate' ).Value :=
          FDataBuffer.FieldByName( 'CreationDate' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductId' ).Value :=
          FDataBuffer.FieldByName( 'ProductId' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ImpulsiveFlag' ).Value :=
          FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ImpulsiveFlag' ).Value :=
          FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductKind' ).Value :=
          FDataBuffer.FieldByName( 'ProductKind' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SpecialPpv' ).Value :=
          FDataBuffer.FieldByName( 'SpecialPpv' ).AsString;
        FDataInsert.Parameters.ParamByName( 'PurchasableFlag' ).Value :=
          FDataBuffer.FieldByName( 'PurchasableFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SoldFlag' ).Value :=
          FDataBuffer.FieldByName( 'SoldFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SuspendedFlag' ).Value :=
          FDataBuffer.FieldByName( 'SuspendedFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductName' ).Value :=
          FDataBuffer.FieldByName( 'ProductName' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductDescription' ).Value :=
          FDataBuffer.FieldByName( 'ProductDescription' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EngProductName' ).Value :=
          FDataBuffer.FieldByName( 'EngProductName' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EngProductDescription' ).Value :=
          FDataBuffer.FieldByName( 'EngProductDescription' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SaleBeginTime' ).Value :=
          FDataBuffer.FieldByName( 'SaleBeginTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SaleEndTime' ).Value :=
          FDataBuffer.FieldByName( 'SaleEndTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ValidityPeriodBeginTime' ).Value :=
          FDataBuffer.FieldByName( 'ValidityPeriodBeginTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ValidityPeriodEndTime' ).Value :=
          FDataBuffer.FieldByName( 'ValidityPeriodEndTime' ).Value;
        FDataInsert.Parameters.ParamByName( 'ProductCategory' ).Value :=
          FDataBuffer.FieldByName( 'ProductCategory' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EpgPrice' ).Value :=
          FDataBuffer.FieldByName( 'EpgPrice' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SubAllowPpvEventFlag' ).Value :=
          FDataBuffer.FieldByName( 'SubAllowPpvEventFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EventId' ).Value :=
          FDataBuffer.FieldByName( 'EventId' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EventPreviewTime' ).Value :=
          FDataBuffer.FieldByName( 'EventPreviewTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UpdTime' ).Value :=
          FDataBuffer.FieldByName( 'UpdTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UpdEn' ).Value :=
          FDataBuffer.FieldByName( 'UpdEn' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ParentalRating' ).Value :=
          FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UTCSaleBeginTime' ).Value :=
          FDataBuffer.FieldByName( 'UTCSaleBeginTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UTCSaleEndTime' ).Value :=
          FDataBuffer.FieldByName( 'UTCSaleEndTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UTCValidityPeriodBeginTime' ).Value :=
          FDataBuffer.FieldByName( 'UTCValidityPeriodBeginTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UTCValidityPeriodEndTime' ).Value :=
          FDataBuffer.FieldByName( 'UTCValidityPeriodEndTime' ).AsString;
        FDataInsert.ExecSQL;
      end;
      FDataBuffer.Next;
      Sleep( 10 );
    end;
  end;
  { 處理刪除的 Product }
  if ( not FDeleteProduct.IsEmpty ) then
  begin
    FDataDelete.SQL.Text :=
      ' UPDATE SO158 SET dStopFlag = 1, Action = 2 ' +
      '  WHERE ProductId = :ProductId ';
    FDeleteProduct.First;
    while not FDeleteProduct.Eof do
    begin
      FDataDelete.Parameters.ParamByName( 'ProductId' ).Value :=
        FDeleteProduct.FieldByName( 'ProductId' ).AsString;
      FDataDelete.ExecSQL;
      FDeleteProduct.Next;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

{ TParseSO163 }

constructor TParseSO163.Create(aXml: TXMLDocument);
begin
  inherited Create( aXml );
  FDeleteProduct := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

destructor TParseSO163.Destroy;
begin
  FDeleteProduct.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO163.UnPrepareDataBuffer;
begin
  inherited;
  if not VarIsNull( FDeleteProduct.Data ) then
    FDeleteProduct.EmptyDataSet;
  FDeleteProduct.Data := Null;  
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO163.PrepareDataBuffer;
begin
  inherited;
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'ProductId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
  end;
  FDataBuffer.CreateDataSet;
  {}
  if FDeleteProduct.FieldDefs.Count <= 0 then
    FDeleteProduct.FieldDefs.Add( 'ProductId', ftString, 20 );
  FDeleteProduct.CreateDataSet;  
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO163.InternalParse;

 type
   TSo163 = record
     ProductId: String;
     ChannelId: String;
 end;

var
  aIndex, aIndex2: Integer;
  a163: TSo163;
  aIDOMList, aIDOMList2: IDOMNodeList;
  aProductNode, aNode: IDOMNode;
begin
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'Product' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    aProductNode := aIDOMList.item[aIndex];
    { 只處理 Sub }
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductType' );
    if WideUpperCase( aNode.attributes.getNamedItem(
      'productKind' ).nodeValue ) <> 'SUB' then Continue;
    {}
    ZeroMemory( @a163, SizeOf( a163 ) );
    {}
    a163.ProductId := ( aProductNode as IDOMNOdeSelect ).selectNode(
      'ProductId' ).firstChild.nodeValue;
    {}
    aNode := ( aProductNode as IDOMNOdeSelect ).selectNode( 'SubscriptionData' );
    aIDOMList2 := ( aNode as IDOMNodeSelect ).selectNodes( 'Channel' );
    for aIndex2 := 0 to aIDOMList2.length - 1 do
    begin
      a163.ChannelId := ( aIDOMList2.item[aIndex2] as IDOMNodeSelect ).selectNode(
        'ChannelId' ).firstChild.nodeValue;
      FDataBuffer.Append;
      FDataBuffer.FieldByName( 'ProductID' ).AsString := a163.ProductId;
      FDataBuffer.FieldByName( 'ChannelID' ).AsString := a163.ChannelId;
      FDataBuffer.Post;
      Sleep( 10 );
    end;
  end;
  { 刪除的 Product }
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'DeletedProductData' );
  if ( aIDOMList.length > 0 ) then
  begin
    aIDOMList := ( aIDOMList.item[0] as IDOMNodeSelect ).selectNodes( 'ProductId' );
    for aIndex := 0 to aIDOMList.length - 1 do
    begin
      FDeleteProduct.Append;
      FDeleteProduct.FieldByName( 'ProductId' ).AsString :=
        aIDOMList.item[aIndex].firstChild.nodeValue;  
      FDeleteProduct.Post;
      Sleep( 10 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO163.InternalApplyUpdate;
var
  aRecordEffect: Integer;
begin
  inherited;
  if ( not FDataBuffer.IsEmpty ) then
  begin
    FDataInsert.SQL.Clear;
    FDataUpdate.SQL.Clear;
    FDataInsert.SQL.Text :=
      ' INSERT INTO SO163 ( ProductId, ChannelId, dStopFlag ) ' +
      ' VALUES ( :ProductId, :ChannelId, 0 )                  ';
    FDataUpdate.SQL.Text :=
      ' UPDATE SO163 SET ChannelId = :ChannelId               ' +
      '  WHERE ProductId = :ProductId                         ';
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      FDataUpdate.Parameters.ParamByName( 'ProductId' ).Value :=
        FDataBuffer.FieldByName( 'ProductId' ).Value;
      FDataUpdate.Parameters.ParamByName( 'ChannelId' ).Value :=
        FDataBuffer.FieldByName( 'ChannelId' ).Value;
      aRecordEffect := FDataUpdate.ExecSQL;
      if ( aRecordEffect <= 0 ) then
      begin
        FDataInsert.Parameters.ParamByName( 'ProductId' ).Value :=
          FDataBuffer.FieldByName( 'ProductId' ).Value;
        FDataInsert.Parameters.ParamByName( 'ChannelId' ).Value :=
          FDataBuffer.FieldByName( 'ChannelId' ).AsString;
        FDataInsert.ExecSQL;
      end;
      FDataBuffer.Next;
      Sleep( 10 );
    end;
  end;
  { 處理刪除的 SUB Product }
  if ( not FDeleteProduct.IsEmpty ) then
  begin
    FDataDelete.SQL.Clear;
    FDataDelete.SQL.Text :=
      ' UPDATE SO163 SET dStopFlag = 1 WHERE ProductId = :ProductId ';
    FDeleteProduct.First;
    while not FDeleteProduct.Eof do
    begin
      FDataDelete.Parameters.ParamByName( 'ProductId' ).Value :=
        FDeleteProduct.FieldByName( 'ProductId' ).AsString;
      FDataDelete.ExecSQL;
      FDeleteProduct.Next;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TDexSubXmlFile }

constructor TDexSubXmlFile.Create;
var
  aIndex: Integer;
  aParse: TXmlParser;
begin
  inherited Create;
  for aIndex := 1 to 2 do
  begin
    case aIndex of
      1: aParse := TParseSO158Sub.Create( Self.FXml );
      2: aParse := TParseSO163.Create( Self.FXml );
    else
      aParse := nil;
    end;
    case aIndex of
      1: aParse.Name := 'SO158';
      2: aParse.Name := 'SO163';
    end;
    if Assigned( aParse ) then
    begin
      aParse.DataConnection := ParseController.DataConnection;
      aParse.DataUpdate := ParseController.DataUpdate;
      aParse.DataInsert := ParseController.DataInsert;
      aParse.DataDelete := ParseController.DataDelete;
      FParerList.Add( aParse );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

destructor TDexSubXmlFile.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TDexSubXmlFile.Parse(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.Parse;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDexSubXmlFile.Save(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.ApplyUpdate;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TAsRunXmlFile }

constructor TAsRunXmlFile.Create;
var
  aParse: TXmlParser;
begin
  inherited Create;
  aParse := TParseAsRun.Create( Self.FXml );
  aParse.Name := 'SO155A';
  aParse.DataConnection := ParseController.DataConnection;
  aParse.DataUpdate := ParseController.DataUpdate;
  aParse.DataInsert := ParseController.DataInsert;
  aParse.DataDelete := ParseController.DataDelete;
  aParse.DataReader := ParseController.DataReader;
  FParerList.Add( aParse );
end;

{ ---------------------------------------------------------------------------- }

destructor TAsRunXmlFile.Destroy;
begin
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TAsRunXmlFile.Parse(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.Parse;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TAsRunXmlFile.Save(var aMsg: String): Boolean;
var
  aParse: TXmlParser;
  aIndex: Integer;
begin
  Result := False;
  for aIndex := 0 to FParerList.Count - 1 do
  begin
    aParse := TXmlParser( FParerList[aIndex] );
    aParse.ApplyUpdate;
    aMsg := aParse.Msg;
    Result := ( aMsg = EmptyStr );
    if not Result then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TParseSO158Sub }

constructor TParseSO158Sub.Create(aXml: TXMLDocument);
begin
  inherited Create( aXml );
  FDeleteProduct := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

destructor TParseSO158Sub.Destroy;
begin
  FDeleteProduct.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Sub.PrepareDataBuffer;
begin
  inherited;
  { SO158 Dex XML資料檔 }
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'FileName', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'CreationDate', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ProductId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ImpulsiveFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'ProductKind', ftString, 3 );
    FDataBuffer.FieldDefs.Add( 'SpecialPpv', ftInteger );
    FDataBuffer.FieldDefs.Add( 'PurchasableFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'SoldFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'SuspendedFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'ProductName', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'ProductDescription', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'EngProductName', ftString, 160 );
    FDataBuffer.FieldDefs.Add( 'EngProductDescription', ftString, 1024 );
    FDataBuffer.FieldDefs.Add( 'SaleBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'SaleEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ValidityPeriodBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ValidityPeriodEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'ProductCategory', ftInteger );
    FDataBuffer.FieldDefs.Add( 'EpgPrice', ftFloat );
    FDataBuffer.FieldDefs.Add( 'SubAllowPpvEventFlag', ftInteger );
    FDataBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FDataBuffer.FieldDefs.Add( 'EventPreviewTime', ftString, 2 );
    FDataBuffer.FieldDefs.Add( 'UpdTime', ftString, 19 );
    FDataBuffer.FieldDefs.Add( 'UpdEn', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ParentalRating', ftInteger );
    FDataBuffer.FieldDefs.Add( 'dStopFlag', ftString, 1 );
    FDataBuffer.FieldDefs.Add( 'eStopFlag', ftString, 1 );
    FDataBuffer.FieldDefs.Add( 'UTCSaleBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'UTCSaleEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'UTCValidityPeriodBeginTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'UTCValidityPeriodEndTime', ftString, 14 );
    FDataBuffer.FieldDefs.Add( 'Action', ftInteger );
  end;
  FDataBuffer.CreateDataSet;
  {}
  if FDeleteProduct.FieldDefs.Count <= 0 then
    FDeleteProduct.FieldDefs.Add( 'ProductId', ftString, 20 );
  FDeleteProduct.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Sub.UnPrepareDataBuffer;
begin
  inherited;
  if not VarIsNull( FDeleteProduct.Data ) then
    FDeleteProduct.EmptyDataSet;
  FDeleteProduct.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Sub.InternalParse;

type
  TSo158 = record
    FileName: String;
    CreationDate: String;
    ProductId: String;
    ImpulsiveFlag: String;
    ProductKind: String;
    SpecialPpv: String;
    PurchasableFlag: String;
    SoldFlag: String;
    SuspendedFlag: String;
    ProductName: String;
    ProductDescription: String;
    EngProductName: String;
    EngProductDescription: String;
    SaleBeginTime: String;
    SaleEndTime: String;
    ValidityPeriodBeginTime: String;
    ValidityPeriodEndTime: String;
    ProductCategory: String;
    EpgPrice: String;
    MoneyUnit: String;
    SubAllowPpvEventFlag: String;
    EventId: String;
    EventPreviewTime: String;
    UpdTime: String;
    UpdEn: String;
    ParentalRating: String;
    UTCSaleBeginTime: String;
    UTCSaleEndTime: String;
    UTCValidityPeriodBeginTime: String;
    UTCValidityPeriodEndTime: String;
  end;

var
  aIDOMList, aIDOMList2: IDOMNodeList;
  aProductNode, aLanguageNode, aNode: IDOMNode;
  aIndex, aIndex2: Integer;
  a158: TSo158;
  aCreationDate, aFileName: String;
begin
  aFileName := FileNameWithoutExt( ExtractFileName( FXml.FileName ) );
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'ProductManagementData' );
  aCreationDate := aIDOMList.item[0].attributes.getNamedItem( 'creationDate' ).nodeValue;
  aIDOMList := nil;
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'Product' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    aProductNode := aIDOMList.item[aIndex];
    { 只處理 Sub }
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductType' );
    if WideUpperCase( aNode.attributes.getNamedItem(
      'productKind' ).nodeValue ) <> 'SUB' then Continue;
    {}
    ZeroMemory( @a158, SizeOf( a158 ) );
    a158.FileName := aFileName;
    a158.CreationDate := aCreationDate;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductId' );
    a158.ProductId := aNode.firstChild.nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductType' );
    a158.ImpulsiveFlag := aNode.attributes.getNamedItem( 'impulsiveFlag' ).nodeValue;
    a158.ProductKind := aNode.attributes.getNamedItem( 'productKind' ).nodeValue;
    a158.SpecialPpv := aNode.attributes.getNamedItem( 'specialPpv' ).nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductStatus' );
    a158.SuspendedFlag := aNode.attributes.getNamedItem( 'suspendedFlag' ).nodeValue;
    a158.SoldFlag := aNode.attributes.getNamedItem( 'soldFlag' ).nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'ProductCategory' );
    if Assigned( aNode ) then
      a158.ProductCategory := aNode.firstChild.nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'EpgPrice' );
    if Assigned( aNode ) then a158.EpgPrice := aNode.firstChild.nodeValue;
    {}
    aIDOMList2 := ( aProductNode as IDOMNodeSelect ).selectNodes( 'ProductText' );
    for aIndex2 := 0 to aIDOMList2.length - 1 do
    begin
      aLanguageNode := aIDOMList2.item[aIndex2];
      if ( aLanguageNode.attributes.getNamedItem( 'language' ).nodeValue = 'chi' ) then
      begin
        a158.ProductName := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductName' ).firstChild.nodeValue;
        a158.ProductDescription := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductDescription' ).firstChild.nodeValue;
      end else
      begin
        a158.EngProductName := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductName' ).firstChild.nodeValue;
        a158.EngProductDescription := ( aLanguageNode as IDOMNodeSelect ).selectNode(
          'ProductDescription' ).firstChild.nodeValue;
      end;
    end;
    {}
    if ( a158.ProductName = EmptyStr ) then a158.ProductName := a158.EngProductName;
    if ( a158.ProductDescription = EmptyStr ) then a158.ProductDescription := a158.EngProductDescription;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode( 'SalePeriod' );
    a158.SaleBeginTime := aNode.attributes.getNamedItem( 'beginTime' ).nodeValue;
    a158.SaleEndTime := aNode.attributes.getNamedItem( 'endTime' ).nodeValue;
    {}
    aNode := ( aProductNode as IDOMNodeSelect ).selectNode(
      'SubscriptionData' );
    a158.SubAllowPpvEventFlag := aNode.attributes.getNamedItem(
      'allowPpvEventsFlag' ).nodeValue;
    {}
    a158.UpdTime := ROCDate;
    a158.UpdEn := SMS_PROGRAM_SCHEDULAR;
    {}
    FDataBuffer.Append;
      FDataBuffer.FieldByName( 'FileName' ).AsString := Copy( a158.FileName, Length( a158.FileName ) - 13, 14 );
      FDataBuffer.FieldByName( 'CreationDate' ).AsString := a158.CreationDate;
      FDataBuffer.FieldByName( 'ProductId' ).AsString := a158.ProductId;
      FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString := a158.ImpulsiveFlag;
      FDataBuffer.FieldByName( 'ProductKind' ).AsString := a158.ProductKind;
      FDataBuffer.FieldByName( 'SpecialPpv' ).AsString := a158.SpecialPpv;
      FDataBuffer.FieldByName( 'PurchasableFlag' ).AsString := a158.PurchasableFlag;
      FDataBuffer.FieldByName( 'SoldFlag' ).AsString := a158.SoldFlag;
      FDataBuffer.FieldByName( 'SuspendedFlag' ).AsString := a158.SuspendedFlag;
      FDataBuffer.FieldByName( 'ProductName' ).AsString := a158.ProductName;
      FDataBuffer.FieldByName( 'ProductDescription' ).AsString := a158.ProductDescription;
      FDataBuffer.FieldByName( 'EngProductName' ).AsString := a158.EngProductName;
      FDataBuffer.FieldByName( 'EngProductDescription' ).AsString := a158.EngProductDescription;
      FDataBuffer.FieldByName( 'SaleBeginTime' ).AsString := a158.SaleBeginTime;
      FDataBuffer.FieldByName( 'SaleEndTime' ).AsString := a158.SaleEndTime;
      FDataBuffer.FieldByName( 'ValidityPeriodBeginTime' ).AsString := a158.ValidityPeriodBeginTime;
      FDataBuffer.FieldByName( 'ValidityPeriodEndTime' ).AsString := a158.ValidityPeriodEndTime;
      FDataBuffer.FieldByName( 'ProductCategory' ).AsString := a158.ProductCategory;
      FDataBuffer.FieldByName( 'EpgPrice' ).AsString := a158.EpgPrice;
      FDataBuffer.FieldByName( 'SubAllowPpvEventFlag' ).AsString := a158.SubAllowPpvEventFlag;
      FDataBuffer.FieldByName( 'EventId' ).AsString := a158.EventId;
      FDataBuffer.FieldByName( 'EventPreviewTime' ).AsString := a158.EventPreviewTime;
      FDataBuffer.FieldByName( 'UpdTime' ).AsString := a158.UpdTime;
      FDataBuffer.FieldByName( 'UpdEn' ).AsString := a158.UpdEn;
      FDataBuffer.FieldByName( 'ParentalRating' ).AsString := a158.ParentalRating;
      FDataBuffer.FieldByName( 'UTCSaleBeginTime' ).AsString := a158.UTCSaleBeginTime;
      FDataBuffer.FieldByName( 'UTCSaleEndTime' ).AsString := a158.UTCSaleEndTime;
      FDataBuffer.FieldByName( 'UTCValidityPeriodBeginTime' ).AsString := a158.UTCValidityPeriodBeginTime;
      FDataBuffer.FieldByName( 'UTCValidityPeriodEndTime' ).AsString := a158.UTCValidityPeriodEndTime;
      FDataBuffer.FieldByName( 'Action' ).AsString := '2';
    FDataBuffer.Post;
  end;
  { 標示為 DeleteProduct 的資料 }
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'DeletedProductData' );
  if ( aIDOMList.length > 0 ) then
  begin
    aIDOMList := ( aIDOMList.item[0] as IDOMNodeSelect ).selectNodes( 'ProductId' );
    for aIndex := 0 to aIDOMList.length - 1 do
    begin
      FDeleteProduct.Append;
      FDeleteProduct.FieldByName( 'ProductId' ).AsString :=
        aIDOMList.item[aIndex].firstChild.nodeValue;
      FDeleteProduct.Post;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseSO158Sub.InternalApplyUpdate;
var
  aRecordEffect: Integer;
begin
  inherited;
  if not ( FDataBuffer.IsEmpty ) then
  begin
    FDataInsert.SQL.Clear;
    FDataUpdate.SQL.Clear;
    FDataInsert.SQL.Text :=
      ' INSERT INTO SO158 ( FileName, CreationDate, ProductId, ImpulsiveFlag,  ' +
      '    ProductKind, SpecialPpv, PurchasableFlag, SoldFlag, SuspendedFlag,  ' +
      '    ProductName, ProductDescription, EngProductName,                    ' +
      '    EngProductDescription, SaleBeginTime, SaleEndTime,                  ' +
      '    ValidityPeriodBeginTime, ValidityPeriodEndTime, ProductCategory,    ' +
      '    EpgPrice, SubAllowPpvEventFlag, EventId,                            ' +
      '    EventPreviewTime, UpdTime, UpdEn, ParentalRating, Action,           ' +
      '    dStopFlag, eStopFlag  )                                             '+
      ' VALUES ( :FileName, :CreationDate, :ProductId, :ImpulsiveFlag,         ' +
      '    :ProductKind, :SpecialPpv, :PurchasableFlag, :SoldFlag,             ' +
      '    :SuspendedFlag, :ProductName, :ProductDescription, :EngProductName, ' +
      '    :EngProductDescription, :SaleBeginTime, :SaleEndTime,               ' +
      '    :ValidityPeriodBeginTime, :ValidityPeriodEndTime, :ProductCategory, ' +
      '    :EpgPrice, :SubAllowPpvEventFlag, :EventId,                         ' +
      '    :EventPreviewTime, :UpdTime, :UpdEn, :ParentalRating, 1, 0, 0  )    ';
    FDataUpdate.SQL.Text :=
      ' UPDATE SO158 SET                                                       ' +
      '    FileName = :FileName, CreationDate = :CreationDate,                 ' +
      '    ImpulsiveFlag = :ImpulsiveFlag, ProductKind = :ProductKind,         ' +
      '    SpecialPpv = :SpecialPpv, PurchasableFlag = :PurchasableFlag,       ' +
      '    SoldFlag = :SoldFlag, SuspendedFlag = :SuspendedFlag,               ' +
      '    ProductName = :ProductName,                                         ' +
      '    ProductDescription = :ProductDescription,                           ' +
      '    EngProductName = :EngProductName,                                   ' +
      '    EngProductDescription = :EngProductDescription,                     ' +
      '    SaleBeginTime = :SaleBeginTime, SaleEndTime = :SaleEndTime,         ' +
      '    ValidityPeriodBeginTime = :ValidityPeriodBeginTime,                 ' +
      '    ValidityPeriodEndTime = :ValidityPeriodEndTime,                     ' +
      '    ProductCategory = :ProductCategory, EpgPrice = :EpgPrice,           ' +
      '    SubAllowPpvEventFlag = :SubAllowPpvEventFlag, EventId = :EventId,   ' +
      '    EventPreviewTime = :EventPreviewTime, UpdTime = :UpdTime,           ' +
      '    UpdEn = :UpdEn, ParentalRating = :ParentalRating, dStopFlag = 0,    ' +
      '    Action = :Action                                                    ' +
      '  WHERE ProductId = :ProductId                                          ';
    FDataBuffer.First;
    while not FDataBuffer.Eof do
    begin
      FDataUpdate.Parameters.ParamByName( 'FileName' ).Value :=
        FDataBuffer.FieldByName( 'FileName' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'CreationDate' ).Value :=
        FDataBuffer.FieldByName( 'CreationDate' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ImpulsiveFlag' ).Value :=
        FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductKind' ).Value :=
        FDataBuffer.FieldByName( 'ProductKind' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SpecialPpv' ).Value :=
        FDataBuffer.FieldByName( 'SpecialPpv' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'PurchasableFlag' ).Value :=
        FDataBuffer.FieldByName( 'PurchasableFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SoldFlag' ).Value :=
        FDataBuffer.FieldByName( 'SoldFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SuspendedFlag' ).Value :=
        FDataBuffer.FieldByName( 'SuspendedFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductName' ).Value :=
        FDataBuffer.FieldByName( 'ProductName' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductDescription' ).Value :=
        FDataBuffer.FieldByName( 'ProductDescription' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EngProductName' ).Value :=
        FDataBuffer.FieldByName( 'EngProductName' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EngProductDescription' ).Value :=
        FDataBuffer.FieldByName( 'EngProductDescription' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SaleBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'SaleBeginTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SaleEndTime' ).Value :=
        FDataBuffer.FieldByName( 'SaleEndTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ValidityPeriodBeginTime' ).Value :=
        FDataBuffer.FieldByName( 'ValidityPeriodBeginTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ValidityPeriodEndTime' ).Value :=
        FDataBuffer.FieldByName( 'ValidityPeriodEndTime' ).Value;
      FDataUpdate.Parameters.ParamByName( 'ProductCategory' ).Value :=
        FDataBuffer.FieldByName( 'ProductCategory' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EpgPrice' ).Value :=
        FDataBuffer.FieldByName( 'EpgPrice' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'SubAllowPpvEventFlag' ).Value :=
        FDataBuffer.FieldByName( 'SubAllowPpvEventFlag' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EventId' ).Value :=
        FDataBuffer.FieldByName( 'EventId' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'EventPreviewTime' ).Value :=
        FDataBuffer.FieldByName( 'EventPreviewTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UpdTime' ).Value :=
        FDataBuffer.FieldByName( 'UpdTime' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'UpdEn' ).Value :=
        FDataBuffer.FieldByName( 'UpdEn' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ParentalRating' ).Value :=
        FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'ProductId' ).Value :=
        FDataBuffer.FieldByName( 'ProductId' ).AsString;
      FDataUpdate.Parameters.ParamByName( 'Action' ).Value :=
        FDataBuffer.FieldByName( 'Action' ).AsString;
      aRecordEffect := FDataUpdate.ExecSQL;
      if ( aRecordEffect <= 0 ) then
      begin
        FDataInsert.Parameters.ParamByName( 'FileName' ).Value :=
          FDataBuffer.FieldByName( 'FileName' ).Value;
        FDataInsert.Parameters.ParamByName( 'CreationDate' ).Value :=
          FDataBuffer.FieldByName( 'CreationDate' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductId' ).Value :=
          FDataBuffer.FieldByName( 'ProductId' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ImpulsiveFlag' ).Value :=
          FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ImpulsiveFlag' ).Value :=
          FDataBuffer.FieldByName( 'ImpulsiveFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductKind' ).Value :=
          FDataBuffer.FieldByName( 'ProductKind' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SpecialPpv' ).Value :=
          FDataBuffer.FieldByName( 'SpecialPpv' ).AsString;
        FDataInsert.Parameters.ParamByName( 'PurchasableFlag' ).Value :=
          FDataBuffer.FieldByName( 'PurchasableFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SoldFlag' ).Value :=
          FDataBuffer.FieldByName( 'SoldFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SuspendedFlag' ).Value :=
          FDataBuffer.FieldByName( 'SuspendedFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductName' ).Value :=
          FDataBuffer.FieldByName( 'ProductName' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ProductDescription' ).Value :=
          FDataBuffer.FieldByName( 'ProductDescription' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EngProductName' ).Value :=
          FDataBuffer.FieldByName( 'EngProductName' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EngProductDescription' ).Value :=
          FDataBuffer.FieldByName( 'EngProductDescription' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SaleBeginTime' ).Value :=
          FDataBuffer.FieldByName( 'SaleBeginTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SaleEndTime' ).Value :=
          FDataBuffer.FieldByName( 'SaleEndTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ValidityPeriodBeginTime' ).Value :=
          FDataBuffer.FieldByName( 'ValidityPeriodBeginTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ValidityPeriodEndTime' ).Value :=
          FDataBuffer.FieldByName( 'ValidityPeriodEndTime' ).Value;
        FDataInsert.Parameters.ParamByName( 'ProductCategory' ).Value :=
          FDataBuffer.FieldByName( 'ProductCategory' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EpgPrice' ).Value :=
          FDataBuffer.FieldByName( 'EpgPrice' ).AsString;
        FDataInsert.Parameters.ParamByName( 'SubAllowPpvEventFlag' ).Value :=
          FDataBuffer.FieldByName( 'SubAllowPpvEventFlag' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EventId' ).Value :=
          FDataBuffer.FieldByName( 'EventId' ).AsString;
        FDataInsert.Parameters.ParamByName( 'EventPreviewTime' ).Value :=
          FDataBuffer.FieldByName( 'EventPreviewTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UpdTime' ).Value :=
          FDataBuffer.FieldByName( 'UpdTime' ).AsString;
        FDataInsert.Parameters.ParamByName( 'UpdEn' ).Value :=
          FDataBuffer.FieldByName( 'UpdEn' ).AsString;
        FDataInsert.Parameters.ParamByName( 'ParentalRating' ).Value :=
          FDataBuffer.FieldByName( 'ParentalRating' ).AsString;
        FDataInsert.ExecSQL;
      end;
      FDataBuffer.Next;
      Sleep( 10 );
    end;
  end;  
  { 處理刪除的 Product }
  if ( not FDeleteProduct.IsEmpty ) then
  begin
    FDataDelete.SQL.Clear;
    FDataDelete.SQL.Text :=
      ' UPDATE SO158 SET dStopFlag = 1 WHERE ProductId = :ProductId ';
    FDeleteProduct.First;
    while not FDeleteProduct.Eof do
    begin
      FDataDelete.Parameters.ParamByName( 'ProductId' ).Value :=
        FDeleteProduct.FieldByName( 'ProductId' ).AsString;
      FDataDelete.ExecSQL;
      FDeleteProduct.Next;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TParseAsRun }

constructor TParseAsRun.Create(aXml: TXMLDocument);
begin
  inherited Create( aXml );
  FOrignalEvent := TClientDataSet.Create( nil );
  FChannelTime := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

destructor TParseAsRun.Destroy;
begin
  FOrignalEvent.Free;
  FChannelTime.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseAsRun.PrepareDataBuffer;
begin
  inherited;
  { AsRun XML 資料 }
  if FDataBuffer.FieldDefs.Count <= 0 then
  begin
    FDataBuffer.FieldDefs.Add( 'Device', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'PlayStart', ftDateTime );
    FDataBuffer.FieldDefs.Add( 'PlayEnd', ftDateTime );
    FDataBuffer.FieldDefs.Add( 'Played', ftBoolean );
    FDataBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FDataBuffer.FieldDefs.Add( 'EventName', ftString, 120 );
  end;
  FDataBuffer.IndexName := EmptyStr;
  FDataBuffer.CreateDataSet;
  FDataBuffer.AddIndex( 'IDX1', 'ChannelId;EventId', []  );
  FDataBuffer.IndexName := 'IDX1';
  { 更新資料時, 須把原 event 資料先 store 起來 }
  if ( FOrignalEvent.FieldDefs.Count <= 0 ) then
  begin
    FOrignalEvent.FieldDefs.Add( 'ProductId', ftString, 20 );
    FOrignalEvent.FieldDefs.Add( 'EventId', ftString, 12 );
    FOrignalEvent.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FOrignalEvent.FieldDefs.Add( 'EventBeginTime', ftDateTime );
    FOrignalEvent.FieldDefs.Add( 'EventEndTime', ftDateTime );
    FOrignalEvent.FieldDefs.Add( 'OldPlayStatus', ftInteger );
    FOrignalEvent.FieldDefs.Add( 'NewPlayStatus', ftInteger );
  end;
  FOrignalEvent.IndexName := EmptyStr;
  FOrignalEvent.CreateDataSet;
  FOrignalEvent.AddIndex( 'IDX1', 'ProductId;EventBeginTime', [] );
  FOrignalEvent.IndexName := 'IDX1';
  {}
  if ( FChannelTime.FieldDefs.Count <= 0 ) then
  begin
    FChannelTime.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FChannelTime.FieldDefs.Add( 'MinTime', ftDateTime );
    FChannelTime.FieldDefs.Add( 'MaxTime', ftDateTime );
  end;
  FChannelTime.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseAsRun.UnPrepareDataBuffer;
begin
  inherited;
  if not VarIsNull( FOrignalEvent.Data ) then
    FOrignalEvent.EmptyDataSet;
  FOrignalEvent.Data := Null;
  {}
  if not VarIsNull( FChannelTime.Data ) then
    FChannelTime.EmptyDataSet;
  FChannelTime.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseAsRun.InternalParse;

type
  TAsRun = record
    Device: String;
    ChannelId: String;
    PlayStart: TDateTime;
    PlayEnd: TDateTime;
    Played: Boolean;
    EventId: String;
    EventName: String;
  end;

  { ------------------------------------------------------ }

  function IsEventFound(const aChannelId, aEventId: String): Boolean;
  begin
    Result := FDataBuffer.Locate( 'ChannelId;EventId', VarArrayOf(
      [aChannelId, aEventId] ), [] );
  end;

  { ------------------------------------------------------ }

  function IsChannelTimeFound(const aChannelId: String): Boolean;
  begin
    Result := FChannelTime.Locate( 'ChannelId', aChannelId, [] );
  end;

  { ------------------------------------------------------ }

var
  aIDOMList: IDOMNodeList;
  aNode: IDOMNode;
  aIndex: Integer;
  aAsRun: TAsRun;
  aPlayStart, aPlayEnd: String;
begin
  inherited;
  aIDOMList := FXml.DOMDocument.getElementsByTagName( 'AsRun' );
  for aIndex := 0 to aIDOMList.length - 1 do
  begin
    aNode := aIDOMList.item[aIndex].attributes.getNamedItem( 'eventID' );
    { EventId = 空字串的是 Promotion 的廣告, 不是正常的 Event }
    if ( aNode.nodeValue = EmptyStr ) then
      Continue;
    {}
    ZeroMemory( @aAsRun, SizeOf( aAsRun ) );
    {}
    aAsRun.Device := aIDOMList.item[aIndex].attributes.getNamedItem( 'device' ).nodeValue;
    {}
    aAsRun.ChannelId := aIDOMList.item[aIndex].attributes.getNamedItem( 'channelID' ).nodeValue;
    {}
    aPlayStart := aIDOMList.item[aIndex].attributes.getNamedItem( 'playStart' ).nodeValue;
    aPlayStart := StringReplace( aPlayStart, 'T', #32, [rfReplaceAll] );
    aPlayStart := StringReplace( aPlayStart, 'Z', EmptyStr, [rfReplaceAll] );
    aPlayStart := StringReplace( aPlayStart, '-', '/', [rfReplaceAll] );
    aAsRun.PlayStart := StrToDateTime( aPlayStart );
    {}
    aPlayEnd := aIDOMList.item[aIndex].attributes.getNamedItem( 'playEnd' ).nodeValue;
    aPlayEnd := StringReplace( aPlayEnd, 'T', #32, [rfReplaceAll] );
    aPlayEnd := StringReplace( aPlayEnd, 'Z', EmptyStr, [rfReplaceAll] );
    aPlayEnd := StringReplace( aPlayEnd, '-', '/', [rfReplaceAll] );
    aAsRun.PlayEnd := StrToDateTime( aPlayEnd );
    {}
    aAsRun.Played := ( aIDOMList.item[aIndex].attributes.getNamedItem( 'played' ).nodeValue = 'true' );
    aAsRun.EventId := aIDOMList.item[aIndex].attributes.getNamedItem( 'eventID' ).nodeValue;
    aAsRun.EventName := aIDOMList.item[aIndex].attributes.getNamedItem( 'eventName' ).nodeValue;
    {}
    if not IsEventFound( aAsRun.ChannelId, aAsRun.EventId ) then
    begin
      FDataBuffer.Append;
      FDataBuffer.FieldByName( 'Device' ).AsString := aAsRun.Device;
      FDataBuffer.FieldByName( 'ChannelId' ).AsString := aAsRun.ChannelId;
      FDataBuffer.FieldByName( 'PlayStart' ).AsDateTime := aAsRun.PlayStart;
      FDataBuffer.FieldByName( 'PlayEnd' ).AsDateTime := aAsRun.PlayEnd;
      FDataBuffer.FieldByName( 'Played' ).AsBoolean := aAsRun.Played;
      FDataBuffer.FieldByName( 'EventId' ).AsString := aAsRun.EventId;
      FDataBuffer.FieldByName( 'EventName' ).AsString := aAsRun.EventName;
      FDataBuffer.Post;
    end else
    begin
      { 沒撥映過的 evnetid, 因為會有2筆Record, Device01 及 Device02, 其中有一筆是 Played=true
        則表示該 event 已撥映, 若 2 筆都是 played=false, 則表示未撥映 }
      if ( aAsRun.Played ) and ( not FDataBuffer.FieldByName( 'Played' ).AsBoolean  ) then
      begin
        FDataBuffer.Edit;
        FDataBuffer.FieldByName( 'Played' ).AsBoolean := True;
        FDataBuffer.Post;
      end;
    end;
    { 把每一個 Channel 裏, 最小及最大的 Event Play Time 記錄起來 }
    if not IsChannelTimeFound( aAsRun.ChannelId ) then
    begin
      FChannelTime.Append;
      FChannelTime.FieldByName( 'ChannelId' ).AsString := aAsRun.ChannelId;
      FChannelTime.FieldByName( 'MinTime' ).AsDateTime := aAsRun.PlayStart;
      FChannelTime.FieldByName( 'MaxTime' ).AsDateTime := aAsRun.PlayEnd;
      FChannelTime.Post;
    end else
    begin
      FChannelTime.Edit;
      if ( FChannelTime.FieldByName( 'MinTime' ).AsDateTime > aAsRun.PlayStart ) then
        FChannelTime.FieldByName( 'MinTime' ).AsDateTime := aAsRun.PlayStart;
      if ( FChannelTime.FieldByName( 'MaxTime' ).AsDateTime < aAsRun.PlayEnd ) then
        FChannelTime.FieldByName( 'MaxTime' ).AsDateTime := aAsRun.PlayEnd;
      FChannelTime.Post;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TParseAsRun.InternalApplyUpdate;

  { ------------------------------------------------------ }

  function IsEventFound(const aProductId, aEventId: String): Boolean;
  begin
    Result := FOrignalEvent.Locate( 'ProductId;EventId', VarArrayOf(
      [aProductId, aEventId] ), [] );
  end;

  { ------------------------------------------------------ }

var
  aSql, aPreProductId: String;
  aPlayStatus: Integer;
  aAllEventPlayed: Boolean;
begin
  inherited;
  if ( FDataBuffer.IsEmpty ) then Exit;
  FDataUpdate.SQL.Clear;
  FDataInsert.SQL.Clear;
  FDataDelete.SQL.Clear;
  FDataReader.SQL.Clear;
  {}
  aSql :=
    ' select a.productid, a.eventid, a.channelid,  a.eventbegintime,  ' +
    '        a.eventendtime, a.playstatus                             ' +
    '   from so155a a,  ( select distinct productid from so155a       ' +
    '  where eventid = ''%s'' ) b where a.productid = b.productid     ' +
    '  order by productid, eventbegintime                             ';
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    FDataReader.SQL.Text := Format( aSql, [
      FDataBuffer.FieldByName( 'eventid' ).AsString] );
    FDataReader.Open;
    FDataReader.First;
    { 先把所有 SO155A 裏面符合該 EventId 的 Product ( 含 Event ) 全抓出來 }
    while not FDataReader.Eof do
    begin
      if not IsEventFound( FDataReader.FieldByName( 'ProductId' ).AsString,
        FDataReader.FieldByName( 'EventId' ).AsString ) then
      begin
        FOrignalEvent.Append;
        FOrignalEvent.FieldByName( 'ProductId' ).AsString := FDataReader.FieldByName( 'productid' ).AsString;
        FOrignalEvent.FieldByName( 'EventId' ).AsString := FDataReader.FieldByName( 'eventid' ).AsString;
        FOrignalEvent.FieldByName( 'ChannelId' ).AsString := FDataReader.FieldByName( 'channelid' ).AsString;
        FOrignalEvent.FieldByName( 'EventBeginTime' ).AsDateTime := FDataReader.FieldByName( 'EventBeginTime' ).AsDateTime;
        FOrignalEvent.FieldByName( 'EventEndTime' ).AsDateTime := FDataReader.FieldByName( 'EventEndTime' ).AsDateTime;
        FOrignalEvent.FieldByName( 'OldPlayStatus' ).AsInteger := FDataReader.FieldByName( 'PlayStatus' ).AsInteger;
        FOrignalEvent.FieldByName( 'NewPlayStatus' ).AsInteger := 0;
        FOrignalEvent.Post;
      end;
      FDataReader.Next;
    end;
    FDataReader.Close;
    FDataBuffer.Next;
  end;
  { 把所有 Event 的 PlayStatus 填回去 }
  FDataBuffer.First;
  while not FDataBuffer.Eof do
  begin
    aPlayStatus := IIF( FDataBuffer.FieldByName( 'Played' ).AsBoolean, 1, 2 );
    FOrignalEvent.Filtered := False;
    FOrignalEvent.Filter := Format( 'EventId=''%s''', [
      FDataBuffer.FieldByName( 'EventId' ).AsString] );
    FOrignalEvent.Filtered := True;
    FOrignalEvent.First;
    while not FOrignalEvent.Eof do
    begin
      FOrignalEvent.Edit;
      FOrignalEvent.FieldByName( 'NewPlayStatus' ).AsInteger := aPlayStatus;
      FOrignalEvent.Post;
      FOrignalEvent.Next;
    end;
    FDataBuffer.Next;
  end;
  { 比對該 Product 下, 所有 Event 的撥映狀態 }
  FOrignalEvent.Filtered := False;
  FOrignalEvent.Filter := EmptyStr;
  FOrignalEvent.First;
  aPreProductId := FOrignalEvent.FieldByName( 'ProductId' ).AsString;
  aAllEventPlayed := True;
  while not FOrignalEvent.Eof do
  begin
    { 1.更新該 Event 的撥映狀態 }
    if ( FOrignalEvent.FieldByName( 'OldPlayStatus' ).AsInteger <>
         FOrignalEvent.FieldByName( 'NewPlayStatus' ).AsInteger ) then
    begin
      FDataUpdate.SQL.Text := Format(
        ' update so155a set playstatus = ''%d''             ' +
        '   where productid = ''%s'' and eventid = ''%s''   ',
        [FOrignalEvent.FieldByName( 'NewPlayStatus' ).AsInteger,
         FOrignalEvent.FieldByName( 'productid' ).AsString,
         FOrignalEvent.FieldByName( 'eventid' ).AsString] );
      FDataUpdate.ExecSQL;
    end;
    if ( FOrignalEvent.FieldByName( 'NewPlayStatus' ).AsInteger <> 1 ) then
      aAllEventPlayed := False;
    FOrignalEvent.Next;
    { 2.換 Product, 更新該 Product 的撥映狀態 }
    if ( aPreProductId <> FOrignalEvent.FieldByName( 'ProductId' ).AsString ) or
       ( FOrignalEvent.Eof ) then
    begin
      if ( aAllEventPlayed ) then
      begin
        FDataUpdate.SQL.Text := Format(
         ' update so155 set playstatus = 1 where productid = ''%s'' ', [aPreProductId] );
        FDataUpdate.ExecSQL;
      end;
      aPreProductId := FOrignalEvent.FieldByName( 'ProductId' ).AsString;
      aAllEventPlayed := True;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TMergeChange }

constructor TMergeChange.Create;
begin
  inherited Create;
  FSo160 := TClientDataSet.Create( nil );
  FProductBuffer := TClientDataSet.Create( nil );
  FEventBuffer := TClientDataSet.Create( nil );
  FProductBuffer2 := TClientDataSet.Create( nil );
  FEventBuffer2 := TClientDataSet.Create( nil );
  FLogText := TClientDataSet.Create( nil );
  FDataDelete := nil;
  FDataWriter := nil;
  FDataReader := nil;
  FDataConnection := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TMergeChange.Destroy;
begin
  FSo160.Free;
  FProductBuffer.Free;
  FEventBuffer.Free;
  FProductBuffer2.Free;
  FEventBuffer2.Free;
  FLogText.Free;
  FDataDelete := nil;
  FDataWriter := nil;
  FDataReader := nil;
  FDataConnection := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TMergeChange.PrepareDataBuffer;
begin
  if ( FSo160.FieldDefs.Count <= 0 ) then
  begin
    FSo160.FieldDefs.Add( 'dFileName', ftString, 14 );
    FSo160.FieldDefs.Add( 'dCreationDate', ftString, 14 );
    FSo160.FieldDefs.Add( 'ProductId', ftString, 20 );
    FSo160.FieldDefs.Add( 'ImpulsiveFlag', ftInteger );
    FSo160.FieldDefs.Add( 'ProductKind', ftString, 3 );
    FSo160.FieldDefs.Add( 'SpecialPpv', ftInteger );
    FSo160.FieldDefs.Add( 'PurchasableFlag', ftInteger );
    FSo160.FieldDefs.Add( 'SoldFlag', ftInteger );
    FSo160.FieldDefs.Add( 'SuspendedFlag', ftInteger );
    FSo160.FieldDefs.Add( 'ProductName', ftString, 160 );
    FSo160.FieldDefs.Add( 'ProductDescription', ftString, 1024 );
    FSo160.FieldDefs.Add( 'EngProductName', ftString, 160 );
    FSo160.FieldDefs.Add( 'EngProductDescription', ftString, 1024 );
    FSo160.FieldDefs.Add( 'SaleBeginTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'SaleEndTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'ValidityPeriodBeginTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'ValidityPeriodEndTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'ProductCategory', ftInteger );
    FSo160.FieldDefs.Add( 'EpgPrice', ftFloat );
    FSo160.FieldDefs.Add( 'SubAllowPpvEventFlag', ftInteger );
    FSo160.FieldDefs.Add( 'eFileName', ftString, 14 );
    FSo160.FieldDefs.Add( 'eCreationDate', ftString, 14 );
    FSo160.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FSo160.FieldDefs.Add( 'EventBeginTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'Duration', ftInteger );
    FSo160.FieldDefs.Add( 'EventId', ftString, 12 );
    FSo160.FieldDefs.Add( 'EventPreviewTime', ftString, 2 );
    FSo160.FieldDefs.Add( 'EventType', ftString, 1 );
    FSo160.FieldDefs.Add( 'Name', ftString, 160 );
    FSo160.FieldDefs.Add( 'EpgDesc', ftString, 1024 );
    FSo160.FieldDefs.Add( 'EngName', ftString, 160 );
    FSo160.FieldDefs.Add( 'EngEpgDesc', ftString, 1024 );
    FSo160.FieldDefs.Add( 'ParentalRating', ftInteger );
    FSo160.FieldDefs.Add( 'ParentalRatingName', ftString, 20 );
    FSo160.FieldDefs.Add( 'ContentNibble1', ftString, 2 );
    FSo160.FieldDefs.Add( 'ContentNibble2', ftString, 2 );
    FSo160.FieldDefs.Add( 'UserNibble1', ftString, 2 );
    FSo160.FieldDefs.Add( 'UserNibble2', ftString, 2 );
    FSo160.FieldDefs.Add( 'ChannelNo', ftInteger );
    FSo160.FieldDefs.Add( 'UpdTime', ftString, 19 );
    FSo160.FieldDefs.Add( 'UpdEn', ftString, 20 );
    FSo160.FieldDefs.Add( 'dParentalRating', ftInteger );
    FSo160.FieldDefs.Add( 'dParentalRatingName', ftString, 20 );
    FSo160.FieldDefs.Add( 'dStopFlag', ftInteger );
    FSo160.FieldDefs.Add( 'eStopFlag', ftInteger );
    FSo160.FieldDefs.Add( 'eeStopFlag', ftInteger );
    FSo160.FieldDefs.Add( 'FilmNo', ftString, 15 );
    FSo160.FieldDefs.Add( 'UTCSaleBeginTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'UTCSaleEndTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'UTCValidityPeriodBeginTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'UTCValidityPeriodEndTime', ftString, 14 );
    FSo160.FieldDefs.Add( 'UTCEventBeginTime', ftString, 14 );
  end;
  FSo160.IndexName := EmptyStr;
  FSo160.CreateDataSet;
  FSo160.AddIndex( 'IDX1', 'ProductId;EventBeginTime', [] );
  FSo160.IndexName := 'IDX1';
  {}
  if ( FProductBuffer.FieldDefs.Count <= 0 ) then
  begin
    FProductBuffer.FieldDefs.Add( 'ProductId', ftString, 20 );
    FProductBuffer.FieldDefs.Add( 'ImpulsiveFlag', ftInteger );
    FProductBuffer.FieldDefs.Add( 'ProductKind', ftString, 3 );
    FProductBuffer.FieldDefs.Add( 'SpecialPpv', ftInteger );
    FProductBuffer.FieldDefs.Add( 'PurchasableFlag', ftInteger );
    FProductBuffer.FieldDefs.Add( 'SoldFlag', ftInteger );
    FProductBuffer.FieldDefs.Add( 'SuspendedFlag', ftInteger );
    FProductBuffer.FieldDefs.Add( 'ProductName', ftString, 160 );
    FProductBuffer.FieldDefs.Add( 'ProductDescription', ftString, 1024 );
    FProductBuffer.FieldDefs.Add( 'EngProductName', ftString, 160 );
    FProductBuffer.FieldDefs.Add( 'EngProductDescription', ftString, 1024 );
    FProductBuffer.FieldDefs.Add( 'SaleBeginTime', ftString, 14 );
    FProductBuffer.FieldDefs.Add( 'SaleEndTime', ftString, 14 );
    FProductBuffer.FieldDefs.Add( 'ValidityPeriodBeginTime', ftString, 14 );
    FProductBuffer.FieldDefs.Add( 'ValidityPeriodEndTime', ftString, 14 );
    FProductBuffer.FieldDefs.Add( 'ProductCategory', ftInteger );
    FProductBuffer.FieldDefs.Add( 'EpgPrice', ftFloat );
    FProductBuffer.FieldDefs.Add( 'SubAllowPpvEventFlag', ftInteger );
    FProductBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FProductBuffer.FieldDefs.Add( 'ChannelNo', ftInteger );
    FProductBuffer.FieldDefs.Add( 'EventBeginTime', ftString, 14 );
    FProductBuffer.FieldDefs.Add( 'EventEndTime', ftString, 14 );
    FProductBuffer.FieldDefs.Add( 'Duration', ftInteger );
    FProductBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FProductBuffer.FieldDefs.Add( 'dParentalRating', ftInteger );
    FProductBuffer.FieldDefs.Add( 'dParentalRatingName', ftString, 20 );
    FProductBuffer.FieldDefs.Add( 'dStopFlag', ftInteger );
  end;
  FProductBuffer.IndexName := EmptyStr;
  FProductBuffer.CreateDataSet;
  FProductBuffer.AddIndex( 'IDX1', 'ProductId', [] );
  FProductBuffer.IndexName := 'IDX1';
  {}
  if ( FEventBuffer.FieldDefs.Count <= 0 ) then
  begin
    FEventBuffer.FieldDefs.Add( 'ProductId', ftString, 20 );
    FEventBuffer.FieldDefs.Add( 'EventId', ftString, 12 );
    FEventBuffer.FieldDefs.Add( 'ChannelId', ftString, 20 );
    FEventBuffer.FieldDefs.Add( 'ChannelNo', ftString, 20 );
    FEventBuffer.FieldDefs.Add( 'EventBeginTime', ftString, 14 );
    FEventBuffer.FieldDefs.Add( 'EventEndTime', ftString, 14 );
    FEventBuffer.FieldDefs.Add( 'EventType', ftString, 1 );
    FEventBuffer.FieldDefs.Add( 'FilmNo', ftString, 1 );
    FEventBuffer.FieldDefs.Add( 'EventPreviewTime', ftString, 2 );
    FEventBuffer.FieldDefs.Add( 'eStopFlag', ftInteger );
    FEventBuffer.FieldDefs.Add( 'Name', ftString, 160 );
  end;
  FEventBuffer.IndexName := EmptyStr;
  FEventBuffer.CreateDataSet;
  FEventBuffer.AddIndex( 'IDX1', 'ProductId;EventBeginTime', [] );
  FEventBuffer.IndexName := 'IDX1';
  {}
  if ( FProductBuffer2.FieldDefs.Count <= 0 ) then
    FProductBuffer2.FieldDefs.Assign( FProductBuffer.FieldDefs );
  FProductBuffer2.IndexName := EmptyStr;
  FProductBuffer2.CreateDataSet;
  FProductBuffer2.AddIndex( 'IDX1', 'ProductId', [] );
  FProductBuffer2.IndexName := 'IDX1';
  {}
  if ( FEventBuffer2.FieldDefs.Count <= 0 ) then
    FEventBuffer2.FieldDefs.Assign( FEventBuffer.FieldDefs );
  FEventBuffer2.IndexName := EmptyStr;
  FEventBuffer2.CreateDataSet;
  FEventBuffer2.AddIndex( 'IDX1', 'ProductId;EventBeginTime', [] );
  FEventBuffer2.IndexName := 'IDX1';
  {}
  if ( FLogText.FieldDefs.Count <= 0 ) then
  begin
    FLogText.FieldDefs.Add( 'ProductId', ftString, 12 );
    FLogText.FieldDefs.Add( 'Note', ftString, 256 );
    FLogText.FieldDefs.Add( 'Flag', ftInteger );
    FLogText.FieldDefs.Add( 'Upden', ftString, 20 );
    FLogText.FieldDefs.Add( 'UpdTime', ftString, 19 );
  end;
  FLogText.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TMergeChange.UnPrepareDataBuffer;
begin
  if not VarIsNull( FSo160.Data ) then FSo160.EmptyDataSet;
  FSo160.Data := Null;
  {}
  if not VarIsNull( FProductBuffer.Data ) then FProductBuffer.EmptyDataSet;
  FProductBuffer.Data := Null;
  if not VarIsNull( FEventBuffer.Data ) then FEventBuffer.EmptyDataSet;
  FEventBuffer.Data := Null;
  {}
  if not VarIsNull( FProductBuffer2.Data ) then FProductBuffer2.EmptyDataSet;
  FProductBuffer2.Data := Null;
  if not VarIsNull( FEventBuffer2.Data ) then FEventBuffer2.EmptyDataSet;
  FEventBuffer2.Data := Null;
  {}
  if not VarIsNull( FLogText.Data ) then FLogText.EmptyDataSet;
  FLogText.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.StoreRecord(const aAction: TStoreAction; var aMsg: String): Boolean;
var
  aProductSql, aEventSql: String;
  aProductStoreDS, aEventStoreDS: TClientDataSet;
begin
  Result := False;
  aProductSql :=
      ' SELECT                                             ' +
      '      A.PRODUCTID,                                  ' +
      '      A.IMPULSIVEFLAG,                              ' +
      '      A.PRODUCTKIND,                                ' +
      '      A.SPECIALPPV,                                 ' +
      '      A.PURCHASABLEFLAG,                            ' +
      '      A.SOLDFLAG,                                   ' +
      '      A.SUSPENDEDFLAG,                              ' +
      '      A.PRODUCTNAME,                                ' +
      '      A.PRODUCTDESCRIPTION,                         ' +
      '      A.ENGPRODUCTNAME,                             ' +
      '      A.ENGPRODUCTDESCRIPTION,                      ' +
      '      A.SALEBEGINTIME,                              ' +
      '      A.SALEENDTIME,                                ' +
      '      A.VALIDITYPERIODBEGINTIME,                    ' +
      '      A.VALIDITYPERIODENDTIME,                      ' +
      '      A.PRODUCTCATEGORY,                            ' +
      '      A.EPGPRICE,                                   ' +
      '      A.SUBALLOWPPVEVENTFLAG,                       ' +
      '      A.CHANNELID,                                  ' +
      '      A.CHANNELNO,                                  ' +
      '      A.EVENTBEGINTIME,                             ' +
      '      A.EVENTENDTIME,                               ' +
      '      A.DURATION,                                   ' +
      '      A.EVENTID,                                    ' +
      '      A.DPARENTALRATING,                            ' +
      '      A.DPARENTALRATINGNAME,                        ' +
      '      A.DSTOPFLAG                                   ' +
      '  FROM SO155 A,                                     ' +
      '      ( SELECT DISTINCT PRODUCTID FROM SO160 ) B    ' +
      ' WHERE A.PRODUCTID = B.PRODUCTID                    ';
  {}
  aEventSql :=
      ' SELECT A.PRODUCTID,                                ' +
      '        A.EVENTID,                                  ' +
      '        A.CHANNELID,                                ' +
      '        A.CHANNELNO,                                ' +
      '        A.EVENTBEGINTIME,                           ' +
      '        A.EVENTENDTIME,                             ' +
      '        A.EVENTTYPE,                                ' +
      '        A.FILMNO,                                   ' +
      '        A.EVENTPREVIEWTIME,                         ' +
      '        A.ESTOPFLAG,                                ' +
      '        C.NAME                                      ' +
      '   FROM SO155A A,                                   ' +
      '       ( SELECT DISTINCT PRODUCTID FROM SO160 ) B,  ' +
      '       SO154 C                                      ' +
      '  WHERE A.PRODUCTID = B.PRODUCTID                   ' +
      '    AND A.FILMNO = C.FILMNO(+)                      ';
  {}
  if ( aAction = saBefore ) then
  begin
    aProductStoreDS := FProductBuffer;
    aEventStoreDS := FEventBuffer;
  end else
  begin
    aProductStoreDS := FProductBuffer2;
    aEventStoreDS := FEventBuffer2;
  end;
  try
    FDataReader.Close;
    FDataReader.SQL.Text := aProductSql;
    FDataReader.Open;
    FDataReader.First;
    while not FDataReader.Eof do
    begin
      aProductStoreDS.Append;
      aProductStoreDS.FieldByName( 'PRODUCTID' ).AsString :=
        FDataReader.FieldByName( 'PRODUCTID' ).AsString;
      aProductStoreDS.FieldByName( 'IMPULSIVEFLAG' ).AsString :=
        FDataReader.FieldByName( 'IMPULSIVEFLAG' ).AsString;
      aProductStoreDS.FieldByName( 'PRODUCTKIND' ).AsString :=
        FDataReader.FieldByName( 'PRODUCTKIND' ).AsString;
      aProductStoreDS.FieldByName( 'SPECIALPPV' ).AsString :=
        FDataReader.FieldByName( 'SPECIALPPV' ).AsString;
      aProductStoreDS.FieldByName( 'PURCHASABLEFLAG' ).AsString :=
        FDataReader.FieldByName( 'PURCHASABLEFLAG' ).AsString;
      aProductStoreDS.FieldByName( 'SOLDFLAG' ).AsString :=
        FDataReader.FieldByName( 'SOLDFLAG' ).AsString;
      aProductStoreDS.FieldByName( 'SUSPENDEDFLAG' ).AsString :=
        FDataReader.FieldByName( 'SUSPENDEDFLAG' ).AsString;
      aProductStoreDS.FieldByName( 'PRODUCTNAME' ).AsString :=
        FDataReader.FieldByName( 'PRODUCTNAME' ).AsString;
      aProductStoreDS.FieldByName( 'PRODUCTDESCRIPTION' ).AsString :=
        FDataReader.FieldByName( 'PRODUCTDESCRIPTION' ).AsString;
      aProductStoreDS.FieldByName( 'ENGPRODUCTNAME' ).AsString :=
        FDataReader.FieldByName( 'ENGPRODUCTNAME' ).AsString;
      aProductStoreDS.FieldByName( 'ENGPRODUCTDESCRIPTION' ).AsString :=
        FDataReader.FieldByName( 'ENGPRODUCTDESCRIPTION' ).AsString;
      aProductStoreDS.FieldByName( 'SALEBEGINTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'SALEBEGINTIME' ).AsDateTime );
      aProductStoreDS.FieldByName( 'SALEENDTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'SALEENDTIME' ).AsDateTime );
      aProductStoreDS.FieldByName( 'VALIDITYPERIODBEGINTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'VALIDITYPERIODBEGINTIME' ).AsDateTime );
      aProductStoreDS.FieldByName( 'VALIDITYPERIODENDTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'VALIDITYPERIODENDTIME' ).AsDateTime );
      aProductStoreDS.FieldByName( 'PRODUCTCATEGORY' ).AsString :=
        FDataReader.FieldByName( 'PRODUCTCATEGORY' ).AsString;
      aProductStoreDS.FieldByName( 'EPGPRICE' ).AsString :=
        FDataReader.FieldByName( 'EPGPRICE' ).AsString;
      aProductStoreDS.FieldByName( 'SUBALLOWPPVEVENTFLAG' ).AsString :=
        FDataReader.FieldByName( 'SUBALLOWPPVEVENTFLAG' ).AsString;
      aProductStoreDS.FieldByName( 'CHANNELID' ).AsString :=
        FDataReader.FieldByName( 'CHANNELID' ).AsString;
      aProductStoreDS.FieldByName( 'CHANNELNO' ).AsString :=
        FDataReader.FieldByName( 'CHANNELNO' ).AsString;
      aProductStoreDS.FieldByName( 'EVENTBEGINTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'EVENTBEGINTIME' ).AsDateTime );
      aProductStoreDS.FieldByName( 'EVENTENDTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'EVENTENDTIME' ).AsDateTime );
      aProductStoreDS.FieldByName( 'DURATION' ).AsString :=
        FDataReader.FieldByName( 'DURATION' ).AsString;
      aProductStoreDS.FieldByName( 'EVENTID' ).AsString :=
        FDataReader.FieldByName( 'EVENTID' ).AsString;
      aProductStoreDS.FieldByName( 'DPARENTALRATING' ).AsString :=
        FDataReader.FieldByName( 'DPARENTALRATING' ).AsString;
      aProductStoreDS.FieldByName( 'DPARENTALRATINGNAME' ).AsString :=
        FDataReader.FieldByName( 'DPARENTALRATINGNAME' ).AsString;
      aProductStoreDS.FieldByName( 'DSTOPFLAG' ).AsString :=
        FDataReader.FieldByName( 'DSTOPFLAG' ).AsString;
      aProductStoreDS.Post;
      FDataReader.Next;
    end;
    FDataReader.Close;
    Sleep( 100 );
    FDataReader.SQL.Text := aEventSql;
    FDataReader.Open;
    FDataReader.First;
    while not FDataReader.Eof do
    begin
      aEventStoreDS.Append;
      aEventStoreDS.FieldByName( 'PRODUCTID' ).AsString :=
        FDataReader.FieldByName( 'PRODUCTID' ).AsString;
      aEventStoreDS.FieldByName( 'EVENTID' ).AsString :=
        FDataReader.FieldByName( 'EVENTID' ).AsString;
      aEventStoreDS.FieldByName( 'CHANNELID' ).AsString :=
        FDataReader.FieldByName( 'CHANNELID' ).AsString;
      aEventStoreDS.FieldByName( 'CHANNELNO' ).AsString :=
        FDataReader.FieldByName( 'CHANNELNO' ).AsString;
      aEventStoreDS.FieldByName( 'EVENTBEGINTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'EVENTBEGINTIME' ).AsDateTime );
      aEventStoreDS.FieldByName( 'EVENTENDTIME' ).AsString :=
        FormatDateTime( 'yyyymmddhhnnss',
          FDataReader.FieldByName( 'EVENTENDTIME' ).AsDateTime );
      aEventStoreDS.FieldByName( 'EVENTTYPE' ).AsString :=
        FDataReader.FieldByName( 'EVENTTYPE' ).AsString;
      aEventStoreDS.FieldByName( 'FILMNO' ).AsString :=
        FDataReader.FieldByName( 'FILMNO' ).AsString;
      aEventStoreDS.FieldByName( 'EVENTPREVIEWTIME' ).AsString :=
        FDataReader.FieldByName( 'EVENTPREVIEWTIME' ).AsString;
      aEventStoreDS.FieldByName( 'ESTOPFLAG' ).AsString :=
        FDataReader.FieldByName( 'ESTOPFLAG' ).AsString;
      aEventStoreDS.FieldByName( 'NAME' ).AsString :=
        FDataReader.FieldByName( 'NAME' ).AsString;
      aEventStoreDS.Post;
      FDataReader.Next;
    end;
    FDataReader.Close;
    Sleep( 100 );
    Result := True;
  except
    on E: Exception do aMsg := Format( '%s - %s', ['SO155-SO155A', E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoSo162(var aMsg: String): Boolean;
var
  aSql: String;
begin
  Result := False;
  FDataConnection.BeginTrans;
  try
    FDataWriter.Close;
    FDataWriter.SQL.Text := ' update so162 set kind = 0 ';
    FDataWriter.ExecSQL;
    aSql := 'update so162 set kind = 1 where channelid = ''%s'' ';
    FDataReader.Close;
    FDataReader.SQL.Text :=
      ' select distinct channelid from so155 where channelid is not null ';
    FDataReader.Open;
    FDataReader.First;
    while not FDataReader.Eof do
    begin
      FDataWriter.SQL.Text := Format( aSql, [
        FDataReader.FieldByName( 'channelid' ).AsString] );
      FDataWriter.ExecSQL;
      FDataReader.Next;
      Sleep( 10 );
    end;
    FDataReader.Close;
    FDataConnection.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      if FDataConnection.InTransaction then FDataConnection.RollbackTrans;
      aMsg := Format( '%s - %s', ['SO162', E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoMergin(var aMsg: String): Boolean;

  { ----------------------------------------------------- }

  function CalcEventEndTime(aEventBegin: String; aDuration: Integer): String;
  var
    aDate: TDateTime;
  begin
    aDate := StrToDateTime( FormatMaskText( '####/##/##;0;_',
      Copy( aEventBegin, 1, 8 ) ) + ' ' + FormatMaskText( '##:##:##;0;-',
      Copy( aEventBegin, 9, 6 ) ) );
    aDate := ( aDate + ( aDuration / 60 / 60 / 24 ) );
    Result := FormatDateTime( 'yyymmddhhnnss', aDate );
  end;

  { ----------------------------------------------------- }

var
  aPreProductId, aEventId, aChannelId, aChannelNo: String;
  aEventBeginTime, aEventEndTime, aEventEndTime2: String;
  aProductCount, aMultiFlag, aDuration, aRecordEffect, aDeleteMark: Integer;
begin
 Result := False;
 FDataConnection.BeginTrans;
 try
   aProductCount := 0;
   aDuration := 0;
   aEventId := EmptyStr;
   aMultiFlag := 0;
   FSo160.First;
   while not FSo160.Eof do
   begin
     if ( aPreProductId <> FSo160.FieldByName( 'ProductId' ).AsString ) then
     begin
       aDeleteMark := FSo160.FieldByName( 'dStopFlag' ).AsInteger;
       if aDeleteMark = 0 then
       begin
         if ( FSo160.FieldByName( 'EventId' ).AsString = EmptyStr ) then
           aDeleteMark := 1;
       end;
       { 更新此次的 Product }
       FDataWriter.SQL.Text := Format(
         ' UPDATE SO155                            ' +
         '    SET ImpulsiveFlag = ''%s'',          ' +
         '        ProductKind = ''%s'',            ' +
         '        SpecialPpv = ''%s'',             ' +
         '        PurchasableFlag = %s,            ' +
         '        SoldFlag = ''%s'',               ' +
         '        SuspendedFlag = ''%s'',          ' +
         '        ProductName = ''%s'',            ' +
         '        ProductDescription = ''%s'',     ' +
         '        EngProductName = ''%s'',         ' +
         '        EngProductDescription = ''%s'',  ' +
         '        SaleBeginTime =  to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),            ' +
         '        SaleEndTime = to_date( ''%s'',  ''YYYY/MM/DD HH24:MI:SS'' ),              ' +
         '        ValidityPeriodBeginTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),   ' +
         '        ValidityPeriodEndTime = to_date( ''%s'',  ''YYYY/MM/DD HH24:MI:SS'' ),    ' +
         '        ProductCategory = %s,            ' +
         '        EpgPrice = ''%s'',               ' +
         '        SubAllowPpvEventFlag = ''%s'',   ' +
         '        UpdTime = ''%s'',                ' +
         '        UpdEn = ''%s'',                  ' +
         '        Pric e = ''%s'',                  ' +
         '        dParentalRating = ''%s'',        ' +
         '        dParentalRatingName = ''%s'',    ' +
         '        dStopFlag = %d                   ' +
         ' WHERE ProductId = ''%s''                ',
         [FSo160.FieldByName( 'ImpulsiveFlag' ).AsString,
          FSo160.FieldByName( 'ProductKind' ).AsString,
          FSo160.FieldByName( 'SpecialPpv' ).AsString,
          Nvl( FSo160.FieldByName( 'PurchasableFlag' ).AsString, 'Null' ),
          FSo160.FieldByName( 'SoldFlag' ).AsString,
          FSo160.FieldByName( 'SuspendedFlag' ).AsString,
          FSo160.FieldByName( 'ProductName' ).AsString,
          FSo160.FieldByName( 'ProductDescription' ).AsString,
          FSo160.FieldByName( 'EngProductName' ).AsString,
          FSo160.FieldByName( 'EngProductDescription' ).AsString,
          UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleBeginTime' ).AsString ) ),
          UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleEndTime' ).AsString ) ),
          UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodBeginTime' ).AsString ) ),
          UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodEndTime' ).AsString ) ),
          Nvl( FSo160.FieldByName( 'ProductCategory' ).AsString, 'Null' ),
          FSo160.FieldByName( 'EpgPrice' ).AsString,
          FSo160.FieldByName( 'SubAllowPpvEventFlag' ).AsString,
          FSo160.FieldByName( 'UpdTime' ).AsString,
          FSo160.FieldByName( 'UpdEn' ).AsString,
          FSo160.FieldByName( 'EpgPrice' ).AsString,
          FSo160.FieldByName( 'dParentalRating' ).AsString,
          FSo160.FieldByName( 'dParentalRatingName' ).AsString,
          aDeleteMark,
          FSo160.FieldByName( 'ProductId' ).AsString] );
       aRecordEffect := FDataWriter.ExecSQL;
       if ( aRecordEffect <= 0 ) then
       begin
         FDataWriter.SQL.Text := Format(
           ' INSERT INTO SO155 ( ProductId, ImpulsiveFlag,             ' +
           '   ProductKind,  SpecialPpv, PurchasableFlag, SoldFlag,    ' +
           '   SuspendedFlag, ProductName, ProductDescription,         ' +
           '   EngProductName, EngProductDescription,                  ' +
           '   SaleBeginTime,                                          ' +
           '   SaleEndTime,                                            ' +
           '   ValidityPeriodBeginTime,                                ' +
           '   ValidityPeriodEndTime,                                  ' +
           '   ProductCategory, EpgPrice,                              ' +
           '   SubAllowPpvEventFlag, MultiEventFlag,                   ' +
           '   UpdTime, UpdEn, Price, dStopFlag, dParentalRating,      ' +
           '   dParentalRatingName )                                   ' +
           '  VALUES ( ''%s'', ''%s'',                                 ' +
           '   ''%s'', ''%s'', %s, ''%s'',                             ' +
           '   ''%s'', ''%s'', ''%s'',                                 ' +
           '   ''%s'', ''%s'',                                         ' +
           '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
           '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
           '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
           '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
           '    %s,  ''%s'',                                           ' +
           '   ''%s'', ''%s'',                                         ' +
           '   ''%s'', ''%s'', ''%s'', %d, ''%s'', ''%s''  )           ',
           [FSo160.FieldByName( 'ProductId' ).AsString,
            FSo160.FieldByName( 'ImpulsiveFlag' ).AsString,
            FSo160.FieldByName( 'ProductKind' ).AsString,
            FSo160.FieldByName( 'SpecialPpv' ).AsString,
            Nvl( FSo160.FieldByName( 'PurchasableFlag' ).AsString, 'Null' ),
            FSo160.FieldByName( 'SoldFlag' ).AsString,
            FSo160.FieldByName( 'SuspendedFlag' ).AsString,
            FSo160.FieldByName( 'ProductName' ).AsString,
            FSo160.FieldByName( 'ProductDescription' ).AsString,
            FSo160.FieldByName( 'EngProductName' ).AsString,
            FSo160.FieldByName( 'EngProductDescription' ).AsString,
            UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleBeginTime' ).AsString ) ),
            UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleEndTime' ).AsString ) ),
            UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodBeginTime' ).AsString ) ),
            UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodEndTime' ).AsString ) ),
            Nvl( FSo160.FieldByName( 'ProductCategory' ).AsString, 'Null' ),
            FSo160.FieldByName( 'EpgPrice' ).AsString,
            FSo160.FieldByName( 'SubAllowPpvEventFlag' ).AsString,
            '0',
            FSo160.FieldByName( 'UpdTime' ).AsString,
            FSo160.FieldByName( 'UpdEn' ).AsString,
            Nvl( FSo160.FieldByName( 'EpgPrice' ).AsString, '0' ),
            aDeleteMark,
            FSo160.FieldByName( 'dParentalRating' ).AsString,
            FSo160.FieldByName( 'dParentalRatingName' ).AsString] );
         FDataWriter.ExecSQL;
       end;
     end;
     {}
     Inc( aProductCount );
     if ( aProductCount > 1 ) then aMultiFlag := 1;
     if ( FSo160.FieldByName( 'EventId' ).AsString <> EmptyStr ) then
     begin
       { 當為第一個該 Product 下的 Event }
       if ( aEventId = EmptyStr ) and
          ( FSo160.FieldByName( 'eStopFlag' ).AsInteger <> 1 ) and
          ( FSo160.FieldByName( 'eeStopFlag' ).AsInteger <> 1 ) then
       begin
         { 直接取第一個 Event 的資料 }
         aEventId := FSo160.FieldByName( 'EventId' ).AsString;
         aEventBeginTime := FSo160.FieldByName( 'EventBeginTime' ).AsString;
         aChannelId := FSo160.FieldByName( 'ChannelId' ).AsString;
         aChannelNo := FSo160.FieldByName( 'ChannelNo' ).AsString;
         aDuration := FSo160.FieldByName( 'Duration' ).AsInteger;
       end;
       if ( FSo160.FieldByName( 'eStopFlag' ).AsInteger = 1 ) or
          ( FSo160.FieldByName( 'eeStopFlag' ).AsInteger = 1 ) then
       begin
         FDataWriter.SQL.Text := Format(
           ' UPDATE SO155A SET eStopFlag = 1    ' +
           '  WHERE EventId = ''%s''            ' +
           '    AND eStopFlag = 0               ',
           [FSo160.FieldByName( 'EventId' ).AsString] );
         FDataWriter.ExecSQL;
       end else
       begin
         aEventEndTime2 := CalcEventEndTime(
           FSo160.FieldByName( 'EventBeginTime' ).AsString,
           FSo160.FieldByName( 'Duration' ).AsInteger );
         FDataWriter.SQL.Text := Format(
           ' UPDATE SO155A SET ChannelId = ''%s'',              ' +
           '   ChannelNo = ''%s'',                              ' +
           '   EventBeginTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
           '   EventEndTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
           '   EventType = ''%s'',                              ' +
           '   FilmNo = ''%s'',                                 ' +
           '   EventPreviewTime = ''%s'',                       ' +
           '   UpdTime = ''%s'',                                ' +
           '   UpdEn = ''%s''                                   ' +
           ' WHERE ProductId = ''%s''                           ' +
           '   AND EventId = ''%s''                             ',
           [FSo160.FieldByName( 'ChannelId' ).AsString,
            FSo160.FieldByName( 'ChannelNo' ).AsString,
            UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'EventBeginTime' ).AsString ) ),
            UTCDateTimeToLocalTime( FormatDateTimeString( aEventEndTime2 ) ),
            FSo160.FieldByName( 'EventType' ).AsString,
            FSo160.FieldByName( 'FilmNo' ).AsString,
            Nvl( FSo160.FieldByName( 'EventPreviewTime' ).AsString, '0' ),
            FSo160.FieldByName( 'UpdTime' ).AsString,
            FSo160.FieldByName( 'UpdEn' ).AsString,
            FSo160.FieldByName( 'ProductId' ).AsString,
            FSo160.FieldByName( 'EventId' ).AsString] );
         aRecordEffect := FDataWriter.ExecSQL;
         if ( aRecordEffect <= 0 ) then
         begin
           FDataWriter.SQL.Text := Format(
             ' INSERT INTO SO155A ( ProductId, EventId, ChannelId,      ' +
             '   ChannelNo, EventBeginTime, EventEndTime, EventType,    ' +
             '   FilmNo, EventPreviewTime, UpdTime, UpdEn, eStopFlag )  ' +
             ' VALUES ( ''%s'', ''%s'', ''%s'',                         ' +
             '   ''%s'',                                                ' +
             '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),          ' +
             '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),          ' +
             '   ''%s'',                                                ' +
             '   ''%s'', ''%s'', ''%s'', ''%s'', 0  )                   ',
             [FSo160.FieldByName( 'ProductId' ).AsString,
              FSo160.FieldByName( 'EventId' ).AsString,
              FSo160.FieldByName( 'ChannelId' ).AsString,
              FSo160.FieldByName( 'ChannelNo' ).AsString,
              UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'EventBeginTime' ).AsString ) ),
              UTCDateTimeToLocalTime( FormatDateTimeString( aEventEndTime2 ) ),
              FSo160.FieldByName( 'EventType' ).AsString,
              FSo160.FieldByName( 'FilmNo' ).AsString,
              FSo160.FieldByName( 'EventPreviewTime' ).AsString,
              FSo160.FieldByName( 'UpdTime' ).AsString,
              FSo160.FieldByName( 'UpdEn' ).AsString] );
           FDataWriter.ExecSQL;
         end;
       end;
       FDataWriter.SQL.Text := Format(
         ' update so159 set action = 0 where eventid = ''%s'' ',
         [ FSo160.FieldByName( 'eventid' ).AsString] );
       FDataWriter.ExecSQL;
     end;  
     {}
     aPreProductId := FSo160.FieldByName( 'ProductId' ).AsString;
     FSo160.Next;
     Sleep( 10 );
     {}
     { 更新上一個 Product 的 EventBeginTime 跟 EventEndTime 及Duration  }
     if ( aPreProductId <> FSo160.FieldByName( 'ProductId' ).AsString ) or
        ( FSo160.Eof ) then
     begin
       { 該 Product 下的 Event 都被 Mark 成刪除了 ?? 奇怪真的發生了 ?? }
       if ( aEventId = EmptyStr ) then
       begin
         FDataWriter.SQL.Text := Format(
           ' UPDATE SO155 SET dStopFlag = 1    ' +
           '  WHERE ProductId = ''%s''         ' +
           '    AND dStopFlag = 0              ', [aPreProductId] );
         FDataWriter.ExecSQL;
         {}
         FDataWriter.SQL.Text := Format(
           ' UPDATE SO158 SET dStopFlag = 1    ' +
           '  WHERE ProductId = ''%s''         ' +
           '    AND dStopFlag = 0              ', [aPreProductId] );
         FDataWriter.ExecSQL;
       end else
       begin
         aEventEndTime := CalcEventEndTime( aEventBeginTime, aDuration );
         FDataWriter.SQL.Text := Format(
           ' UPDATE SO155                            ' +
           '    SET MultiEventFlag = ''%d'',         ' +
           '        EventId = ''%s'',                ' +
           '        EventBeginTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
           '        EventEndTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),    ' +
           '        ChannelId = ''%s'',              ' +
           '        ChannelNo = ''%s'',              ' +
           '        Duration = ''%d''                ' +
           '  WHERE ProductId = ''%s''               ',
           [aMultiFlag, aEventId,
            UTCDateTimeToLocalTime( FormatDateTimeString( aEventBeginTime ) ),
            UTCDateTimeToLocalTime( FormatDateTimeString( aEventEndTime ) ),
            aChannelId, aChannelNo, aDuration, aPreProductId] );
         FDataWriter.ExecSQL;
       end;
       {}
       FDataWriter.SQL.Text := Format(
         ' update so158 set action = 0 where productid = ''%s'' ', [aPreProductId] );
       FDataWriter.ExecSQL;
       { 重置 }
       aProductCount := 0;
       aMultiFlag := 0;
       aEventId := EmptyStr;
       aEventBeginTime := EmptyStr;
       aEventEndTime := EmptyStr;
       aChannelId := EmptyStr;
       aChannelNo := EmptyStr;
       aDuration := 0;
       {}
     end;
   end;
   FDataConnection.CommitTrans;
   Result := True;
  except
    on E: Exception do
    begin
      if FDataConnection.InTransaction then FDataConnection.RollbackTrans;
      aMsg := Format( '%s - %s', ['SO155', E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoMerginEx(var aMsg: String): Boolean;
var
  aPreProductId, aEventId, aChannelId, aChannelNo: String;
  aEventBeginTime, aEventEndTime, aEventEndTime2: String;
  aProductCount, aMultiFlag, aDuration, aDeleteMark: Integer;
  
  { ----------------------------------------------------- }

  function CalcEventEndTime(aEventBegin: String; aDuration: Integer): String;
  var
    aDate: TDateTime;
  begin
    aDate := StrToDateTime( FormatMaskText( '####/##/##;0;_',
      Copy( aEventBegin, 1, 8 ) ) + ' ' + FormatMaskText( '##:##:##;0;-',
      Copy( aEventBegin, 9, 6 ) ) );
    aDate := ( aDate + ( aDuration / 60 / 60 / 24 ) );
    Result := FormatDateTime( 'yyymmddhhnnss', aDate );
  end;

  { ----------------------------------------------------- }

  function ProductRecordFound(const aProdutcId: String): Boolean;
  begin
    Result := False;
    if FProductBuffer.IsEmpty then Exit;
    Result := FProductBuffer.Locate( 'PRODUCTID', aProdutcId, [] );
  end;

  { ----------------------------------------------------- }

  function EventRecordFound(const aProdutcId, aEventId: String): Boolean;
  begin
    Result := False;
    if FEventBuffer.IsEmpty then Exit;
    Result := FEventBuffer.Locate( 'PRODUCTID;EVENTID', VarArrayOf(
      [aProdutcId, aEventId ] ), [] );
  end;

  { ----------------------------------------------------- }

  function UpdateSo155: Integer;
  begin
    FDataWriter.SQL.Text := Format(
      ' UPDATE SO155                            ' +
      '    SET ImpulsiveFlag = ''%s'',          ' +
      '        ProductKind = ''%s'',            ' +
      '        SpecialPpv = ''%s'',             ' +
      '        PurchasableFlag = %s,            ' +
      '        SoldFlag = ''%s'',               ' +
      '        SuspendedFlag = ''%s'',          ' +
      '        ProductName = ''%s'',            ' +
      '        ProductDescription = ''%s'',     ' +
      '        EngProductName = ''%s'',         ' +
      '        EngProductDescription = ''%s'',  ' +
      '        SaleBeginTime =  to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),            ' +
      '        SaleEndTime = to_date( ''%s'',  ''YYYY/MM/DD HH24:MI:SS'' ),              ' +
      '        ValidityPeriodBeginTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),   ' +
      '        ValidityPeriodEndTime = to_date( ''%s'',  ''YYYY/MM/DD HH24:MI:SS'' ),    ' +
      '        ProductCategory = %s,            ' +
      '        EpgPrice = ''%s'',               ' +
      '        SubAllowPpvEventFlag = ''%s'',   ' +
      '        UpdTime = ''%s'',                ' +
      '        UpdEn = ''%s'',                  ' +
      '        Price = ''%s'',                  ' +
      '        dParentalRating = ''%s'',        ' +
      '        dParentalRatingName = ''%s'',    ' +
      '        dStopFlag = %d                   ' +
      ' WHERE ProductId = ''%s''                ',
      [FSo160.FieldByName( 'ImpulsiveFlag' ).AsString,
       FSo160.FieldByName( 'ProductKind' ).AsString,
       FSo160.FieldByName( 'SpecialPpv' ).AsString,
       Nvl( FSo160.FieldByName( 'PurchasableFlag' ).AsString, 'Null' ),
       FSo160.FieldByName( 'SoldFlag' ).AsString,
       FSo160.FieldByName( 'SuspendedFlag' ).AsString,
       FSo160.FieldByName( 'ProductName' ).AsString,
       FSo160.FieldByName( 'ProductDescription' ).AsString,
       FSo160.FieldByName( 'EngProductName' ).AsString,
       FSo160.FieldByName( 'EngProductDescription' ).AsString,
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleBeginTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleEndTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodBeginTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodEndTime' ).AsString ) ),
       FormatDateTimeString( FSo160.FieldByName( 'SaleBeginTime' ).AsString ),
       FormatDateTimeString( FSo160.FieldByName( 'SaleEndTime' ).AsString ),
       FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodBeginTime' ).AsString ),
       FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodEndTime' ).AsString ),
       Nvl( FSo160.FieldByName( 'ProductCategory' ).AsString, 'Null' ),
       FSo160.FieldByName( 'EpgPrice' ).AsString,
       FSo160.FieldByName( 'SubAllowPpvEventFlag' ).AsString,
       FSo160.FieldByName( 'UpdTime' ).AsString,
       FSo160.FieldByName( 'UpdEn' ).AsString,
       FSo160.FieldByName( 'EpgPrice' ).AsString,
       FSo160.FieldByName( 'dParentalRating' ).AsString,
       FSo160.FieldByName( 'dParentalRatingName' ).AsString,
       aDeleteMark,
       FSo160.FieldByName( 'ProductId' ).AsString] );
    Result := FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  function InsertSo155: Integer;
  begin
    FDataWriter.SQL.Text := Format(
      ' INSERT INTO SO155 ( ProductId, ImpulsiveFlag,             ' +
      '   ProductKind,  SpecialPpv, PurchasableFlag, SoldFlag,    ' +
      '   SuspendedFlag, ProductName, ProductDescription,         ' +
      '   EngProductName, EngProductDescription,                  ' +
      '   SaleBeginTime,                                          ' +
      '   SaleEndTime,                                            ' +
      '   ValidityPeriodBeginTime,                                ' +
      '   ValidityPeriodEndTime,                                  ' +
      '   ProductCategory, EpgPrice,                              ' +
      '   SubAllowPpvEventFlag, MultiEventFlag,                   ' +
      '   UpdTime, UpdEn, Price, dStopFlag, dParentalRating,      ' +
      '   dParentalRatingName )                                   ' +
      '  VALUES ( ''%s'', ''%s'',                                 ' +
      '   ''%s'', ''%s'', %s, ''%s'',                             ' +
      '   ''%s'', ''%s'', ''%s'',                                 ' +
      '   ''%s'', ''%s'',                                         ' +
      '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
      '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
      '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
      '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),           ' +
      '    %s,  ''%s'',                                           ' +
      '   ''%s'', ''%s'',                                         ' +
      '   ''%s'', ''%s'', ''%s'', %d, ''%s'', ''%s''  )           ',
      [FSo160.FieldByName( 'ProductId' ).AsString,
       FSo160.FieldByName( 'ImpulsiveFlag' ).AsString,
       FSo160.FieldByName( 'ProductKind' ).AsString,
       FSo160.FieldByName( 'SpecialPpv' ).AsString,
       Nvl( FSo160.FieldByName( 'PurchasableFlag' ).AsString, 'Null' ),
       FSo160.FieldByName( 'SoldFlag' ).AsString,
       FSo160.FieldByName( 'SuspendedFlag' ).AsString,
       FSo160.FieldByName( 'ProductName' ).AsString,
       FSo160.FieldByName( 'ProductDescription' ).AsString,
       FSo160.FieldByName( 'EngProductName' ).AsString,
       FSo160.FieldByName( 'EngProductDescription' ).AsString,
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleBeginTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'SaleEndTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodBeginTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodEndTime' ).AsString ) ),
       FormatDateTimeString( FSo160.FieldByName( 'SaleBeginTime' ).AsString ),
       FormatDateTimeString( FSo160.FieldByName( 'SaleEndTime' ).AsString ),
       FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodBeginTime' ).AsString ),
       FormatDateTimeString( FSo160.FieldByName( 'ValidityPeriodEndTime' ).AsString ),
       Nvl( FSo160.FieldByName( 'ProductCategory' ).AsString, 'Null' ),
       FSo160.FieldByName( 'EpgPrice' ).AsString,
       FSo160.FieldByName( 'SubAllowPpvEventFlag' ).AsString,
       '0',
       FSo160.FieldByName( 'UpdTime' ).AsString,
       FSo160.FieldByName( 'UpdEn' ).AsString,
       Nvl( FSo160.FieldByName( 'EpgPrice' ).AsString, '0' ),
       aDeleteMark,
       FSo160.FieldByName( 'dParentalRating' ).AsString,
       FSo160.FieldByName( 'dParentalRatingName' ).AsString] );
    Result := FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  function UpdateSo155A: Integer;
  begin
    FDataWriter.SQL.Text := Format(
      ' UPDATE SO155A SET ChannelId = ''%s'',              ' +
      '   ChannelNo = ''%s'',                              ' +
      '   EventBeginTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '   EventEndTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '   EventType = ''%s'',                              ' +
      '   FilmNo = ''%s'',                                 ' +
      '   EventPreviewTime = ''%s'',                       ' +
      '   UpdTime = ''%s'',                                ' +
      '   UpdEn = ''%s''                                   ' +
      ' WHERE ProductId = ''%s''                           ' +
      '   AND EventId = ''%s''                             ',
      [FSo160.FieldByName( 'ChannelId' ).AsString,
       FSo160.FieldByName( 'ChannelNo' ).AsString,
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'EventBeginTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( aEventEndTime2 ) ),
       FormatDateTimeString( FSo160.FieldByName( 'EventBeginTime' ).AsString ),
       FormatDateTimeString( aEventEndTime2 ),
       FSo160.FieldByName( 'EventType' ).AsString,
       FSo160.FieldByName( 'FilmNo' ).AsString,
       Nvl( FSo160.FieldByName( 'EventPreviewTime' ).AsString, '0' ),
       FSo160.FieldByName( 'UpdTime' ).AsString,
       FSo160.FieldByName( 'UpdEn' ).AsString,
       FSo160.FieldByName( 'ProductId' ).AsString,
       FSo160.FieldByName( 'EventId' ).AsString] );
    Result := FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  function InsertSo155A: Integer;
  begin
    FDataWriter.SQL.Text := Format(
      ' INSERT INTO SO155A ( ProductId, EventId, ChannelId,      ' +
      '   ChannelNo, EventBeginTime, EventEndTime, EventType,    ' +
      '   FilmNo, EventPreviewTime, UpdTime, UpdEn, eStopFlag )  ' +
      ' VALUES ( ''%s'', ''%s'', ''%s'',                         ' +
      '   ''%s'',                                                ' +
      '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),          ' +
      '   to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),          ' +
      '   ''%s'',                                                ' +
      '   ''%s'', ''%s'', ''%s'', ''%s'', 0  )                   ',
      [FSo160.FieldByName( 'ProductId' ).AsString,
       FSo160.FieldByName( 'EventId' ).AsString,
       FSo160.FieldByName( 'ChannelId' ).AsString,
       FSo160.FieldByName( 'ChannelNo' ).AsString,
//       UTCDateTimeToLocalTime( FormatDateTimeString( FSo160.FieldByName( 'EventBeginTime' ).AsString ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( aEventEndTime2 ) ),
       FormatDateTimeString( FSo160.FieldByName( 'EventBeginTime' ).AsString ),
       FormatDateTimeString( aEventEndTime2 ),
       FSo160.FieldByName( 'EventType' ).AsString,
       FSo160.FieldByName( 'FilmNo' ).AsString,
       FSo160.FieldByName( 'EventPreviewTime' ).AsString,
       FSo160.FieldByName( 'UpdTime' ).AsString,
       FSo160.FieldByName( 'UpdEn' ).AsString] );
    Result := FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  procedure SetProductStop(const aProductId: String);
  begin
    FDataWriter.SQL.Text := Format(
      ' UPDATE SO155 SET dStopFlag = 1    ' +
      '  WHERE ProductId = ''%s''         ' +
      '    AND dStopFlag = 0              ', [aProductId] );
    FDataWriter.ExecSQL;
    {}
    FDataWriter.SQL.Text := Format(
      ' UPDATE SO158 SET dStopFlag = 1    ' +
      '  WHERE ProductId = ''%s''         ' +
      '    AND dStopFlag = 0              ', [aProductId] );
    FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  procedure SetEventStop(const aEventId: String);
  begin
    FDataWriter.SQL.Text := Format(
      ' UPDATE SO155A SET eStopFlag = 1    ' +
      '  WHERE EventId = ''%s''            ' +
      '    AND eStopFlag = 0               ', [aEventId] );
    FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  procedure ResetProductChangeFlag(const aProductId: String);
  begin
    FDataWriter.SQL.Text := Format(
      ' update so158 set action = 0 where productid = ''%s'' ', [aProductId] );
    FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  procedure ResetEventChangeFlag(const aEventId: String);
  begin
    FDataWriter.SQL.Text := Format(
      ' update so159 set action = 0 where eventid = ''%s'' ', [aEventId] );
    FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

  procedure SetSo155;
  begin
    FDataWriter.SQL.Text := Format(
      ' UPDATE SO155                            ' +
      '    SET MultiEventFlag = ''%d'',         ' +
      '        EventId = ''%s'',                ' +
      '        EventBeginTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),  ' +
      '        EventEndTime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),    ' +
      '        ChannelId = ''%s'',              ' +
      '        ChannelNo = ''%s'',              ' +
      '        Duration = ''%d''                ' +
      '  WHERE ProductId = ''%s''               ',
      [aMultiFlag, aEventId,
//       UTCDateTimeToLocalTime( FormatDateTimeString( aEventBeginTime ) ),
//       UTCDateTimeToLocalTime( FormatDateTimeString( aEventEndTime ) ),
       FormatDateTimeString( aEventBeginTime ),
       FormatDateTimeString( aEventEndTime ),
       aChannelId, aChannelNo, aDuration, aPreProductId] );
    FDataWriter.ExecSQL;
  end;

  { ----------------------------------------------------- }

begin
 Result := False;
 FDataConnection.BeginTrans;
 try
   aProductCount := 0;
   aDuration := 0;
   aEventId := EmptyStr;
   aMultiFlag := 0;
   FSo160.First;
   while not FSo160.Eof do
   begin
     if ( aPreProductId <> FSo160.FieldByName( 'ProductId' ).AsString ) then
     begin
       aDeleteMark := FSo160.FieldByName( 'dStopFlag' ).AsInteger;
       if aDeleteMark = 0 then
       begin
         if ( FSo160.FieldByName( 'EventId' ).AsString = EmptyStr ) then
           aDeleteMark := 1;
       end;
       { 更新此次的 Product }
       if ProductRecordFound( FSo160.FieldByName( 'ProductId' ).AsString ) then
         UpdateSo155
       else
         InsertSo155;
     end;
     {}
     Inc( aProductCount );
     if ( aProductCount > 1 ) then aMultiFlag := 1;
     if ( FSo160.FieldByName( 'EventId' ).AsString <> EmptyStr ) then
     begin
       { 當為第一個該 Product 下的 Event }
       if ( aEventId = EmptyStr ) and
          ( FSo160.FieldByName( 'eStopFlag' ).AsInteger <> 1 ) and
          ( FSo160.FieldByName( 'eeStopFlag' ).AsInteger <> 1 ) then
       begin
         { 直接取第一個 Event 的資料 }
         aEventId := FSo160.FieldByName( 'EventId' ).AsString;
         aEventBeginTime := FSo160.FieldByName( 'EventBeginTime' ).AsString;
         aChannelId := FSo160.FieldByName( 'ChannelId' ).AsString;
         aChannelNo := FSo160.FieldByName( 'ChannelNo' ).AsString;
         aDuration := FSo160.FieldByName( 'Duration' ).AsInteger;
       end;
       { 判斷該 Event 是否被停用 }
       if ( FSo160.FieldByName( 'eStopFlag' ).AsInteger = 1 ) or
          ( FSo160.FieldByName( 'eeStopFlag' ).AsInteger = 1 ) then
       begin
         SetEventStop( FSo160.FieldByName( 'EventId' ).AsString );
       end else
       begin
         { 計算 Event 的結束時間 }
         aEventEndTime2 := CalcEventEndTime(
           FSo160.FieldByName( 'EventBeginTime' ).AsString,
           FSo160.FieldByName( 'Duration' ).AsInteger );
         {}  
         if EventRecordFound( FSo160.FieldByName( 'ProductId' ).AsString,
           FSo160.FieldByName( 'EventId' ).AsString ) then
           UpdateSo155A
         else
           InsertSo155A;
       end;
       { 重置 event 的 action flag }
       ResetEventChangeFlag( FSo160.FieldByName( 'EventId' ).AsString );
     end;
     {}
     aPreProductId := FSo160.FieldByName( 'ProductId' ).AsString;
     FSo160.Next;
     Sleep( 10 );
     {}
     { 更新上一個 Product 的 EventBeginTime 跟 EventEndTime 及Duration  }
     if ( aPreProductId <> FSo160.FieldByName( 'ProductId' ).AsString ) or
        ( FSo160.Eof ) then
     begin
       { 該 Product 下的 Event 都被 Mark 成刪除了 ?? 奇怪真的發生了 ?? }
       if ( aEventId = EmptyStr ) then
       begin
         SetProductStop( aPreProductId );
       end else
       begin
         aEventEndTime := CalcEventEndTime( aEventBeginTime, aDuration );
         { 回頭更新 so155 }
         SetSo155;
       end;
       { 重置 product 的 action flag }
       ResetProductChangeFlag( aPreProductId );
       { 重置 }
       aProductCount := 0;
       aMultiFlag := 0;
       aEventId := EmptyStr;
       aEventBeginTime := EmptyStr;
       aEventEndTime := EmptyStr;
       aChannelId := EmptyStr;
       aChannelNo := EmptyStr;
       aDuration := 0;
       {}
     end;
   end;
   FDataConnection.CommitTrans;
   Result := True;
  except
    on E: Exception do
    begin
      if FDataConnection.InTransaction then FDataConnection.RollbackTrans;
      aMsg := Format( '%s - %s', ['SO155', E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoSo154(var aMsg: String): Boolean;

  { ------------------------------------------ }

  function GetNewFilmNo: String;
  begin
    Result := EmptyStr;
    FDataReader.SQL.Text :=
      ' select to_char( sysdate, ''yyyymmdd'' ) || ' +
      '        lpad( to_char( s_so154_filmno.nextval ), 7, ''0'' ) as seq ' +
      ' from dual ';
    FDataReader.Open;
    Result := FDataReader.FieldByName( 'seq' ).AsString;
    FDataReader.Close;
  end;

  { ------------------------------------------ }

  function GetExistsFilmNo(const aFilmName: String): String;
  begin
    Result := EmptyStr;
    FDataReader.Close;
    FDataReader.SQL.Text := Format(
      ' select FilmNo from so154 where Name = ''%s'' ', [aFilmName] );
    FDataReader.Open;
    Result := FDataReader.FieldByName( 'FilmNo' ).AsString;
    FDataReader.Close;
  end;

  { ------------------------------------------ }

var
  aFlimNo: String;
begin
  Result := False;
  {}
  FDataInsert.SQL.Text :=
    ' INSERT INTO SO154 ( FilmNo, CreationDate, Name, EpgDesc,           ' +
    '    EngName, EngEpgDesc, ParentalRating, ParentalRatingName,        ' +
    '    Duration, UpdTime, UpdEn  )                                     ' +
    ' VALUES ( :FilmNo, :CreationDate, :Name, :EpgDesc,                  ' +
    '    :EngName, :EngEpgDesc, :ParentalRating, :ParentalRatingName,    ' +
    '    :Duration, :UpdTime, :UpdEn  )                                  ';
  {}
  FDataUpdate.SQL.Text :=
    ' UPDATE SO154 SET EpgDesc = :EpgDesc, EngName = :EngName,           ' +
    '    EngEpgDesc = :EngEpgDesc, ParentalRating = :ParentalRating,     ' +
    '    ParentalRatingName = :ParentalRatingName, Duration = :Duration, ' +
    '    UpdTime = :UpdTime, UpdEn = :UpdEn                              ' +
    ' WHERE FilmNo = :FilmNo                                             ';
  FDataConnection.BeginTrans;
  try
    FSo160.First;
    while not FSo160.Eof do
    begin
      if ( FSo160.FieldByName( 'EventId').AsString <> EmptyStr ) then
      begin
        aFlimNo := GetExistsFilmNo( FSo160.FieldByName( 'Name' ).AsString );
        if ( aFlimNo <> EmptyStr ) then
        begin
          FDataUpdate.Parameters.ParamByName( 'EpgDesc' ).Value :=
            FSo160.FieldByName( 'EpgDesc' ).AsString;
          FDataUpdate.Parameters.ParamByName( 'EngName' ).Value :=
            FSo160.FieldByName( 'EngName' ).AsString;
          FDataUpdate.Parameters.ParamByName( 'EngEpgDesc' ).Value :=
            FSo160.FieldByName( 'EngEpgDesc' ).AsString;
          FDataUpdate.Parameters.ParamByName( 'ParentalRating' ).Value :=
            FSo160.FieldByName( 'ParentalRating' ).AsString;
          FDataUpdate.Parameters.ParamByName( 'ParentalRatingName' ).Value :=
            FSo160.FieldByName( 'ParentalRatingName' ).AsString;;
          FDataUpdate.Parameters.ParamByName( 'Duration' ).Value :=
            FSo160.FieldByName( 'Duration' ).AsString;
          FDataUpdate.Parameters.ParamByName( 'UpdTime' ).Value :=
            FSo160.FieldByName( 'UpdTime' ).AsString;
          FDataUpdate.Parameters.ParamByName( 'UpdEn' ).Value := SMS_PROGRAM_SCHEDULAR;
          FDataUpdate.Parameters.ParamByName( 'FilmNo' ).Value := aFlimNo;
          FDataUpdate.ExecSQL;
        end else
        begin
          aFlimNo := GetNewFilmNo;
          FDataInsert.Parameters.ParamByName( 'FilmNo' ).Value := aFlimNo;
          FDataInsert.Parameters.ParamByName( 'CreationDate' ).Value :=
            StrToDate(
            Copy( FSo160.FieldByName( 'eCreationDate' ).AsString, 1, 4 ) + '/' +
            Copy( FSo160.FieldByName( 'eCreationDate' ).AsString, 5, 2 ) + '/' +
            Copy( FSo160.FieldByName( 'eCreationDate' ).AsString, 7, 2 ) );
          FDataInsert.Parameters.ParamByName( 'Name' ).Value :=
            FSo160.FieldByName( 'Name' ).AsString;
          FDataInsert.Parameters.ParamByName( 'EpgDesc' ).Value :=
            FSo160.FieldByName( 'EpgDesc' ).AsString;
          FDataInsert.Parameters.ParamByName( 'EngName' ).Value :=
            FSo160.FieldByName( 'EngName' ).AsString;
          FDataInsert.Parameters.ParamByName( 'EngEpgDesc' ).Value :=
            FSo160.FieldByName( 'EngEpgDesc' ).AsString;
          FDataInsert.Parameters.ParamByName( 'ParentalRating' ).Value :=
            FSo160.FieldByName( 'ParentalRating' ).AsString;
          FDataInsert.Parameters.ParamByName( 'ParentalRatingName' ).Value :=
            FSo160.FieldByName( 'ParentalRatingName' ).AsString;
          FDataInsert.Parameters.ParamByName( 'Duration' ).Value :=
            FSo160.FieldByName( 'Duration' ).AsString;
          FDataInsert.Parameters.ParamByName( 'UpdTime' ).Value :=
            FSo160.FieldByName( 'UpdTime' ).AsString;
          FDataInsert.Parameters.ParamByName( 'UpdEn' ).Value :=
            FSo160.FieldByName( 'UpdEn' ).AsString;
          FDataInsert.ExecSQL;
        end;
        FSo160.Edit;
        FSo160.FieldByName( 'FilmNo' ).AsString := aFlimNo;
        FSo160.Post;
      end;
      FSo160.Next;
      Sleep( 10 );
    end;  
    FDataConnection.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      if FDataConnection.InTransaction then FDataConnection.RollbackTrans;
      FDataReader.Close;
      aMsg := Format( '%s - %s', ['SO154', E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoSo160(var aMsg: String): Boolean;
var
  aPRCodeWrapper, aPRCodeDex: String;
  aPRNameWrapper, aPRNameDex: String; 
begin
  Result := False;
  try
    FDataDelete.Close;
    FDataDelete.SQL.Text := ' delete from so160 ';
    FDataDelete.ExecSQL;
    FDataWriter.Close;
    FDataWriter.SQL.Text := Format(
      '    insert into so160                                      ' +
      '      (  dfilename,                                        ' +
      '         dcreationdate,                                    ' +
      '         productid,                                        ' +
      '         impulsiveflag,                                    ' +
      '         productkind,                                      ' +
      '         specialppv,                                       ' +
      '         purchasableflag,                                  ' +
      '         soldflag,                                         ' +
      '         suspendedflag,                                    ' +
      '         productname,                                      ' +
      '         productdescription,                               ' +
      '         engproductname,                                   ' +
      '         engproductdescription,                            ' +
      '         salebegintime,                                    ' +
      '         saleendtime,                                      ' +
      '         validityperiodbegintime,                          ' +
      '         validityperiodendtime,                            ' +
      '         productcategory,                                  ' +
      '         epgprice,                                         ' +
      '         suballowppveventflag,                             ' +
      '         efilename,                                        ' +
      '         ecreationdate,                                    ' +
      '         channelid,                                        ' +
      '         eventbegintime,                                   ' +
      '         duration,                                         ' +
      '         eventid,                                          ' +
      '         eventpreviewtime,                                 ' +
      '         eventtype,                                        ' +
      '         name,                                             ' +
      '         epgdesc,                                          ' +
      '         engname,                                          ' +
      '         engepgdesc,                                       ' +
      '         parentalrating,                                   ' +
      '         contentnibble1,                                   ' +
      '         contentnibble2,                                   ' +
      '         usernibble1,                                      ' +
      '         usernibble2,                                      ' +
      '         channelno,                                        ' +
      '         updtime,                                          ' +
      '         upden,                                            ' +
      '         dparentalrating,                                  ' +
      '         dstopflag,                                        ' +
      '         estopflag,                                        ' +
      '         eestopflag,                                       ' +
      '         utcsalebegintime,                                 ' +
      '         utcsaleendtime,                                   ' +
      '         utcvalidityperiodbegintime,                       ' +
      '         utcvalidityperiodendtime,                         ' +
      '         utceventbegintime     )                           ' +
      ' select a.FileName,                                        ' +
      '        a.CreationDate,                                    ' +
      '        a.ProductId,                                       ' +
      '        a.ImpulsiveFlag,                                   ' +
      '        a.ProductKind,                                     ' +
      '        a.SpecialPpv,                                      ' +
      '        a.PurchasableFlag,                                 ' +
      '        a.SoldFlag,                                        ' +
      '        a.SuspendedFlag,                                   ' +
      '        a.ProductName,                                     ' +
      '        a.ProductDescription,                              ' +
      '        a.EngProductName,                                  ' +
      '        a.EngProductDescription,                           ' +
      '        a.SaleBeginTime,                                   ' +
      '        a.SaleEndTime,                                     ' +
      '        a.ValidityPeriodBeginTime,                         ' +
      '        a.ValidityPeriodEndTime,                           ' +
      '        a.ProductCategory,                                 ' +
      '        a.EpgPrice,                                        ' +
      '        a.SubAllowPpvEventFlag,                            ' +
      '        b.FileName,                                        ' +
      '        b.CreationDate,                                    ' +
      '        b.ChannelId,                                       ' +
      '        b.EventBeginTime,                                  ' +
      '        b.Duration,                                        ' +
      '        b.EventId,                                         ' +
      '        a.EventPreviewTime,                                ' +
      '        b.EventType,                                       ' +
      '        b.Name,                                            ' +
      '        b.EpgDesc,                                         ' +
      '        b.EngName,                                         ' +
      '        b.EngEpgDesc,                                      ' +
      '        b.ParentalRating,                                  ' +
      '        b.ContentNibble1,                                  ' +
      '        b.ContentNibble2,                                  ' +
      '        b.UserNibble1,                                     ' +
      '        b.UserNibble2,                                     ' +
      '        c.ChannelNo,                                       ' +
      '        to_char( to_number( to_char( sysdate, ''yyyy'' ) ) - 1911 ) || ' +
      '        to_char( sysdate, ''/mm/dd hh24:mi:ss'' ),         ' +
      '        ''%s'',                                            ' +
      '        a.ParentalRating,                                  ' +
      '        a.dStopFlag,                                       ' +
      '        a.eStopFlag,                                       ' +
      '        b.eStopFlag,                                       ' +
      '        a.utcsalebegintime,                                ' +
      '        a.utcsaleendtime,                                  ' +
      '        a.utcvalidityperiodbegintime,                      ' +
      '        a.utcvalidityperiodendtime,                        ' +
      '        b.utceventbegintime                                ' +
      '   from so158 a, so159 b, so162 c                          ' +
      '  where a.eventid = b.eventid(+)                           ' +
      '    and b.channelid = c.channelid(+)                       ' +
      '    and a.productkind = ''PPV''                            ' +
      '    and ( ( a.Action <> 0 ) or ( b.Action <> 0  ) )        ',
      [SMS_PROGRAM_SCHEDULAR] );
    FDataConnection.BeginTrans;
    FDataWriter.ExecSQL;
    FDataConnection.CommitTrans;
    FDataReader.Close;
    FDataReader.SQL.Text :=
      ' select * from so160 order by productid, eventbegintime ';
    FDataReader.Open;
    FDataReader.First;
    while not FDataReader.Eof do
    begin
      aPRCodeWrapper := FDataReader.FieldByName( 'ParentalRating' ).AsString;
      aPRCodeDex := FDataReader.FieldByName( 'dParentalRating' ).AsString;
      GetParentalratingName( FDataWriter,  aPRCodeWrapper, aPRNameWrapper );
      GetParentalratingName( FDataWriter, aPRCodeDex, aPRNameDex );
      FSo160.AppendRecord( [
        FDataReader.FieldByName( 'dfilename' ).AsString,
        FDataReader.FieldByName( 'dcreationdate' ).AsString,
        FDataReader.FieldByName( 'productid' ).AsString,
        FDataReader.FieldByName( 'impulsiveflag' ).AsString,
        FDataReader.FieldByName( 'ProductKind' ).AsString,
        FDataReader.FieldByName( 'SpecialPpv' ).AsString,
        FDataReader.FieldByName( 'PurchasableFlag' ).AsString,
        FDataReader.FieldByName( 'SoldFlag' ).AsString,
        FDataReader.FieldByName( 'SuspendedFlag' ).AsString,
        FDataReader.FieldByName( 'ProductName' ).AsString,
        FDataReader.FieldByName( 'ProductDescription' ).AsString,
        FDataReader.FieldByName( 'EngProductName' ).AsString,
        FDataReader.FieldByName( 'EngProductDescription' ).AsString,
        FDataReader.FieldByName( 'salebegintime' ).AsString,
        FDataReader.FieldByName( 'saleendtime' ).AsString,
        FDataReader.FieldByName( 'ValidityPeriodBeginTime' ).AsString,
        FDataReader.FieldByName( 'ValidityPeriodEndTime' ).AsString,
        FDataReader.FieldByName( 'productcategory' ).AsString,
        FDataReader.FieldByName( 'EpgPrice' ).AsString,
        FDataReader.FieldByName( 'SubAllowPpvEventFlag' ).AsString,
        FDataReader.FieldByName( 'eFileName' ).AsString,
        FDataReader.FieldByName( 'eCreationDate' ).AsString,
        FDataReader.FieldByName( 'ChannelId' ).AsString,
        FDataReader.FieldByName( 'EventBeginTime' ).AsString,
        FDataReader.FieldByName( 'Duration' ).AsString,
        FDataReader.FieldByName( 'EventId' ).AsString,
        FDataReader.FieldByName( 'EventPreviewTime' ).AsString,
        FDataReader.FieldByName( 'EventType' ).AsString,
        FDataReader.FieldByName( 'Name' ).AsString,
        FDataReader.FieldByName( 'EpgDesc' ).AsString,
        FDataReader.FieldByName( 'EngName' ).AsString,
        FDataReader.FieldByName( 'EngEpgDesc' ).AsString,
        aPRCodeWrapper,
        aPRNameWrapper,
        FDataReader.FieldByName( 'ContentNibble1' ).AsString,
        FDataReader.FieldByName( 'ContentNibble2' ).AsString,
        FDataReader.FieldByName( 'UserNibble1' ).AsString,
        FDataReader.FieldByName( 'UserNibble2' ).AsString,
        FDataReader.FieldByName( 'ChannelNo' ).AsString,
        FDataReader.FieldByName( 'UpdTime' ).AsString,
        FDataReader.FieldByName( 'UpdEn' ).AsString,
        aPRCodeDex,
        aPRNameDex,
        FDataReader.FieldByName( 'dStopFlag' ).AsInteger,
        FDataReader.FieldByName( 'eStopFlag' ).AsInteger,
        FDataReader.FieldByName( 'eeStopFlag' ).AsInteger,
        EmptyStr,
        FDataReader.FieldByName( 'UTCSaleBeginTime' ).AsString,
        FDataReader.FieldByName( 'UTCSaleEndTime' ).AsString,
        FDataReader.FieldByName( 'UTCValidityPeriodBeginTime' ).AsString,
        FDataReader.FieldByName( 'UTCValidityPeriodEndTime' ).AsString,
        FDataReader.FieldByName( 'UTCEventBeginTime' ).AsString] );
      FDataReader.Next;
    end;
    FDataReader.Close;
    Result := True;
  except
    on E: Exception do
    begin
      if FDataConnection.InTransaction then FDataConnection.RollbackTrans;
      if FDataReader.Active then FDataReader.Close;
      aMsg := Format( '%s - %s', ['SO160', E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.Merge(var aMsg: String): Boolean;
begin
  UnPrepareDataBuffer;
  PrepareDataBuffer;
  Result := DoSo160( aMsg );
  if not Result then Exit;
  Result := StoreRecord( saBefore, aMsg );
  if not Result then Exit;
  Result := DoSo154( aMsg );
  if not Result then Exit;
  Result := DoMerginEx( aMsg );
  if not Result then Exit;
  Result := StoreRecord( saAftrer, aMsg );
  if not Result then Exit;
  Result := DoSo162( aMsg );
  if not Result then Exit;
  Result := DoMakeLog( aMsg );
  if not Result then Exit;
  Result := DoWriteLog( aMsg );
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DetectChangeCount(var aMsg: String; var aCount: Integer): Boolean;
begin
  aCount := 0;
  FDataReader.Close;
  FDataReader.SQL.Text :=
    ' select count(1) as counts                          ' +
    '   from so158 a, so159 b, so162 c                   ' +
    ' where a.eventid = b.eventid(+)                     ' +
    '   and b.channelid = c.channelid(+)                 ' +
    '   and a.productkind = ''PPV''                      ' +
    '   and ( ( a.Action <> 0 ) or ( b.Action <> 0  ) )  ';
  try
    FDataReader.Open;
    aCount := FDataReader.FieldByName( 'counts' ).AsInteger;
  except
    on E: Exception do aMsg := E.Message;
  end;
  FDataReader.Close;
  Result := ( aCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.MakeLogText(const aType: TLogType): String;
begin
  Result := EmptyStr;
  case aType of
     ltProductDelete:
       begin
         Result := '【節目取消】 該時段節目:%s。';
       end;
     ltEventDelete:
       begin
         Result := '【時段異動】 該時段節目:%s, 原時段播映(%s)%s 已被取消。';
       end;
     ltEventTimeChange:
       begin
         Result := '【時段異動】 該時段節目:%s, 原時段播映(%s)%s 從 %s 改為 %s。';
       end;
     ltEpgPriceChange:
       begin
         Result := '【價格異動】 該時段節目:%s, 原價:%f 變更為 %f。';
       end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoMakeLog(var aMsg: String): Boolean;

  { ------------------------------------------------------- }

  function ProductRecordFound(aFindDS: TClientDataSet; const aProductId: String): Boolean;
  begin
    Result := False;
    if not Assigned( aFindDS ) then Exit;
    Result := aFindDS.Locate( 'PRODUCTID', aProductId, [] );
  end;

  { ------------------------------------------------------- }

  function EventRecordFound(aFindDS: TClientDataSet; const aProductId, aEventId: String): Boolean;
  begin
    Result := False;
    if not Assigned( aFindDS ) then Exit;
    Result := aFindDS.Locate( 'PRODUCTID;EVENTID', VarArrayOf( [aProductId, aEventId ] ), [] );
  end;

  { ------------------------------------------------------- }

  procedure AppendLog(const aProductId, aText: String);
  var
    aFlag: Integer;
  begin
    if ( aText <> EmptyStr ) then
    begin
      aFlag := 0;
      if AnsiPos( '節目取消', aText ) > 0 then aFlag := 1;
      FLogText.Append;
      FLogText.FieldByName( 'ProductId' ).AsString := aProductId;
      FLogText.FieldByName( 'Note' ).AsString := aText;
      FLogText.FieldByName( 'Flag' ).AsInteger := aFlag;
      FLogText.FieldByName( 'Upden' ).AsString := SMS_PROGRAM_SCHEDULAR;
      FLogText.FieldByName( 'UpdTime' ).AsString := ROCDate;
      FLogText.Post;
    end;
  end;

  { ------------------------------------------------------- }

  procedure SetEventFilter(aFindDS: TClientDataSet; const aProductId: String);
  begin
    aFindDS.Filter := EmptyStr;
    aFindDS.Filtered := False;
    if ( aProductId <> EmptyStr ) then
    begin
      aFindDS.Filter := Format( 'ProductId=''%s''', [aProductId] );
      aFindDS.Filtered := True;
    end;  
  end;

  { ------------------------------------------------------- }

label
  DoNothing1, DoNothing2;
var
  aProductId, aEventId, aLogText, aProductName, aEventName: String;
  aBeforeEventBeginTime, aAfterEventBeginTIme: String;
  aBeforeEpgPrice, aAfterEpgPrice: Double;
  aDelete, aIgnor: Boolean;
  aCount: Integer;
begin
  Result := False;
  aCount := 0;
  { 合併前的資料 FProductBuffer, FEventBuffer }
  { 合併後的資料 FProductBuffer2, FEventBuffer2 }
  FProductBuffer2.First;
  try
    while not FProductBuffer2.Eof do
    begin
      aProductId := FProductBuffer2.FieldByName( 'ProductId' ).AsString;
      aProductName := FProductBuffer2.FieldByName( 'ProductName' ).AsString;
      { 沒合併前的資料沒找到, 表示本次該筆 Product 是新增 }
      if not ProductRecordFound( FProductBuffer, aProductId ) then
        goto DoNothing1;
      { Product 本次異動是刪除 }
      aDelete :=
        ( FProductBuffer2.FieldByName( 'dStopFlag' ).AsInteger = 1 ) and
        ( FProductBuffer.FieldByName( 'dStopFlag' ).AsInteger = 0 );
      { Product 本次比較結果, 原本就已經是刪除 }
      aIgnor :=
        ( FProductBuffer2.FieldByName( 'dStopFlag' ).AsInteger = 1 ) and
        ( FProductBuffer.FieldByName( 'dStopFlag' ).AsInteger = 1 );
      if ( aIgnor ) then goto DoNothing1;
      // 1.節目刪除
      if ( aDelete ) then
      begin
        aLogText := MakeLogText( ltProductDelete );
        aLogText := Format( aLogText, [aProductName] );
        AppendLog( aProductId, aLogText);
        goto DoNothing1;
      end;
      // 2.價格異動, 只檢核 EpgPrice, Price 欄位是開放給使用者變動, 不是 Nagra
      //   匯出的, 所以不管
      aAfterEpgPrice := FProductBuffer2.FieldByName( 'EpgPrice' ).AsFloat;
      aBeforeEpgPrice := FProductBuffer.FieldByName( 'EpgPrice' ).AsFloat;
      if ( aAfterEpgPrice <> aBeforeEpgPrice ) then
      begin
        aLogText := MakeLogText( ltEpgPriceChange );
        aLogText := Format( aLogText, [aProductName,
          aBeforeEpgPrice, aAfterEpgPrice] );
        AppendLog( aProductId, aLogText );
      end;
      // 3.時段異動
      SetEventFilter( FEventBuffer2, aProductId );
      FEventBuffer2.First;
      while not FEventBuffer2.Eof do
      begin
        aEventId := FEventBuffer2.FieldByName( 'EventId' ).AsString;
        aEventName := FEventBuffer2.FieldByName( 'Name' ).AsString;
        if not EventRecordFound( FEventBuffer, aProductId, aEventId ) then
          goto DoNothing2;
        { Event 本次異動是刪除 }
        aDelete :=
          ( FEventBuffer2.FieldByName( 'eStopFlag' ).AsInteger = 1 ) and
          ( FEventBuffer.FieldByName( 'eStopFlag' ).AsInteger = 0 );
        { Event 本次比較結果, 原本就已經是刪除 }
        aIgnor :=
          ( FEventBuffer2.FieldByName( 'eStopFlag' ).AsInteger = 1 ) and
          ( FEventBuffer.FieldByName( 'eStopFlag' ).AsInteger = 1 );
        if ( aIgnor ) then goto DoNothing2;
        // 1.Event 刪除
        if ( aDelete ) then
        begin
          aLogText := MakeLogText( ltEventDelete );
          aLogText := Format( aLogText, [aProductName, aEventId, aEventName] );
          AppendLog( aProductId, aLogText );
          goto DoNothing2;
        end;
        // 2.時間變動
        aAfterEventBeginTime := FEventBuffer2.FieldByName( 'EventBeginTime' ).AsString;
        aBeforeEventBeginTime := FEventBuffer.FieldByName( 'EventBeginTime' ).AsString;
        if ( aAfterEventBeginTime <> aBeforeEventBeginTime ) then
        begin
          aBeforeEventBeginTime := FormatMaskText( '####/##/##;0;_',
            Copy( aBeforeEventBeginTime, 1, 8 ) ) + ' ' + FormatMaskText( '##:##:##;0;-',
            Copy( aBeforeEventBeginTime, 9, 6 ) );
          aAfterEventBeginTime := FormatMaskText( '####/##/##;0;_',
            Copy( aAfterEventBeginTime, 1, 8 ) ) + ' ' + FormatMaskText( '##:##:##;0;-',
            Copy( aAfterEventBeginTime, 9, 6 ) );
          aLogText := MakeLogText( ltEventTimeChange );
          aLogText := Format( aLogText, [aProductName, aEventId,
            aEventName, aBeforeEventBeginTime, aAfterEventBeginTime] );
          AppendLog( aProductId, aLogText );
          goto DoNothing2;
        end;
  DoNothing2:
        FEventBuffer2.Next;
      end;
      SetEventFilter( FEventBuffer2, EmptyStr );
  DoNothing1:
      Inc( aCount );
      FProductBuffer2.Next;
      if ( aCount mod 10 ) = 0 then Sleep( 10 );
    end;
    Result := True;
  except
    on E: Exception do aMsg := Format( '%s - %s', ['SO154Log', E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TMergeChange.DoWriteLog(var aMsg: String): Boolean;
begin
  Result := False;
  FDataWriter.SQL.Text :=
    ' insert into SO155Log ( SeqNo, ProductId, Note, Flag, UpdEn, UpdTime )  ' +
    '   values( :SeqNo, :ProductId, :Note, :Flag, :UpdEn, :UpdTime )         ';
  FDataConnection.BeginTrans;
  FLogText.First;
  try
    while not FLogText.Eof do
    begin
      FDataWriter.Parameters.ParamByName( 'SeqNo' ).Value := EmptyStr;
      FDataWriter.Parameters.ParamByName( 'ProductId' ).Value :=
        FLogText.FieldByName( 'ProductId' ).AsString;
      FDataWriter.Parameters.ParamByName( 'Note' ).Value :=
        FLogText.FieldByName( 'Note' ).AsString;
      FDataWriter.Parameters.ParamByName( 'Flag' ).Value :=
        FLogText.FieldByName( 'Flag' ).AsInteger;
      FDataWriter.Parameters.ParamByName( 'UpdEn' ).Value :=
        FLogText.FieldByName( 'UpdEn' ).AsString;
      FDataWriter.Parameters.ParamByName( 'UpdTime' ).Value :=
        FLogText.FieldByName( 'UpdTime' ).AsString;
      FDataWriter.ExecSQL;  
      FLogText.Next;
    end;
    FDataConnection.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      if FDataConnection.InTransaction then FDataConnection.RollbackTrans;
      aMsg := Format( '%s - %s', ['SO154Log', E.Message] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }


end.
