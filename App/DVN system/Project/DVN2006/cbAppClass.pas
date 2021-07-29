unit cbAppClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, SyncObjs, IniFiles,
  Forms, DB, ADODB, DBClient;


const
  WM_SOCKET = WM_USER + 1;
  WM_DATABASE = WM_USER + 2;

type
  TDbState = ( dbClose, dbOpen, dbActive, dbError );
  TCmdType = ( ctHigh, ctLow );
  TSocketMsgReason = ( smrStop, smrStart, smrAck, smrNack, smrSocketError, smrIdleTimeReach );
  TThreadType = ( tdSo, tdSend, tdRecv );

  TWMSocket = packed record
    Msg: Cardinal;
    ThreadType: Byte;
    SocketMsgReason: Byte;
    ErrorCode: Smallint;
    ErrorDescription: PChar;
    Result: Longint;
  end;


  { DVN-CA 介接 API 請參考 801-SMSGWYCMD-406_TBC.pdf }

  { TCasEnv }

  TCasEnv = class(TPersistent)
  private
    FIp: String;
    FSendPort: String;
    FRecvPort: String;
    FProtocol: String;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    { DVN-CA 主機 IP }
    property Ip: String read FIp write FIp;
    { DVN-CA 主機傳送指令 prot }
    property SendPort: String read FSendPort write FSendPort;
    { DVN-CA 主機接收回傳資料 prot }
    property RecvPort: String read FRecvPort write FRecvPort;
    { DVN-CA 主機 protocol, 傳指令時會須要至此值, 現固定為 2 }
    property Protocol: String read FProtocol write FProtocol;
  end;

  { TCommEnv }
  
  TCommEnv = class(TPersistent)
  private
    FCAEnableRecv: Boolean;
    FDbWriteError: Boolean;
    FDbProcIPPV: String;
    FCAEnableSend: Boolean;
    FDbProcBatch: String;
    FCASendDelay: Double;
    FNorTimeReadFreq: Integer;
    FDbProcRecords: Integer;
    FDbWaterMark: Integer;
    FCACheckFreq: Integer;
    FDbRetryFreq: Integer;
    FCAMaxError: Integer;
    FBusyTimeReadFreq: Integer;
    FCARetryFreq: Integer;
    FBusyTimeEnd: String;
    FBusyTimeStart: String;
    FCAProductDefine: String;
    FCAIdleTime: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    { 當資料庫斷線時, 每多少秒重試連線 }
    property DbRetryFreq: Integer read FDbRetryFreq write FDbRetryFreq;
    { 忙碌時段定義(起) }
    property BusyTimeStart: String read FBusyTimeStart write FBusyTimeStart;
    { 忙碌時段定義(迄) }
    property BusyTimeEnd: String read FBusyTimeEnd write FBusyTimeEnd;
    { 忙碌時段每多少秒讀一取次資料 }
    property BusyTimeReadFreq: Integer read FBusyTimeReadFreq write FBusyTimeReadFreq;
    { 一般時段每多少秒讀一取次資料 }
    property NorTimeReadFreq: Integer read FNorTimeReadFreq write FNorTimeReadFreq;
    { 讀取系統台待處理指令時, 一次讀取幾筆( SEND_DVN Table ) }
    property DbProcRecords: Integer read FDbProcRecords write FDbProcRecords;
    { 當無法傳送指令給 DVN-CA 時, 待處理指令直接標示成錯誤 }
    property DbWriteError: Boolean read FDbWriteError write FDbWriteError;
    { 是否處理 IPPV 資料, 值:A--> 全部, N--> 不處理IPPV, O--> 只處理IPPV }
    property DbProcIPPV: String read FDbProcIPPV write FDbProcIPPV;
    { 是否處理標示為 [批次] 的指令, 值:A--> 全部, N--> 不處理批次, O--> 只處理批次 }
    property DbProcBatch: String read FDbProcBatch write FDbProcBatch;
    { 當傳送給 DVN-CA 待處理指令序列超出否一基準值後,即 hold 住 不再抓取各系統待處理指令 }
    property DbWaterMark: Integer read FDbWaterMark write FDbWaterMark;
    { 當 DVN-CA 連不上時或是斷線時, 每多少秒重試連線 }
    property CARetryFreq: Integer read FCARetryFreq write FCARetryFreq;
    { 每多少秒送一次 虛擬的指令, PS: DVN-CA 用不到, 保留 }
    property CACheckFreq: Integer read FCACheckFreq write FCACheckFreq;
    { 啟用傳送指令 }
    property CAEnableSend: Boolean read FCAEnableSend write FCAEnableSend;
    { 啟用接收 Callback 的資料 }
    property CAEnableRecv: Boolean read FCAEnableRecv write FCAEnableRecv;
    { 每傳送一筆低階指令給 DVN-CA 後, 暫停多久(秒), 用來控制傳送指令速度 }
    property CASendDelay: Double read FCASendDelay write FCASendDelay;
    { 當傳送指令或接收回傳資料, 發生多少次錯誤自動斷線 }
    property CAMaxError: Integer read FCAMaxError write FCAMaxError;
    { CA 頭端對於 Product 的定義 --> GP, PP, 或是由客服拋過來 }
    property CAProductDefine: String read FCAProductDefine write FCAProductDefine;
    { 當 Gateway 連線到 CA 後, 若一段時間沒有連線, 則自動斷線再重連 CA }
    property CAIdleTime: Integer read FCAIdleTime write FCAIdleTime;
  end;

  { TNotify }

  TTextNotify = class(TPersistent)
  private
    FText: String;
    FFlag: Integer;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    property Text: String read FText write FText;
    property Flag: Integer read FFlag write FFlag;
  end;

  { TSo }

  TSo = class(TPersistent)
  private
    FSelected: Boolean;
    FRecordCount: Integer;
    FErrorCount: Integer;
    FDisplayNode: Pointer;
    FDbAliase: String;
    FNetworkId: String;
    FDbUserId: String;
    FPosName: String;
    FDbUserPass: String;
    FAreaCode: String;
    FCompName: String;
    FSoType: String;
    FCompCode: String;
    FErrorTime: TDateTime;
    FDbState: TDbState;
    FConnection: TADOConnection;
    FReader: TADOQuery;
    FWriter: TADOQuery;
    FBuffer: TClientDataSet;
    procedure SetDbState(const Value: TDbState);
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    property Selected: Boolean read FSelected write FSelected;
    property PosName: String read FPosName write FPosName;
    property SoType: String read FSoType write FSoType;
    property CompCode: String read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property NetworkId: String read FNetworkId write FNetworkId;
    property AreaCode: String read FAreaCode write FAreaCode;
    property DbUserId: String read FDbUserId write FDbUserId;
    property DbUserPass: String read FDbUserPass write FDbUserPass;
    property DbAliase: String read FDbAliase write FDbAliase;
    property DbState: TDbState read FDbState write SetDbState;
    property RecordCount: Integer read FRecordCount write FRecordCount;
    property ErrorCount: Integer read FErrorCount;
    property ErrorTime: TDateTime read FErrorTime;
    property DisplayNode: Pointer read FDisplayNode write FDisplayNode;
    property Connection: TADOConnection read FConnection write FConnection;
    property Reader: TADOQuery read FReader write FReader;
    property Writer: TADOQuery read FWriter write FWriter;
    property Buffer: TClientDataSet read FBuffer write FBuffer;
  end;

  { TSoList }

  TSoList = class(TList)
  private
    function GetSelectCount: Integer;
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TSo;
    procedure SetItem(Index: Integer; aObj: TSo);
  public
    procedure Insert(Index: Integer; aObj: TSo);
    procedure Assign(aSource: TSoList);
    function Add(aObj: TSo): Integer;
    function Extract(aObj: TSo): TSo;
    function Remove(aObj: TSo): Integer;
    function IndexOf(aObj: TSo): Integer; overload;
    function IndexOf(aCompCode: String): Integer; overload;
    function First: TSo;
    function Last: TSo;
    property Items[Index: Integer]: TSo read GetItem write SetItem; default;
    property SelectCount: Integer read GetSelectCount;
  end;


  TObServer = class;

  { TSubject }

  TSubject = class(TObject)
  private
    FList: TList;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddObServer(aObServer: TObServer);
    procedure RemoveObServer; overload;
    procedure RemoveObServer(aObServer: TObServer); overload;
    procedure NotifyObServer; overload;
    procedure NotifyObServer(aObServer: TObServer); overload;
  end;

  TTextSubject = class(TSubject)
  private
    FNotify: TTextNotify;
    procedure InternalNotify(const aMsg: String; const aFlag: Integer);
  public
    constructor Create;
    destructor Destroy; override;
    procedure OK(const aMsg: String);
    procedure Warning(const aMsg: String);
    procedure Error(const aMsg: String);
    procedure Info(const aMsg: String);
    procedure OKFmt(const aMsg: String; const Args: array of const);
    procedure WarningFmt(const aMsg: String; const Args: array of const);
    procedure ErrorFmt(const aMsg: String; const Args: array of const);
    procedure InfoFmt(const aMsg: String; const Args: array of const);
    property Notify: TTextNotify read FNotify;
  end;

  { TSoSubject }

  TSoSubject = class(TSubject)
  private
     FSo: TSo;
  public
    procedure Notify(aSo: TSo);
    property So: TSo read FSo;
  end;

  TSendDvn = class;

  { TSendDvnSubject }

  TSendDvnSubject = class(TSubject)
  private
     FSendDvn: TSendDvn;
  public
    procedure Notify(aSendDvn: TSendDvn);
    property SendDvn: TSendDvn read FSendDvn;
  end;


  TSubjectEvent = procedure (aSubject: TSubject) of object;

  { TObServer }

  TObServer = class(TObject)
  private
    FPause: Boolean;
    FOnUpdate: TSubjectEvent;
    FSource: TSubject;
    FLock: TCriticalSection;
    procedure CallOnUpdate;
  protected
    procedure Update(aSubject: TSubject); virtual;
  public
    constructor Create;
    destructor Destroy; override;
    property Pause: Boolean read FPause write FPause;
    property OnUpdate: TSubjectEvent read FOnUpdate write FOnUpdate;
  end;

  { TMessageThread }

  TMessageThread = class(TThread)
  private
    FMsgHnd: THandle;
  protected
    procedure WndProc(var Msg: TMessage); virtual;
  public
    constructor Create(CreateSuspended: Boolean); virtual;
    destructor Destroy; override;
    property MsgHandle: THandle read FMsgHnd;
  end;

  { TAppThread }

  TAppThread = class(TMessageThread)
  private
    FCasEnv: TCasEnv;
    FCommEnv: TCommEnv;
    FHighCmdEnv: TClientDataSet;
    FLowCmdEnv: TClientDataSet;
    FCmdErrCnv: TClientDataSet;
    FMainFormHandle: THandle;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property CasEnv: TCasEnv read FCasEnv;
    property CommEnv: TCommEnv read FCommEnv;
    property HighCmdEnv: TClientDataSet read FHighCmdEnv;
    property LowCmdEnv: TClientDataSet read FLowCmdEnv;
    property CmdErrCnv: TClientDataSet read FCmdErrCnv;
    property MainFormHandle: THandle read FMainFormHandle write FMainFormHandle;
  end;
  
  { TLanguageManager }

  TLanguageManager = class(TObject)
  private
    FResNames: TStringList;
    FResValues: TStringList;
  public
    constructor Create;
    destructor Destroy; override;
    function Get(const aResName: String): String;
    procedure Add(const aResName, aResValue: String);
    procedure Clear;
    procedure LoadFromFile(const aFileName: String); overload;
    procedure LoadFromFile; overload;
  end;


  { TSendDvn }

  TSendDvn = class(TPersistent)
  private
    FCredit: Double;
    FLowCmdCount: Integer;
    FLowCmdNote: String;
    FFrameNo: String;
    FVersion: String;
    FHighSeqNo: String;
    FParentNode: Pointer;
    FDisplayNode: Pointer;
    FConsoleNode: Pointer;
    FStartingTime: String;
    FCreditCfCode: String;
    FCompCode: String;
    FCompName: String;
    FExpiryTime: String;
    FProductId: String;
    FRecvCmdText: String;
    FHighCmdStatus: String;
    FConfigValue: String;
    FProductType: String;
    FUniqueId: String;
    FSNo: String;
    FIccNo: String;
    FLowCmdStatus: String;
    FOperator: String;
    FErrMsg: String;
    FSendCmdText: String;
    FStbNo: String;
    FAreaCode: String;
    FLowCmd: String;
    FErrCode: String;
    FBankType: String;
    FProcessingDate: String;
    FHighCmd: String;
    FCreditAction: String;
    FMailAction: String;
    FNotes: String;
    FCmdType: TCmdType;
    FRecvTime: TDateTime;
    FUpdTime: TDateTime;
    FSendTime: TDateTime;
    FIsBatch: Boolean;
    FAck: String;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create(const aCmdType: TCmdType);
    procedure Reset;
    { 高階指令的 UniqueId = SendDvn Table 的 SeqNo }
    { 低階指令的 UniqueId = FrameNo }
    property UniqueId: String read FUniqueId write FUniqueId;
    property CompCode: String read FCompCode write FCompCode;
    property CompName: String read FCompName write FCompName;
    property HighSeqNo: String read FHighSeqNo write FHighSeqNo;
    property FrameNo: String read FFrameNo write FFrameNo;
    property Version: String read FVersion write FVersion;
    property HighCmd: String read FHighCmd write FHighCmd;
    property LowCmd: String read FLowCmd write FLowCmd;
    property LowCmdCount: Integer read FLowCmdCount write FLowCmdCount;
    property LowCmdNote: String read FLowCmdNote write FLowCmdNote;
    property CmdType: TCmdType read FCmdType;
    property IccNo: String read FIccNo write FIccNo;
    property StbNo: String read FStbNo write FStbNo;
    property AreaCode: String read FAreaCode write FAreaCode;
    property Notes: String read FNotes write FNotes;
    property HighCmdStatus: String read FHighCmdStatus write FHighCmdStatus;
    property LowCmdStatus: String read FLowCmdStatus write FLowCmdStatus;
    property Operator: String read FOperator write FOperator;
    property UpdTime: TDateTime read FUpdTime write FUpdTime;
    property SendTime: TDateTime read FSendTime write FSendTime;
    property RecvTime: TDateTime read FRecvTime write FRecvTime;
    property ErrCode: String read FErrCode write FErrCode;
    property ErrMsg: String read FErrMsg write FErrMsg;
    property StartingTime: String read FStartingTime write FStartingTime;
    property ExpiryTime: String read FExpiryTime write FExpiryTime;
    property ProductId: String read FProductId write FProductId;
    property Credit: Double read FCredit write FCredit;
    property CreditAction: String read FCreditAction write FCreditAction;
    property CreditCfCode: String read FCreditCfCode write FCreditCfCode;
    property BankType: String read FBankType write FBankType;
    property ProductType: String read FProductType write FProductType;
    property ConfigValue: String read FConfigValue write FConfigValue;
    property MailAction: String read FMailAction write FMailAction;
    property SNo: String read FSNo write FSNo;
    property ProcessingDate: String read FProcessingDate write FProcessingDate;
    property SendCmdText: String read FSendCmdText write FSendCmdText;
    property RecvCmdText: String read FRecvCmdText write FRecvCmdText;
    property Ack: String read FAck write FAck;
    property DisplayNode: Pointer read FDisplayNode write FDisplayNode;
    property ParentNode: Pointer read FParentNode write FParentNode;
    property ConsoleNode: Pointer read FConsoleNode write FConsoleNode;
    property IsBatch: Boolean read FIsBatch write FIsBatch;
  end;

  { TDvnAckParser }

  TDvnAckParser = class(TObject)
  private
    FList: TStrings;
    FCmdErrCnv: TClientDataSet;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Parse(const aAckObj: TSendDvn);
    property CmdErrCnv: TClientDataSet read FCmdErrCnv;
  end;


  { TSyncList }

  TSyncList = class(TObject)
  private
    FLock: TMREWSync;
    FList: TStringList;
  protected
    function Get(Index: Integer): string;
    procedure Put(Index: Integer; const UniqueId: string);
    function GetObject(Index: Integer): TSendDvn;
    procedure PutObject(Index: Integer; AObject: TSendDvn);
    function GetSorted: Boolean;
    procedure SetSorted(const Value: Boolean);
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(const UniqueId: string): Integer;
    function AddObject(const UniqueId: string; AObject: TSendDvn): Integer;
    procedure Clear;
    procedure Delete(Index: Integer); overload;
    procedure Delete(AObject: TSendDvn); overload;
    procedure FreeObject(Index: Integer); overload;
    procedure FreeObject(AObject: TSendDvn); overload;
    function IndexOf(const UniqueId: string): Integer;
    function IndexOfObject(AObject: TSendDvn): Integer;
    property UniqueId[Index: Integer]: string read Get write Put; default;
    property Objects[Index: Integer]: TSendDvn read GetObject write PutObject;
    property Sorted: Boolean read GetSorted write SetSorted;
    property Count: Integer read GetCount;
  end;

  { TCurValueLoader }

  TCurValueLoader = class(TObject)
  public
    function Restore: Variant; virtual; abstract;
    procedure Save(const aCurValue: Variant); virtual; abstract;
  end;


  { TIniCurValueLoader }

  TIniCurValueLoader = class(TCurValueLoader)
  private
    FIni: TIniFile;
    FSection: String;
    FIdent: String;
  public
    constructor Create(const aFileName, aSection, aIdent: String);
    destructor Destroy; override;
    function Restore: Variant; override;
    procedure Save(const aCurValue: Variant); override;
  end;


  { TGenerateRule }

  TGenerateRule = class(TObject)
  protected
    function GetMinValue: Variant; virtual; abstract;
    function GetMaxValue: Variant; virtual; abstract;
    function GetNextValue: Variant; virtual; abstract;
    function GetCurValue: Variant; virtual; abstract;
  end;

  { TNumberGenerateRule }

  TNumberGenerateRule = class(TGenerateRule)
  private
    FMinValue: Integer;
    FMaxValue: Integer;
    FCurValue: Integer;
    FLoader: TCurValueLoader;
  protected
    function GetMinValue: Variant; override;
    function GetMaxValue: Variant; override;
    function GetNextValue: Variant; override;
    function GetCurValue: Variant; override;
    procedure ResetValue; virtual;
    property Loader: TCurValueLoader read FLoader;
  public
    constructor Create(aMinValue, aMaxValue, aCurValue: Integer); overload;
    constructor Create(aMinValue, aMaxValue: Variant; aLoader: TCurValueLoader); overload;
    destructor Destroy; override;
  end;

  { TIdGenerator }

  TIdGenerator = class(TObject)
  private
    FLock: TCriticalSection;
    FThreadSafe: Boolean;
    FRule: TGenerateRule;
    function GetMinValue: Variant;
    function GetMaxValue: Variant;
    function GetNextValue: Variant;
    function GetCurValue: Variant;
  public
    constructor Create(const aThreadSafe: Boolean; aRule: TGenerateRule);
    destructor Destroy; override;
    property MinValue: Variant read GetMinValue;
    property MaxValue: Variant read GetMaxValue;
    property NextValue: Variant read GetNextValue;
    property CurrentValue: Variant read GetCurValue;
    procedure SaveCurValue;
  end;


  { TDvnCommandBuilder }

  TDvnCommandBuilder = class(TObject)
  private
    FHighCmd: TSendDvn;
    FHighCmdEnv: TClientDataSet;
    FLowCmdEnv: TClientDataSet;
    FCommEnv: TCommEnv;
    FCommandList: TStringList;
    FTempList: TStringList;
    FCasEnv: TCasEnv;
    FFrameNoGenerator: TIdGenerator;
    FCreditCfCodeGenerator: TIdGenerator;
    function GetLowCmdCount: Integer;
    function GetLowCmd(Index: Integer): TSendDvn;
    function GetValue(const aFindField, aResultField: String; const Args: array of Variant): String;
    function IsSpecialHighCmd(const aHighCmdId: String): Boolean;
    procedure ParseNormal;
    procedure ParseSpecial;
    procedure HighCmdExtractToLowCmd;
    procedure MoveList;
    procedure AssemblyCommand;
    procedure EachLowCmd(aObj: TSendDvn);
    function GetCommandCode(const aObj: TSendDvn): String;
    function GetCommandHeader(const aObj: TSendDvn): String;
    function GetCommandParam(const aObj: TSendDvn): String;
  public
    constructor Create;
    destructor Destroy; override;
    procedure BuildLowCommand(aHighCmd: TSendDvn); overload;
    procedure Clear;
    property CasEnv: TCasEnv read FCasEnv;
    property HighCmdEnv: TClientDataSet read FHighCmdEnv;
    property CommEnv: TCommEnv read FCommEnv;
    property LowCmdEnv: TClientDataSet read FLowCmdEnv;
    property LowCmdCount: Integer read GetLowCmdCount;
    property LowCmd[Index: Integer]: TSendDvn read GetLowCmd;
  end;

  var LanguageManager: TLanguageManager;


implementation

uses cbUtilis;

{ TCasEnv }

procedure TCasEnv.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TCasEnv ) then
  begin
    TCasEnv( Dest ).Ip := FIp;
    TCasEnv( Dest ).SendPort := FSendPort;
    TCasEnv( Dest ).RecvPort := FSendPort;
    TCasEnv( Dest ).Protocol := FProtocol;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TCommEnv }

procedure TCommEnv.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TCommEnv ) then
  begin
    TCommEnv( Dest ).DbRetryFreq := FDbRetryFreq;
    TCommEnv( Dest ).BusyTimeStart := FBusyTimeStart;
    TCommEnv( Dest ).BusyTimeEnd := FBusyTimeEnd;
    TCommEnv( Dest ).BusyTimeReadFreq := FBusyTimeReadFreq;
    TCommEnv( Dest ).NorTimeReadFreq := FNorTimeReadFreq;
    TCommEnv( Dest ).DbProcRecords := FDbProcRecords;
    TCommEnv( Dest ).DbWriteError := FDbWriteError;
    TCommEnv( Dest ).DbProcIPPV := FDbProcIPPV;
    TCommEnv( Dest ).DbProcBatch := FDbProcBatch;
    TCommEnv( Dest ).DbWaterMark := FDbWaterMark;
    TCommEnv( Dest ).CARetryFreq := FCARetryFreq;
    TCommEnv( Dest ).CACheckFreq := FCACheckFreq;
    TCommEnv( Dest ).CAEnableSend := FCAEnableSend;
    TCommEnv( Dest ).CAEnableRecv := FCAEnableRecv;
    TCommEnv( Dest ).CASendDelay := FCASendDelay;
    TCommEnv( Dest ).CAMaxError := FCAMaxError;
    TCommEnv( Dest ).CAProductDefine := FCAProductDefine;
    TCommEnv( Dest ).CAIdleTime := FCAIdleTime;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TNotify }

constructor TTextNotify.Create;
begin
  FFlag := 0;
  FText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TTextNotify.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TTextNotify ) then
  begin
    TTextNotify( Dest ).Text := FText;
    TTextNotify( Dest ).Flag := FFlag;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

{ TSo }

constructor TSo.Create;
begin
  FSelected := False;
  FRecordCount := 0;
  FErrorCount := 0;
  FErrorTime := 0;
  FDisplayNode := nil;
  FConnection := nil;
  FReader := nil;
  FWriter := nil;
  FBuffer := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TSo.Destroy;
begin
  FDisplayNode := nil;
  if Assigned( FBuffer ) then FBuffer.Free;
  if Assigned( FReader ) then FReader.Free;
  if Assigned( FWriter ) then FWriter.Free;
  if Assigned( FConnection ) then FConnection.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSo.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TSo ) then
  begin
    TSo( Dest ).Selected := FSelected;
    TSo( Dest ).PosName := FPosName;
    TSo( Dest ).SoType := FSoType;
    TSo( Dest ).CompCode := FCompCode;
    TSo( Dest ).CompName := FCompName;
    TSo( Dest ).NetworkId := FNetworkId;
    TSo( Dest ).AreaCode := FAreaCode;
    TSo( Dest ).DbUserId := FDbUserId;
    TSo( Dest ).DbUserPass := FDbUserPass;
    TSo( Dest ).DbAliase := FDbAliase;
    TSo( Dest ).DbState := FDbState;
    TSo( Dest ).DisplayNode := FDisplayNode;
    TSo( Dest ).RecordCount := FRecordCount;
  end else
    inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

procedure TSo.SetDbState(const Value: TDbState);
begin
  FDbState := Value;
  case FDbState of
    dbClose:
      begin
        FErrorCount := 0;
      end;
    dbOpen, dbActive:
      begin
        FErrorCount := 0;
        FErrorTime := 0;
      end;
    dbError:
      begin
        Inc( FErrorCount );
        FErrorTime := Now;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSoList }

procedure TSoList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if ( Action = lnDeleted ) and Assigned( Ptr ) then
    TSo( Ptr ).Free;
  inherited Notify( Ptr, Action );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.GetItem(Index: Integer): TSo;
begin
  Result := TSo( inherited Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoList.SetItem(Index: Integer; aObj: TSo);
begin
  inherited Items[Index] := aObj;
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Add(aObj: TSo): Integer;
begin
  Result := inherited Add( aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Extract(aObj: TSo): TSo;
begin
  Result := TSo( inherited Extract( aObj ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoList.Insert(Index: Integer; aObj: TSo);
begin
  inherited Insert( Index, aObj );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoList.Assign(aSource: TSoList);
var
  aIndex: Integer;
  aSo: TSo;
begin
  if not Assigned( aSource ) then Exit;
  Self.Clear;
  for aIndex := 0 to aSource.Count - 1 do
  begin
    aSo := TSo.Create;
    aSo.Assign( aSource[aIndex] );
    Self.Add( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Remove(aObj: TSo): Integer;
begin
  Result := inherited Remove( aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.IndexOf(aCompCode: String): Integer;
begin
  Result := 0;
  while ( Result < Count ) and ( TSo( List^[Result] ).CompCode <> aCompCode ) do
    Inc( Result );
  if Result = Count then
    Result := -1;
end;

{ ---------------------------------------------------------------------------- }

function TSoList.IndexOf(aObj: TSo): Integer;
begin
  Result := inherited IndexOf( aObj );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.First: TSo;
begin
  Result := TSo( inherited First );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.Last: TSo;
begin
  Result := TSo( inherited Last );
end;

{ ---------------------------------------------------------------------------- }

function TSoList.GetSelectCount: Integer;
var
  aIndex: Integer;
begin
  Result:= 0;
  for aIndex := 0 to Self.Count - 1 do
  begin
    if ( Self.Items[aIndex].Selected ) then
      Inc( Result );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TMessageQueueThread }

constructor TMessageThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FMsgHnd := Classes.AllocateHWnd( WndProc );
end;

{ ---------------------------------------------------------------------------- }

destructor TMessageThread.Destroy;
begin
  Classes.DeallocateHWnd( FMsgHnd );
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageThread.WndProc(var Msg: TMessage);
begin
  Msg.Result := DefWindowProc( FMsgHnd, Msg.Msg, Msg.wParam, Msg.lParam );
end;

{ ---------------------------------------------------------------------------- }

{ TAppThread }

constructor TAppThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FCasEnv := TCasEnv.Create;
  FCommEnv := TCommEnv.Create;
  FHighCmdEnv := TClientDataSet.Create( nil );
  FLowCmdEnv := TClientDataSet.Create( nil );
  FCmdErrCnv := TClientDataSet.Create( nil );
  FMainFormHandle := 0;
end;

{ ---------------------------------------------------------------------------- }

destructor TAppThread.Destroy;
begin
  FCmdErrCnv.Free;
  FLowCmdEnv.Free;
  FHighCmdEnv.Free;
  FCommEnv.Free;
  FCasEnv.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

{ TSubject }

constructor TSubject.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TSubject.Destroy;
begin
  RemoveObServer;
  FList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.AddObServer(aObServer: TObServer);
begin
  if FList.IndexOf( aObServer ) < 0 then
    FList.Add( aObServer );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer(aObServer: TObServer);
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aObServer );
  if aIndex >= 0 then
    FList.Delete( aIndex );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer;
begin
  while FList.Count > 0 do
    FList.Delete( 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.NotifyObServer(aObServer: TObServer);
var
  aIndex: Integer;
begin
  aIndex := FList.IndexOf( aObServer );
  if aIndex >= 0 then
  begin
    if not TObServer( FList[aIndex] ).Pause then
      TObServer( FList[aIndex] ).Update( Self );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.NotifyObServer;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
  begin
    if not TObServer( FList[aIndex] ).Pause then
      TObServer( FList[aIndex] ).Update( Self );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TObServer }

constructor TObServer.Create;
begin
  inherited Create;
  FPause := False;
  FSource := nil;
  FLock := TCriticalSection.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TObServer.Destroy;
begin
  FLock.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.CallOnUpdate;
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
    FOnUpdate( FSource );
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.Update(aSubject: TSubject);
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
  begin
    if ( GetCurrentThreadId <> MainThreadID ) then
    begin
      FLock.Enter;
      try
        FSource := aSubject;
        TThread.Synchronize( nil, CallOnUpdate )
      finally
        FLock.Leave;
      end;  
    end else
      FOnUpdate( aSubject );
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TTextSubject }

constructor TTextSubject.Create;
begin
  inherited Create;
  FNotify := TTextNotify.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TTextSubject.Destroy;
begin
  FNotify.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.InternalNotify(const aMsg: String;
  const aFlag: Integer);
begin
  FNotify.Text := aMsg;
  FNotify.Flag := aFlag;
  inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.Error(const aMsg: String);
begin
  InternalNotify( aMsg, MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.Info(const aMsg: String);
begin
  InternalNotify( aMsg, MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.OK(const aMsg: String);
begin
  InternalNotify( aMsg, MB_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.Warning(const aMsg: String);
begin
  InternalNotify( aMsg, MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.ErrorFmt(const aMsg: String;
  const Args: array of const);
begin
  InternalNotify( Format( aMsg, Args ), MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.InfoFmt(const aMsg: String;
  const Args: array of const);
begin
  InternalNotify( Format( aMsg, Args ), MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.OKFmt(const aMsg: String;
  const Args: array of const);
begin
  InternalNotify( Format( aMsg, Args ), MB_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure TTextSubject.WarningFmt(const aMsg: String;
  const Args: array of const);
begin
  InternalNotify( Format( aMsg, Args ), MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

{ TSoSubject }

procedure TSoSubject.Notify(aSo: TSo);
begin
  FSo := aSo;
  if Assigned( FSo ) then inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TSendDvnSubject }

procedure TSendDvnSubject.Notify(aSendDvn: TSendDvn);
begin
  FSendDvn := aSendDvn;
  if Assigned( FSendDvn ) then inherited NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TLanguageManager }

constructor TLanguageManager.Create;
begin
  inherited;
  FResNames := TStringList.Create;
  FResValues := TStringList.Create;
  FResNames.Sorted := True;
end;

{ ---------------------------------------------------------------------------- }

destructor TLanguageManager.Destroy;
begin
  FResNames.Free;
  FResValues.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TLanguageManager.Add(const aResName, aResValue: String);
var
  aIndex: Integer;
begin
  aIndex := FResNames.IndexOf( aResName );
  if ( aIndex < 0 ) then
  begin
    FResNames.AddObject( aResName, Pointer( FResValues.Count ) );
    FResValues.Add( aResValue );
  end else
    FResValues[Integer( FResNames.Objects[aIndex] )] := aResValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TLanguageManager.Clear;
begin
  FResNames.Clear;
  FResValues.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TLanguageManager.Get(const aResName: String): String;
var
  aIndex: Integer;
begin
  aIndex := FResNames.IndexOf( aResName );
  if ( aIndex >= 0 ) then
    Result := FResValues[Integer( FResNames.Objects[aIndex] )]
  else
    Result := aResName;
end;

{ ---------------------------------------------------------------------------- }

procedure TLanguageManager.LoadFromFile(const aFileName: String);
var
  aList: TStringList;
  aIndex, aPos: Integer;
  aName, aValue: String;
begin
  aList := TStringList.Create;
  try
    aList.LoadFromFile( aFileName );
    Self.Clear;
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
end;

{ ---------------------------------------------------------------------------- }

procedure TLanguageManager.LoadFromFile;
var
  aFileName, aLanExt: String;
  aLanIndex: Integer;
begin
  aLanIndex := Languages.IndexOf( SysLocale.DefaultLCID );
  aLanExt := Languages.Ext[aLanIndex];
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + Format( 'Lang.%s', [aLanExt] );
  Self.LoadFromFile( aFileName );
end;

{ ---------------------------------------------------------------------------- }

{ TSendDvn }

constructor TSendDvn.Create(const aCmdType: TCmdType);
begin
  inherited Create;
  FCmdType := aCmdType;
  Reset;
end;

{ ---------------------------------------------------------------------------- }

procedure TSendDvn.AssignTo(Dest: TPersistent);
begin
  if ( Dest is TSendDvn ) then
  begin
    TSendDvn( Dest ).UniqueId := FUniqueId;
    TSendDvn( Dest ).CompCode := FCompCode;
    TSendDvn( Dest ).CompName := FCompName;
    TSendDvn( Dest ).HighSeqNo := FHighSeqNo;
    TSendDvn( Dest ).FrameNo := FFrameNo;
    TSendDvn( Dest ).Version := FVersion;
    TSendDvn( Dest ).HighCmd := FHighCmd;
    TSendDvn( Dest ).LowCmd := FLowCmd;
    TSendDvn( Dest ).LowCmdCount := FLowCmdCount;
    TSendDvn( Dest ).LowCmdNote := FLowCmdNote;
    TSendDvn( Dest ).IccNo := FIccNo;
    TSendDvn( Dest ).StbNo := FStbNo;
    TSendDvn( Dest ).AreaCode := FAreaCode;
    TSendDvn( Dest ).Notes := FNotes;
    TSendDvn( Dest ).HighCmdStatus := FHighCmdStatus;
    TSendDvn( Dest ).LowCmdStatus := FLowCmdStatus;
    TSendDvn( Dest ).Operator := FOperator;
    TSendDvn( Dest ).UpdTime := FUpdTime;
    TSendDvn( Dest ).SendTime := FSendTime;
    TSendDvn( Dest ).RecvTime := FRecvTime;
    TSendDvn( Dest ).ErrCode := FErrCode;
    TSendDvn( Dest ).ErrMsg := FErrMsg;
    TSendDvn( Dest ).StartingTime := FStartingTime;
    TSendDvn( Dest ).ExpiryTime := FExpiryTime;
    TSendDvn( Dest ).ProductId := FProductId;
    TSendDvn( Dest ).Credit := FCredit;
    TSendDvn( Dest ).CreditAction := FCreditAction;
    TSendDvn( Dest ).CreditCfCode := FCreditCfCode;
    TSendDvn( Dest ).BankType := FBankType;
    TSendDvn( Dest ).ProductType := FProductType;
    TSendDvn( Dest ).ConfigValue := FConfigValue;
    TSendDvn( Dest ).MailAction := FMailAction;
    TSendDvn( Dest ).SNo := FSNo;
    TSendDvn( Dest ).ProcessingDate := FProcessingDate;
    TSendDvn( Dest ).SendCmdText := FSendCmdText;
    TSendDvn( Dest ).RecvCmdText := FRecvCmdText;
    TSendDvn( Dest ).ParentNode := nil;
    TSendDvn( Dest ).DisplayNode := nil;
    {}
    if ( TSendDvn( Dest ).CmdType = ctLow ) and ( Self.CmdType = ctHigh ) then
    begin
      TSendDvn( Dest ).ParentNode := FDisplayNode;
      TSendDvn( Dest ).DisplayNode := TSendDvn( Dest ).DisplayNode;
    end else
    if ( TSendDvn( Dest ).CmdType = ctLow ) and ( Self.CmdType = ctLow ) then
    begin
      TSendDvn( Dest ).ParentNode := FParentNode;
      TSendDvn( Dest ).DisplayNode := FDisplayNode;
    end else
    begin
      TSendDvn( Dest ).ParentNode := FParentNode;
      TSendDvn( Dest ).DisplayNode := FDisplayNode;
    end;
    {}
    TSendDvn( Dest ).IsBatch := FIsBatch;
    TSendDvn( Dest ).Ack := FAck;
    {}
    { ConsolNode 不須要 Clone }
    { TSendDvn( Dest ).ConsoleNode := FConsoleNode }
  end else
  inherited AssignTo( Dest );
end;

{ ---------------------------------------------------------------------------- }

procedure TSendDvn.Reset;
begin
  FIsBatch := False;
  FCredit := 0;
  FLowCmdCount := 0;
  LowCmdNote := EmptyStr;
  FFrameNo := EmptyStr;
  FHighSeqNo := EmptyStr;
  FParentNode := nil;
  FDisplayNode := nil;
  FConsoleNode := nil;
  FStartingTime := EmptyStr;
  FCreditCfCode := EmptyStr;
  FCompCode := EmptyStr;
  FExpiryTime := EmptyStr;
  FProductId := EmptyStr;
  FRecvCmdText := EmptyStr;
  FHighCmdStatus := 'W';
  FConfigValue := EmptyStr;
  FProductType := EmptyStr;
  FUniqueId := EmptyStr;
  FSNo := EmptyStr;
  FIccNo := EmptyStr;
  FLowCmdStatus := 'W';
  FOperator := EmptyStr;
  FErrMsg := EmptyStr;
  FSendCmdText := EmptyStr;
  FStbNo := EmptyStr;
  FAreaCode := EmptyStr;
  FLowCmd := EmptyStr;
  FErrCode := EmptyStr;
  FBankType := EmptyStr;
  FProcessingDate := EmptyStr;
  FHighCmd := EmptyStr;
  FCreditAction := EmptyStr;
  FMailAction := EmptyStr;
  FNotes := EmptyStr;
  FRecvTime := 0;
  FUpdTime := 0;
  FSendTime := 0;
  FAck := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

{ TDvnAckParser }

constructor TDvnAckParser.Create;
begin
  inherited Create;
  FList := TStringList.Create;
  FList.Delimiter := ';';
  FCmdErrCnv := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

destructor TDvnAckParser.Destroy;
begin
  FList.Free;
  FCmdErrCnv.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnAckParser.Parse(const aAckObj: TSendDvn);
var
  aPos1, aPos2: Integer;
  aData: String;
begin
  if not Assigned( aAckObj ) then Exit;
  if ( aAckObj.RecvCmdText = EmptyStr ) then Exit;
  aPos1 := AnsiPos( '(', aAckObj.RecvCmdText );
  aPos2 := AnsiPos( ')', aAckObj.RecvCmdText );
  if ( aPos1 > 0 ) and ( aPos2 > 0 ) then
  begin
    aData := Copy( aAckObj.RecvCmdText, aPos1 + 1, aPos2 - 1 );
    FList.DelimitedText := aData;
    aAckObj.UniqueId := FList[0];
    aAckObj.FrameNo := FList[0];
    aAckObj.LowCmd := FList[3];
    aAckObj.ErrMsg := EmptyStr;
    aAckObj.ErrCode := EmptyStr;
    aAckObj.LowCmdStatus := 'C';
    aAckObj.Ack := FList[4];
    if FCmdErrCnv.Locate( 'ErrorCode', aAckObj.Ack, [loCaseInsensitive] ) then
    begin
      if ( FCmdErrCnv.FieldByName( 'ErrorFlag' ).AsInteger = 1 ) then
      begin
        aAckObj.ErrCode := FCmdErrCnv.FieldByName( 'ErrorCode' ).AsString;
        aAckObj.ErrMsg := FCmdErrCnv.FieldByName( 'ErrorDesc' ).AsString;
        aAckObj.LowCmdStatus := 'E';
      end;
    end else
    begin
      aAckObj.ErrCode := Nvl( aAckObj.Ack, LanguageManager.Get( 'SNull' ) );
      aAckObj.ErrMsg := LanguageManager.Get( 'SUnknowAck' );
      aAckObj.LowCmdStatus := 'E';
    end;
    FList.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TSyncList }

constructor TSyncList.Create;
begin
  inherited Create;
  FLock := TMREWSync.Create;
  FList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TSyncList.Destroy;
begin
  FLock.BeginWrite;
  try
    FList.Free;
    inherited Destroy;
  finally
    FLock.EndWrite;
    FLock.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.Add(const UniqueId: string): Integer;
begin
  FLock.BeginWrite;
  try
    Result := FList.Add( UniqueId );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.AddObject(const UniqueId: string; AObject: TSendDvn): Integer;
begin
  FLock.BeginWrite;
  try
    Result := FList.AddObject( UniqueId, AObject);
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Clear;
begin
  FLock.BeginWrite;
  try
    FList.Clear;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Delete(AObject: TSendDvn);
begin
  FLock.BeginWrite;
  try
    FList.Delete( FList.IndexOf( AObject.UniqueId ) );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Delete(Index: Integer);
begin
  FLock.BeginWrite;
  try
    FList.Delete( Index );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.IndexOf(const UniqueId: string): Integer;
begin
  FLock.BeginRead;
  try
    Result := FList.IndexOf( UniqueId );
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.IndexOfObject(AObject: TSendDvn): Integer;
begin
  FLock.BeginRead;
  try
    Result := FList.IndexOfObject( AObject );
  finally
    FLock.EndRead;
  end;
end;
{ ---------------------------------------------------------------------------- }

function TSyncList.GetSorted: Boolean;
begin
  Result := FList.Sorted;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.SetSorted(const Value: Boolean);
begin
  FLock.BeginWrite;
  try
    FList.Sorted := Value;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.Get(Index: Integer): string;
begin
  FLock.BeginRead;
  try
    Result := FList[Index];
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.GetObject(Index: Integer): TSendDvn;
begin
  FLock.BeginRead;
  try
    Result := TSendDvn( FList.Objects[Index] );
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.Put(Index: Integer; const UniqueId: string);
begin
  FLock.BeginWrite;
  try
    FList[Index] := UniqueId;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.PutObject(Index: Integer; AObject: TSendDvn);
begin
  FLock.BeginWrite;
  try
    FList.Objects[Index] := AObject;
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.FreeObject(Index: Integer);
begin
  FLock.BeginWrite;
  try
    TSendDvn( FList.Objects[Index] ).Free;
    FList.Delete( Index );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSyncList.FreeObject(AObject: TSendDvn);
begin
  FLock.BeginWrite;
  try
    FList.Delete( FList.IndexOfObject( AObject ) );
    FreeAndNil( AObject );
  finally
    FLock.EndWrite;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSyncList.GetCount: Integer;
begin
  FLock.BeginRead;
  try
    Result := FList.Count;
  finally
    FLock.EndRead;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TDvnCommandBuilder }

constructor TDvnCommandBuilder.Create;
var
  aIniLoader: TCurValueLoader;
  aGenerateRule: TNumberGenerateRule;
  aFileName: String;
begin
  inherited Create;
  FCasEnv := TCasEnv.Create;
  FCommEnv := TCommEnv.Create;
  FTempList := TStringList.Create;
  FCommandList := TStringList.Create;
  FHighCmd := nil;
  FHighCmdEnv := TClientDataSet.Create( nil );
  FLowCmdEnv := TClientDataSet.Create( nil );
  {}
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'Sequence.ini';
  {}
  aIniLoader := TIniCurValueLoader.Create( aFileName, 'Sequence', 'FrameNo' );
  { frameno range 0 - 999999999 }
  aGenerateRule := TNumberGenerateRule.Create( 1, 999999999, aIniLoader );
  FFrameNoGenerator := TIdGenerator.Create( False, aGenerateRule );
  {}
  aIniLoader := TIniCurValueLoader.Create( aFileName, 'Sequence', 'CreditCfCode' );
  aGenerateRule := TNumberGenerateRule.Create( 1, 999999, aIniLoader );
  FCreditCfCodeGenerator := TIdGenerator.Create( False, aGenerateRule );
end;

{ ---------------------------------------------------------------------------- }

destructor TDvnCommandBuilder.Destroy;
begin
  FCasEnv.Free;
  FCommEnv.Free;
  FFrameNoGenerator.SaveCurValue;
  FCreditCfCodeGenerator.SaveCurValue;
  FTempList.Clear;
  FCommandList.Clear;
  FHighCmdEnv.Free;
  FLowCmdEnv.Free;
  FFrameNoGenerator.Free;
  FCreditCfCodeGenerator.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.BuildLowCommand(aHighCmd: TSendDvn);
begin
  FHighCmd := aHighCmd;
  HighCmdExtractToLowCmd;
  if IsSpecialHighCmd( FHighCmd.HighCmd ) then
    ParseSpecial
  else
    ParseNormal;
  MoveList;
  AssemblyCommand;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.Clear;
begin
  FCommandList.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.GetLowCmd(Index: Integer): TSendDvn;
begin
  Result := TSendDvn( FCommandList.Objects[Index] );
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.GetLowCmdCount: Integer;
begin
  Result := FCommandList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.GetValue(const aFindField,
  aResultField: String; const Args: array of Variant): String;
begin
  Result := EmptyStr;
  FHighCmdEnv.First;
  if FHighCmdEnv.Locate( aFindField, VarArrayOf( Args ), [] ) then
    Result := FHighCmdEnv.FieldByName( aResultField ).AsString;
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.IsSpecialHighCmd(const aHighCmdId: String): Boolean;
begin
  Result := ( aHighCmdId = 'B8' );
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.HighCmdExtractToLowCmd;
var
  aNumber: Integer;
  aValue: String;
  aLowCmd: TSendDvn;
begin
  aNumber := 0;
  aValue := GetValue( 'highlevelcmd', 'lowlevelcmd', [FHighCmd.HighCmd] );
  FTempList.Clear;
  repeat
    aLowCmd := TSendDvn.Create( ctLow );
    { 將所有高階指令的參數, 抄一份給低階指令 }
    aLowCmd.Assign( FHighCmd );
    aLowCmd.LowCmd := ExtractValue( aValue );
    { 先加到 Temp List, 稍後再處理, Temp List 先將每個低階指令以 100 做間隔,
      先編號碼, Ex : 100, 200, 300 .....  }
    Inc( aNumber, 100 );
    FTempList.AddObject( Lpad( IntToStr( aNumber ), 5, '0' ), aLowCmd );
  until ( aValue = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.MoveList;
var
  aIndex: Integer;
begin
  FCommandList.Clear;
  { 將 temp list 搬移至正式傳指令 list 裏 }
  for aIndex := 0 to FTempList.Count - 1 do
    FCommandList.AddObject( FTempList[aIndex], FTempList.Objects[aIndex] );
  FTempList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.ParseNormal;
var
  aIndex, aShift, aNumber: Integer;
  aLowCmd, aLowCmd2: TSendDvn;
  aNoteText: String;
  aIsFirst: Boolean;
begin
  aIndex := 0;
  while ( aIndex < FTempList.Count ) do
  begin
    aShift := 0;
    aNumber := StrToInt( FTempList[aIndex] );
    aLowCmd := TSendDvn( FTempList.Objects[aIndex] );
    { Notes 欄位有填值, 就是要再多拆低階指令 }
    { 發送訊息的指令不用拆解 }
    if ( aLowCmd.Notes <> EmptyStr ) and ( aLowCmd.HighCmd <> 'E2' ) then
    begin
      { 假如該低階指令是使用 Notes 的話 }
      if ( GetValue( 'highlevelcmd;cmdbynote', 'cmdbynote', [
        aLowCmd.HighCmd, aLowCmd.LowCmd] ) <> EmptyStr ) then
      begin
        aNoteText := aLowCmd.Notes;
        aIsFirst := True;
        repeat
          { 本身就是第一個低階, 直接將第一個解出來的 note 抄進去 }
          if ( aIsFirst ) then
          begin
            aLowCmd.Notes := ExtractValue( aNoteText );
            aIsFirst := False;
          end else
          { 第 2 個以後, 就必須要產生新的低階指令 }
          begin
            aLowCmd2 := TSendDvn.Create( ctLow );
            aLowCmd2.Assign( aLowCmd );
            aLowCmd2.Notes := ExtractValue( aNoteText );
            Inc( aNumber );
            FTempList.InsertObject( aIndex + 1, Lpad( IntToStr( aNUmber ), 5, '0' ), aLowCmd2 );
            { 位移 --> 在原本指令後插入了幾個新的低階指令 }
            Inc( aShift );
          end;
        until ( aNoteText = EmptyStr );
      end else
        aLowCmd.Notes := EmptyStr;
    end;
    { 換下一個指令, 必須跳過從中間插入的 }
    Inc( aIndex, aShift + 1 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.ParseSpecial;

   { -------------------------------  }

   function Find(const aLowCmdId: String): Integer;
   var
     aIndex: Integer;
   begin
     Result := -1;
     for aIndex := 0 to FTempList.Count - 1 do
     begin
       if ( TSendDvn( FTempList.Objects[aIndex] ).LowCmd = aLowCmdId ) then
       begin
         Result := aIndex;
         Break;
       end;
     end;
   end;

   { -------------------------------  }

var
  aNumber, aCmdIndex, aIndex: Integer;
  aLowCmd, aLowCmd2: TSendDvn;
  aNoteText, aValue: String;
  aIsFirst: Boolean;
begin
  // B8- 熱重開頻道
  if ( FHighCmd.HighCmd = 'B8' ) then
  begin
    // B8 --> 1.REMOVE_PRODUCT  2.ADD_PRODUCT
    for aIndex := 1 to 2 do
    begin
      if aIndex = 1 then
        aCmdIndex := Find( 'REMOVE_PRODUCT' )
      else
        aCmdIndex := Find( 'ADD_PRODUCT' );
      aLowCmd := TSendDvn( FTempList.Objects[aCmdIndex] );
      aNumber := StrToInt( FTempList[aCmdIndex] );
      aNoteText := aLowCmd.Notes;
      aIsFirst := True;
      repeat
        aValue := ExtractValue( aNoteText );
        if ( aIsFirst ) then
        begin
          aLowCmd.Notes := aValue;
          aIsFirst := False;
        end else
        begin
          aLowCmd2 := TSendDvn.Create( ctLow );
          aLowCmd2.Assign( aLowCmd );
          aLowCmd2.Notes := aValue;
          Inc( aNumber );
          Inc( aCmdIndex );
          FTempList.InsertObject( aCmdIndex, Lpad( IntToStr( aNumber ), 5, '0' ), aLowCmd2 );
        end;
      until ( aNoteText = EmptyStr );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.AssemblyCommand;
var
  aIndex, aSeq: Integer;
  aObj: TSendDvn;
  aCode, aHeader, aParam, aCmdNote: String;
begin
  aCmdNote := EmptyStr;
  for aIndex := 0 to FCommandList.Count - 1 do
  begin
    aObj := TSendDvn( FCommandList.Objects[aIndex] );
    { 指令序號編號 }
    aSeq := FFrameNoGenerator.NextValue;
    aObj.FrameNo := Format( '0-%d', [aSeq] );
    aCmdNote := ( aCmdNote + aObj.FFrameNo + ',' );
    { CAS 版本 }
    aObj.Version := FCasEnv.Protocol;
    { 將參數搬到定位 }
    EachLowCmd( aObj );
    { 指令格式 }
    { COMMAND_CODE(FrameNo;Version;STB_ID;SmartCardNo;StartingTime;ExpiryTime;Parameters) }
    aCode := GetCommandCode( aObj );
    aHeader := GetCommandHeader( aObj );
    aParam := GetCommandParam( aObj );
    { 重新給予 UniqueId, 該低階指令現在的 unique id 還是高階指令的 SeqNo }
    aObj.UniqueId := aObj.FrameNo;
    aObj.SendCmdText := Format( '%s(%s%s)', [aCode, aHeader, aParam] );
  end;
  Delete( aCmdNote, Length( aCmdNote ), 1 );
  FHighCmd.LowCmdNote := aCmdNote;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnCommandBuilder.EachLowCmd(aObj: TSendDvn);
var
  aNote, aAddType: String;
begin
  if ( aObj.LowCmd = 'PAIRING_STB' ) or
     ( aObj.LowCmd = 'MAIL_MESSAGE' ) then
  begin
    aObj.StartingTime := FormatDateTime( 'yyyy/mm/dd', Now ) + ' 00:00:00';
    aObj.ExpiryTime := EmptyStr;
  end else
    { 因為 ADD_TOKEN 及 DEDUCT_TOKEN 的 StartingTime
      表示是 Transaction Time, 所以後面時分秒用實際值代入 }
  if ( aObj.LowCmd = 'ADD_TOKEN' ) or
     ( aObj.LowCmd = 'DEDUCT_TOKEN' ) then
  begin
    aObj.StartingTime := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
    aObj.ExpiryTime := EmptyStr;
  end;   
  if ( aObj.LowCmd = 'STB_ON' ) or
     ( aObj.LowCmd = 'STB_OFF' ) or
     ( aObj.LowCmd = 'USER_DEFINED_D1' ) or
     ( aObj.LowCmd = 'USER_DEFINED_D2' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_ZIPCODE' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_NETWORKID' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_AREACODE' ) or
     ( aObj.LowCmd = 'UPLOAD_RECORDS' ) then
  begin
    aObj.StartingTime := EmptyStr;
    aObj.ExpiryTime := EmptyStr;
  end else
  if ( aObj.LowCmd = 'ADD_PRODUCT' ) then
  begin
    aNote := aObj.Notes;
    aAddType := ExtractValue( aNote, '~' );
    aObj.ProductId := ExtractValue( aNote, '~' );
    if ( aAddType = 'B' ) then
    begin
      aObj.StartingTime := FormatMaskText( '####/##/##;0;_', ExtractValue( aNote, '~' ) ) + ' 00:00:00';
      aObj.ExpiryTime := FormatMaskText( '####/##/##;0;_', aNote ) + ' 23:59:59';
    end;
    aObj.Notes := EmptyStr;
  end else
  { REMOVE_PRODUCT 的 StartTime 跟 ExpireTime 不可以相同 }
  if ( aObj.LowCmd = 'REMOVE_PRODUCT' ) then
  begin
    aObj.StartingTime := FormatDateTime( 'yyyy/mm/dd', Now ) + ' 00:00:00';
    aObj.ExpiryTime := FormatDateTime( 'yyyy/mm/dd', Now ) + ' 00:00:01';
    aNote := aObj.Notes;
    if ( aObj.HighCmd = 'B8' ) then
    begin
      aAddType := ExtractValue( aNote, '~' );
      aObj.ProductId := ExtractValue( aNote, '~' );
    end else
    begin
      aObj.ProductId := aObj.Notes;
    end;
    aObj.Notes := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.GetCommandCode(const aObj: TSendDvn): String;
begin
  if ( aObj.LowCmd = 'USER_DEFINED_D1' ) or
     ( aObj.LowCmd = 'USER_DEFINED_D2' ) then
    Result := 'USER_DEFINED'
  else
  if ( aObj.LowCmd = 'SYS_CONFIG_ZIPCODE' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_NETWORKID' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_AREACODE' ) then
    Result := 'SYS_CONFIG'
  else
    Result := aObj.LowCmd;
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.GetCommandHeader(const aObj: TSendDvn): String;
var
  aBoradCase: String;
begin
  if ( aObj.LowCmd = 'PAIRING_STB' ) or
     ( aObj.LowCmd = 'ADD_PRODUCT' ) or
     ( aObj.LowCmd = 'REMOVE_PRODUCT' ) or
     ( aObj.LowCmd = 'ADD_TOKEN' ) or
     ( aObj.LowCmd = 'DEDUCT_TOKEN' ) or
     ( aObj.LowCmd = 'USER_DEFINED_D1' ) or
     ( aObj.LowCmd = 'USER_DEFINED_D2' ) then
  begin
    Result :=
      ( aObj.FrameNo + ';' +
        aObj.Version + ';' +
        '0-' + '0-0-0-0' + '-' + aObj.StbNo + ';' +
        aObj.IccNo + ';' +
        aObj.StartingTime + ';' +
        aObj.ExpiryTime + ';' );
  end else
  { UPLOAD_RECORDS --> 這個DVN 的指令, 在 4.0.6 的 API 手冊被拿掉,
    所以是看 4.0.4 版本做的 }
  if ( aObj.LowCmd = 'STB_ON' ) or
     ( aObj.LowCmd = 'STB_OFF' ) or
     ( aObj.LowCmd = 'UPLOAD_RECORDS' ) then
  begin
    Result :=
      ( aObj.FrameNo + ';' +
        aObj.Version + ';' +
        '0-' + '0-0-0-0' + '-' + aObj.StbNo + ';' +
        aObj.IccNo + ';' +
        aObj.StartingTime + ';' +
        aObj.ExpiryTime );
  end else
  { 不須要 StartingTime 跟 ExpiryTime }
  if ( aObj.LowCmd = 'SYS_CONFIG_ZIPCODE' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_NETWORKID' ) or
     ( aObj.LowCmd = 'SYS_CONFIG_AREACODE' ) then
  begin
    Result :=
      ( aObj.FrameNo + ';' +
        aObj.Version + ';' +
        '0-' + '0-0-0-0' + '-' + aObj.StbNo + ';' +
        aObj.IccNo + ';' );
  end else
  if ( aObj.LowCmd = 'MAIL_MESSAGE' ) then
  begin
    aBoradCase := Copy( aObj.MailAction, 1, 1 );
    if ( aBoradCase <> '0' ) or ( aBoradCase <> '1' ) then aBoradCase := '0';
    Result :=
      ( aObj.FrameNo + ';' +
        aObj.Version + ';' +
        aBoradCase + '-' + '0-0-0-0' + '-' + aObj.StbNo + ';' +
        aObj.IccNo + ';' +
        aObj.StartingTime + ';' +
        aObj.ExpiryTime + ';' );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDvnCommandBuilder.GetCommandParam(const aObj: TSendDvn): String;
var
  aPassword, aClientName, aAccountNo, aTelephoneNo, aEMailId, aAddress: String;
  aBankType, aProductType, aNumberOfProduct, aGatewayName: String;
  aAccountCode: String;
  aPos: Integer;
begin
  if ( aObj.LowCmd = 'PAIRING_STB' ) then
  begin
    // 開機指令須要一個參數:Password
    // 等同於 STB NO
    aPassword := ( '0' + '-' + '0-0-0-0' + '-' + aObj.StbNo );
    { 以下可以不用填, DVN 說的 }
    aClientName := EmptyStr;
    { PARING_STB 的 Note 欄位填的是客編 }
    // 2007/2/15 會議決定, AccountNo 不用填客編, 因為在倉管開機時,
    // 還不知道會被裝到那個客戶手上, 所以客編沒辦法預先知道
    //aAccountNo := aObj.Notes;
    aAccountNo := EmptyStr;
    aTelephoneNo := EmptyStr;
    aEMailId := EmptyStr;
    aAddress := EmptyStr;
    Result := (
      aPassword + ';' +
      aClientName + ';' +
      aAccountNo + ';' +
      aTelephoneNo + ';' +
      aEMailId + ';' +
      aAddress );
  end else
  { 這 2 個不用參數 }
  if ( aObj.LowCmd = 'STB_ON' ) or
     ( aObj.LowCmd = 'STB_OFF' ) or
     ( aObj.LowCmd = 'UPLOAD_RECORDS' ) then
  begin
    Result := EmptyStr;
  end else
  { ADD 跟 REMOVE 參數一樣 }
  if ( aObj.LowCmd = 'ADD_PRODUCT' ) or
     ( aObj.LowCmd = 'REMOVE_PRODUCT' ) then
  begin
    aBankType := '1';
    if ( aObj.BankType <> EmptyStr ) then aBankType := aObj.BankType;
    aProductType := EmptyStr;
    if ( aObj.ProductType = EmptyStr ) then
      aProductType := FCommEnv.CAProductDefine
    else
      aProductType := aObj.ProductType;
    {}  
    aNumberOfProduct := '1';
    aGatewayName := EmptyStr;
    Result := (
      aBankType + ';' +
      aProductType + ';' +
      { 注意, 這裏是用逗號分隔 }
      aNumberOfProduct + ',' +
      aObj.ProductId + ';' +
      aGatewayName );
  end else
  if ( aObj.LowCmd = 'MAIL_MESSAGE' ) then
  begin
    Result := Copy( aObj.MailAction, 3, 1 ) + Copy( aObj.Notes, 1, 99 );
  end else
  if ( aObj.LowCmd = 'ADD_TOKEN' ) or
     ( aObj.LowCmd = 'DEDUCT_TOKEN' ) then
  begin
    { AccountCode 不須給值 }
    aAccountCode := EmptyStr;
    { aObj.CreditCfCode 有值表示須要重送 Credit }
    { 沒有值, 表示要加新的 Credit 或扣 Credit, 這時候 ConfirmCode
       必須由 Gateway 產生序號 }
    if ( aObj.CreditCfCode = EmptyStr ) then
    begin
      aObj.CreditCfCode :=
        Lpad( aObj.CompCode, 2, '0' ) +
        FormatDateTime( 'yyyymmdd', Date ) +
        Lpad( FCreditCfCodeGenerator.NextValue, 6, '0' );
      { 高階指令完成後, 必須回填 CreditCfCode 到 Send_DVN 介接 Table }  
      FHighCmd.CreditCfCode := aObj.CreditCfCode;
    end;
    Result := (
      aAccountCode + ';' +
      aObj.CreditAction + ';' +
      aObj.CreditCfCode + ';' +
      FloatToStr( aObj.Credit ) );
  end else
  if ( aObj.LowCmd = 'USER_DEFINED_D1' ) or
     ( aObj.LowCmd = 'USER_DEFINED_D2' ) then
  begin
    aPos := LastDelimiter( '_', aObj.LowCmd );
    Result := Copy( aObj.LowCmd, aPos + 1, Length( aObj.LowCmd ) - aPos );
  end else
  if ( aObj.LowCmd = 'SYS_CONFIG_AREACODE' ) then
  begin
    Result := '2;1,' + aObj.AreaCode + ';';
  end else
  if ( aObj.LowCmd = 'SYS_CONFIG_ZIPCODE' ) then
  begin
    Result := '6;1,' + aObj.ConfigValue + ';';
  end else
  if ( aObj.LowCmd = 'SYS_CONFIG_NETWORKID' ) then
  begin
    Result := '3;1,' + aObj.ConfigValue + ';';
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TNumberGenerateRule }

constructor TNumberGenerateRule.Create(aMinValue, aMaxValue,
  aCurValue: Integer);
begin
  inherited Create;
  FMinValue := aMinValue;
  FMaxValue := aMaxValue;
  FCurValue := aCurValue;
end;

{ ---------------------------------------------------------------------------- }

constructor TNumberGenerateRule.Create(aMinValue, aMaxValue: Variant;
  aLoader: TCurValueLoader);
begin
  inherited Create;
  FMinValue := aMinValue;
  FMaxValue := aMaxValue;
  FLoader := aLoader;
  FCurValue := FLoader.Restore;
end;

{ ---------------------------------------------------------------------------- }

destructor TNumberGenerateRule.Destroy;
begin
  FLoader.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetCurValue: Variant;
begin
  Result := FCurValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetMaxValue: Variant;
begin
  Result := FMaxValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetMinValue: Variant;
begin
  Result := FMinValue;
end;

{ ---------------------------------------------------------------------------- }

function TNumberGenerateRule.GetNextValue: Variant;
begin
  if ( FCurValue >= FMaxValue ) then
    ResetValue
  else begin
    Inc( FCurValue );
  end;
  Result := FCurValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TNumberGenerateRule.ResetValue;
begin
  FCurValue := FMinValue;
end;

{ ---------------------------------------------------------------------------- }

{ TIdGenerator }

constructor TIdGenerator.Create(const aThreadSafe: Boolean;
  aRule: TGenerateRule);
begin
  inherited Create;
  FThreadSafe := aThreadSafe;
  if ( FThreadSafe ) then FLock := TCriticalSection.Create;
  FRule := aRule;
end;

{ ---------------------------------------------------------------------------- }

destructor TIdGenerator.Destroy;
begin
  if ( FThreadSafe ) then FLock.Free;
  FRule.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetCurValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FRule.GetCurValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetMaxValue: Variant;
begin
  Result := FRule.GetMaxValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetMinValue: Variant;
begin
  Result := FRule.GetMinValue;
end;

{ ---------------------------------------------------------------------------- }

function TIdGenerator.GetNextValue: Variant;
begin
  if ( FThreadSafe ) then FLock.Enter;
  try
    Result := FRule.GetNextValue;
  finally
    if ( FThreadSafe ) then FLock.Leave;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TIdGenerator.SaveCurValue;
begin
  if ( FRule is TNumberGenerateRule ) then
    TNumberGenerateRule( FRule ).Loader.Save( FRule.GetCurValue );
end;

{ ---------------------------------------------------------------------------- }

{ TIniCurValueLoader }

constructor TIniCurValueLoader.Create(const aFileName, aSection, aIdent: String);
begin
  FIni := TIniFile.Create( aFileName );
  FSection := aSection;
  FIdent := aIdent;
end;

{ ---------------------------------------------------------------------------- }

destructor TIniCurValueLoader.Destroy;
begin
  FIni.Destroy;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TIniCurValueLoader.Restore: Variant;
begin
  Result := FIni.ReadString( FSection, FIdent, '0' );
end;

{ ---------------------------------------------------------------------------- }

procedure TIniCurValueLoader.Save(const aCurValue: Variant);
begin
  FIni.WriteString( FSection, FIdent, VarToStrDef( aCurValue, '0' ) );
  FIni.UpdateFile;
end;

{ ---------------------------------------------------------------------------- }

initialization
  LanguageManager := TLanguageManager.Create;

finalization
  FreeAndNil( LanguageManager );
end.



