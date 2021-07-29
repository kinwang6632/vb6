unit cbClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Contnrs, SyncObjs, IniFiles,
  DB, ADODB, DBClient,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;


const

  WM_PLAY = WM_USER + 110;
  WM_STOP = WM_USER + 120;
  WM_SOCKET = WM_USER + 130;
  WM_DATABASE = WM_USER + 140;
  WM_UPDATEGUI = WM_USER + 150;

  DEFAULT_CONFIGFILE_EXT = '.CFG';
  DEFAULT_TRANSACTIONFILE = 'TranNum.ini';
  DEFAULT_APPFILE = 'AppInfo.ini';
  DEFAULT_DOCKLAYOUTFILE = 'DockLayout.ini';

  COMMON_DELAY = 100;
  CONTROL_COMPCODE = '98';
  FEEDBACK_COMPCODE = '99';


type

  TWMSocket = packed record
    Msg: Cardinal;
    ThreadType: Byte;
    NotifyCommandType: Byte;
    Unused: Word;
    ErrorCode: Smallint;
    ErrorMsg: Smallint;
    Result: Longint;
  end;

  TWMDataBase = packed record
    Msg: Cardinal;
    So: Longint;
    Unused: Longint;
    Result: Longint;
  end;


  TDbConnectStatus = ( dbNoSelect, dbNone, dbOK, dbError, dbActive );

  TThreadType = ( tdControlDb, tdFeedbackDb, tdControlSend, tdControlRecv, tdFeedbackSend, tdFeedbackRecv );

  TRunState = ( rsStartup, rsRunning, rsStop );

  TNotifyCommandType = ( ncCommand, ncCommunication, ncAck, ncNack, ncSocketError );

  { �{���Ұʰ����T }

  PAppStartupInfo = ^TAppStartupInfo;
  TAppStartupInfo = record
    AutoExecute: Boolean;                { �۰ʰ��� }
    StartupAssignedConfigFile: String;   { �Ұʮ�, �ǤJ�����ҳ]�w�� }
  end;

  { �t�Υx��T }

  PSoInfo = ^TSoInfo;
  TSoInfo = record
    Selected: Boolean;                 { �O�_���  }
    Pos: Integer;                      { ���x�Ϥ� }
    PosName: String;                   { ���x�W�� }
    SoType: String;                    { �ǰe���O���t�Υx�Φ^�Ǿ���t�Υx}
    CompCode: String;                  { �t�Υx�N�X }
    CompName: String;                  { �t�Υx�W�� }
    NetworkId: String;                 { Netword Id }
    LoginUser: String;                 { �n�J��Ʈw�b�� }
    LoginPass: String;                 { �n�J��Ʈw�K�X }
    DbAliase: String;                  { �n�J��Ʈw }
    DbConnectStatus: TDbConnectStatus; { ��Ʈw�s�u���A }
    RecordCount: Integer;              { �q�t�ΥxŪ�����֭p��Ƶ��� }
    BufferCount: Integer;              { �b�ݳB�z���O��C�����O�� }
    Data: Pointer;                     { �Өt�Υx���`�I }
    NotifyType: Integer;               { �q���T���˦�,ĵ�i,���~�� }
    NotifyMsg: String;                 { �q�����T�� }
    CriticalErrorCount: Integer;       { ��Ʈw���~���� }
    LastCriticalErrorTime: TDateTime;  { �̫�@����Ʈw���~�ɶ� }
  end;


  { �U�t�Υx����Ʈw�s�����n�ϥΨ쪺 ADO ���� }
  PSoADOControl = ^TSoADOControl;
  TSoADOControl = record
    CompCode: String;
    Connection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    DataSet: TClientDataSet;
  end;

  { STB �u�{�K�X:007911 }

  { Nagra CAS �]�w }

  { 2006/2/20
    EMC Labe �i�� Source ID  1,2,3,4,5,6,7,8,9,10,50,256,1234 }

  { �`�N, �{���بϥ� CompCode ���ѧO�O�_�O�^�Ǹ�ƪ� Source Id,
    Ex: �p�G CompCode = FEEDBACK_COMPCODE = '99', ��ܬ� �^�Ǹ��,
        ���� Build �C�@�����O�� Header ��, ���Өϥ�
        CASEnv.FeedbackSourceId,
        CASEnv.FeedbackDestId,
        CASEnv.FeedbackMopPPId }

  { �]���ǰe���O���C�@�����O Header, �����@�� Source Id �N�X,
    �@�� Source Id �u�঳�@�ӳs�W�h�� TCP/IP Conneciton �Ӷǫ��O�򦬦^�Ǹ�� }

  { EMC �{�b Callback �^�Ӫ� Nagra SourcId �T�w�]������ 10 }

  { 2006/07/26 EMC�b�s�x�_�����l, ��e��v�O���b�s�x�_�t�Υx�U
    �쥻 networkid �O 2,  �i�O��e��v�O���Ʀ�H��, �ҥH networkid �O 10  }  

  PCASEnv = ^TCASEnv;
  TCASEnv = record
    IP: String;                        { CAS IP }
    ControlPort: Integer;              { �ǰe���O�� Port }
    FeedbackPort: Integer;             { ReturnPath �� Port }
    ControlSourceId: String;           { CAS �W�� �ǰe���O�Ϊ� Source Id }
    ControlDestId: String;             { CAS �W�� �ǰe���O�Ϊ� Dest Id }
    ControlMopPPId: String;            { CAS �W�� �ǰe���O�Ϊ� Mop PPID }
    FeedbackSourceId: String;          { * CAS �W�� �����^�Ǹ�ƥΪ� Source Id }
    FeedbackDestId: String;            { * CAS �W�� �����^�Ǹ�ƥΪ� Dest Id }
    FeedbackMopPPId: String;           { * CAS �W�� �����^�Ǹ�ƥΪ� Mop PPID }
    BoardCastMode: String;             { CAS �W�s����, N = Normal, B = Batch }
    AddressType: String;               { EMM ���O, U = Unique ICC No, G = Global }
    ControlTransId: String;            { �ǰe���O�ΨӲ��� Transsaction number ���ѧO�X }
    FeedbackTransId: String;           { �^�Ǿ���ΨӲ��� Transsaction number ���ѧO�X }
    MailId: String;                    { Ird Command �إΨӶǰe Mail �T�����ѧO�X }
    CmdMaxCounter: Integer;            { �ǰe�� CAS �����O�W���p�� }
  end;

  { ���O�ǰe����]�w }

  PCmdTransEnv = ^TCmdTransEnv;
  TCmdTransEnv = record
    UseSimulator: Boolean;             { �ϥ� NAGRA �������� }
    EnableControlSend: Boolean;        { �ҥζǰe���O }
    EnableFeedbackRecv: Boolean;       { �ҥα����^�Ǹ�� }
    SendCommandDelay: Double;          { �ǰe�@�����O��Ȱ��X�� }
  end;

  { �䥦�t�γ]�w }

  PCommEnv = ^TCommonEnv;
  TCommonEnv = record
    ProcessRecordCount: Integer;         { Ū����Ʈ�, �@��Ū���X�� }
    ProcessIPPV: String;                 { �B�z IPPV ���, A--> ����, N--> ���B�z, O--> �u�B�z }
    ProcessBatch: String;                { �B�z�妸���, A--> ����, N--> ���B�z, O--> �u�B�z }
    DbMultiThread: Boolean;              { �ϥοW�ߪ� Thread �h�s����Ʈw }
    DbRetryFrequence: Integer;           { ���Ʈw�_�u��, �C�h�֬��ճs�u }
    DbResendCount: Integer;              { ��ƭ��e���ƭ��� }
    DbWriteErrorWhenSocketFail: Boolean; { ��L�k�ǰe���O��, �ݳB�z���O�����Хܦ����~ }
    BusyTimeStart: String;               { ���L�ɬq�w�q(�_) }
    BusyTimeEnd: String;                 { ���L�ɬq�w�q(��) }
    BusyTimeDbReadFrequence: Double;     { ���L�ɬq�C�h�[Ū�@������� }
    NormalTimeDbReadFrequence: Double;   { �@��ɬq�C�h�[Ū�@������� }
    BatchOperator: String;               { �妸����ѧO�b�� }
    CASRetryFrequence: Integer;          { �� Nagra �s���W��, �C�h�֬��ճs�u }
    CASSendErrMax: Integer;              { �ǰe���O��, �o�ͦh�֦����~�۰��_�u }
    CASRecvErrMax: Integer;              { �������O��, �o�ͦh�֦����~�۰��_�u }
    CASCommCheck: Integer;               { �C�h�֬�e�@�� Communication Check }
    DbWaterMark: Integer;                { ��ݶǰe��C�W�L�h�ֵ���, �h��������O, �j�q�妸��Ʈɥ� }
  end;


  { Send_Nagar Table ����������� }
  { SendStatus --> 'W' �|���ǰe,
                   'P' �w�ǰe, �|������^��,
                   'E' �ǰe����, TCP/IP �άO����^�������~,
                   'C' �w�ǰe, �w����^��, �B�榡���T }

  PSendNagra = ^TSendNagra;
  TSendNagra = record
    HighCmdId: String;
    LowCmdId: String;
    IccNo: String;
    StbNo: String;
    SubBeginDate: String;
    SubEndDate: String;
    Notes: String;
    CmdStatus: String;
    Operator: String;
    UpdTime: TDateTime;
    ErrCode: String;
    ErrMsg: String;
    ImdProductId: String;
    ZipCode: String;
    CreditMode: String;
    Credit: String;
    CreditLimit: String;
    ThreholdCredit: String;
    EventName: String;
    Price: String;
    CCNumber1: String;
    IPAddr: String;
    CCPort: String;
    CallbackDate: String;
    CallbackTime: String;
    CallbackFrequency: String;
    FirstCallbackDate: String;
    PhoneNum1: String;
    PhoneNum2: String;
    PhoneNum3: String;
    MisIrdCmd: String;
    MisIrdData: String;
    SeqNo: String;
    CompCode: String;
    ResentTimes: Integer;
    CleanupDate: String;
    ConditionDate: String;
    CollectDate: String;
    STBFlag: String;
    NetWorkId: String;
    PinCode: String;
    DVRQuota: String;
    VodMoppId: String;    
    { �P�_�O�_����s�����O }
    IsGlobalCommand: Boolean;
    {}
    SmsSendStatus: String;
    SmsSendTime: TDateTime;
    SmsTransactionNumber: String;
    SmsCommandText: String;
    SmsFullCommandText: String;
    CaAckTransactionNumber: String; { NAGRA �ǹL�Ӫ� TransactionNumber }
    CaAckCommandId: String;
    CaAckCommandText: String;
    CaAckStatus: String;
    CaAckTime: TDateTime;
    SoCompListIndex: Integer;
    ReferenceLowCmdCount: Integer;
    HasbeenLog: Boolean;
    { ��ܫ��O�� }
    Data: Pointer;
    ParentData: Pointer;
    { ��ܩR�O�Ҧ��� }
    Index: Integer;
    ParentIndex: Integer;
  end;

  PRecvNagra = ^TRecvNagra;
  TRecvNagra = record
    CompCode: String;
    TransactionNumber: String;
    ResponseTransactionNumber: String; { NAGRA �ǹL�Ӫ� TransactionNumber }
    ResponseCommandText: String;
    HighLevelCmd: String;
    LowLevelCmd: String;
    IccNo: String;
    StbNo: String;
    CallbackDate: String;
    CallbackTime: String;
    NumberOfIppv: String;
    Credit: String;
    Debit: String;
    ImdProductId: String;
    PurchaseDate: String;
    WatchStatus: String;
    ProductSuspended: String;
    IccSuspended: String;
    PhoneNum1: String;
    PhoneNum2: String;
    PhoneNum3: String;
    AbnormalPhone: String;
    StbResponding: String;
    CmdStatus: String;
    CommandText: String;
    FullCommandText: String;
    SendStatus: String;   { W-->�ݳB�z, C--> Ack��Nagra, E-->Nack��Nagra}
    NackStatus: String;
    ErrCode: String;
    ErrMsg: String;
    UpdTime: TDateTime;
    { 2 �� }
    Data: Pointer;
    ParentData: Pointer;
    Index: Integer;
    ParentIndex: Integer;
  end;


  { For IRD Command --> Command 69 �|�Ψ� }

  PIrdInfo = ^TIrdInfo;
  TIrdInfo = record
    IrdCmdId: String;
    IrdOperation: String;
    IrdDataLength: String;
    IrdData: String;
  end;


  { TThreadQueue }

  TThreadQueue = class(TQueue)
  private
    FLock: TCriticalSection;
    FThreadSafe: Boolean;
    procedure LockQueue;
    procedure UnlockQueue;
  protected
    procedure PushItem(AItem: Pointer); override;
    function PopItem: Pointer; override;
    function PeekItem: Pointer; override;
  public
    constructor Create(const AThreadSafe: Boolean);
    destructor Destroy; override;
    property IsThreadSafe: Boolean read FThreadSafe;
    procedure Lock;
    procedure Unlock;
  end;

  { TThreadStringList }

  TThreadStringList = class(TStringList)
  private
    FLock: TMREWSync;
  public
    constructor Create;
    destructor Destroy; override;
    procedure BeginRead;
    procedure EndRead;
    procedure BeginWrite;
    procedure EndWrite;
  end;

  { TTransactionNumberGenerator }

  TTransactionNumberGenerator = class(TObject)
  private
    FName: String;
    FLock: TCriticalSection;
    FThreadSafe: Boolean;
    FInitValue: Integer;
    FMaxValue: Integer;
    FSequence: Integer;
    function GetNextValue: Integer;
    procedure LockGenerator;
    procedure UnlockGenerator;
    function GetCurrentValue: Integer;
  protected
    function Generate: Integer; virtual;
  public
    constructor Create(const aName: String; const aThreadSafe: Boolean; const aInitValue,
      aMaxValue: Integer);
    destructor Destroy; override;
    property InitValue: Integer read FInitValue;
    property MaxValue: Integer read FMaxValue;
    property CurrentValue: Integer read GetCurrentValue;
    property NextValue: Integer read GetNextValue;
    property ThreadSafe: Boolean read FThreadSafe;
    procedure Restore(const aFileName: String); overload;
    procedure Restore(const aIniFile: TIniFile); overload;
    procedure SaveTo(const aFileName: String); overload;
    procedure SaveTo(const aIniFile: TIniFile); overload;
  end;

  { TThreadManager } 

  TThreadManager = class(TObject)
  private
    FList: TList;
    FLock: TCriticalSection;
  protected
    function GetThread(Index: Integer): TThread;
    function GetCount: Integer;
    function CheckEmpty: Boolean;
    procedure Clear; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Add(aThread: TThread);
    function IndexOf(aThread: TThread): Integer;
    procedure StartThread(aIndex: Integer); virtual;
    procedure StopThread(aIndex: Integer); virtual;
    procedure Lock;
    procedure Unlock;
    property Count: Integer read GetCount;
    property Items[Index: Integer]: TThread read GetThread; default;
  end;


  { TCommandListMagager }

  TCommandListMagager = class(TObject)
  private
    FManagerList: TList;
    function GetItem(aIndex: Integer): TThreadStringList;
    function GetItemCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Add(aItem: TThreadStringList);
    property Items[Index: Integer]: TThreadStringList read GetItem; default;
    property ItemCount: Integer read GetItemCount;
    procedure Clear;
  end;



  TObServer = class;

  { TSubject }

  TSubject = class(TObject)
  private
    FList: TList;
    FAutoNotify: Boolean;
    FRunState: TRunState;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddObServer(aObServer: TObServer);
    procedure RemoveObServer(aObServer: TObServer);
    procedure NotifyObServer; virtual;
    property AutoNotify: Boolean read FAutoNotify write FAutoNotify;
    property RunState: TRunState read FRunState write FRunState;
  end;


  { TMessageSubject }

  TMessageSubject = class(TSubject)
  private
    FNotifyMessage: String;
    FNotifyType: Longint;
  protected
    function GetMsgType: Longint; virtual;
    procedure SetNotifyMessage(aMsg: String); virtual;
  public
    property NotifyType: Longint read FNotifyType;
    property NotifyMessage: String read FNotifyMessage write SetNotifyMessage;
  end;


  { TMessageExSubject }

  TMessageExSubject = class(TSubject)
  private
    FNotifyMessage: String;
    FNotifyType: Longint;
  protected
    procedure InternalNotify(const AMsg: String; const AType: Integer);
  public
    procedure Warning(const AMsg: String);
    procedure Error(const AMsg: String);
    procedure Info(const AMsg: String);
    procedure OK(const AMsg: String);
    property NotifyType: Longint read FNotifyType;
    property NotifyMessage: String read FNotifyMessage;
  end;


  { TCommandSubject }

  TCommandSubject = class(TSubject)
  private
    FThreadType: TThreadType;
    FUpdateType: TNotifyCommandType;
    FUpdateCommandObj: Pointer;
  protected
    procedure SetUpdateCommandObject(aObj: Pointer);
  public
    property ThreadType: TThreadType read FThreadType write FThreadType;
    property UpdateCommandType: TNotifyCommandType read FUpdateType;
    property UpdateCommandObject: Pointer read FUpdateCommandObj
      write SetUpdateCommandObject;
  end;


  TSubjectNotifyEvent = procedure (aSubject: TSubject) of object;


  { TObServer }

  TObServer = class(TObject)
  private
    FPause: Boolean;
    FSource: TSubject;
    FLock: TCriticalSection;
    FOnUpdate: TSubjectNotifyEvent;
    procedure CallOnUpdate;
  protected
    procedure Update(ASubject: TSubject); virtual;
  public
    constructor Create;
    destructor Destroy; override;
    property Pause: Boolean read FPause write FPause;
    property OnUpdate: TSubjectNotifyEvent read FOnUpdate write FOnUpdate;
  end;



  { TCommonThread }

  TCommonThread = class(TThread)
  private
    FWndHandle: THandle;
    FMainFormHandle: THandle;
    FCommEnv: TCommonEnv;
    FHighCmdEnv: TClientDataSet;
    FCASEnv: TCASEnv;
    FPlay: Boolean;
    FStop: Boolean;
    FMsgSubject: TMessageExSubject;
    FCmdTransEnv: TCmdTransEnv;
    FCmdError: TClientDataSet;
  protected
    procedure SetCommEnv(const aValue: TCommonEnv); virtual;
    procedure SetHighCmdEnv(aValue: TClientDataSet); virtual;
    procedure SetCmdTransEnv(const aValue: TCmdTransEnv); virtual;
    procedure SetCASEnv(const aValue: TCASEnv); virtual;
    procedure SetCmdError(aValue: TClientDataSet);
    procedure WndProc(var Msg: TMessage); virtual;
    procedure BeginActive; virtual; abstract;
    procedure EndActive; virtual; abstract;
    procedure Update; virtual; abstract;
    procedure WaitForPlaySignal; virtual;
    procedure WaitForTerminalSignal; virtual;
    function GetDelayFrquence: Cardinal; virtual; abstract;
    function GetWaitWhileFrquence: Cardinal; virtual;
    function GetBetweenFrquence(aValue: Cardinal): Cardinal;
    property Play: Boolean read FPlay write FPlay;
    property Stop: Boolean read FStop write FStop;
  public
    constructor Create;
    destructor Destroy; override;
    property WndHandle: THandle read FWndHandle;
    property MainFormHandle: THandle read FMainFormHandle write FMainFormHandle;
    property CommEnv: TCommonEnv read FCommEnv write SetCommEnv;
    property HighCmdEnv: TClientDataSet read FHighCmdEnv write SetHighCmdEnv;
    property CmdTransEnv: TCmdTransEnv read FCmdTransEnv write SetCmdTransEnv;
    property CASEnv: TCASEnv read FCASEnv write SetCASEnv;
    property CmdError: TClientDataSet read FCmdError write SetCmdError;
    property MessageSubject: TMessageExSubject read FMsgSubject;
  end;


  { TCommandThread }

  TSMSCommandThread = class(TCommonThread)
  private
    FUpdatedObject: Pointer;
    FCommandSubject: TCommandSubject;
  protected
    property UpdatedObject: Pointer read FUpdatedObject write FUpdatedObject;
  public
    constructor Create;
    destructor Destroy; override;
    property CommandSubject: TCommandSubject read FCommandSubject;
  end;


  { TGateWaySocketThread }

  TSMSSocketThread = class(TSMSCommandThread)
  private
    FSocket: TIdTCPClient;
  protected
    procedure WndProc(var Msg: TMessage); override;
    function BeginConnection: Boolean; virtual;
    function EndConnection: Boolean; virtual;
  public
    constructor Create(aSocket: TIdTCPClient);
    destructor Destroy; override;
    property GateWaySocket: TIdTCPClient read FSocket;
  end;


  { TConfigFileManager }

  TConfigFileManager = class(TObject)
  private
    FList: TStringList;
    FActiveFile: String;
    function GetFiles(Index: Integer): String;
    function GetFileCount: Integer;
  protected
    procedure SearchFiles;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Refresh;
    procedure Restore(const aFileName: String);
    procedure SaveTo(const aFileName: String);
    property Files[Index: Integer]: String read GetFiles;
    property FileCount: Integer read GetFileCount;
    property ActiveFile: String read FActiveFile write FActiveFile;
  end;


implementation

uses StrUtils, cbNagraDevice, cbUtilis, cbResStr;


{ TThreadQueue }

constructor TThreadQueue.Create(const AThreadSafe: Boolean);
begin
  inherited Create;
  if AThreadSafe then FLock := TCriticalSection.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TThreadQueue.Destroy;
begin
  LockQueue;
  try
    inherited Destroy;
  finally
    UnlockQueue;
    FLock.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadQueue.Lock;
begin
  LockQueue;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadQueue.LockQueue;
begin
  if FThreadSafe then FLock.Enter;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadQueue.Unlock;
begin
  UnlockQueue;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadQueue.UnlockQueue;
begin
  if FThreadSafe then FLock.Leave;
end;

{ ---------------------------------------------------------------------------- }

function TThreadQueue.PeekItem: Pointer;
begin
  LockQueue;
  try
    Result := inherited PeekItem;
  finally
    UnlockQueue;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TThreadQueue.PopItem: Pointer;
begin
  LockQueue;
  try
    Result := inherited PopItem;
  finally
    UnlockQueue;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadQueue.PushItem(AItem: Pointer);
begin
  LockQueue;
  try
    inherited PushItem( AItem );
  finally
    UnlockQueue;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TThreadStringList }

constructor TThreadStringList.Create;
begin
  inherited Create;
  FLock := TMREWSync.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TThreadStringList.Destroy;
begin
  FLock.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadStringList.BeginRead;
begin
  FLock.BeginRead;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadStringList.BeginWrite;
begin
  FLock.BeginWrite;
end;

{ ---------------------------------------------------------------------------- }


procedure TThreadStringList.EndRead;
begin
  FLock.EndRead;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadStringList.EndWrite;
begin
  FLock.EndWrite;
end;

{ ---------------------------------------------------------------------------- }

{ TTransactionNumberGenerator }

constructor TTransactionNumberGenerator.Create(const aName: String;
  const aThreadSafe: Boolean; const aInitValue, aMaxValue: Integer);
begin
  inherited Create;
  FName := aName;
  FThreadSafe := aThreadSafe;
  if FThreadSafe then FLock := TCriticalSection.Create;
  FInitValue := aInitValue;
  FMaxValue := aMaxValue;
  FSequence := FInitValue;
end;

{ ---------------------------------------------------------------------------- }

destructor TTransactionNumberGenerator.Destroy;
begin
  LockGenerator;
  try
    inherited Destroy;
  finally
    UnlockGenerator;
    if FThreadSafe then FLock.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.LockGenerator;
begin
  if FThreadSafe then FLock.Enter;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.UnlockGenerator;
begin
  if FThreadSafe then FLock.Leave;
end;

{ ---------------------------------------------------------------------------- }

function TTransactionNumberGenerator.Generate: Integer;
begin
  LockGenerator;
  try
    if FSequence >= FMaxValue then
      FSequence := FInitValue
    else
      Inc( FSequence );
    Result := FSequence;
  finally
    UnlockGenerator;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TTransactionNumberGenerator.GetCurrentValue: Integer;
begin
  Result := FSequence;
end;

{ ---------------------------------------------------------------------------- }

function TTransactionNumberGenerator.GetNextValue: Integer;
begin
  Result := Generate;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.Restore(const aFileName: String);
var
  aIniFile: TIniFile;
begin
  if Trim( aFileName ) = EmptyStr then
    raise Exception.Create( 'The FileName Is Empty!' );
  aIniFile := TIniFile.Create( aFileName );
  try
    Self.Restore( aIniFile );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.Restore(const aIniFile: TIniFile);
begin
  if Trim( FName ) = EmptyStr then
    raise Exception.Create( 'The Object Name Is Empty!' );
  FSequence := aIniFile.ReadInteger( 'TransactionNumber', FName, 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.SaveTo(const aFileName: String);
var
  aIniFile: TIniFile;
begin
  if Trim( aFileName ) = EmptyStr then
    raise Exception.Create( 'The FileName Is Empty!' );
  aIniFile := TIniFile.Create( aFileName );
  try
    Self.SaveTo( aIniFile );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.SaveTo(const aIniFile: TIniFile);
begin
  if not Assigned( aIniFile ) then
    raise Exception.Create( 'The IniFile object is nil!' );
  if Trim( FName ) = EmptyStr then
    raise Exception.Create( 'The Object Name Is Empty!' );
  aIniFile.WriteInteger( 'TransactionNumber', FName, FSequence );
end;

{ ---------------------------------------------------------------------------- }

{ TThreadManager }

constructor TThreadManager.Create;
begin
  inherited Create;
  FList := TList.Create;
  FLock := TCriticalSection.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TThreadManager.Destroy;
begin
  FLock.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadManager.Add(aThread: TThread);
begin
  FList.Add( aThread );
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadManager.Clear;
begin
  FList.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TThreadManager.GetCount: Integer;
begin
  Result := FList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TThreadManager.GetThread(Index: Integer): TThread;
begin
  Result := TThread( FList.Items[Index] );
end;

{ ---------------------------------------------------------------------------- }

function TThreadManager.IndexOf(aThread: TThread): Integer;
begin
  Result := FList.IndexOf( aThread );
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadManager.StartThread(aIndex: Integer);
var
  aThread: TThread;
begin
  aThread := TThread( FList.Items[aIndex] );
  if aThread.Suspended then
    aThread.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadManager.StopThread(aIndex: Integer);
var
  aThread: TThread;
begin
  try  //By Kin
    aThread := TThread( FList.Items[aIndex] );
    if aThread.Suspended then aThread.Resume;
    aThread.Terminate;
    aThread.WaitFor;
    aThread.Free;
    FList.Items[aIndex] := nil;
    if CheckEmpty then FList.Clear;
  except
  {}
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadManager.Lock;
begin
  FLock.Enter;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadManager.Unlock;
begin
  FLock.Leave;
end;

{ ---------------------------------------------------------------------------- }

function TThreadManager.CheckEmpty: Boolean;
var
  aIndex, aEmptyCount: Integer;
begin
  aEmptyCount := 0;
  for aIndex := 0 to FList.Count - 1 do
  begin
    if not Assigned( FList.Items[aIndex] ) then
      Inc( aEmptyCount );
  end;
  Result := ( aEmptyCount = FList.Count );
end;

{ ---------------------------------------------------------------------------- }

{ TCommandListMagager }

constructor TCommandListMagager.Create;
begin
  FManagerList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TCommandListMagager.Destroy;
begin
  FManagerList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandListMagager.Add(aItem: TThreadStringList);
begin
  FManagerList.Add( aItem );
end;

{ ---------------------------------------------------------------------------- }

function TCommandListMagager.GetItem(aIndex: Integer): TThreadStringList;
begin
  Result := TThreadStringList( FManagerList[aIndex] );
end;

{ ---------------------------------------------------------------------------- }

function TCommandListMagager.GetItemCount: Integer;
begin
  Result := FManagerList.Count;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommandListMagager.Clear;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FManagerList.Count - 1 do
    TThreadStringList( FManagerList[aIndex] ).Free;
  FManagerList.Clear;
end;

{ ---------------------------------------------------------------------------- }

{ TSubject }

constructor TSubject.Create;
begin
  inherited Create;
  FList := TList.Create;
  FAutoNotify := False;
  FRunState := rsStartup;
end;

{ ---------------------------------------------------------------------------- }

destructor TSubject.Destroy;
begin
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

procedure TSubject.NotifyObServer;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FList.Count - 1 do
    TObServer( FList[aIndex] ).Update( Self );
end;

{ ---------------------------------------------------------------------------- }

{ TMessageSubject }

function TMessageSubject.GetMsgType: Longint;
const
  aErrChar: array [0..2] of String = ( '����', '���~', '��]' );
  aSuccessChar: array [0..2] of String = ( '���\', '����', '�N��' );
var
  aIndex: Integer;
begin
  Result := MB_ICONINFORMATION;
  for aIndex := Low( aErrChar ) to High( aErrChar ) do
  begin
    if AnsiPos( aErrChar[aIndex], FNotifyMessage ) > 0 then
    begin
      Result := MB_ICONERROR;
      Break;
    end;
  end;
  if Result <> MB_ICONINFORMATION then Exit;
  for aIndex := Low( aSuccessChar ) to High( aSuccessChar ) do
  begin
    if AnsiPos( aSuccessChar[aIndex], FNotifyMessage ) > 0 then
    begin
      Result := MB_OK;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageSubject.SetNotifyMessage(aMsg: String);
begin
  FNotifyMessage := aMsg;
  FNotifyType := GetMsgType;
  if AutoNotify then NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TMessageExSubject }

procedure TMessageExSubject.InternalNotify(const AMsg: String; const AType: Integer);
begin
  FNotifyType := AType;
  FNotifyMessage := AMsg;
  if AutoNotify then NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageExSubject.Info(const AMsg: String);
begin
  InternalNotify( AMsg, MB_ICONINFORMATION );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageExSubject.Warning(const AMsg: String);
begin
  InternalNotify( AMsg, MB_ICONWARNING );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageExSubject.Error(const AMsg: String);
begin
  InternalNotify( AMsg, MB_ICONERROR );
end;

{ ---------------------------------------------------------------------------- }

procedure TMessageExSubject.OK(const AMsg: String);
begin
  InternalNotify( AMsg, MB_OK );
end;

{ ---------------------------------------------------------------------------- }

{ TCommandSubject }

procedure TCommandSubject.SetUpdateCommandObject(aObj: Pointer);
var
  aCmdId: Integer;
begin
  FUpdateCommandObj := aObj;
  if FThreadType in [tdControlSend, tdControlRecv] then
    aCmdId := StrToInt( Nvl( PSendNagra( aObj ).LowCmdId, '1' ) )
  else
    aCmdId := StrToInt( Nvl( PRecvNagra( aObj ).LowLevelCmd, '1' ) );
  case aCmdId of
    1000: FUpdateType := ncAck;
    1001: FUpdateType := ncNAck;
    1002: FUpdateType := ncCommunication;
  else
    FUpdateType := ncCommand;
  end;
  if AutoNotify then NotifyObServer;
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
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TObServer.CallOnUpdate;
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
    FOnUpdate( FSource );
end;

{ ---------------------------------------------------------------------------- }
procedure TObServer.Update(ASubject: TSubject);
begin
  if Assigned( FOnUpdate ) and ( not FPause ) then
  begin
    FLock.Enter;
    try
      if ( GetCurrentThreadId <> MainThreadID ) then
      begin
        //FLock.Enter;
        try
          if Assigned( ASubject ) then
          begin
            FSource := ASubject;
            TThread.Synchronize( nil, CallOnUpdate );
          end;
        finally
          //FLock.Leave;
        end;
      end else
      begin
        if Assigned( ASubject ) then
          FOnUpdate( ASubject );
      end;
    finally
      FLock.Leave;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TCommonThread }

constructor TCommonThread.Create;
begin
  inherited Create( True );
  FWndHandle := Classes.AllocateHWnd( WndProc );
  FHighCmdEnv := TClientDataSet.Create( nil );
  FCmdError := TClientDataSet.Create( nil );
  FMsgSubject := TMessageExSubject.Create;
  FMsgSubject.AutoNotify := True;
  FPlay := False;
  FStop := False;
  Priority := tpNormal;
  FreeOnTerminate := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TCommonThread.Destroy;
begin
  Classes.DeallocateHWnd( FWndHandle ) ;
  FCmdError.Free;
  FHighCmdEnv.Free;
  FMsgSubject.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_PLAY: FPlay := True;
    WM_STOP: FStop := True;
  else
    Msg.Result := DefWindowProc( FWndHandle, Msg.Msg, Msg.wParam, Msg.lParam );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommonThread.GetWaitWhileFrquence: Cardinal;
begin
  Result := 100;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.SetCommEnv(const aValue: TCommonEnv);
begin
  FCommEnv.ProcessRecordCount := aValue.ProcessRecordCount;
  FCommEnv.ProcessIPPV := aValue.ProcessIPPV;
  FCommEnv.ProcessBatch := aValue.ProcessBatch;
  FCommEnv.DbMultiThread := aValue.DbMultiThread;
  FCommEnv.DbRetryFrequence := aValue.DbRetryFrequence;
  FCommEnv.DbResendCount := aValue.DbResendCount;
  FCommEnv.DbWriteErrorWhenSocketFail := aValue.DbWriteErrorWhenSocketFail;
  FCommEnv.BusyTimeStart := aValue.BusyTimeStart;
  FCommEnv.BusyTimeEnd := aValue.BusyTimeEnd;
  FCommEnv.BusyTimeDbReadFrequence := aValue.BusyTimeDbReadFrequence;
  FCommEnv.NormalTimeDbReadFrequence := aValue.NormalTimeDbReadFrequence;
  FCommEnv.BatchOperator := aValue.BatchOperator;
  FCommEnv.CASRetryFrequence := aValue.CASRetryFrequence;
  FCommEnv.CASSendErrMax := aValue.CASSendErrMax;
  FCommEnv.CASRecvErrMax := aValue.CASRecvErrMax;
  FCommEnv.CASCommCheck := aValue.CASCommCheck;
  FCommEnv.DbWaterMark := aValue.DbWaterMark;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.SetHighCmdEnv(aValue: TClientDataSet);
begin
  FHighCmdEnv.Data := aValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.SetCmdTransEnv(const aValue: TCmdTransEnv);
begin
  FCmdTransEnv.UseSimulator := aValue.UseSimulator;
  FCmdTransEnv.EnableControlSend := aValue.EnableControlSend;
  FCmdTransEnv.EnableFeedbackRecv := aValue.EnableFeedbackRecv;
  FCmdTransEnv.SendCommandDelay := aValue.SendCommandDelay;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.SetCASEnv(const aValue: TCASEnv);
begin
  FCASEnv.IP := aValue.IP;
  FCASEnv.ControlPort := aValue.ControlPort;
  FCASEnv.FeedbackPort := aValue.FeedbackPort;
  FCASEnv.ControlSourceId := aValue.ControlSourceId;
  FCASEnv.ControlDestId := aValue.ControlDestId;
  FCASEnv.ControlMopPPId := aValue.ControlMopPPId;
  FCASEnv.FeedbackSourceId := aValue.FeedbackSourceId;
  FCASEnv.FeedbackDestId := aValue.FeedbackDestId;
  FCASEnv.FeedbackMopPPId := aValue.FeedbackMopPPId;
  FCASEnv.BoardCastMode := aValue.BoardCastMode;
  FCASEnv.AddressType := aValue.AddressType;
  FCASEnv.ControlTransId := aValue.ControlTransId;
  FCASEnv.FeedbackTransId := aValue.FeedbackTransId;
  FCASEnv.MailId := aValue.MailId;
  FCASEnv.CmdMaxCounter := aValue.CmdMaxCounter;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.SetCmdError(aValue: TClientDataSet);
begin
  FCmdError.Data := aValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.WaitForPlaySignal;
begin
  while not FPlay  do
  begin
    Sleep( 300 );
    if Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.WaitForTerminalSignal;
begin
  while FStop  do
  begin
    Sleep( 150 );
    if Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCommonThread.GetBetweenFrquence(aValue: Cardinal): Cardinal;
begin
  Result := 0;
  if ( GetTickCount - aValue ) > 0 then
    Result := ( GetTickCount - aValue );
end;

{ ---------------------------------------------------------------------------- }


{ TSMSCommandThread }

constructor TSMSCommandThread.Create;
begin
  inherited Create;
  FCommandSubject := TCommandSubject.Create;
  FCommandSubject.AutoNotify := True;
  FUpdatedObject := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TSMSCommandThread.Destroy;
begin
  FCommandSubject.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

{ TSMSSocketThread }

constructor TSMSSocketThread.Create(aSocket: TIdTCPClient);
begin
  inherited Create;
  FSocket := aSocket;
end;

{ ---------------------------------------------------------------------------- }

destructor TSMSSocketThread.Destroy;
begin
  FSocket := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

{ �аѦ� EMC SasGwyStdSpe020705.pdf �� 7 �� --> 2.3.7 }

function TSMSSocketThread.BeginConnection: Boolean;
const
  aName = 'SMSGW';
var
  aWriteBuffer: PDeviceCall;
  aReadBuffer: PDeviceAnswer;
  aIndex, aWriteSize: Integer;
  aThreadName: String;
  aSaveTimeout: Integer;
begin
  aThreadName := EmptyStr;
  if CommandSubject.FThreadType = tdControlSend then
    aThreadName := SControlSend
  else if CommandSubject.FThreadType = tdFeedbackSend then
    aThreadName := SFeedbackSend;
  aSaveTimeout := FSocket.ReadTimeout;
  FSocket.ReadTimeout := 10000;
  if FSocket.Connected then FSocket.Disconnect;
  FSocket.Connect;
  New( aWriteBuffer );
  try
    ZeroMemory( aWriteBuffer, SizeOf( aWriteBuffer^ ) );
    aWriteBuffer.ObNameLen := Length( aName );
    for aIndex := 1 to Length( aName ) do
      aWriteBuffer.ObName[aIndex] := aName[aIndex];
    aWriteBuffer.OpMode := 0;
    aWriteBuffer.Len[1] := 0;
    aWriteBuffer.Len[2] := SizeOf( aWriteBuffer.OpMode ) +
      SizeOf( aWriteBuffer.ObNameLen ) + aWriteBuffer.ObNameLen;
    aWriteSize := ( Length( aName ) + SizeOf( aWriteBuffer.OpMode ) +
      SizeOf( aWriteBuffer.ObNameLen ) + SizeOf( aWriteBuffer.Len ) );
    FSocket.WriteBuffer( aWriteBuffer^, aWriteSize, True );
  finally
    Dispose( aWriteBuffer );
  end;
  Sleep( 3000 );
  New( aReadBuffer );
  try
    ZeroMemory( aReadBuffer, SizeOf( aReadBuffer^ ) );
    FSocket.ReadBuffer( aReadBuffer^, SizeOf( aReadBuffer^ ) );
    Result := ( aReadBuffer.Status = 6 ) and ( aReadBuffer.AnswerCode = 0 );
    if ( Result ) then
      MessageSubject.OK( Format( SSocketConnectionSuccess, [aThreadName] ) )
    else begin
      if ( aReadBuffer.Status <> 6 ) then
        MessageSubject.Error( Format( SSocketConnectionFailure, [aThreadName] ) )
      else if ( aReadBuffer.AnswerCode <> 0  ) then
        MessageSubject.Error( Format( SSocketConnectionReject, [aThreadName] ) );
    end;
  finally
    FSocket.ReadTimeout := aSaveTimeout;
    Dispose( aReadBuffer );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSMSSocketThread.EndConnection: Boolean;
begin
  if FSocket.Connected then
    FSocket.Disconnect;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TSMSSocketThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_UPDATEGUI:
      begin
        if Msg.LParam = Ord( False ) then
          CommandSubject.RemoveObServer( TObServer( Msg.WParam ) )
        else
          CommandSubject.AddObServer( TObServer( Msg.WParam ) );
      end;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

{ TConfigFileManager }

constructor TConfigFileManager.Create;
begin
  FList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TConfigFileManager.Destroy;
begin
  FList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TConfigFileManager.GetFileCount: Integer;
begin
  Result := FList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TConfigFileManager.GetFiles(Index: Integer): String;
begin
  Result := FList.Strings[Index];
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigFileManager.Refresh;
begin
  SearchFiles;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigFileManager.Restore(const aFileName: String);
var
  aIniFile: TIniFile;
begin
  if Trim( aFileName ) = EmptyStr then
    raise Exception.Create( 'The FileName Is Empty!' );
  aIniFile := TIniFile.Create( aFileName );
  try
    FActiveFile := aIniFile.ReadString( 'ConfigFile', 'ActiveFile', EmptyStr );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigFileManager.SaveTo(const aFileName: String);
var
  aIniFile: TIniFile;
begin
  if Trim( aFileName ) = EmptyStr then
    raise Exception.Create( 'The FileName Is Empty!' );
  aIniFile := TIniFile.Create( aFileName );
  try
    aIniFile.WriteString( 'ConfigFile', 'ActiveFile', FActiveFile );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigFileManager.SearchFiles;
var
  aFound: Integer;
  aSearchPath: String;
  aSearchRect: TSearchRec;
begin
  FList.Clear;
  aSearchPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + Format( '*%s', [DEFAULT_CONFIGFILE_EXT] );
  aFound := FindFirst( aSearchPath, faAnyFile, aSearchRect );
  while ( aFound = 0 ) do
  begin
    FList.Add( aSearchRect.Name );
    aFound := FindNext( aSearchRect );
  end;
end;
{ ---------------------------------------------------------------------------- }

end.
