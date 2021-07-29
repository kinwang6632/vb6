unit cbClass;

interface

uses Windows, Messages, SysUtils, Variants, Classes, Contnrs, SyncObjs, IniFiles,
     DB, ADODB, DBClient,
     IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;


const

  WM_PLAY = WM_USER + 10;
  WM_STOP = WM_USER + 20;

  CONFIG_FILE_NAME = 'Config.ini';
  COMMON_DELAY = 100;

type


  TDbConnectStatus = ( dbNoSelect, dbNone, dbOK, dbError, dbActive);

  TReadyType = ( tdSo, tdContrlSend, tdControlRecv, tdFeedbackSend, tdFeedbackRecv );

  TRunState = ( rsStartup, rsRunning, rsStop );

  { 程式啟動執行資訊 }

  PAppStartupInfo = ^TAppStartupInfo;
  TAppStartupInfo = record
    AutoExecute: Boolean;              { 自動執行 }
    ConfigFileName: String;            { 指定的環境設定檔 }
  end;

  { 系統台資訊 }

  PSoInfo = ^TSoInfo;
  TSoInfo = record
    Selected: Boolean;                 { 是否選擇  }
    Pos: Integer;                      { 平台區分 }
    PosName: String;                   { 平台名稱 }
    CompCode: String;                  { 系統台代碼 }
    CompName: String;                  { 系統台名稱 }
    NetworkId: String;                 { Netword Id, 通常等於系統台代碼 }
    LoginUser: String;                 { 登入資料庫帳號 }
    LoginPass: String;                 { 登入資料庫密碼 }
    DbAliase: String;                  { 登入資料庫 }
    DbConnectStatus: TDbConnectStatus; { 資料庫連線狀態 }
    RecordCount: Integer;              { 從系統台讀取的資料筆數 }
    ItemIndex: Integer;                { 該系統台的節點 }
    NotifyMsg: String;                 { 通知的訊息 }
    CriticalError: Integer;            { 資料庫錯誤次數 }
    LastCriticalErrorTickCount: Cardinal; { 最後一次資料庫錯誤時間 --> TickCount }
  end;


  { 各系統台的資料庫連接須要使用到的 ADO 元件 }
  PSoADOControl = ^TSoADOControl;
  TSoADOControl = record
    CompCode: String;
    Connection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    DataSet: TClientDataSet;
    {
    StoredProc1: TADOStoredProc;
    StoredProc2: TADOStoredProc;
    StoredProc3: TADOStoredProc;
    StoredProc4: TADOStoredProc;
    StoredProc5: TADOStoredProc;
    StoredProc6: TADOStoredProc;
    }
  end;

  { Nagra CAS 設定 }

  PCASEnv = ^TCASEnv;
  TCASEnv = record
    IP: String;                        { CAS IP }
    ControlPort: Integer;              { 傳送指令的 Port }
    FeedbackPort: Integer;             { ReturnPath 的 Port }
    SourceId: String;                  { CAS 上的 Source Id }
    DestId: String;                    { CAS 上的 Dest Id }
    MopPPId: String;                   { CAS 上的 Mop PPID }
  end;

  { 指令傳送控制設定 }
  
  PCmdTransEnv = ^TCmdTransEnv;
  TCmdTransEnv = record
    BusyTimeStart: String;             { 忙碌時段定義(起) }
    BusyTimeEnd: String;               { 忙碌時段定義(迄) }
    BusyTimeReadFrequence: Integer;    { 忙碌時段每多久讀一取次資料 }
    NormalTimeReadFrequence: Integer;  { 一般時段每多久讀一取次資料 }
  end;

  { 其它系統設定 }

  PCommEnv = ^TCommonEnv;
  TCommonEnv = record
    ProcessRecordCount: Integer;         { 讀取資料時, 一次讀取幾筆 }
    ProcessIPPV: String;                 { 處理 IPPV 資料, A--> 全部, N--> 不處理, O--> 只處理 }
    ProcessBatch: String;                { 處理批次資料, A--> 全部, N--> 不處理, O--> 只處理 }
    DbMultiThread: Boolean;              { 使用獨立的 Thread 去連接資料庫 }
    DbRetryFrequence: Integer;           { 當資料庫斷線時, 每多少秒重試連線 }
    DbResendCount: Integer;              { 資料重送次數限制 }
    BusyTimeStart: String;               { 忙碌時段定義(起) }
    BusyTimeEnd: String;                 { 忙碌時段定義(迄) }
    BusyTimeDbReadFrequence: Double;     { 忙碌時段每多久讀一取次資料 }
    NormalTimeDbReadFrequence: Double;   { 一般時段每多久讀一取次資料 }
    BatchOperator: String;               {  批次資料識別帳號 }
  end;


  { Send_Nagar Table 欄位對應的資料 }

  PSendNagra = ^TSendNagra;
  TSendNagra = record
    SendStatus: String;
    TransactionNumber: String;
    HighLevelCmd: String;
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
    FullCommandText: String;
    LowLevelCmd: String;
    ResponseTransactionNumber: String;
    ResponseCommandId: String;
    FullReponseCommandText: String;
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
    FLock: TMultiReadExclusiveWriteSynchronizer;
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
    procedure Restore(const aFileName: String);
    procedure SaveTo(const aFileName: String);
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
    procedure SetNotifyMessage(aMsg: String);
  protected
    function GetMsgType: Longint; virtual;
  public
    property NotifyType: Longint read FNotifyType;
    property NotifyMessage: String read FNotifyMessage write SetNotifyMessage;
  end;


  { TObServer }

  TSubjectNotifyEvent = procedure (aSubject: TSubject) of object;

  TObServer = class(TObject)
  private
    FOnUpdate: TSubjectNotifyEvent;
  protected
    procedure Update(aSubject: TSubject); virtual;
  public
    property OnUpdate: TSubjectNotifyEvent read FOnUpdate write FOnUpdate;
  end;



  { TCommonThread }

  TCommonThread = class(TThread)
  private
    FWndHandle: THandle;
    FMainFormHandle: THandle;
    FCommEnv: TCommonEnv;
    FCmdEnv: TClientDataSet;
    FPlay: Boolean;
    FStop: Boolean;
    FMsgSubject: TMessageSubject;
    procedure SetCommEnv(const aValue: TCommonEnv);
    procedure SetCmdEnv(aValue: TClientDataSet);
  protected
    property Play: Boolean read FPlay write FPlay;
    property Stop: Boolean read FStop write FStop;
    procedure WndProc(var Msg: TMessage); virtual;
    procedure BeginActive; virtual; abstract;
    procedure EndActive; virtual; abstract;
    procedure Notify; virtual; abstract;
    procedure Update; virtual; abstract;
    procedure WaitForPlaySignal; virtual;
    procedure WaitForTerminalSignal; virtual;
    function GetDelayFrquence: Cardinal; virtual; abstract;
    function GetWaitWhileFrquence: Cardinal; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    property WndHandle: THandle read FWndHandle;
    property MainFormHandle: THandle read FMainFormHandle write FMainFormHandle;
    property CommEnv: TCommonEnv read FCommEnv write SetCommEnv;
    property CmdEnv: TClientDataSet read FCmdEnv write SetCmdEnv;
    property MessageSubject: TMessageSubject read FMsgSubject;
  end;


  { TGateWaySocketThread }

  TGateWaySocketThread = class(TCommonThread)
  private
    FSocket: TIdTCPClient;
    FCASEnv: TCASEnv;
    FSocketMessage: String;
  protected
    procedure SetCASEnv(const aValue: TCASEnv); virtual;
    procedure Notify; override;
    function BeginConnection: Boolean; virtual;
    function EndConnection: Boolean; virtual;
    property GateWaySocket: TIdTCPClient read FSocket;
    property SocketMessage: String read FSocketMessage write FSocketMessage;
  public
    constructor Create(aSocket: TIdTCPClient);
    destructor Destroy; override;
    property CASEnv: TCASEnv read FCASEnv write SetCASEnv;
  end;


implementation

uses StrUtils, cbNagraDevice, cbResStr;


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
  FLock := TMultiReadExclusiveWriteSynchronizer.Create;
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
  Generate;
  Result := GetCurrentValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.Restore(const aFileName: String);
var
  aIniFile: TIniFile;
begin
  if Trim( aFileName ) = EmptyStr then
    raise Exception.Create( 'The FileName Is Empty!' );
  if Trim( FName ) = EmptyStr then
    raise Exception.Create( 'The Object Name Is Empty!' );
  aIniFile := TIniFile.Create( aFileName );
  try
    FSequence := aIniFile.ReadInteger( 'TransactionNumber', FName, 0 );
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransactionNumberGenerator.SaveTo(const aFileName: String);
var
  aIniFile: TIniFile;
begin
  if Trim( aFileName ) = EmptyStr then
    raise Exception.Create( 'The FileName Is Empty!' );
  if Trim( FName ) = EmptyStr then
    raise Exception.Create( 'The Object Name Is Empty!' );
  aIniFile := TIniFile.Create( aFileName );
  try
    aIniFile.ReadInteger( 'TransactionNumber', FName, FSequence );
  finally
    aIniFile.Free;
  end;
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
  aThread := TThread( FList.Items[aIndex] );
  aThread.Terminate;
  if aThread.Suspended then aThread.Resume;
  aThread.WaitFor;
  aThread.Free;
  FList.Items[aIndex] := nil;
  if CheckEmpty then FList.Clear;
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
  FList.Add( aObServer );
end;

{ ---------------------------------------------------------------------------- }

procedure TSubject.RemoveObServer(aObServer: TObServer);
begin
  FList.Delete( FList.IndexOf( aObServer ) );
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
  aErrChar: array [0..2] of String = ( '失敗', '錯誤', '原因' );
  aSuccessChar: array [0..2] of String = ( '成功', '完成', '就序' );
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
  if FAutoNotify then NotifyObServer;
end;

{ ---------------------------------------------------------------------------- }

{ TObServer }

procedure TObServer.Update(aSubject: TSubject);
begin
  if Assigned( FOnUpdate ) then FOnUpdate( aSubject );
end;

{ ---------------------------------------------------------------------------- }

{ TCommonThread }

constructor TCommonThread.Create;
begin
  inherited Create( True );
  FWndHandle := Classes.AllocateHWnd( WndProc );
  FCmdEnv := TClientDataSet.Create( nil );
  FMsgSubject := TMessageSubject.Create;
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
  FCmdEnv.Free;
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
  FCommEnv.BusyTimeStart := aValue.BusyTimeStart;
  FCommEnv.BusyTimeEnd := aValue.BusyTimeEnd;
  FCommEnv.BusyTimeDbReadFrequence := aValue.BusyTimeDbReadFrequence;
  FCommEnv.NormalTimeDbReadFrequence := aValue.NormalTimeDbReadFrequence;
  FCommEnv.BatchOperator := aValue.BatchOperator;
end;

{ ---------------------------------------------------------------------------- }

procedure TCommonThread.SetCmdEnv(aValue: TClientDataSet);
begin
  FCmdEnv.Data := aValue.Data;
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
    Sleep( 300 );
    if Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TGateWaySocketThread }

constructor TGateWaySocketThread.Create(aSocket: TIdTCPClient);
begin
  inherited Create;
  FSocket := aSocket;
  FSocketMessage := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

destructor TGateWaySocketThread.Destroy;
begin
  FSocket := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TGateWaySocketThread.SetCASEnv(const aValue: TCASEnv);
begin
  FCASEnv.IP := aValue.IP;
  FCASEnv.ControlPort := aValue.ControlPort;
  FCASEnv.FeedbackPort := aValue.FeedbackPort;
  FCASEnv.SourceId := aValue.SourceId;
  FCASEnv.DestId := aValue.DestId;
  FCASEnv.MopPPId := aValue.MopPPId;
end;

{ ---------------------------------------------------------------------------- }

function TGateWaySocketThread.BeginConnection: Boolean;
const
  aName = 'SMWGW';
var
  aWriteBuffer: PDeviceCall;
  aReadBuffer: PDeviceAnswer;
  aIndex, aWriteSize: Integer;
begin
  FSocket.Connect( 10 );
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
    Sleep( 1000 );
  finally
    Dispose( aWriteBuffer );
  end;
  New( aReadBuffer );
  try
    ZeroMemory( aReadBuffer, SizeOf( aReadBuffer^ ) );
    FSocket.ReadBuffer( aReadBuffer^, SizeOf( aReadBuffer^ ) );
    Result := ( aReadBuffer.Status = 6 ) and ( aReadBuffer.AnswerCode = 0 );
  finally
    Dispose( aReadBuffer );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TGateWaySocketThread.EndConnection: Boolean;
begin
  if FSocket.Connected then
    FSocket.Disconnect;
  Result := True;  
end;

{ ---------------------------------------------------------------------------- }

procedure TGateWaySocketThread.Notify;
begin
  MessageSubject.NotifyMessage := FSocketMessage;
end;

{ ---------------------------------------------------------------------------- }

end.
