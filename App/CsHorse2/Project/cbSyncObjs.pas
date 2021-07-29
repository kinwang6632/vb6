unit cbSyncObjs;

{$H+,X+}

interface

uses
  Windows,  Messages, SysUtils, Classes;

type
  PSecurityAttributes = pointer;

  TSynchroObject = class(TObject)
  public
    procedure Acquire; virtual;
    procedure Release; virtual;
  end;

  TWaitResult = (wrSignaled, wrTimeout, wrAbandoned, wrError);

  THandleObject = class(TSynchroObject)
  protected
    FHandle: THandle;
    FLastError: Integer;
    FUseCOMWait: Boolean;
  public
    { Specify UseCOMWait to ensure that when blocked waiting for the object
      any STA COM calls back into this thread can be made. }
    constructor Create(UseCOMWait: Boolean = False);
    destructor Destroy; override;
  public
    function WaitFor(Timeout: LongWord): TWaitResult; virtual;
    property LastError: Integer read FLastError;
    property Handle: THandle read FHandle;
  end;

  TEvent = class(THandleObject)
  public
    constructor Create(EventAttributes: PSecurityAttributes; ManualReset,
      InitialState: Boolean; const Name: string; UseCOMWait: Boolean = False); overload;
    constructor Create(UseCOMWait: Boolean = False); overload;
    procedure SetEvent;
    procedure ResetEvent;
  end;

  TSimpleEvent = class(TEvent);

  TMutex = class(THandleObject)
  public
    constructor Create(UseCOMWait: Boolean = False); overload;
    constructor Create(MutexAttributes: PSecurityAttributes; InitialOwner: Boolean; const Name: string; UseCOMWait: Boolean = False); overload;
    constructor Create(DesiredAccess: LongWord; InheritHandle: Boolean; const Name: string; UseCOMWait: Boolean = False); overload;
    procedure Acquire; override;
    procedure Release; override;
  end;

  TCriticalSection = class(TSynchroObject)
  protected
    FSection: TRTLCriticalSection;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Acquire; override;
    procedure Release; override;
    function TryEnter: Boolean;
    procedure Enter;
    procedure Leave;
  end;

implementation

type
  TCoWaitForMultipleHandlesProc = function (dwFlags: DWORD; dwTimeOut: DWORD;
    cHandles: LongWord; var Handles; var lpdwIndex: DWORD): HRESULT; stdcall;

var
  CoWaitForMultipleHandlesProc: TCoWaitFormultipleHandlesProc;

threadvar
  OleThreadWnd: HWND;

const
  OleThreadWndClassName = 'OleMainThreadWndClass'; //do not localize
  COWAIT_WAITALL = $00000001;
  COWAIT_ALERTABLE = $00000002;

  RPC_E_TIMEOUT = HRESULT($8001011F);
  {$EXTERNALSYM RPC_E_TIMEOUT}


function GetOleThreadWindow: HWND;
var
  ChildWnd: HWND;
  ParentWnd: HWND;
begin
  if (OleThreadWnd = 0) or not IsWindow(OleThreadWnd) then
  begin
    if (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 5) then
      ParentWnd := HWND(HWND_MESSAGE)
    else
      ParentWnd := 0;
    ChildWnd := 0;
    repeat
      OleThreadWnd := FindWindowEx(ParentWnd, ChildWnd, OleThreadWndClassName, nil);
      ChildWnd := OleThreadWnd;
    until (OleThreadWnd = 0) or (GetWindowThreadProcessId(OleThreadWnd, nil) = GetCurrentThreadId);
  end;
  Result := OleThreadWnd;
end;

function InternalCoWaitForMultipleHandles(dwFlags: DWORD; dwTimeOut: DWORD;
  cHandles: LongWord; var Handles; var lpdwIndex: DWORD): HRESULT; stdcall;
var
  WaitResult: DWORD;
  OleThreadWnd: HWnd;
  Msg: TMsg;
begin
  WaitResult := 0; // supress warning
  OleThreadWnd := GetOleThreadWindow;
  if OleThreadWnd <> 0 then
    while True do
    begin
      WaitResult := MsgWaitForMultipleObjectsEx(cHandles, Handles, dwTimeOut, QS_ALLEVENTS, dwFlags);
      if WaitResult = WAIT_OBJECT_0 + cHandles then
      begin
        if PeekMessage(Msg, OleThreadWnd, 0, 0, PM_REMOVE) then
        begin
          TranslateMessage(Msg);
          DispatchMessage(Msg);
        end;
      end else
        Break;
    end
  else
    WaitResult := WaitForMultipleObjectsEx(cHandles, @Handles,
      dwFlags and COWAIT_WAITALL <> 0, dwTimeOut, dwFlags and COWAIT_ALERTABLE <> 0);
  if WaitResult = WAIT_TIMEOUT then
    Result := RPC_E_TIMEOUT
  else if WaitResult = WAIT_IO_COMPLETION then
    Result := RPC_S_CALLPENDING
  else
  begin
    Result := S_OK;
    if (WaitResult >= WAIT_ABANDONED_0) and (WaitResult < WAIT_ABANDONED_0 + cHandles) then
      lpdwIndex := WaitResult - WAIT_ABANDONED_0
    else
      lpdwIndex := WaitResult - WAIT_OBJECT_0;
  end;
end;

function CoWaitForMultipleHandles(dwFlags: DWORD; dwTimeOut: DWORD;
  cHandles: LongWord; var Handles; var lpdwIndex: DWORD): HRESULT;

  procedure LookupProc;
  var
    Ole32Handle: HMODULE;
  begin
    Ole32Handle := GetModuleHandle('ole32.dll'); //do not localize
    if Ole32Handle <> 0 then
      CoWaitForMultipleHandlesProc := GetProcAddress(Ole32Handle, 'CoWaitForMultipleHandles'); //do not localize
    if not Assigned(CoWaitForMultipleHandlesProc) then
      CoWaitForMultipleHandlesProc := InternalCoWaitForMultipleHandles;
  end;

begin
  if not Assigned(CoWaitForMultipleHandlesProc) then
    LookupProc;
  Result := CoWaitForMultipleHandlesProc(dwFlags, dwTimeOut, cHandles, Handles, lpdwIndex)
end;

{ TSynchroObject }

procedure TSynchroObject.Acquire;
begin
end;

procedure TSynchroObject.Release;
begin
end;

{ THandleObject }

constructor THandleObject.Create(UseComWait: Boolean);
begin
  inherited Create;
  FUseCOMWait := UseCOMWait;
end;

destructor THandleObject.Destroy;
begin
  CloseHandle(FHandle);
  inherited Destroy;
end;

function THandleObject.WaitFor(Timeout: LongWord): TWaitResult;
var
  Index: DWORD;
begin
  if FUseCOMWait then
  begin
    case CoWaitForMultipleHandles(0, TimeOut, 1, FHandle, Index) of
      S_OK: Result := wrSignaled;
      RPC_S_CALLPENDING,
      RPC_E_TIMEOUT: Result := wrTimeout;
    else
      Result := wrError;
      FLastError := GetLastError;
    end;
  end else
  begin
    case WaitForSingleObject(Handle, Timeout) of
      WAIT_ABANDONED: Result := wrAbandoned;
      WAIT_OBJECT_0: Result := wrSignaled;
      WAIT_TIMEOUT: Result := wrTimeout;
      WAIT_FAILED:
        begin
          Result := wrError;
          FLastError := GetLastError;
        end;
    else
      Result := wrError;
    end;
  end;
end;

{ TEvent }

constructor TEvent.Create(EventAttributes: PSecurityAttributes; ManualReset,
  InitialState: Boolean; const Name: string; UseCOMWait: Boolean);
begin
  inherited Create(UseCOMWait);
  FHandle := CreateEvent(EventAttributes, ManualReset, InitialState, PChar(Name));
end;

constructor TEvent.Create(UseCOMWait: Boolean);
begin
  Create(nil, True, False, '', UseCOMWait);
end;

procedure TEvent.SetEvent;
begin
  Windows.SetEvent(Handle);
end;

procedure TEvent.ResetEvent;
begin
  Windows.ResetEvent(Handle);
end;

{ TCriticalSection }

constructor TCriticalSection.Create;
begin
  inherited Create;
  InitializeCriticalSection(FSection);
end;

destructor TCriticalSection.Destroy;
begin
  DeleteCriticalSection(FSection);
  inherited Destroy;
end;

procedure TCriticalSection.Acquire;
begin
  EnterCriticalSection(FSection);
end;

procedure TCriticalSection.Release;
begin
  LeaveCriticalSection(FSection);
end;

function TCriticalSection.TryEnter: Boolean;
begin
  Result := TryEnterCriticalSection(FSection);
end;

procedure TCriticalSection.Enter;
begin
  Acquire;
end;

procedure TCriticalSection.Leave;
begin
  Release;
end;

{ TMutex }

procedure TMutex.Acquire;
begin
  if WaitFor(INFINITE) = wrError then
    RaiseLastOSError;
end;

constructor TMutex.Create(UseCOMWait: Boolean);
begin
  Create(nil, False, '', UseCOMWait);
end;

constructor TMutex.Create(MutexAttributes: PSecurityAttributes;
  InitialOwner: Boolean; const Name: string; UseCOMWait: Boolean);
var
  lpName: PChar;
begin
  inherited Create(UseCOMWait);
  if Name <> '' then
    lpName := PChar(Name)
  else
    lpName := nil;
  FHandle := CreateMutex(MutexAttributes, InitialOwner, lpName);
  if FHandle = 0 then
    RaiseLastOSError;
end;

constructor TMutex.Create(DesiredAccess: LongWord; InheritHandle: Boolean;
  const Name: string; UseCOMWait: Boolean);
var
  lpName: PChar;
begin
  inherited Create(UseCOMWait);
  if Name <> '' then
    lpName := PChar(Name)
  else
    lpName := nil;
  FHandle := OpenMutex(DesiredAccess, InheritHandle, lpName);
  if FHandle = 0 then
    RaiseLastOSError;
end;

procedure TMutex.Release;
begin
  if not ReleaseMutex(FHandle) then
    RaiseLastOSError;
end;

end.
