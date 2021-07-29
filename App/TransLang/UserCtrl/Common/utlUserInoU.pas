unit utlUserInoU;

{ Inter-Process Communication Thread Classes }

{$DEFINE DEBUG}

interface

uses
  SysUtils, Classes, Windows;

type

{ This is a generic class for all encapsulated WinAPI's which need to call
  CloseHandle when no longer needed.  This code eliminates the need for
  3 identical destructors in the TEvent, TMutex, and TSharedMem classes
  which are descended from this class. }

  THandledObject = class(TObject)
  protected
    FHandle: THandle;
  public
    destructor Destroy; override;
    property Handle: THandle read FHandle;
  end;

{ TSharedMem }

{ This class simplifies the process of creating a region of shared memory.
  In Win32, this is accomplished by using the CreateFileMapping and
  MapViewOfFile functions. }

  TSharedMem = class(THandledObject)
  private
    FName: string;
    FSize: Integer;
    FCreated: Boolean;
    FFileView: Pointer;
  public
    constructor Create(const Name: string; Size: Integer);
    destructor Destroy; override;
    property Name: string read FName;
    property Size: Integer read FSize;
    property Buffer: Pointer read FFileView;
    property Created: Boolean read FCreated;
  end;

{TUserInfo}

  UserInfo = record
    UserID: array[0..20] of Char;
    UserName: array[0..20] of Char;
    UserGrp: array[0..20] of Char;
    Password: array[0..20] of Char;
  end;

  PUserInfo = ^UserInfo;

procedure GetUserInfo(var UserId, UserName, Password, UserGrp: string);
procedure SaveUserInfo(UserId, UserName, Password, UserGrp: string);

implementation

procedure Error(const Msg: string);
begin
  raise Exception.Create(Msg);
end;

{ THandledObject }

destructor THandledObject.Destroy;
begin
  if FHandle <> 0 then
    CloseHandle(FHandle);
end;

{ TSharedMem }

constructor TSharedMem.Create(const Name: string; Size: Integer);
begin
  try
    FName := Name;
    FSize := Size;
    { CreateFileMapping, when called with $FFFFFFFF for the hanlde value,
      creates a region of shared memory }
    FHandle := CreateFileMapping($FFFFFFFF, nil, PAGE_READWRITE, 0,
        Size, PChar(Name));
    if FHandle = 0 then abort;
    FCreated := GetLastError = 0;
    { We still need to map a pointer to the handle of the shared memory region }
    FFileView := MapViewOfFile(FHandle, FILE_MAP_WRITE, 0, 0, Size);
    if FFileView = nil then abort;
  except
    Error(Format('Error creating shared memory %s (%d)', [Name, GetLastError]));
  end;
end;

destructor TSharedMem.Destroy;
begin
  if FFileView <> nil then
    UnmapViewOfFile(FFileView);
  inherited Destroy;
end;

procedure GetUserInfo(var UserId, UserName, Password, UserGrp: string);
var
  UserInfoBuf: PUserInfo;
  Shm: TSharedMem;
begin
  Shm := TSharedMem.Create('UCINFO', SizeOf(UserInfo));
  UserInfoBuf := Shm.Buffer;
  UserId := UserInfoBuf.UserId;
  UserName := UserInfoBuf.UserName;
  Password := UserInfoBuf.Password;
  UserGrp := UserInfoBuf.UserGrp;
  Shm.Free;
end;

procedure SaveUserInfo(UserId, UserName, Password, UserGrp: string);
var
  UserInfoBuf: PUserInfo;
  Shm: TSharedMem;
begin
  Shm := TSharedMem.Create('UCINFO', SizeOf(UserInfo));
  UserInfoBuf := Shm.Buffer;
  StrPCopy(UserInfoBuf.UserId, UserId);
  StrPCopy(UserInfoBuf.UserName, UserName);
  StrPCopy(UserInfoBuf.Password, Password);
  StrPCopy(UserInfoBuf.UserGrp, UserGrp);
  Shm.Free;
end;

end.
