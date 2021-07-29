unit TrayIcon;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, ExtCtrls, Forms,
  Menus, ShellAPI;

{ TTrayIcon }

const

  STrayIconRemoveError = 'Cannot remove shell notification icon';
  STrayIconCreateError = 'Cannot create shell notification icon';

  WM_SYSTEM_TRAY_MESSAGE = WM_USER + 1;

  {$EXTERNALSYM NIF_STATE}
  NIF_STATE       = $00000008;
  {$EXTERNALSYM NIF_INFO}
  NIF_INFO        = $00000010;

  {$EXTERNALSYM NIIF_NONE}
  NIIF_NONE       = $00000000;
  {$EXTERNALSYM NIIF_INFO}
  NIIF_INFO       = $00000001;
  {$EXTERNALSYM NIIF_WARNING}
  NIIF_WARNING    = $00000002;
  {$EXTERNALSYM NIIF_ERROR}
  NIIF_ERROR      = $00000003;
  {$EXTERNALSYM NIIF_ICON_MASK}
  NIIF_ICON_MASK  = $0000000F;

  {$EXTERNALSYM NIN_SELECT}
  NIN_SELECT      = $0400;
  {$EXTERNALSYM NINF_KEY}
  NINF_KEY        =  $1;
  {$EXTERNALSYM NIN_KEYSELECT}
  NIN_KEYSELECT   = NIN_SELECT or NINF_KEY;

  {$EXTERNALSYM NIN_BALLOONSHOW}
  NIN_BALLOONSHOW       = $0400 + 2;
  {$EXTERNALSYM NIN_BALLOONHIDE}
  NIN_BALLOONHIDE       = $0400 + 3;
  {$EXTERNALSYM NIN_BALLOONTIMEOUT}
  NIN_BALLOONTIMEOUT    = $0400 + 4;
  {$EXTERNALSYM NIN_BALLOONUSERCLICK}
  NIN_BALLOONUSERCLICK  = $0400 + 5;

type
  PNotifyIconDataA = ^TNotifyIconDataA;
  PNotifyIconDataW = ^TNotifyIconDataW;
  PNotifyIconData = PNotifyIconDataA;
  _NOTIFYICONDATAA = record
    cbSize: DWORD;
    Wnd: HWND;
    uID: UINT;
    uFlags: UINT;
    uCallbackMessage: UINT;
    hIcon: HICON;
    szTip: array [0..127] of AnsiChar;
    dwState: DWORD;
    dwStateMask: DWORD;
    szInfo: array [0..255] of AnsiChar;
    uTimeout: UINT;
    szInfoTitle: array [0..63] of AnsiChar;
    dwInfoFlags: DWORD;
  end;
  _NOTIFYICONDATAW = record
    cbSize: DWORD;
    Wnd: HWND;
    uID: UINT;
    uFlags: UINT;
    uCallbackMessage: UINT;
    hIcon: HICON;
    szTip: array [0..127] of WideChar;
    dwState: DWORD;
    dwStateMask: DWORD;
    szInfo: array [0..255] of WideChar;
    uTimeout: UINT;
    szInfoTitle: array [0..63] of WideChar;
    dwInfoFlags: DWORD;
  end;
  _NOTIFYICONDATA = _NOTIFYICONDATAA;
  TNotifyIconDataA = _NOTIFYICONDATAA;
  TNotifyIconDataW = _NOTIFYICONDATAW;
  TNotifyIconData = TNotifyIconDataA;
  {$EXTERNALSYM NOTIFYICONDATAA}
  NOTIFYICONDATAA = _NOTIFYICONDATAA;
  {$EXTERNALSYM NOTIFYICONDATAW}
  NOTIFYICONDATAW = _NOTIFYICONDATAW;
  {$EXTERNALSYM NOTIFYICONDATA}
  NOTIFYICONDATA = NOTIFYICONDATAA;

 var
   RM_TaskbarCreated: DWORD;

type
  TBalloonFlags = (bfNone = NIIF_NONE, bfInfo = NIIF_INFO,
    bfWarning = NIIF_WARNING, bfError = NIIF_ERROR);

  TCustomTrayIcon = class(TComponent)
  private
    FAnimate: Boolean;
    FData: TNotifyIconData;
    FIsClicked: Boolean;
    FCurrentIcon: TIcon;
    FIcon: TIcon;
    FIconList: TImageList;
    FPopupMenu: TPopupMenu;
    FTimer: TTimer;
    FHint: String;
    FIconIndex: Integer;
    FVisible: Boolean;
    FOnMouseMove: TMouseMoveEvent;
    FOnClick: TNotifyEvent;
    FOnDblClick: TNotifyEvent;
    FOnMouseDown: TMouseEvent;
    FOnMouseUp: TMouseEvent;
    FOnAnimate: TNotifyEvent;
    FBalloonHint: string;
    FBalloonTitle: string;
    FBalloonFlags: TBalloonFlags;
  protected
    procedure SetHint(const Value: string);
    function GetAnimateInterval: Cardinal;
    procedure SetAnimateInterval(Value: Cardinal);
    procedure SetAnimate(Value: Boolean);
    procedure SetBalloonHint(const Value: string);
    function GetBalloonTimeout: Integer;
    procedure SetBalloonTimeout(Value: Integer);
    procedure SetBalloonTitle(const Value: string);
    procedure SetVisible(Value: Boolean); virtual;
    procedure SetIconIndex(Value: Integer); virtual;
    procedure SetIcon(Value: TIcon);
    procedure SetIconList(Value: TImageList);
    procedure WindowProc(var Message: TMessage); virtual;
    procedure DoOnAnimate(Sender: TObject); virtual;
    property Data: TNotifyIconData read FData;
    function Refresh(Message: Integer): Boolean; overload;
  public
    constructor Create(Owner: TComponent); override;
    destructor Destroy; override;
    procedure Refresh; overload;
    procedure SetDefaultIcon;
    procedure ShowBalloonHint; virtual;
    property Animate: Boolean read FAnimate write SetAnimate default False;
    property AnimateInterval: Cardinal read GetAnimateInterval write SetAnimateInterval default 1000;
    property Hint: string read FHint write SetHint;
    property BalloonHint: string read FBalloonHint write SetBalloonHint;
    property BalloonTitle: string read FBalloonTitle write SetBalloonTitle;
    property BalloonTimeout: Integer read GetBalloonTimeout write SetBalloonTimeout default 3000;
    property BalloonFlags: TBalloonFlags read FBalloonFlags write FBalloonFlags default bfNone;
    property Icon: TIcon read FIcon write SetIcon;
    property Icons: TImageList read FIconList write SetIconList;
    property IconIndex: Integer read FIconIndex write SetIconIndex default 0;
    property PopupMenu: TPopupMenu read FPopupMenu write FPopupMenu;
    property Visible: Boolean read FVisible write SetVisible default False;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
    property OnDblClick: TNotifyEvent read FOnDblClick write FOnDblClick;
    property OnMouseMove: TMouseMoveEvent read FOnMouseMove write FOnMouseMove;
    property OnMouseUp: TMouseEvent read FOnMouseUp write FOnMouseUp;
    property OnMouseDown: TMouseEvent read FOnMouseDown write FOnMouseDown;
    property OnAnimate: TNotifyEvent read FOnAnimate write FOnAnimate;
  end;

  TTrayIcon = class(TCustomTrayIcon)
  published
    property Animate;
    property AnimateInterval;
    property Hint;
    property BalloonHint;
    property BalloonTitle;
    property BalloonTimeout;
    property BalloonFlags;
    property Icon;
    property Icons;
    property IconIndex;
    property PopupMenu;
    property Visible;
    property OnClick;
    property OnDblClick;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseDown;
    property OnAnimate;
  end;


implementation

{ TTrayIcon}

constructor TCustomTrayIcon.Create(Owner: TComponent);
begin
  inherited;
  FAnimate := False;
  FBalloonFlags := bfNone;
  BalloonTimeout := 3000;
  FIcon := TIcon.Create;
  FCurrentIcon := TIcon.Create;
  FTimer := TTimer.Create(Nil);
  FIconIndex := 0;
  FVisible := False;
  FIsClicked := False;
  FTimer.Enabled := False;
  FTimer.OnTimer := DoOnAnimate;
  FTimer.Interval := 1000;

  if not (csDesigning in ComponentState) then
  begin
    FillChar(FData, SizeOf(FData), 0);
    FData.cbSize := SizeOf(FData);
    FData.Wnd := Classes.AllocateHwnd(WindowProc);
    FData.uID := FData.Wnd;
    FData.uTimeout := 3000;
    FData.hIcon := FCurrentIcon.Handle;
    FData.uFlags := NIF_ICON or NIF_MESSAGE;
    FData.uCallbackMessage := WM_SYSTEM_TRAY_MESSAGE;
    StrPLCopy(FData.szTip, Application.Title, SizeOf(FData.szTip) - 1);

    if Length(Application.Title) > 0 then
       FData.uFlags := FData.uFlags or NIF_TIP;

    Refresh;
  end;
end;

destructor TCustomTrayIcon.Destroy;
begin
  if not (csDesigning in ComponentState) then
    Refresh(NIM_DELETE);

  FCurrentIcon.Free;
  FIcon.Free;
  FTimer.Free;
  Classes.DeallocateHWnd(FData.Wnd);
  inherited;
end;

procedure TCustomTrayIcon.SetVisible(Value: Boolean);
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    if (not FAnimate) or (FAnimate and FCurrentIcon.Empty) then
      SetDefaultIcon;

    if not (csDesigning in ComponentState) then
    begin
      if FVisible then
      begin
        if not Refresh(NIM_ADD) then
          raise EOutOfResources.Create(STrayIconCreateError);
      end
      else if not (csLoading in ComponentState) then
      begin
        if not Refresh(NIM_DELETE) then
          raise EOutOfResources.Create(STrayIconRemoveError);
      end;
      if FAnimate then
        FTimer.Enabled := Value;
    end;
  end;
end;

procedure TCustomTrayIcon.SetIconList(Value: TImageList);
begin
  if FIconList <> Value then
  begin
    FIconList := Value;
    if not (csDesigning in ComponentState) then
    begin
      if Assigned(FIconList) then
        FIconList.GetIcon(FIconIndex, FCurrentIcon)
      else
        SetDefaultIcon;
      Refresh;
    end;
  end;
end;

procedure TCustomTrayIcon.SetHint(const Value: string);
begin
  if CompareStr(FHint, Value) <> 0 then
  begin
    FHint := Value;
    StrPLCopy(FData.szTip, FHint, SizeOf(FData.szTip) - 1);

    if Length(Hint) > 0 then
      FData.uFlags := FData.uFlags or NIF_TIP
    else
      FData.uFlags := FData.uFlags and not NIF_TIP;

    Refresh;
  end;
end;

function TCustomTrayIcon.GetAnimateInterval: Cardinal;
begin
  Result := FTimer.Interval;
end;

procedure TCustomTrayIcon.SetAnimateInterval(Value: Cardinal);
begin
  FTimer.Interval := Value;
end;

procedure TCustomTrayIcon.SetAnimate(Value: Boolean);
begin
  if FAnimate <> Value then
  begin
    FAnimate := Value;
    if not (csDesigning in ComponentState) then
    begin
      if (FIconList <> nil) and (FIconList.Count > 0) and Visible then
        FTimer.Enabled := Value;
      if (not FAnimate) and (not FCurrentIcon.Empty) then
        FIcon.Assign(FCurrentIcon);
    end;
  end;
end;

{ Message handler for the hidden shell notification window. Most messages
  use WM_SYSTEM_TRAY_MESSAGE as the Message ID, with WParam as the ID of the
  shell notify icon data. LParam is a message ID for the actual message, e.g.,
  WM_MOUSEMOVE. Another important message is WM_ENDSESSION, telling the shell
  notify icon to delete itself, so Windows can shut down.

  Send the usual events for the mouse messages. Also interpolate the OnClick
  event when the user clicks the left button, and popup the menu, if there is
  one, for right click events. }
procedure TCustomTrayIcon.WindowProc(var Message: TMessage);

  { Return the state of the shift keys. }
  function ShiftState: TShiftState;
  begin
    Result := [];

    if GetKeyState(VK_SHIFT) < 0 then
      Include(Result, ssShift);
    if GetKeyState(VK_CONTROL) < 0 then
      Include(Result, ssCtrl);
    if GetKeyState(VK_MENU) < 0 then
      Include(Result, ssAlt);
  end;

var
  Point: TPoint;
  Shift: TShiftState;
begin
  case Message.Msg of
    WM_QUERYENDSESSION:
      Message.Result := 1;

    WM_ENDSESSION:
    begin
      if TWmEndSession(Message).EndSession then
        Refresh(NIM_DELETE);
    end;

    WM_SYSTEM_TRAY_MESSAGE:
    begin
      case Message.lParam of
        WM_MOUSEMOVE:
        begin
          if Assigned(FOnMouseMove) then
          begin
            Shift := ShiftState;
            GetCursorPos(Point);
            FOnMouseMove(Self, Shift, Point.X, Point.Y);
          end;
        end;

        WM_LBUTTONDOWN:
        begin
          if Assigned(FOnMouseDown) then
          begin
            Shift := ShiftState + [ssLeft];
            GetCursorPos(Point);
            FOnMouseDown(Self, mbLeft, Shift, Point.X, Point.Y);
          end;

          FIsClicked := True;
        end;

        WM_LBUTTONUP:
        begin
          Shift := ShiftState + [ssLeft];
          GetCursorPos(Point);
          if FIsClicked and Assigned(FOnClick) then
          begin
            FOnClick(Self);
            FIsClicked := False;
          end;
          if Assigned(FOnMouseUp) then
            FOnMouseUp(Self, mbLeft, Shift, Point.X, Point.Y);
        end;

        WM_RBUTTONDOWN:
        begin
          if Assigned(FOnMouseDown) then
          begin
            Shift := ShiftState + [ssRight];
            GetCursorPos(Point);
            FOnMouseDown(Self, mbRight, Shift, Point.X, Point.Y);
          end;
        end;

        WM_RBUTTONUP:
        begin
          Shift := ShiftState + [ssRight];
          GetCursorPos(Point);
          if Assigned(FOnMouseUp) then
            FOnMouseUp(Self, mbRight, Shift, Point.X, Point.Y);
          if Assigned(FPopupMenu) then
          begin
            SetForegroundWindow(Application.Handle);
            Application.ProcessMessages;
            FPopupMenu.AutoPopup := False;
            FPopupMenu.PopupComponent := Owner;
            FPopupMenu.Popup(Point.x, Point.y);
          end;
        end;

        WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK:
          if Assigned(FOnDblClick) then
            FOnDblClick(Self);

        WM_MBUTTONDOWN:
        begin
          if Assigned(FOnMouseDown) then
          begin
            Shift := ShiftState + [ssMiddle];
            GetCursorPos(Point);
            FOnMouseDown(Self, mbMiddle, Shift, Point.X, Point.Y);
          end;
        end;

        WM_MBUTTONUP:
        begin
          if Assigned(FOnMouseUp) then
          begin
            Shift := ShiftState + [ssMiddle];
            GetCursorPos(Point);
            FOnMouseUp(Self, mbMiddle, Shift, Point.X, Point.Y);
          end;
        end;

        NIN_BALLOONHIDE, NIN_BALLOONTIMEOUT:
        begin
          FData.uFlags := FData.uFlags and not NIF_INFO;
        end;
      end;
    end;

    else if (Message.Msg = RM_TaskBarCreated) and Visible then
      Refresh(NIM_ADD);
  end;
end;

procedure TCustomTrayIcon.Refresh;
begin
  if not (csDesigning in ComponentState) then
  begin
    FData.hIcon := FCurrentIcon.Handle;

    if Visible then
      Refresh(NIM_MODIFY);
  end;
end;

function TCustomTrayIcon.Refresh(Message: Integer): Boolean;
begin
  Result := Shell_NotifyIcon(Message, @FData);
end;

procedure TCustomTrayIcon.SetIconIndex(Value: Integer);
begin
  if FIconIndex <> Value then
  begin
    FIconIndex := Value;
    if not (csDesigning in ComponentState) then
    begin
      if Assigned(FIconList) then
        FIconList.GetIcon(FIconIndex, FCurrentIcon);
      Refresh;
    end;
  end;
end;

procedure TCustomTrayIcon.DoOnAnimate(Sender: TObject);
begin
  if Assigned(FOnAnimate) then
    FOnAnimate(Self);
  if Assigned(FIconList) and (FIconIndex < FIconList.Count - 1) then
    IconIndex := FIconIndex + 1
  else
    IconIndex := 0;
  Refresh;
end;

procedure TCustomTrayIcon.SetIcon(Value: TIcon);
begin
  FIcon.Assign(Value);
  FCurrentIcon.Assign(Value);
  Refresh;
end;

procedure TCustomTrayIcon.SetBalloonHint(const Value: string);
begin
  if CompareStr(FBalloonHint, Value) <> 0 then
  begin
    FBalloonHint := Value;
    StrPLCopy(FData.szInfo, FBalloonHint, SizeOf(FData.szInfo) - 1);
    Refresh(NIM_MODIFY);
  end;
end;

procedure TCustomTrayIcon.SetDefaultIcon;
begin
  if not FIcon.Empty then
    FCurrentIcon.Assign(FIcon)
  else
    FCurrentIcon.Assign(Application.Icon);
  Refresh;
end;

procedure TCustomTrayIcon.SetBalloonTimeout(Value: Integer);
begin
  FData.uTimeout := Value;
end;

function TCustomTrayIcon.GetBalloonTimeout: Integer;
begin
  Result := FData.uTimeout;
end;

procedure TCustomTrayIcon.ShowBalloonHint;
begin
  FData.uFlags := FData.uFlags or NIF_INFO;
  FData.dwInfoFlags := Integer(FBalloonFlags);
  Refresh(NIM_MODIFY);
end;

procedure TCustomTrayIcon.SetBalloonTitle(const Value: string);
begin
  if CompareStr(FBalloonTitle, Value) <> 0 then
  begin
    FBalloonTitle := Value;
    StrPLCopy(FData.szInfoTitle, FBalloonTitle, SizeOf(FData.szInfoTitle) - 1);
    Refresh(NIM_MODIFY);
  end;
end;

initialization
  RM_TaskBarCreated := RegisterWindowMessage('TaskbarCreated');

end.
 