unit cbClass;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, ExtCtrls;

type
  TThreadTimer = class(TComponent)
  private
    FEnabled: Boolean;
    FInterval: Cardinal;
    FOnTimer: TNotifyEvent;
    FWindowHandle: HWND;
    FSyncEvent: Boolean;
    FThreaded: Boolean;
    FTimerThread: TThread;
    FThreadPriority: TThreadPriority;
    procedure UpdateTimer;
    procedure WndProc(var Message: TMessage);
    procedure SetThreaded(Value: Boolean);
    procedure SetThreadPriority(Value: TThreadPriority);
    procedure SetEnabled(Value: Boolean);
    procedure SetInterval(Value: Cardinal);
    procedure SetOnTimer(Value: TNotifyEvent);
  protected
    procedure Timer; dynamic;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Synchronize(Method: TThreadMethod);
  published
    property Enabled: Boolean read FEnabled write SetEnabled default True;
    property Interval: Cardinal read FInterval write SetInterval default 1000;
    property SyncEvent: Boolean read FSyncEvent write FSyncEvent default True;
    property Threaded: Boolean read FThreaded write SetThreaded default True;
    property ThreadPriority: TThreadPriority read FThreadPriority write
      SetThreadPriority default tpNormal;
    property OnTimer: TNotifyEvent read FOnTimer write SetOnTimer;
  end;
  

  TScrollTextBox = class(TPaintBox)
  private
    FTimer: TThreadTimer;
    FMemoryImage: TBitmap;
    FActive: Boolean;
    FAlignment: TAlignment;
    FLines: TStrings;
    FCycled: Boolean;
    FScrollCnt: Integer;
    FMaxScroll: Integer;
    FFirstLine: Integer;
    FTxtDivider: Byte;
    FTxtRect: TRect;
    FPaintRect: TRect;
    function GetInflateWidth: Integer;
    procedure UpdateMemoryImage;
    procedure RecalcDrawRect;
    procedure TimerExpired(Sender: TObject);
    procedure SetActive(Value: Boolean);
    procedure SetAlignment(Value: TAlignment);
    procedure SetLines(Value: TStrings);
    procedure PaintText;
    procedure Play;
    procedure Stop;
    procedure WMSize(var Message: TMessage); message WM_SIZE;
  protected  
    procedure Paint; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Active: Boolean read FActive write SetActive default False;
    property Alignment: TAlignment read FAlignment write SetAlignment default taCenter;
    property Cycled: Boolean read FCycled write FCycled default False;
    property Lines: TStrings read FLines write SetLines;
  end;


implementation

uses Forms;

type
  TTimerThread = class(TThread)
  private
    FOwner: TThreadTimer;
    FInterval: Cardinal;
    FException: Exception;
    procedure HandleException;
  protected
    procedure Execute; override;
  public
    constructor Create(Timer: TThreadTimer; Enabled: Boolean);
  end;

const
  Alignments: array [TAlignment] of Word = (DT_LEFT, DT_RIGHT, DT_CENTER);


{ ---------------------------------------------------------------------------- }

function WidthOf(R: TRect): Integer;
begin
  Result := R.Right - R.Left;
end;

{ ---------------------------------------------------------------------------- }

function HeightOf(R: TRect): Integer;
begin
  Result := R.Bottom - R.Top;
end;

{ ---------------------------------------------------------------------------- }

{ TTimerThread }

constructor TTimerThread.Create(Timer: TThreadTimer; Enabled: Boolean);
begin
  FOwner := Timer;
  inherited Create( not Enabled );
  FInterval := 1000;
  FreeOnTerminate := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TTimerThread.HandleException;
begin
  if not ( FException is EAbort ) then begin
    if Assigned( Application.OnException ) then
      Application.OnException( Self, FException )
    else
      Application.ShowException( FException );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TTimerThread.Execute;

  { ------------------------------------------------- }

  function ThreadClosed: Boolean;
  begin
    Result := Terminated or Application.Terminated or (FOwner = nil);
  end;

  { ------------------------------------------------- }

begin
  repeat
    if ( not ThreadClosed ) then
    begin
      if SleepEx( FInterval, False ) = 0 then
      begin
        if not ( ThreadClosed ) and ( FOwner.FEnabled ) then
        begin
          if FOwner.SyncEvent then
            FOwner.Synchronize( FOwner.Timer )
          else begin
            try
              FOwner.Timer;
            except
              on E: Exception do
              begin
                FException := E;
                HandleException;
              end;
            end;
          end;
        end;
      end;
    end;
  until Terminated;
end;

{ ---------------------------------------------------------------------------- }

{ TThreadTimer }

constructor TThreadTimer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FEnabled := True;
  FInterval := 1000;
  FSyncEvent := True;
  FThreaded := True;
  FThreadPriority := tpNormal;
  FTimerThread := TTimerThread.Create(Self, False);
end;

{ ---------------------------------------------------------------------------- }

destructor TThreadTimer.Destroy;
begin
  Destroying;
  FEnabled := False;
  FOnTimer := nil;
  while FTimerThread.Suspended do
    FTimerThread.Resume;
  FTimerThread.Terminate;
  if ( FWindowHandle ) <> 0 then
  begin
    KillTimer( FWindowHandle, 1 );
    Classes.DeallocateHWnd( FWindowHandle );
  end;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.Synchronize(Method: TThreadMethod);
begin
  if ( FTimerThread <> nil ) then
  begin
    with TTimerThread( FTimerThread ) do
    begin
      if Suspended or Terminated then
        Method
      else
        TTimerThread( FTimerThread ).Synchronize( Method );
    end;
  end else
    Method;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.UpdateTimer;
begin
  if FThreaded then
  begin
    if ( FWindowHandle <> 0 ) then
    begin
      KillTimer( FWindowHandle, 1 );
      Classes.DeallocateHWnd( FWindowHandle );
      FWindowHandle := 0;
    end;
    if not FTimerThread.Suspended then
      FTimerThread.Suspend;
    TTimerThread( FTimerThread ).FInterval := FInterval;
    if ( FInterval <> 0 ) and ( FEnabled ) and ( Assigned( FOnTimer ) ) then
    begin
      FTimerThread.Priority := FThreadPriority;
      while FTimerThread.Suspended do
        FTimerThread.Resume;
    end;
  end else
  begin
    if not FTimerThread.Suspended then FTimerThread.Suspend;
    if FWindowHandle = 0 then
      FWindowHandle := Classes.AllocateHWnd( WndProc )
    else
      KillTimer( FWindowHandle, 1 );
    if ( FInterval <> 0 ) and FEnabled and Assigned( FOnTimer ) then
      if SetTimer( FWindowHandle, 1, FInterval, nil ) = 0 then
        raise EOutOfResources.Create( 'Not enough timers available' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.WndProc(var Message: TMessage);
begin
  if ( Message.Msg = WM_TIMER ) then
  begin
    try
      Timer;
    except
      Application.HandleException( Self );
    end
  end else
    Message.Result := DefWindowProc( FWindowHandle, Message.Msg,
      Message.wParam, Message.lParam );
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.SetEnabled(Value: Boolean);
begin
  if ( Value <> FEnabled ) then
  begin
    FEnabled := Value;
    UpdateTimer;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.SetInterval(Value: Cardinal);
begin
  if ( Value <> FInterval ) then
  begin
    FInterval := Value;
    UpdateTimer;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.SetOnTimer(Value: TNotifyEvent);
begin
  if ( Assigned( FOnTimer ) <> Assigned( Value ) ) then
  begin
    FOnTimer := Value;
    UpdateTimer;
  end else
    FOnTimer := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.SetThreaded(Value: Boolean);
begin
  if ( Value <> FThreaded ) then
  begin
    FThreaded := Value;
    UpdateTimer;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.SetThreadPriority(Value: TThreadPriority);
begin
  if ( Value <> FThreadPriority ) then
  begin
    FThreadPriority := Value;
    if FThreaded then UpdateTimer;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TThreadTimer.Timer;
begin
  if FEnabled and not ( csDestroying in ComponentState ) and
    Assigned( FOnTimer ) then FOnTimer( Self );
end;

{ ---------------------------------------------------------------------------- }

{ TScrollTextBox }

constructor TScrollTextBox.Create(AOwner: TComponent);
begin
  inherited Create( AOwner );
  FTimer := TThreadTimer.Create( Self );
  FTimer.Enabled := False;
  FTimer.SyncEvent := True;
  FTimer.OnTimer := Self.TimerExpired;
  FTimer.Interval := 30;
  FScrollCnt := 0;
  FAlignment := taCenter;
  FActive := False;
  FLines := TStringList.Create;
  FTxtDivider := 1;
  {}
end;

{ ---------------------------------------------------------------------------- }

destructor TScrollTextBox.Destroy;
begin
  FTimer.Free;
  FLines.Free;
  inherited Destroy;
end;

function TScrollTextBox.GetInflateWidth: Integer;
begin
  Result := 1;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.UpdateMemoryImage;
var
  Metrics: TTextMetric;
begin
  if FMemoryImage = nil then FMemoryImage := TBitmap.Create;
  FMemoryImage.Canvas.Lock;
  try
    FFirstLine := 0;
    while (FFirstLine < FLines.Count) and (Trim(FLines[FFirstLine]) = '') do
      Inc(FFirstLine);
    Canvas.Font := Self.Font;
    GetTextMetrics(Canvas.Handle, Metrics);
    FTxtDivider := Metrics.tmHeight + Metrics.tmExternalLeading;
    Inc( FTxtDivider );
    RecalcDrawRect;
    FMaxScroll := ((FLines.Count - FFirstLine) * FTxtDivider) +
      HeightOf( FTxtRect );
    FMemoryImage.Width := WidthOf(FTxtRect);
    FMemoryImage.Height := HeightOf(FTxtRect);
    FMemoryImage.Canvas.Font := Self.Font;
    FMemoryImage.Canvas.Brush.Color := Color;
    SetBkMode( FMemoryImage.Canvas.Handle, TRANSPARENT );
  finally
    FMemoryImage.Canvas.UnLock;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.WMSize(var Message: TMessage);
begin
  inherited;
  if Active then
  begin
    UpdateMemoryImage;
    Invalidate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.RecalcDrawRect;
const
  MinOffset = 3;
var
  InflateWidth: Integer;
  LastLine: Integer;
begin
  FTxtRect := GetClientRect;
  FPaintRect := FTxtRect;
  InflateWidth := GetInflateWidth;
  InflateRect(FPaintRect, -InflateWidth, -InflateWidth);
  Inc(InflateWidth, MinOffset);
  InflateRect(FTxtRect, -InflateWidth, -InflateWidth);
  with FTxtRect do
    if (Left >= Right) or (Top >= Bottom) then FTxtRect := Rect(0, 0, 0, 0);
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.TimerExpired(Sender: TObject);
begin
  if ( FScrollCnt < FMaxScroll ) then
  begin
    Inc( FScrollCnt );
    if Assigned( FMemoryImage ) then PaintText;
  end else
  if ( Cycled ) then
  begin
    FScrollCnt := 0;
    if Assigned( FMemoryImage ) then PaintText;
  end else
  begin
    FTimer.Synchronize(Stop);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.SetActive(Value: Boolean);
begin
  if ( Value <> FActive ) then
  begin
    FActive := Value;
    if FActive then begin
      FScrollCnt := 0;
      UpdateMemoryImage;
      try
        FTimer.Enabled := True;
      except
        FActive := False;
        FTimer.Enabled := False;
        raise;
      end;
    end
    else begin
      FMemoryImage.Canvas.Lock;
      FTimer.Enabled := False;
      FScrollCnt := 0;
      FMemoryImage.Free;
      FMemoryImage := nil;
//      if (csDesigning in ComponentState) and
//        not (csDestroying in ComponentState) then
//        ValidParentForm(Self).Designer.Modified;
    end;
//    if not (csDestroying in ComponentState) then
//      for I := 0 to Pred(ControlCount) do begin
//        if FActive then begin
//          if Controls[I].Visible then FHiddenList.Add(Controls[I]);
//          if not (csDesigning in ComponentState) then
//            Controls[I].Visible := False
//        end
//        else if FHiddenList.IndexOf(Controls[I]) >= 0 then begin
//          Controls[I].Visible := True;
//          Controls[I].Invalidate;
//          if (csDesigning in ComponentState) then Controls[I].Update;
//        end;
//      end;
//    if not FActive then FHiddenList.Clear;
    Invalidate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.SetAlignment(Value: TAlignment);
begin
  if FAlignment <> Value then
  begin
    FAlignment := Value;
    if Active then Invalidate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.SetLines(Value: TStrings);
begin
  FLines.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.Paint;
begin
  //inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.PaintText;
var
  STmp: array[0..255] of Char;
  R: TRect;
  I: Integer;
  Flags: Longint;
begin
  if (FLines.Count = 0) or IsRectEmpty(FTxtRect) then //or ( not Canvas.HandleAllocated ) then
    Exit;
  FMemoryImage.Canvas.Lock;
  try
//    with FMemoryImage.Canvas do
//    begin
//      I := SaveDC(Handle);
//      try
//        with FTxtRect do
//          MoveWindowOrg(Handle, -Left, -Top);
//        Brush.Color := Self.Color;
//      finally
//        RestoreDC(Handle, I);
//        SetBkMode(Handle, Transparent);
//      end;
//    end;
    R := Bounds(0, 0, WidthOf(FTxtRect), HeightOf(FTxtRect));
    R.Top := R.Bottom - FScrollCnt;
    R.Bottom := R.Top + FTxtDivider;
    Flags := DT_EXPANDTABS or Alignments[FAlignment] or DT_SINGLELINE or
      DT_NOCLIP or DT_NOPREFIX;
    Flags := DrawTextBiDiModeFlags(Flags);
    for I := FFirstLine to FLines.Count do
    begin
      if I = FLines.Count then
        StrCopy(STmp, ' ' )
      else
        StrPLCopy(STmp, FLines[I], SizeOf(STmp) - 1);
      if ( R.Top >= HeightOf(FTxtRect) ) then
        Break
      else if R.Bottom > 0 then
        DrawText(FMemoryImage.Canvas.Handle, STmp, -1, R, Flags);
      OffsetRect(R, 0, FTxtDivider);
    end;
    Canvas.Lock;
    try
      BitBlt(Canvas.Handle, FTxtRect.Left, FTxtRect.Top, FMemoryImage.Width,
        FMemoryImage.Height, FMemoryImage.Canvas.Handle, 0, 0, SRCCOPY);
    finally
      Canvas.Unlock;
    end;
  finally
    FMemoryImage.Canvas.Unlock;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.Play;
begin
  SetActive( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TScrollTextBox.Stop;
begin
  SetActive( False );
end;

{ ---------------------------------------------------------------------------- }

end.
