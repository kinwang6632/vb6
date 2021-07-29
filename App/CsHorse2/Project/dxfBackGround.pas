{**************************************************}
{                                                  }
{         Delphi VCL Extensions                    }
{                                                  }
{ Copyright (c) 1998-2001 Developer Express Inc.   }
{           All Rights Reserved                    }
{                                                  }
{**************************************************}

unit dxfBackGround;

{$I dxFLVer.inc}

interface

uses Windows, SysUtils, Classes, Controls, Forms, Graphics, ExtCtrls{$IFDEF DELPHI4}, ImgList{$ENDIF};

type
  TdxfBackGround = class;

  TdxfFillStyle = (fsHorz, fsVert);

  TdxfBkColor = class(TPersistent)
  private
    FBeginColor: TColor;
    FEndColor: TColor;
    FFillStyle: TdxfFillStyle;
    FBG: TdxfBackGround;
    procedure SetBeginColor(Value: TColor);
    procedure SetEndColor(Value: TColor);
    procedure SetFillStyle(Value: TdxfFillStyle);
  public
    constructor Create(ABackGround: TdxfBackGround);
  published
    property BeginColor: TColor read FBeginColor write SetBeginColor;
    property EndColor: TColor read FEndColor write SetEndColor;
    property FillStyle: TdxfFillStyle read FFillStyle write SetFillStyle;
  end;

  TdxfBkAnimate = class(TPersistent)
  private
    FImageList: TImageList;
    FImageChangeLink: TChangeLink;
    FSpeed: Integer;
    FTimer: TTimer;
    FBitmap: TBitmap;
    FStaff: Integer;
    FForm: TForm;
    procedure SetImageList(Value: TImageList);
    procedure SetSpeed(Value: Integer);
  public
    constructor Create;
    destructor Destroy; override;
  published
    property ImageList: TImageList read FImageList write SetImageList;
    property Speed: Integer read FSpeed write SetSpeed;
  end;

  TdxfBackGround = class(TComponent)
  private
    FForm: TForm;
    FBkAnimate: TdxfBkAnimate;
    FBkColor: TdxfBkColor;
    FPicture: TPicture;
    FOldPaint: TNotifyEvent;
    FOldResize: TNotifyEvent;

    procedure FillBackGround(Sender: TObject);
    procedure RefreshBackGround(Sender: TObject);
    procedure Animate(Sender: TObject);
    procedure SetPicture(Value: TPicture);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property BkColor: TdxfBkColor read FBkColor write FBkColor;
    property BkPicture: TPicture read FPicture write SetPicture;
    property BkAnimate: TdxfBkAnimate read FBkAnimate write FBkAnimate;
  end;

procedure DrawGradient(Canvas: TCanvas; const ARect: TRect;
  FromColor, ToColor: TColor; IsVertical: Boolean);

implementation

procedure DrawGradient(Canvas: TCanvas; const ARect: TRect;
  FromColor, ToColor: TColor; IsVertical: Boolean);
var
  SR: TRect;
  H, I: Integer;
  R, G, B: Byte;
  FromR, ToR, FromG, ToG, FromB, ToB: Byte;
begin
  FromR := GetRValue(FromColor);
  FromG := GetGValue(FromColor);
  FromB := GetBValue(FromColor);
  ToR := GetRValue(ToColor);
  ToG := GetGValue(ToColor);
  ToB := GetBValue(ToColor);
  SR := ARect;
  with ARect do
    if IsVertical then
      H := Bottom - Top
    else
      H := Right - Left;

  for I := 0 to 255 do
  begin
    if IsVertical then
      SR.Bottom := ARect.Top + MulDiv(I + 1, H, 256)
    else
      SR.Right := ARect.Left + MulDiv(I + 1, H, 256);
    with Canvas do
    begin
      R := FromR + MulDiv(I, ToR - FromR, 255);
      G := FromG + MulDiv(I, ToG - FromG, 255);
      B := FromB + MulDiv(I, ToB - FromB, 255);
      Brush.Color := RGB(R, G, B);
      FillRect(SR);
    end;
    if IsVertical then
      SR.Top := SR.Bottom
    else
      SR.Left := SR.Right;
  end;
end;

{ TdxfBkAnimate }
constructor TdxfBkAnimate.Create;
begin
  inherited Create;
  FImageChangeLink := TChangeLink.Create;
  FBitmap := TBitmap.Create;
  FTimer:= TTimer.Create(nil);
  FSpeed := 700;
  FTimer.Interval := 301;
end;

destructor TdxfBkAnimate.Destroy;
begin
  FImageChangeLink.Free;
  FTimer.Free;
  FBitmap.Free;
  inherited;
end;

procedure TdxfBkAnimate.SetImageList(Value: TImageList);
begin
  if FImageList <> Value then
  begin
    if (FImageList <> nil) and not (csDestroying in FImageList.ComponentState) then
      FImageList.UnRegisterChanges(FImageChangeLink);
    FImageList := Value;
    if FImageList <> nil then
    begin
      FImageList.RegisterChanges(FImageChangeLink);
      FImageList.BkColor := FForm.Color;
      FImageList.BlendColor := FForm.Color;
    end;
  end;
end;

procedure TdxfBkAnimate.SetSpeed(Value: Integer);
begin
  if FSpeed < 1 then
    FSpeed := 1;
  if FSpeed > 1000 then
    FSpeed := 1000;
  FSpeed := Value;
  FTimer.Interval := 1001 - FSpeed;
end;

{ TdxfBkColor }
constructor TdxfBkColor.Create(ABackGround: TdxfBackGround);
begin
  inherited Create;
  FBG := ABackGround;
  FBeginColor := clBlue;
  FEndColor := clBlack;
  FFillStyle := fsVert;
end;

procedure TdxfBkColor.SetBeginColor(Value: TColor);
begin
  if FBeginColor <> Value then
  begin
    FBeginColor := Value;
    FBG.RefreshBackGround(FBG.FForm);
  end;
end;

procedure TdxfBkColor.SetEndColor(Value: TColor);
begin
  if FEndColor <> Value then
  begin
    FEndColor := Value;
    FBG.RefreshBackGround(FBG.FForm);
  end;
end;

procedure TdxfBkColor.SetFillStyle(Value: TdxfFillStyle);
begin
  if FFillStyle <> Value then
  begin
    FFillStyle := Value;
    FBG.RefreshBackGround(FBG.FForm);
  end;
end;

{ TdxfBackGround }
constructor TdxfBackGround.Create(AOwner: TComponent);
var
  I: Integer;
begin
  inherited Create(AOwner);

  if not (AOwner is TForm) then
    raise Exception.Create('TdxfBackGround can be placed only on a TForm descendant');
  FForm := TForm(AOwner);

  for I := 0 to FForm.ComponentCount - 1 do
    if (FForm.Components[I] is TdxfBackGround) and (FForm.Components[I] <> Self) then
      raise Exception.Create('Form should have only a single TdxfBackGround');

  FBkColor := TdxfBkColor.Create(Self);
  FPicture := TPicture.Create;
  FBkAnimate := TdxfBkAnimate.Create;
  FBkAnimate.FTimer.OnTimer := Animate;
  FBkAnimate.FForm := FForm;

  if not (csDesigning in ComponentState) then
  begin
    if Assigned(FForm.OnPaint) then
      FOldPaint := FForm.OnPaint;
    if Assigned(FForm.OnResize) then
      FOldResize := FForm.OnResize;
    FForm.OnPaint := FillBackGround;
    FForm.OnResize := RefreshBackGround;
  end;
end;

destructor TdxfBackGround.Destroy;
begin
  if (FForm <> nil) and not (csDesigning in ComponentState) then
  begin
    FForm.OnPaint := FOldPaint;
    FForm.OnResize := FOldResize;
  end;
  FBkColor.Free;
  FPicture.Free;
  FBkAnimate.Free;
  inherited;
end;

procedure TdxfBackGround.Animate(Sender: TObject);
begin
  if not (csDesigning in ComponentState) and (FBkAnimate.ImageList <> nil) and
    (FBkAnimate.ImageList.Count > 0) then
  begin
    InvalidateRect(FForm.Handle, nil, False);
    FBkAnimate.FStaff := (FBkAnimate.FStaff + 1) mod FBkAnimate.ImageList.Count;
  end;
end;

procedure TdxfBackGround.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (FBkAnimate <> nil) and (AComponent = FBkAnimate.FImageList) then
    FBkAnimate.FImageList := nil;
end;

procedure TdxfBackGround.FillBackGround(Sender: TObject);
var
  X, Y, AHeight, AWidth: Integer;
begin
  AHeight := FForm.ClientHeight;
  AWidth  := FForm.ClientWidth;
  try
    if (FBkAnimate.FImageList <> nil) and (FBkAnimate.FImageList.Count > 0) then
    begin
      FBkAnimate.FImageList.GetBitmap(FBkAnimate.FStaff, FBkAnimate.FBitmap);
      Y := 0;
      while Y < AHeight do
      begin
        X := 0;
        while X < AWidth do
        begin
          FForm.Canvas.Draw(X, Y, FBkAnimate.FBitmap);
          Inc(X, FBkAnimate.FBitmap.Width);
        end;
        Inc(Y, FBkAnimate.FBitmap.Height);
      end;
      Exit;
    end;

    if (FPicture.Graphic <> nil) and not FPicture.Graphic.Empty then
    begin
      Y := 0;
      while Y < AHeight do
      begin
        X := 0;
        while X < AWidth do begin
          FForm.Canvas.Draw(X, Y, FPicture.Bitmap);
          Inc(X, FPicture.Width);
        end;
        Inc(Y, FPicture.Height);
      end;
      Exit;
    end;

    with FBkColor do
      DrawGradient(FForm.Canvas, FForm.ClientRect,
        FBeginColor, FEndColor, FFillStyle = fsVert);
  finally
    if Assigned(FOldPaint) then
      FOldPaint(Sender);
  end;
end;

procedure TdxfBackGround.RefreshBackGround(Sender: TObject);
begin
  if Assigned(FOldResize) then
    FOldResize(Sender);
  if ((FPicture.Graphic = nil) or (FPicture.Graphic.Empty)) and
    ((FBkAnimate.ImageList = nil) or (FBkAnimate.ImageList.Count = 0)) then
    InvalidateRect(FForm.Handle, nil, False);
end;

procedure TdxfBackGround.SetPicture(Value: TPicture);
begin
  FPicture.Assign(Value);
end;

end.
