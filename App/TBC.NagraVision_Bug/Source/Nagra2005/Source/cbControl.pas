unit cbControl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
  ExtCtrls, Dialogs;


type

  TTransparentControl = class(TCustomControl)
  private
    procedure WMEraseBkgnd(var Msg: TMessage); message WM_ERASEBKGND;
  protected
    procedure Paint; override;
  public
    constructor Create(AOwner: TComponent); override;
  end;

  TFlickerControl = class(TTransparentControl)
  private
    FTimer: TTimer;
    FBitmap: TBitmap;
    FImage: TImage;
    FMode: Integer;
    FResName: String;
    FFlickered: Boolean;
    FFlickerSpeed: Cardinal;
    procedure SetFlickered(const Value: Boolean);
    procedure SetFlickerSpeed(const Value: Cardinal);
    procedure SetMode(const Value: Integer);
    procedure LoadResourceBmp;
    procedure TimerExpired(Sender: TObject);
  protected
    procedure SetParent(AParent: TWinControl); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Flickered: Boolean read FFlickered write SetFlickered;
    property FlickerSpeed: Cardinal read FFlickerSpeed write SetFlickerSpeed;
    property Mode: Integer read FMode write SetMode;
  end;


implementation

{$R Flicker.res}

{ TTransparentControl }

constructor TTransparentControl.Create(AOwner: TComponent);
begin
  inherited Create( AOwner );
  Width := 100;
  Height := 100;
end;

{ ---------------------------------------------------------------------------- }

procedure TTransparentControl.Paint;
begin
  inherited;
  Canvas.Brush.Style := bsClear;
  Canvas.FillRect( GetClientRect );
end;

{ ---------------------------------------------------------------------------- }

procedure TTransparentControl.WMEraseBkgnd(var Msg: TMessage);
begin
  Msg.Result := 1;
end;

{ ---------------------------------------------------------------------------- }

{ TFlickerControl }

constructor TFlickerControl.Create(AOwner: TComponent);
begin
  inherited Create( AOwner );
  FTimer := TTimer.Create( Self );
  FTimer.OnTimer := TimerExpired;
  FBitmap := TBitmap.Create;
  FImage := TImage.Create( nil );
  FImage.Parent := Self;
  FImage.AutoSize := True;
  FImage.Transparent := True;
  FTimer.Enabled := False;
  Mode := MB_ICONWARNING;
  FFlickered := False;
  FFlickerSpeed := 500;
  FTimer.Interval := FFlickerSpeed;
  AutoSize := True;
  Visible := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TFlickerControl.Destroy;
begin
  FTimer.Enabled := False;
  FImage.Free;
  FBitmap.Free;
  FTimer.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TFlickerControl.SetFlickered(const Value: Boolean);
begin
  if Value <> FFlickered then
  begin
    FFlickered := Value;
    FTimer.Enabled := FFlickered;
    Visible := FFlickered;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFlickerControl.SetFlickerSpeed(const Value: Cardinal);
begin
  if Value <> FFlickerSpeed then
  begin
    FFlickerSpeed := Value;
    FTimer.Interval := FFlickerSpeed;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFlickerControl.SetMode(const Value: Integer);
begin
  if ( FMode <> Value ) and
     ( Value in [MB_ICONWARNING] ) then
  begin
    FMode := Value;
    case FMode of
      MB_ICONWARNING: FResName := 'WARNING';
    end;
    LoadResourceBmp;  
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFlickerControl.LoadResourceBmp;
begin
  FBitmap.FreeImage;
  FBitmap.LoadFromResourceName( HInstance, FResName );
  FImage.Picture.Graphic := FBitmap;
end;

{ ---------------------------------------------------------------------------- }

procedure TFlickerControl.TimerExpired(Sender: TObject);
begin
  FImage.Visible := not FImage.Visible;
end;

{ ---------------------------------------------------------------------------- }

procedure TFlickerControl.SetParent(AParent: TWinControl);
begin
  inherited;
  if Assigned( AParent ) then
  begin
    Left := ( AParent.Width - Width ) div 2;
    Top := ( AParent.Height - Height ) div 2;
    BringToFront;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
