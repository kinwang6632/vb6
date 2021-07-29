unit cbBusy;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, GIFImg,
{ Developer Express: }
  dxSkinsCore,
  cxContainer, cxEdit, cxControls,
{ Developer Express Skin: }
  dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinGlassOceans, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinSilver, dxSkinStardust,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinXmas2008Blue;

type
  TfmBusy = class(TForm)
    imgBusy: TImage;
    DisplayText: TLabel;
    BorderShape: TShape;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    FBackGroundColor: TColor;
    FBusyText: String;
    procedure SetBackGroundColor(const AColor: TColor);
    procedure SetBusyText(const AText: String);
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowBusy; overload;
    class procedure ShowBusy(const AText: String); overload;
    class procedure CloseBusy;
    property BusyText: String read FBusyText write SetBusyText;
    property BackGroundColor: TColor read FBackGroundColor write SetBackGroundColor;
  end;


implementation

var
  fmBusy: TfmBusy;
  AWindowList: Pointer;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmBusy.FormCreate(Sender: TObject);
begin
  imgBusy.AutoSize := True;
  TGIFImage( imgBusy.Picture.Graphic ).AnimationSpeed := 500;
  TGIFImage( imgBusy.Picture.Graphic ).Animate := True;
  DisplayText.Alignment := taCenter;
  Self.BorderStyle := bsNone;
  Self.Position := poMainFormCenter;
  Self.Color := clWhite;
  FBackGroundColor := Self.Color;
  FBusyText := DisplayText.Caption;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmBusy.FormDestroy(Sender: TObject);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmBusy.SetBackGroundColor(const AColor: TColor);
begin
  if ( AColor <> FBackGroundColor ) then
  begin
    FBackGroundColor := AColor;
    Self.Color := AColor;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmBusy.SetBusyText(const AText: String);
begin
  FBusyText := AText;
  fmBusy.DisplayText.Caption := FBusyText;
end;

{ ---------------------------------------------------------------------------- }

class procedure TfmBusy.ShowBusy;
begin
  TfmBusy.ShowBusy( EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

class procedure TfmBusy.ShowBusy(const AText: String);
begin
  if not Assigned( fmBusy ) then
    fmBusy := TfmBusy.Create( nil );
  if ( AText <> EmptyStr ) then
    fmBusy.SetBusyText( AText );
  if not Assigned( AWindowList ) then
    AWindowList := DisableTaskWindows( 0 );
  Screen.Cursor := crHourGlass;
  fmBusy.Show;
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

class procedure TfmBusy.CloseBusy;
begin
  EnableTaskWindows( AWindowList );
  Screen.Cursor := crDefault;
  FreeAndNil( fmBusy );
  AWindowList := nil;
end;

{ ---------------------------------------------------------------------------- }

end.
