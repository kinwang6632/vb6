unit cbStyleModule;

interface

uses
  Windows, SysUtils, Classes, Controls, Graphics, ImgList, Forms,
{ App: }
  cbClientClass,
{ Developer Express: }
  dxSkinsCore, dxSkinsDefaultPainters, dxBarSkinConsts, dxRibbonSkins,
  cxContainer, cxGraphics, cxStyles, cxLookAndFeels, cxEdit,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkiniMaginary, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Silver;

type
  TStyleModule = class(TDataModule)
    SmallImages: TcxImageList;
    LargeImages: TcxImageList;
    TrayImageList: TImageList;
    LookAndFeelController: TcxLookAndFeelController;
    StyleRepository: TcxStyleRepository;
    cxStyleSkinBlueLargeGroup: TcxStyle;
    EditStyle: TcxEditStyleController;
    TrayAnimateImageList: TImageList;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FColorName: String;
    procedure ApplyFormColorSchema(const AColorName: String);
  public
    { Public declarations }
    procedure ChangeColorSchema(const AColorName: String);
    function GetUserWorkClassImageIndex: Integer;
    function GetUserStatusImageIndex(const AStatus: Integer): Integer;
    function GetUserDisplayButtonImageIndex(const ADisplayType: TUserListDisplayType): Integer;
    function GetMsgItemsStateIndex(const APriority: String): Integer;
    function GetMsgItemImageIndex(const AReaded: Boolean): Integer;
    property ColorName: String read FColorName;
  end;

var
  StyleModule: TStyleModule;

implementation

uses cbUtilis, cbMain, cbSendMsg, cbDataClientModule;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TStyleModule.DataModuleCreate(Sender: TObject);
begin
  LookAndFeelController.NativeStyle := False;
  FColorName := 'Blue';
end;

{ ---------------------------------------------------------------------------- }

procedure TStyleModule.DataModuleDestroy(Sender: TObject);
begin
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TStyleModule.ApplyFormColorSchema(const AColorName: String);
var
  aIndex: Integer;
  aForm: THrClientForm;
begin
  for aIndex := 0 to MsgFormList.Count - 1 do
  begin
    aForm := THrClientForm( MsgFormList[aIndex] );
    aForm.ChangeColorSchema( AColorName );
  end;
  {}
  for aIndex := 0 to AnnFormList.Count - 1 do
  begin
    aForm := THrClientForm( AnnFormList[aIndex] );
    aForm.ChangeColorSchema( AColorName );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TStyleModule.ChangeColorSchema(const AColorName: String);
begin
  FColorName := AColorName;
  if ( AColorName = 'Blue' ) then
  begin
    fmHrMain.RibbonBar.ColorSchemeName := AColorName;
    LookAndFeelController.SkinName := 'MoneyTwins';
    EditStyle.Style.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_NORMAL );
    EditStyle.StyleDisabled.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_DISABLED );
    EditStyle.StyleFocused.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_FOCUSED );
    EditStyle.StyleHot.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_HOT );
    fmHrMain.GridPanel.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      rfspRibbonForm );
    fmHrMain.GridPanel.Font.Color := clWindowText;
    fmHrMain.AnnGrid.LookAndFeel.SkinName := 'Office2007Blue';
    fmHrMain.SoPage.LookAndFeel.SkinName := 'MoneyTwins';
  end else
  if ( AColorName = 'Black' ) then
  begin
    fmHrMain.RibbonBar.ColorSchemeName := AColorName;
    LookAndFeelController.SkinName := 'Caramel';
    EditStyle.Style.Color := TdxBlackRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_NORMAL );
    EditStyle.StyleDisabled.Color := TdxBlackRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_DISABLED );
    EditStyle.StyleFocused.Color := TdxBlackRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_FOCUSED );
    EditStyle.StyleHot.Color := TdxBlackRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_HOT );
    fmHrMain.GridPanel.Color := TdxBlackRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      rfspRibbonForm );
    fmHrMain.GridPanel.Font.Color := clWhite;
    fmHrMain.AnnGrid.LookAndFeel.SkinName := 'Office2007Black';
    fmHrMain.SoPage.LookAndFeel.SkinName := 'Caramel';
  end else
  if ( AColorName = 'Silver' ) then
  begin
    fmHrMain.RibbonBar.ColorSchemeName := AColorName;
    LookAndFeelController.SkinName := 'Office2007Silver';
    fmHrMain.GridPanel.Color := TdxSilverRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      rfspRibbonForm );
    fmHrMain.GridPanel.Font.Color := clWindowText;
    fmHrMain.AnnGrid.LookAndFeel.SkinName := 'Office2007Silver';
    fmHrMain.SoPage.LookAndFeel.SkinName := 'Office2007Silver';
  end else
  if ( AColorName = EmptyStr ) then
  begin
    EditStyle.Style.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_NORMAL );
    EditStyle.StyleDisabled.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_DISABLED );
    EditStyle.StyleFocused.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_FOCUSED );
    EditStyle.StyleHot.Color := TdxBlueRibbonSkin( fmHrMain.RibbonBar.ColorScheme ).GetPartColor(
      DXBAR_EDIT_BACKGROUND, DXBAR_HOT );
  end;
  ApplyFormColorSchema( AColorName );
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetUserWorkClassImageIndex: Integer;
begin
  Result := -1;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetUserStatusImageIndex(const AStatus: Integer): Integer;
begin
  case  AStatus of
    0: Result := 49;
    1: Result := 50;
    2: Result := -1;
  else
    Result := -1;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetMsgItemsStateIndex(const APriority: String): Integer;
begin
  Result := 51;
  if ( APriority = '1' ) then Result := 47;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetMsgItemImageIndex(const AReaded: Boolean): Integer;
begin
  Result := 37;
  if ( AReaded ) then Result := 38
end;

{ ---------------------------------------------------------------------------- }

{ ---------------------------------------------------------------------------- }

function TStyleModule.GetUserDisplayButtonImageIndex(
  const ADisplayType: TUserListDisplayType): Integer;
begin
  case ADisplayType of
    gkWorkClass: Result := 45;
    gkStatus: Result := 46;
  else
    Result := -1;  
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
