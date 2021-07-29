unit cbStyleModule;

interface

uses
  Windows, SysUtils, Classes, ImgList, Controls,
{ App }
  cbSo, cbSrvClass,
{ Developer Express }
  dxSkinsCore, dxSkinsDefaultPainters,  cxLookAndFeels, cxContainer, cxEdit,
  cxGraphics, cxClasses, cxStyles, cxGridTableView, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinGlassOceans, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinSilver, dxSkinStardust,
  dxSkinValentine, dxSkinXmas2008Blue, dxSkinSummer2008;

type
  TStyleModule = class(TDataModule)
    DefaultStyle: TcxLookAndFeelController;
    ToolBarImgList: TImageList;
    SoImgList: TImageList;
    ImgList16: TImageList;
    MsgImgList: TImageList;
    ImgList32: TImageList;
    MsgStyle: TcxEditStyleController;
    LabelFontStyle: TcxEditStyleController;
    EditorStyle: TcxEditStyleController;
    StyleRepository: TcxStyleRepository;
    Sunny: TcxStyle;
    Dark: TcxStyle;
    Golden: TcxStyle;
    Summer: TcxStyle;
    Autumn: TcxStyle;
    Bright: TcxStyle;
    Cold: TcxStyle;
    Spring: TcxStyle;
    Light: TcxStyle;
    Winter: TcxStyle;
    Depth: TcxStyle;
    UserStyleSheet: TcxGridTableViewStyleSheet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function MsgImageIndex(const AFlag: Integer): Integer;
    function SoImageIndex(const ASo: TAppSo): Integer;
    function SoStateIndex(const ASo: TAppSo): Integer;
    function SocketConnectImageIndex: Integer;
    function SocketDisconnectImageIndex: Integer;
    function SocketWarningImageIndex: Integer;
    function SocketNullImageIndex: Integer;
    function ConfigFileImageIndex: Integer;
    function UserStateImageIndex(const AState: Integer): Integer;
    function ConfigHeaderImage(const APageIndex: Integer): Integer;
    function GroupSetHeaderImage: Integer;
    function GroupImageIndex: Integer;
  end;

var
  StyleModule: TStyleModule;

implementation

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

{ TStyleController }

procedure TStyleModule.DataModuleCreate(Sender: TObject);
begin
  MsgStyle.Style.Color := Dark.Color; 
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.MsgImageIndex(const AFlag: Integer): Integer;
begin
  case AFlag of
    MB_ICONINFORMATION: Result := 0;
    MB_OK: Result := 1;
    MB_ICONERROR: Result := 2;
    MB_ICONWARNING: Result := 3;
  else
    Result := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SoImageIndex(const ASo: TAppSo): Integer;
begin
  Result := 1;
  if ( ASo.Selected ) or ( ASo.SynData ) then
  begin
    case ASo.DbState of
      dbClose: Result := 2;
      dbOpen, dbActive: Result := 3;
      dbError: Result := 4;
      dbWarning: Result := 5; 
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SoStateIndex(const ASo: TAppSo): Integer;
begin
  Result := 6;
  if ( ASo.Selected ) or ( ASo.SynData ) then
  begin
    case ASo.DbState of
      dbClose: Result := 6;
      dbOpen, dbError, dbWarning: Result := 7;
      dbActive: Result := 8;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketConnectImageIndex: Integer;
begin
  Result := 18;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketDisconnectImageIndex: Integer;
begin
  Result := 19;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketWarningImageIndex: Integer;
begin
  Result := 25;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.SocketNullImageIndex: Integer;
begin
  Result := 27;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.ConfigFileImageIndex: Integer;
begin
  Result := 7;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.UserStateImageIndex(const AState: Integer): Integer;
begin
  case AState of
    1,2: Result := 4;
  else
    Result := 19;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.ConfigHeaderImage(const APageIndex: Integer): Integer;
begin
  case APageIndex of
    0: Result := 8;
    1: Result := 9;
    2: Result := 10;
  else
    Result := 0;    
  end;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GroupSetHeaderImage: Integer;
begin
  Result := 11;
end;

{ ---------------------------------------------------------------------------- }

function TStyleModule.GroupImageIndex: Integer;
begin
  Result := 20;
end;

{ ---------------------------------------------------------------------------- }

end.
