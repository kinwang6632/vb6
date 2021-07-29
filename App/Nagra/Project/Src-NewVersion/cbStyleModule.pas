unit cbStyleModule;

interface

uses
  SysUtils, Classes, ImgList, Controls, cxContainer, cxEdit, cxStyles;

type
  TStyleModule = class(TDataModule)
    BarImageList: TImageList;
    TreeImageList: TImageList;
    CommonStyle: TcxEditStyleController;
    MsgStyle: TcxEditStyleController;
    DockPanelImageList: TImageList;
    MsgImageList: TImageList;
    cxStyleRepository: TcxStyleRepository;
    cxStyle1: TcxStyle;
    TitleCaptionStyle: TcxEditStyleController;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  StyleModule: TStyleModule;

implementation

{$R *.dfm}

end.
