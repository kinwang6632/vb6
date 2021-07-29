unit cbStyleModule;

interface

uses
  SysUtils, Classes, cxClasses, cxStyles, cxTL, cxContainer, cxEdit,
  ImgList, Controls;

type
  TStyleModule = class(TDataModule)
    BarImageList: TImageList;
    SoTreeImageList: TImageList;
    CommonStyle: TcxEditStyleController;
    MsgStyle: TcxEditStyleController;
    DockPanelImageList: TImageList;
    MsgImageList: TImageList;
    cxStyleRepository: TcxStyleRepository;
    cxStyle1: TcxStyle;
    cxStyle2: TcxStyle;
    cxStyle3: TcxStyle;
    cxStyle4: TcxStyle;
    cxStyle5: TcxStyle;
    cxStyle6: TcxStyle;
    cxStyle7: TcxStyle;
    cxStyle8: TcxStyle;
    cxStyle9: TcxStyle;
    cxStyle10: TcxStyle;
    cxStyle11: TcxStyle;
    cxStyle12: TcxStyle;
    cxStyle13: TcxStyle;
    cxStyle14: TcxStyle;
    cxStyle15: TcxStyle;
    cxStyle16: TcxStyle;
    TreeListStyleSheetStormVGA: TcxTreeListStyleSheet;
    TreeListStyleSheetConsoleBlack: TcxTreeListStyleSheet;
    CmdTreeImageList: TImageList;
    ConfigImageList: TImageList;
    cxEditStyle: TcxEditStyleController;
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
