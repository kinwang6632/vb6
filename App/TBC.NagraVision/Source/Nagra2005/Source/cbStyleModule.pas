unit cbStyleModule;

interface

uses
  SysUtils, Classes, ImgList, Controls,
  { Developer Express Common }
  cxContainer, cxClasses, cxEdit, cxStyles,
  { Develoepr Express TreeList }
  cxTL;

type
  TStyleModule = class(TDataModule)
    BarImageList: TImageList;
    SoTreeImageList: TImageList;
    CmdTreeImageList: TImageList;
    CommonStyle: TcxEditStyleController;
    MsgStyle: TcxEditStyleController;
    DockPanelImageList: TImageList;
    cxStyleRepository: TcxStyleRepository;
    TreeListStyleSheetStormVGA: TcxTreeListStyleSheet;
    TreeListStyleSheetConsoleBlack: TcxTreeListStyleSheet;
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
    ConfigImageList: TImageList;
    cxEditStyle: TcxEditStyleController;
    cxStyle16: TcxStyle;
    MsgImageList: TImageList;
  private
    { Private declarations }
  public
    { Public declarations }
  end;


var StyleModule: TStyleModule;

implementation

{$R *.dfm}

end.
