unit cbStyleModule;

interface

uses
  SysUtils, Classes, ImgList, Controls, cxContainer, cxEdit, cxStyles,
  cxClasses, cxTL, cxLookAndFeels;

type
  TStyleModule = class(TDataModule)
    BarImageList: TImageList;
    ActionImageList: TImageList;
    MsgImageList: TImageList;
    PageImageList: TImageList;
    NavBarImageList: TImageList;
    OptionImageList: TImageList;
    cxStyleRepository: TcxStyleRepository;
    cxEditStyle: TcxEditStyleController;
    CommonStyle: TcxEditStyleController;
    MsgStyle: TcxEditStyleController;
    cxLookAndFeelController: TcxLookAndFeelController;
    cxGroupBoxStyle: TcxEditStyleController;

  private
    { Private declarations }
  public
    { Public declarations }
  end;


var StyleModule: TStyleModule;

implementation

{$R *.dfm}

end.
