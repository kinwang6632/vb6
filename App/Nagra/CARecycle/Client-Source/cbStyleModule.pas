unit cbStyleModule;

interface

uses
  SysUtils, Classes, ImgList, Controls, cxContainer, cxEdit, cxLookAndFeels;

type
  TStyleModule = class(TDataModule)
    cxLookAndFeelController: TcxLookAndFeelController;
    cxEditStyle: TcxEditStyleController;
    BarLargeImageList: TImageList;
    MsgImageList: TImageList;
    ActionImageList: TImageList;
    BarSmallImageList: TImageList;
    CheckImageList: TImageList;
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
