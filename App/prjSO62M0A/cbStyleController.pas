unit cbStyleController;

interface

uses
  SysUtils, Classes, cxContainer, cxEdit, cxStyles, cxHint, 
  cxGridTableView;

type
  TStyleController = class(TDataModule)
    EditorStyle: TcxEditStyleController;
    CheckBoxStyle: TcxEditStyleController;
    GridStyleRepository: TcxStyleRepository;
    HeaderStyle: TcxStyle;
    RowActiveStyle: TcxStyle;
    GridBackgroundStyle: TcxStyle;
    HintStyleController: TcxHintStyleController;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    GroupStyle: TcxStyle;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  StyleController: TStyleController;

implementation

{$R *.dfm}

end.
