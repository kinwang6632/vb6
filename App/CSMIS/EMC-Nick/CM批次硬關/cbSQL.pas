unit cbSQL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, cxLookAndFeelPainters, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMemo, StdCtrls, cxButtons, ExtCtrls, SynEdit,
  SynEditHighlighter, SynHighlighterSQL;

type
  TfmSQL = class(TForm)
    Panel1: TPanel;
    btnOK: TcxButton;
    Panel2: TPanel;
    txtSQL: TSynEdit;
    SynSQLSyn1: TSynSQLSyn;
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSQL: TfmSQL;

implementation

{$R *.dfm}

procedure TfmSQL.btnOKClick(Sender: TObject);
begin
  Self.ModalResult := mrOK;
end;

end.
