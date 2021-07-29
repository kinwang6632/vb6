unit frmAppAttrU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, utlUCU;

type
  TfrmAppAttr = class(TForm)
    Label1: TLabel;
    edtDesc: TEdit;
    Bevel1: TBevel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  function EditAppAttr(App: TApp): TModalResult;

var
  frmAppAttr: TfrmAppAttr;

implementation

{$R *.DFM}

function EditAppAttr(App: TApp): TModalResult;
begin
  frmAppAttr := TfrmAppAttr.Create(Application);
  frmAppAttr.edtDesc.Text := App.Desc;
  if App.Id = 0 then frmAppAttr.Caption := '新增系統'
  else frmAppAttr.Caption := '系統異動';
  Result := frmAppAttr.ShowModal;
  if Result = mrOk then App.Desc := frmAppAttr.edtDesc.Text;
  frmAppAttr.Free;
end;

procedure TfrmAppAttr.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and (edtDesc.Text = '') then CanClose := False;
end;

end.
