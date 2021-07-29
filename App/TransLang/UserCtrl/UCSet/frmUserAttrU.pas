unit frmUserAttrU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, utlUCU;

type
  TfrmUserAttr = class(TForm)
    Label1: TLabel;
    edtCode: TEdit;
    Bevel1: TBevel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label2: TLabel;
    edtDesc: TEdit;
    Label3: TLabel;
    edtPassword: TEdit;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  function EditUserAttr(User: TUser): TModalResult;

var
  frmUserAttr: TfrmUserAttr;

implementation

{$R *.DFM}

function EditUserAttr(User: TUser): TModalResult;
begin
  frmUserAttr := TfrmUserAttr.Create(Application);
  frmUserAttr.edtCode.Text := User.Code;
  frmUserAttr.edtDesc.Text := User.Desc;
  frmUserAttr.edtPassword.Text := User.Password;
  if User.Id = 0 then frmUserAttr.Caption := '新增使用者'
  else frmUserAttr.Caption := '使用者異動';
  Result := frmUserAttr.ShowModal;
  if Result = mrOk then
  begin
    User.Code := frmUserAttr.edtCode.Text;
    User.Desc := frmUserAttr.edtDesc.Text;
    User.Password := frmUserAttr.edtPassword.Text;
  end;
  frmUserAttr.Free;
end;

procedure TfrmUserAttr.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and ((edtCode.Text = '') or (edtDesc.Text = '') or
    (edtPassword.Text = '')) then CanClose := False;
end;

end.

