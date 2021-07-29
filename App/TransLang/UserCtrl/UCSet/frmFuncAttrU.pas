unit frmFuncAttrU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, utlUCU;

type
  TfrmFuncAttr = class(TForm)
    Bevel1: TBevel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    edtParam: TEdit;
    edtDesc: TEdit;
    edtExeName: TEdit;
    Label4: TLabel;
    edtIconName: TEdit;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  function EditFuncAttr(Func: TFunc): TModalResult;

var
  frmFuncAttr: TfrmFuncAttr;

implementation

{$R *.DFM}

function EditFuncAttr(Func: TFunc): TModalResult;
begin
  frmFuncAttr := TfrmFuncAttr.Create(Application);
  frmFuncAttr.edtDesc.Text := Func.Desc;
  frmFuncAttr.edtExeName.Text := Func.ExeName;
  frmFuncAttr.edtParam.Text := Func.Param;
  frmFuncAttr.edtIconName.Text := Func.IconName;
  if Func.Id = 0 then frmFuncAttr.Caption := '新增功能'
  else frmFuncAttr.Caption := '功能異動';
  Result := frmFuncAttr.ShowModal;
  if Result = mrOk then
  begin
    Func.Desc := frmFuncAttr.edtDesc.Text;
    Func.ExeName := frmFuncAttr.edtExeName.Text;
    Func.Param := frmFuncAttr.edtParam.Text;
    Func.IconName := frmFuncAttr.edtIconName.Text;
  end;
  frmFuncAttr.Free;
end;

procedure TfrmFuncAttr.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and ((edtDesc.Text = '') or
    (edtExeName.Text = '')) then CanClose := False;
end;

end.

