unit frmUserGrpAttrU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, utlUCU;

type
  TfrmUserGrpAttr = class(TForm)
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

  function EditUserGrpAttr(UserGrp: TUserGrp): TModalResult;

var
  frmUserGrpAttr: TfrmUserGrpAttr;

implementation

{$R *.DFM}

function EditUserGrpAttr(UserGrp: TUserGrp): TModalResult;
begin
  frmUserGrpAttr := TfrmUserGrpAttr.Create(Application);
  frmUserGrpAttr.edtDesc.Text := UserGrp.Desc;
  if UserGrp.Id = 0 then frmUserGrpAttr.Caption := '新增群組'
  else frmUserGrpAttr.Caption := '群組異動';
  Result := frmUserGrpAttr.ShowModal;
  if Result = mrOk then UserGrp.Desc := frmUserGrpAttr.edtDesc.Text;
  frmUserGrpAttr.Free;
end;

procedure TfrmUserGrpAttr.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and (edtDesc.Text = '') then CanClose := False;
end;

end.
