unit frmDataFileAttrU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, utlUCU;

type
  TfrmDataFileAttr = class(TForm)
    Bevel1: TBevel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    edtFileName: TEdit;
    Label5: TLabel;
    edtDesc: TEdit;
    ckbAutoReg: TCheckBox;
    Label4: TLabel;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  function EditDataFileAttr(DataFile: TDataFile): TModalResult;

var
  frmDataFileAttr: TfrmDataFileAttr;

implementation

{$R *.DFM}

function EditDataFileAttr(DataFile: TDataFile): TModalResult;
begin
  frmDataFileAttr := TfrmDataFileAttr.Create(Application);
  frmDataFileAttr.edtFileName.Text := DataFile.FileName;
  frmDataFileAttr.edtDesc.Text := DataFile.Desc;
  frmDataFileAttr.ckbAutoReg.Checked := DataFile.AutoReg;
  if DataFile.Id = 0 then frmDataFileAttr.Caption := '新增資料檔'
  else frmDataFileAttr.Caption := '資料檔異動';
  Result := frmDataFileAttr.ShowModal;
  if Result = mrOk then
  begin
    DataFile.FileName := frmDataFileAttr.edtFileName.Text;
    DataFile.Desc := frmDataFileAttr.edtDesc.Text;
    DataFile.AutoReg := frmDataFileAttr.ckbAutoReg.Checked;
  end;
  frmDataFileAttr.Free;
end;

procedure TfrmDataFileAttr.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and ((edtDesc.Text = '') or
    (edtFileName.Text = '')) then CanClose := False;
end;

end.

