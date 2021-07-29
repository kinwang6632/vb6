unit frmInvA04_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMemo;

type
  TfrmInvA04_2 = class(TForm)
    txtErrMsg: TcxMemo;
    procedure FormShow(Sender: TObject);

  private
    { Private declarations }
    FErrMsg: String;
  public
    { Public declarations }
  property ImportErr: String read FErrMsg write FErrMsg;
  end;

var
  frmInvA04_2: TfrmInvA04_2;

implementation

{$R *.dfm}



procedure TfrmInvA04_2.FormShow(Sender: TObject);
begin
  txtErrMsg.Text := FErrMsg;
end;

end.
