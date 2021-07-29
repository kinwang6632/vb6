unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses frmSO18B5U;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
    L_Frm : TfrmSO18B5;
begin
    L_Frm := TfrmSO18B5.Create(Application);
    L_Frm.sG_CompCode := '9';// 暫時的
    L_Frm.sG_User := 'abc';  // 暫時的
    L_Frm.ShowModal;
    L_Frm.Free;
end;

end.
