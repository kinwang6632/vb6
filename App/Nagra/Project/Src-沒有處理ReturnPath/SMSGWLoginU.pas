unit SMSGWLoginU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Db, DBTables;

type
  TfrmSMSGW0000A = class(TForm)
    bvlLogin: TBevel;
    lblUserID: TLabel;
    lblPassword: TLabel;
    edtUserID: TEdit;
    edtPassword: TEdit;
    btnOK: TBitBtn;
    btnCancel: TBitBtn;
    lblTitle1: TLabel;
    lblTitle2: TLabel;
    qryUserInfo: TQuery;
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
  public
    loginFlag : integer;
    user_ID: string;
    user_Name: string;
    user_group: string;
    { Public declarations }
  end;

var
  frmSMSGW0000A: TfrmSMSGW0000A;

implementation

uses SMSGWLibU;

{$R *.DFM}

procedure TfrmSMSGW0000A.btnOKClick(Sender: TObject);
begin
  with qryUserInfo do
  begin
    close;
    ParamByName('USER_ID').asstring := trim(edtUserID.text);
    ParamByName('PASSWD').asstring := TSMSGWLib.Encrypt(trim(edtPassword.text));
//    showmessage(ParamByName('PASSWD').asstring);
    open;
    if RecordCount > 0 then
    begin
      loginFlag := 1;  //�b���αK�X�ŦX�A���\�n�J
      user_ID := fieldbyname('USER_ID').asstring;
      user_Name:= fieldbyname('USER_NAME').asstring;
      user_group:= fieldbyname('GROUP_ID').asstring;
    end
    else
      showmessage('�ڵ��n�J�A�Х��T�{�z���b���αK�X�I');
  end;
end;

procedure TfrmSMSGW0000A.btnCancelClick(Sender: TObject);
begin
  loginFlag := 3; //�ϥΪ̫��U������
end;

end.
