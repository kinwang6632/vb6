unit frmCheckUserU;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, utlUCU;

type
  TfrmCheckUser = class(TForm)
    btnOk: TButton;
    btnCancel: TButton;
    Bevel1: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    edtID: TEdit;
    edtPassword: TEdit;
    procedure btnOkClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    FUserMgr: TUserMgr;
  public
    { Public declarations }
  end;

var
  frmCheckUser: TfrmCheckUser;

implementation

uses Dialogs, utlUserInoU;

{$R *.DFM}


procedure TfrmCheckUser.btnOkClick(Sender: TObject);
var
  User: TUser;
begin
  User := FUserMgr.HasUser(edtID.Text, edtPassword.Text);
  if Assigned(User) then
  begin
    SaveUserInfo(User.Code, User.Desc, User.Password, '');
    Close;
  end
  else MessageDlg('使用者代號或密碼錯誤', mtError, [mbOk], 0);
end;

procedure TfrmCheckUser.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmCheckUser.FormCreate(Sender: TObject);
begin
  SaveUserInfo('', '', '', '');
  FUserMgr := TUserMgr.Create;
end;

procedure TfrmCheckUser.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  FUserMgr.Free;
end;

end.
