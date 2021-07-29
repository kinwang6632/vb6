unit frmLoginU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, StdCtrls, Buttons, ExtCtrls, Variants;

type
  TfrmLogin = class(TForm)
    bvlLogin: TBevel;
    lblUserID: TLabel;
    lblPassword: TLabel;
    lblTitle1: TLabel;
    lblTitle2: TLabel;
    edtUserID: TEdit;
    edtPassword: TEdit;
    btnOK: TBitBtn;
    btnCancel: TBitBtn;
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    procedure btnOKClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
    nG_LoginCount : Integer;
  public
    { Public declarations }
    sG_UserID: string;
    sG_UserName: string;
    
  end;

var
  frmLogin: TfrmLogin;

implementation

uses ConstObjU;

{$R *.DFM}

procedure TfrmLogin.btnOKClick(Sender: TObject);
var
    sL_ID, sL_Password : String;
begin
    sL_ID := edtUserID.text;
    sL_Password := edtPassword.text;
    if (sL_ID=SUPER_USER_ID) and (sL_Password=SUPER_USER_PASSWD) then
    begin
      sG_UserID := SUPER_USER_ID;
      sG_UserName:= SUPER_USER_PASSWD;
      ModalResult := mrOK;
    end
    else
    begin
      with cdsUser do
      begin
        if (Locate('sID;sPassword', VarArrayOf([sL_ID, sL_Password]), [loPartialKey])) then
        begin
          sG_UserID := fieldbyname('sID').asstring;
          sG_UserName:= fieldbyname('sName').asstring;
          ModalResult := mrOK;
        end
        else
        begin
          INC(nG_LoginCount);
          if (nG_LoginCount=PASSWORD_CHECKING_TIMES) then
          begin
            MessageDlg('密碼錯誤超過 ' + IntToStr(PASSWORD_CHECKING_TIMES) + ' 次！程式即將結束.',mtError,[mbOK],0);
            ModalResult := mrCancel;
          end
          else
          begin
            MessageDlg('拒絕登入，請先確認您的帳號及密碼！',mtError,[mbOK],0);
            edtUserID.SetFocus;
          end;
        end;
      end;
    end;
end;

procedure TfrmLogin.FormActivate(Sender: TObject);
begin

    if (cdsUser.State = dsInactive) then
      cdsUser.CreateDataSet;
    if (FileExists(USER_INFO_FILENAME)) then
      cdsUser.LoadFromFile(USER_INFO_FILENAME);

end;

end.


