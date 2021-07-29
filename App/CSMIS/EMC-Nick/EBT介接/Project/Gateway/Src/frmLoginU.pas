unit frmLoginU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, DB, DBClient;

type
  TfrmLogin = class(TForm)
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    edtUserID: TEdit;
    edtUserPasswd: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    lblTitle2: TLabel;
    lblTitle1: TLabel;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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

uses dtmMainU, ConstObjU;

{$R *.dfm}

procedure TfrmLogin.BitBtn1Click(Sender: TObject);
var

    sL_ID, sL_Password : String;
begin
    sL_ID := edtUserID.text;
    sL_Password := edtUserPasswd.text;
    if (sL_ID=SUPER_USER_ID) and (sL_Password=SUPER_USER_PASSWD) then
    begin
      sG_UserID := SUPER_USER_ID;
      sG_UserName:= SUPER_USER_PASSWD;
      ModalResult := mrOK;
    end
    else
    begin
      with dtmMain.cdsUser do
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

procedure TfrmLogin.FormCreate(Sender: TObject);
begin
    nG_LoginCount := 0;
end;

end.
