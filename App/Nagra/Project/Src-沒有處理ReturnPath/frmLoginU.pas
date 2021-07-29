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

uses ConstObjU, dtmMainU, frmMainU;

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
          
            MessageDlg(frmMain.getLanguageSettingInfo('frmMain_Msg_49_Content') + IntToStr(PASSWORD_CHECKING_TIMES) + frmMain.getLanguageSettingInfo('frmMain_Msg_50_Content'),mtError,[mbOK],0);
            ModalResult := mrCancel;
          end
          else
          begin
            MessageDlg(frmMain.getLanguageSettingInfo('frmMain_Msg_51_Content'),mtError,[mbOK],0);
            edtUserID.SetFocus;
          end;
        end;
      end;
    end;
end;

procedure TfrmLogin.FormActivate(Sender: TObject);
begin

    if (dtmMain.cdsUser.State = dsInactive) then
      dtmMain.cdsUser.CreateDataSet;
    if (FileExists(USER_INFO_FILENAME)) then
      dtmMain.cdsUser.LoadFromFile(USER_INFO_FILENAME);

end;

end.


