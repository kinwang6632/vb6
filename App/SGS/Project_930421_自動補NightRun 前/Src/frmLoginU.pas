unit frmLoginU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls;

type
  TfrmLogin = class(TForm)
    bvlLogin: TBevel;
    lblUserID: TLabel;
    lblPassword: TLabel;
    lblTitle: TLabel;
    lblTitle21: TLabel;
    edtUserID: TEdit;
    edtPasswd: TEdit;
    btnOK: TBitBtn;
    btnCancel: TBitBtn;
    lblTitle1: TLabel;
    lblTitle2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    function IsDataOK : boolean;
    procedure setLanguageString;
  public
    { Public declarations }
    sG_UserID : String;
    function checkPasswd(sI_UserID,sI_Passwd : String; var sI_UserName:STring) : Boolean;
  end;

var
  frmLogin: TfrmLogin;

implementation

uses dtmMainU, frmMainU;

{$R *.dfm}

function TfrmLogin.checkPasswd(sI_UserID, sI_Passwd: String;
  var sI_UserName: STring): Boolean;
var
    sL_ExePath : String;
begin
    sG_UserID := sI_UserID;

    if dtmMain.cdsUser.Locate('SID;SPASSWORD',VarArrayOf([sI_UserID, sI_Passwd]) , []) then
    begin
      sI_UserName:= dtmMain.cdsUser.fieldbyname('sName').asstring;    
      Result := true;
    end
    else
      Result := false;
end;

procedure TfrmLogin.FormCreate(Sender: TObject);
begin
    dtmMain.activeCDS;
    dtmMain.getAllUserID;

    setLanguageString;
end;

function TfrmLogin.IsDataOK: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    if edtUserID.Text = '' then
    begin
      bL_Result := false;
    end;
    if edtPasswd.Text = '' then
    begin
      bL_Result := false;
    end;
    result := bL_Result;
end;

procedure TfrmLogin.btnOKClick(Sender: TObject);
var
  sL_UserID,sL_Passwd, sL_UserName : String;
  bL_IsPasswordOK : Boolean;
begin
    if IsDataOK then
    begin
      sL_UserID := edtUserID.Text;
      sL_Passwd := edtPasswd.Text;

      bL_IsPasswordOK := checkPasswd(sL_UserID,sL_Passwd, sL_UserName);

      if bL_IsPasswordOK then
      begin
        //frmLogin.Visible := false;
        //frmMain := TfrmMain.Create(Application);
        frmMain.sG_OperatorName :=  sL_UserName;
        ModalResult := mrOK;
        //frmMain.ShowModal;
        //frmMain.Free;
      end
      else
      begin
        edtUserID.SetFocus;
        edtPasswd.Text := '';
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLogin_Msg_1_Content'),mtError, [mbOK],0);
        //ModalResult := mrCancel;
        exit;
      end;
    end
    else
    begin
      edtUserID.SetFocus;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmLogin_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

end;

procedure TfrmLogin.btnCancelClick(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmLogin.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmLogin_Caption');

    lblUserID.Caption := dtmMain.getLanguageSettingInfo('frmLogin_LblUserID_Caption');
    lblPassword.Caption := dtmMain.getLanguageSettingInfo('frmLogin_LblPassword_Caption');

    btnOK.Caption := dtmMain.getLanguageSettingInfo('frmLogin_BtnOK_Caption');
    btnCancel.Caption := dtmMain.getLanguageSettingInfo('frmLogin_BtnCancel_Caption');

    lblTitle1.Caption := dtmMain.getLanguageSettingInfo('frmLogin_lblTitle1_Caption');
    lblTitle2.Caption := dtmMain.getLanguageSettingInfo('frmLogin_lblTitle2_Caption');
end;

end.


