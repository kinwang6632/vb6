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
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    bG_Login : boolean;
    G_UserNameStrList : TStringList;
    G_UserPasswdStrList : TStringList;        
    nG_LoginCount : Integer;
  public
    { Public declarations }
    sG_UserID: string;
    sG_UserName: string;
//    constructor create(var sI_LoginID, sI_LoginName:String);    
  end;

var
  frmLogin: TfrmLogin;

implementation

uses dtmMainU;


{$R *.DFM}

procedure TfrmLogin.btnOKClick(Sender: TObject);
var
    nL_Ndx : Integer;
    bL_Login : boolean;
    sL_UserID, sL_Password : String;
begin
    sL_UserID := edtUserID.text;
    sL_Password := edtPassword.text;
    bL_Login := false;

    nL_Ndx := G_UserNameStrList.IndexOf(sL_UserID);
    if nL_Ndx<>-1 then
    begin
      if sL_Password=G_UserPasswdStrList.Strings[nL_Ndx] then
        bL_Login := true
      else
        bL_Login := false;
    end
    else
      bL_Login := false;

    if bL_Login then
    begin
      dtmMain.sG_UserID := sL_UserID;
      ModalResult := mrOK;
    end
    else
    begin
      edtUserID.SetFocus;
      MessageDlg('請輸入正確的帳號與密碼!', mtError, [mbOK],0);
    end;


end;
{
constructor TfrmLogin.create(var sI_LoginID, sI_LoginName: String);
begin
    bG_Login := false;
    G_UserNameStrList := TStringList.Create;
    G_UserPasswdStrList := TStringList.Create;

    G_UserNameStrList.Add('a1');
    G_UserPasswdStrList.Add('1688');

    G_UserNameStrList.Add('b1');
    G_UserPasswdStrList.Add('9988');

    G_UserNameStrList.Add('c1');
    G_UserPasswdStrList.Add('3702');

    G_UserNameStrList.Add('d1');
    G_UserPasswdStrList.Add('2581');

    inherited Create(Application);
end;
}
procedure TfrmLogin.FormCreate(Sender: TObject);
begin
    bG_Login := false;
    G_UserNameStrList := TStringList.Create;
    G_UserPasswdStrList := TStringList.Create;

    G_UserNameStrList.Add('a1');
    G_UserPasswdStrList.Add('1688');

    G_UserNameStrList.Add('b1');
    G_UserPasswdStrList.Add('9988');

    G_UserNameStrList.Add('c1');
    G_UserPasswdStrList.Add('3702');

    G_UserNameStrList.Add('d1');
    G_UserPasswdStrList.Add('2581');


end;

end.



