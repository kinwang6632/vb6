unit frmSO8F10_4U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls;

type
  TfrmSO8F10_4 = class(TForm)
    Panel1: TPanel;
    edtUserID: TEdit;
    lblUserID: TLabel;
    lblPassword: TLabel;
    edtPassword: TEdit;
    btnCancel: TBitBtn;
    btnOK: TBitBtn;
    procedure btnOKClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    G_UserNameStrList : TstringList;
    G_UserPasswdStrList : TstringList;
  public
    { Public declarations }
    bG_IsSuperUser : Boolean;
  end;

var
  frmSO8F10_4: TfrmSO8F10_4;

implementation

uses frmSO8F10_1U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8F10_4.btnOKClick(Sender: TObject);
var
    sL_UserID,sL_Password : String;
    bL_Login : Boolean;
    nL_Ndx : Integer;
begin
    bG_IsSuperUser := false;
    sL_UserID := Trim(edtUserID.text);
    sL_Password := Trim(edtPassword.text);
    bL_Login := false;

    nL_Ndx := G_UserNameStrList.IndexOf(sL_UserID);
    if nL_Ndx<>-1 then
    begin
      if sL_Password=G_UserPasswdStrList.Strings[nL_Ndx] then
      begin
        bL_Login := true;
        bG_IsSuperUser := true;
      end
      else
      begin
        bL_Login := false;
        bG_IsSuperUser := false;
      end;
    end
    else
    begin
      bL_Login := false;
    end;

    if bL_Login then
    begin
      Close;
    end
    else
    begin
      MessageDlg('密碼錯誤請重新輸入!',mtError, [mbOK],0);
      edtPassword.Text := '';
      edtPassword.SetFocus;
    end;

end;

procedure TfrmSO8F10_4.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    G_UserNameStrList.Free;
    G_UserPasswdStrList.Free;
    
end;

procedure TfrmSO8F10_4.FormCreate(Sender: TObject);
begin
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

procedure TfrmSO8F10_4.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8F10','[二階代碼資料同步密碼確認]');

    bG_IsSuperUser := false;
end;

procedure TfrmSO8F10_4.btnCancelClick(Sender: TObject);
begin
    bG_IsSuperUser := false;
end;

end.
