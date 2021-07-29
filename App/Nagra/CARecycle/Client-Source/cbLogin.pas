unit cbLogin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, cxGraphics, cxTextEdit, cxControls, cxContainer,
  cxEdit, cxMaskEdit, cxDropDownEdit, StdCtrls, cbStyleModule, Menus,
  cxLookAndFeelPainters, cxButtons;

type
  TfmLogin = class(TForm)
    Panel1: TPanel;
    Bevel1: TBevel;
    ScrollBox1: TScrollBox;
    Label1: TLabel;
    cmbSo: TcxComboBox;
    Label2: TLabel;
    txtUserId: TcxTextEdit;
    Label3: TLabel;
    txtPassword: TcxTextEdit;
    btnOk: TcxButton;
    btnCancel: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmLogin: TfmLogin;

implementation

uses cbUtilis, cbAppClass, cbSoModule;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmLogin.FormCreate(Sender: TObject);
var
  aIndex: Integer;
  aText: String;
begin
  cmbSo.Properties.Items.Clear;
  for aIndex := 0 to SoDataModule.SoList.Count - 1 do
  begin
    if ( TSo( SoDataModule.SoList[aIndex] ).Active ) then
    begin
      aText := Format( '%s-%s', [
        TSo( SoDataModule.SoList[aIndex] ).CompCode,
        TSo( SoDataModule.SoList[aIndex] ).CompName] );
      cmbSo.Properties.Items.Add( aText );  
    end;
  end;
  if ( cmbSo.Properties.Items.Count > 0 ) then
    cmbSo.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmLogin.FormShow(Sender: TObject);
begin
end;

{ ---------------------------------------------------------------------------- }

procedure TfmLogin.FormDestroy(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmLogin.btnOkClick(Sender: TObject);
var
  aCompCode, aCompName, aMsg: String;
begin
  if ( cmbSo.ItemIndex < 0 ) then
  begin
    WarningMsg( '請選擇系統台。' );
    if ( cmbSo.CanFocus ) then cmbSo.SetFocus;
    Exit;
  end;
  {}
  if ( Trim( txtUserId.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入登入帳號。' );
    if ( txtUserId.CanFocus ) then txtUserId.SetFocus;
    Exit;
  end;
  if ( Trim( txtPassword.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入登入密碼。' );
    if ( txtPassword.CanFocus ) then txtPassword.SetFocus;
    Exit;
  end;
  Screen.Cursor := crSQLWait;
  try
    { 系統台 }
    aCompCode := Copy( cmbSo.Properties.Items[cmbSo.ItemIndex],
      1, Pos( '-', cmbSo.Properties.Items[cmbSo.ItemIndex] ) - 1 );
    aCompName := Copy( cmbSo.Properties.Items[cmbSo.ItemIndex],
      Pos( '-', cmbSo.Properties.Items[cmbSo.ItemIndex] ) + 1,
       Length( cmbSo.Properties.Items[cmbSo.ItemIndex] ) );
    {}
    if not SoDataModule.DbLogin( aCompCode, aMsg ) then
    begin
      WarningMsg( Format( '無法連接系統台, 訊息:%s。',  [aMsg] ) );
      Exit;
    end;
    { COM 區 }
    if not SoDataModule.DbLogin( '-1', aMsg ) then
    begin
      WarningMsg( Format( '無法連接資料回收作業區(COM), 訊息:%s。',  [aMsg] ) );
      Exit;
    end;
    { 檢查權限 }
    SoDataModule.User.UserId := txtUserId.Text;
    SoDataModule.User.Password := txtPassword.Text;
    SoDataModule.User.CompCode := aCompCode;
    SoDataModule.User.CompName := aCompName;
    if not SoDataModule.UserAuthorize( aMsg ) then
    begin
      WarningMsg( aMsg );
      Exit;
    end;
    Self.ModalResult := mrOk;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
