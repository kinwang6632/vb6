unit frmInvE04U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, cxLookAndFeelPainters, cxControls, cxContainer, cxEdit,
  cxTextEdit, StdCtrls, cxButtons, ExtCtrls, Encryption_TLB, dxSkinsCore,
  dxSkinsDefaultPainters;

type
  TfrmInvE04 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    btnOk: TcxButton;
    btnClose: TcxButton;
    txtPassword1: TcxTextEdit;
    txtPassword2: TcxTextEdit;
    Label2: TLabel;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
  private
    { Private declarations }
    FPassObj: _Password;
    function ChangePassword(var aErrMsg: String): Boolean;
  public
    { Public declarations }
  end;

var
  frmInvE04: TfrmInvE04;

implementation

uses cbUtilis, frmMainU, dtmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE04.FormCreate(Sender: TObject);
begin
  FPassObj := CoPassword.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE04.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'E04', '�ܧ�K�X' )
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE04.FormDestroy(Sender: TObject);
begin
  FPassObj := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE04.btnOkClick(Sender: TObject);
var
  aMsg: String;
begin
  if Trim( txtPassword1.Text ) = EmptyStr then
  begin
    WarningMsg( '�п�J�s�K�X�C' );
    if txtPassword1.CanFocusEx then txtPassword1.SetFocus;
    Exit;
  end;
  {}
  if Trim( txtPassword2.Text ) = EmptyStr then
  begin
    WarningMsg( '�п�J�T�{�s���K�X�C' );
    if txtPassword2.CanFocusEx then txtPassword2.SetFocus;
    Exit;
  end;
  if ( Trim( txtPassword1.Text ) <> Trim( txtPassword2.Text ) ) then
  begin
    WarningMsg( '��J���s�K�X�����T, �нT�{�C' );
    if txtPassword2.CanFocusEx then txtPassword2.SetFocus;
    Exit;
  end;
  {}
  Screen.Cursor := crSQLWait;
  try
    if not ChangePassword( aMsg ) then
    begin
      WarningMsg( Format( '�ܧ�K�X���~, �T��:%s�C', [aMsg] ) );
      if txtPassword1.CanFocusEx then txtPassword1.SetFocus;
      Exit;
    end;
    InfoMsg( '�K�X�ܧ󧹦��C' );
    Self.Close;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvE04.ChangePassword(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  dtmMain.adoComm2.SQL.Clear;
  dtmMain.adoComm2.Parameters.Clear;
  try
    dtmMain.adoComm2.SQL.Text := Format(
      ' update inv004 set password = :1 ' +
      '  where identifyid1 = ''%s''     ' +
      '    and identifyid2 = ''%s''     ' +
      '    and userid = ''%s''          ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getLoginUser] );
    dtmMain.adoComm2.Parameters.ParamByName( '1' ).Value :=
      FPassObj.Encrypt( txtPassword1.Text, PASSKEY );
    try
      dtmMain.adoComm2.ExecSQL;
    except
      on E: Exception do aErrMsg := E.Message;
    end;
  finally
    dtmMain.adoComm2.SQL.Clear;
    dtmMain.adoComm2.Parameters.Clear;
  end;
  Result := ( aErrMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

end.
