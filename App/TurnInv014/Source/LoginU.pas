unit LoginU ;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DBCtrls ,IniFiles, DB,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, jpeg, cxGraphics, Menus, cxLookAndFeelPainters,
  cxButtons, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

type
  TfrmLogin = class(TForm)
    bvlLogin: TBevel;
    lblTitle: TLabel;
    lblTitle21: TLabel;
    lblTitle1: TLabel;
    cmbSMSDb: TcxImageComboBox;
    btnConnect: TcxButton;
    Label3: TLabel;
    Label2: TLabel;
    btnCancel: TcxButton;
    cmbCompany: TcxImageComboBox;
    Image1: TImage;
    btnOk: TcxButton;
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
    FDbSID : String;
    FCompany: String;
    FCompanyList: TList;
    procedure TransTmpIniFile(const aSourceFileName, aTempFileName: String);
    procedure BuildDataArear(const aFileName: String);
    procedure GetInvoiceCompany;
    procedure WarningMsg(const s:string);
    procedure SetConnStatus(AConnecting: Boolean);
  public
    { Public declarations }
    property CompanyList: TList read FCompanyList write FCompanyList;
  end;

var
  frmLogin: TfrmLogin;

implementation

uses  Encryption_TLB,dtmInv014AU,frmInvTurnInv014U ;

{$R *.dfm}

procedure TfrmLogin.btnCancelClick(Sender: TObject);
begin
  Application.Terminate;
end;



procedure TfrmLogin.BuildDataArear(const aFileName: String);
var
  aIndex: Integer;
  aSoInfo: PSMSDb;
  aItem: TcxImageComboBoxItem;

begin

  dtmInv014.GetSMSDb( aFileName );
  cmbSMSDb.Clear;
  for aIndex := 0 to dtmInv014.SMSDbList.Count -1 do
  begin
    aSoInfo := PSMSDB( dtmInv014.SMSDbList[aIndex] );
    aItem := cmbSMSDb.Properties.Items.Add;
    aItem.Value := IntToStr( aSoInfo.ConnSeq );
    aItem.Description := aSoInfo.Description;
    aItem.ImageIndex := 16;
  end;

end;

procedure TfrmLogin.FormCreate(Sender: TObject);
var
  aSourceFileName, aTempFileName: String;
  aTempDir: array [0..MAX_PATH-1] of Char;
begin
  FCompanyList := TList.Create;
  aSourceFileName := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_SYS_INFO_INI_FILE;
  ZeroMemory( @aTempDir, SizeOf( aTempDir ) );
  GetTempPath( SizeOf( aTempDir ), aTempDir );
  aTempFileName := IncludeTrailingPathDelimiter( String( aTempDir ) ) +
    TMP_INV_SYS_INFO_INI_FILE;
  if ( FileExists( aSourceFileName ) ) then
  begin
    try
      TransTmpIniFile( aSourceFileName, aTempFileName );
      BuildDataArear( aTempFileName );
    finally
      DeleteFile( aTempFileName );
    end;
  end;
end;
procedure TfrmLogin.TransTmpIniFile(const aSourceFileName,
  aTempFileName: String);
var
  aStrList, aTmpStrList : TStringList;
  aIndex: Integer;
  aIntf: _Password;
  aEncKey: WideString;
begin
  aEncKey := 'CS';
  aStrList := TStringList.Create;
  try
    aTmpStrList := TStringList.Create;
    try
      aIntf := CoPassword.Create;
      try
        aStrList.LoadFromFile( aSourceFileName );
        for AIndex := 0 to aStrList.Count - 1 do
        begin
          if ( Copy( aStrList.Strings[AIndex], 1, 2 ) <> '//' ) then
            aTmpStrList.Add( aIntf.Decrypt( aStrList.Strings[aIndex],
              aEncKey ) )
          else
            aTmpStrList.Add( aStrList.Strings[AIndex] );
        end;
        aTmpStrList.SaveToFile( aTempFileName );
      finally
        aIntf := nil;
      end;
    finally
      aTmpStrList.Free;
    end;
  finally
     aStrList.Free;
  end;

end;

procedure TfrmLogin.btnConnectClick(Sender: TObject);
var
  aConnected: Boolean;
begin
  if cmbSMSDb.ItemIndex < 0 then Exit;
  SetConnStatus( True );
  try
    aConnected := dtmInv014.ConnectToINV( PSMSDB( dtmInv014.SMSDbList.Items[ cmbSMSDb.ItemIndex] ));
    if aConnected then
    begin
      FDbSID := cmbSMSDb.Properties.Items[ cmbSMSDb.ItemIndex].Value;
      GetInvoiceCompany;
      cmbCompany.SetFocus;
    end;
  finally
    SetConnStatus( False );
  end;

 
end;

procedure TfrmLogin.GetInvoiceCompany;
var
  aIndex: Integer;
  aItem: TcxImageComboBoxItem;
begin
  dtmInv014.GetInvoiceCompany( FCompanyList );
  cmbCompany.Properties.Items.Clear;
  for aIndex := 0 to FCompanyList.Count -1 do
  begin
    aItem := cmbCompany.Properties.Items.Add;
    aItem.Value := PCompany( FCompanyList[aIndex] ).CompanyId;
    aItem.Description := PCompany( FCompanyList[aIndex] ).CompanyName;
    aItem.ImageIndex := 8;
    aItem.Tag := Integer( FCompanyList[aIndex] );
  end;
 
end;

procedure TfrmLogin.btnOkClick(Sender: TObject);
var
  aConnected: Boolean;
begin
  if ( FDbSID = EmptyStr ) then
    begin
      WarningMsg( '?????s???i?o???j???????C' );
      cmbSMSDb.SetFocus;
      Exit;
    end;
    if ( cmbCompany.Text = '' ) then
    begin
      WarningMsg( '???????i???q?O?j?C' );
      cmbCompany.SetFocus;
      Exit;
    end;
    aConnected := dtmInv014.ConnectToSO(
      dtmInv014.GetSoInfo((PCompany(FCompanyList.Items[cmbCompany.ItemIndex]).ConnSeq)));
    if not aConnected then
      Exit;
    FCompany := cmbCompany.Properties.Items[cmbCompany.ItemIndex].Description;
    Application.CreateForm(TfrmTurnInv014, frmTurnInv014);
    frmTurnInv014.ShowModal;

end;

procedure TfrmLogin.WarningMsg(const s:string);
begin
  MessageDlg(s,mtWarning,[mbOK],0);
end;

procedure TfrmLogin.SetConnStatus(AConnecting: Boolean);
begin
  if AConnecting then
  begin
    Screen.Cursor := crSQLWait;
    btnCancel.Enabled := False;
    btnOK.Enabled := False;
  end
  else begin
    Screen.Cursor := crDefault;
    btnCancel.Enabled := True;
    btnOK.Enabled := True;
  end;
end;

procedure TfrmLogin.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  FCompanyList.Free;
end;

end.
