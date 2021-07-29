unit frmLoginU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DBCtrls ,IniFiles, DB,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, jpeg, cxGraphics, Menus, cxLookAndFeelPainters,
  cxButtons, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, cxLookAndFeels;

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
    Label1: TLabel;
    Label4: TLabel;
    btnCancel: TcxButton;
    txtUserId: TcxTextEdit;
    txtPassword: TcxTextEdit;
    cmbCompany: TcxImageComboBox;
    Image1: TImage;
    btnOk: TcxButton;
    procedure btnCancelClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cmbCompanyPropertiesChange(Sender: TObject);
  private
    { Private declarations }
    FDbSID : String;
    FUserId: String;
    FUserName: String;
    FPass: String;
    FCompany: String;
    FCompanyList: TList;
    procedure TransTmpIniFile(const aSourceFileName, aTempFileName: String);
    procedure BuildDataArear(const aFileName: String); overload;
    procedure BuildDataArear; overload;
    procedure BuildComArear( const aFileName: String);
    function HasExpire(const aFileName: String): Boolean;
    procedure GetInvoiceCompany;
    procedure SetConnStatus(AConnecting: Boolean);
  public
    { Public declarations }
    property CompanyList: TList read FCompanyList write FCompanyList;
  end;

var
  frmLogin: TfrmLogin;

implementation

uses frmMainU, dtmMainU, cbUtilis, Encryption_TLB, DateUtils;

{$R *.dfm}


{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.FormCreate(Sender: TObject);
var
  aSourceFileName, aTempFileName: String;
  aTempDir: array [0..MAX_PATH-1] of Char;
  aComFileName,aTempFileName2 : String;
begin
  {$IFDEF DEBUG}
    if ( ParamCount >0 ) then
    begin
      if ParamStr(1) ='kintest' then
      begin
        cmbSMSDb.Clear;
        BuildDataArear;
        Exit;
      end;
    end;
  {$ENDIF}

  aSourceFileName := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_SYS_INFO_INI_FILE;
  ZeroMemory( @aTempDir, SizeOf( aTempDir ) );
  GetTempPath( SizeOf( aTempDir ), aTempDir );
  aTempFileName := IncludeTrailingPathDelimiter( String( aTempDir ) ) +
    TMP_INV_SYS_INFO_INI_FILE;

  aComFileName := IncludeTrailingPathDelimiter(
    ExtractFilePath( Application.ExeName ) ) + INV_COM_INI_FILE;
  aTempFileName2 :=IncludeTrailingPathDelimiter( String( aTempDir ) ) +
    TMP_INV_COM_INFO_INI_FILE;

  if ( FileExists( aSourceFileName ) ) then
  begin
    try
      TransTmpIniFile( aSourceFileName, aTempFileName );
      { 檢查是否有設定使用期限, 若有設定則檢查否已過, 已過使用期限的話
        把連線設定檔砍掉 }
      if HasExpire( aTempFileName ) then
        DeleteFile( aSourceFileName )
      else begin
        BuildDataArear( aTempFileName );
      end;
    finally
      DeleteFile( aTempFileName );
    end;
    try
      if FileExists( aComFileName ) then
      begin
        TransTmpIniFile( aComFileName ,aTempFileName2 );
        BuildComArear( aTempFileName2 );
      end else
      begin
        BuildComArear( EmptyStr );
      end;
    finally
      DeleteFile( aTempFileName2 );
    end;


    {$IFDEF DEBUG}
    if ( ParamCount >= 2 ) then
    begin
      txtUserId.Text :=  ParamStr(1);
      txtPassword.Text :=  ParamStr(2);
    end;
    {$ENDIF}
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.FormDestroy(Sender: TObject);
begin
  { ... }
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.btnOKClick(Sender: TObject);
var
  aGroupId: String;
  aConnected: Boolean;
begin
    if ( FDbSID = EmptyStr ) then
    begin
      WarningMsg( '請先連接【發票】資料區。' );
      cmbSMSDb.SetFocus;
      Exit;
    end;
    if ( cmbCompany.Text = '' ) then
    begin
      WarningMsg( '請選擇【公司別】。' );
      cmbCompany.SetFocus;
      Exit;
    end;
    if Trim( txtUserId.Text ) = '' then
    begin
      WarningMsg( '請輸入【帳號】。' );
      txtUserId.SetFocus;
      Exit;
    end;
    if Trim( txtPassword.Text ) = '' then
    begin
      WarningMsg( '請輸入【密碼】。' );
      txtPassword.SetFocus;
      Exit;
    end;
    {}
    if dtmMain.GetLinkToMIS then
    begin
      aConnected := dtmMain.ConnectToSO(  dtmMain.GetSoInfo(
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).ConnSeq ) );
      if not aConnected then Exit;
    end;  
    {}
    FUserId := txtUserId.Text;
    FPass := txtPassword.Text;
    FCompany := cmbCompany.Properties.Items[cmbCompany.ItemIndex].Description;
    if dtmMain.checkPasswd( FUserId, FPass , aGroupId, FUserName ) then
    begin
      dtmMain.sG_LoginUser := FUserId;
      dtmMain.sG_LoginUserName := FUserName;
      dtmMain.sG_GroupId := aGroupId;
      dtmMain.sG_ServiceTypeStr :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).ServiceType;
      dtmMain.sG_CompID :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).CompanyId;
      dtmMain.sG_CompName :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).CompanyName;
      dtmMain.sG_LinkToMIS :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).LinkToMIS;
      dtmMain.sG_ConnSeq :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).ConnSeq;
      dtmMain.sG_AutoCreateNum :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).AutoCreateNum;
      dtmMain.sG_IfPrintTitle :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).IfPrintTitle;
      dtmMain.sG_IfPrintAddr :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).IfPrintAddr;
      dtmMain.sG_IfPrintCheck :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).IfPrintCheck;
      dtmMain.sG_ExpAddrType :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).ExpAddrType;
      dtmMain.sG_CompTel :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).Tel;
      dtmMain.sG_StarEInvoice :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).StarEInvoice;
      dtmMain.sG_StarEmail :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).StarEmail;
      dtmMain.sG_StarMessage :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).StarMessage;
      dtmMain.sG_StarAutoNotify :=
        PCompany( FCompanyList.Items[cmbCompany.ItemIndex] ).StartAutoNotify;
      ModalResult := mrOK;
    end
    else
    begin
      txtUserId.SetFocus;
      txtPassword.Text := EmptyStr;
      WarningMsg( '密碼錯誤。' );
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.btnCancelClick(Sender: TObject);
begin
  Application.Terminate;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.TransTmpIniFile(const aSourceFileName, aTempFileName: String);
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

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.BuildDataArear(const aFileName: String);
var
  aIndex: Integer;
  aSoInfo: PSMSDb;
  aItem: TcxImageComboBoxItem;
begin
  dtmMain.GetSMSDb( aFileName );
  cmbSMSDb.Clear;
  for aIndex := 0 to dtmMain.SMSDbList.Count - 1 do
  begin
    aSoInfo := PSMSDb( dtmMain.SMSDbList[aIndex] );
    aItem := cmbSMSDb.Properties.Items.Add;
    aItem.Value := IntToStr( aSoInfo.ConnSeq );
    aItem.Description := aSoInfo.Description;
    aItem.ImageIndex := 16;
  end;
end;

{ ---------------------------------------------------------------------------- }


function TfrmLogin.HasExpire(const aFileName: String): Boolean;
var
  aIniFile: TIniFile;
begin
  Result := True;
  aIniFile := TIniFile.Create( aFileName );
  try
    frmMain.ExpireDate := IncYear( Date, 100 );
    if aIniFile.SectionExists( UpperCase( EXPIRE_SECTION ) ) then
    begin
      frmMain.ExpireDate := aIniFile.ReadDate( UpperCase( EXPIRE_SECTION ),
        UpperCase( EXPIRE_DATE ), frmMain.ExpireDate );
    end;
  finally
    aIniFile.Free;
  end;
  try
    Result := ( Date > frmMain.ExpireDate );
  except
    { ... }
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.GetInvoiceCompany;
var
  aIndex: Integer;
  aItem: TcxImageComboBoxItem;
begin
  dtmMain.GetInvoiceCompany( FCompanyList );
  cmbCompany.Properties.Items.Clear;
  for aIndex := 0 to FCompanyList.Count - 1 do
  begin
    aItem := cmbCompany.Properties.Items.Add;
    aItem.Value := PCompany( FCompanyList[aIndex] ).CompanyId;
    aItem.Description := PCompany( FCompanyList[aIndex] ).CompanyName;
    aItem.ImageIndex := 8;
    aItem.Tag := Integer( FCompanyList[aIndex] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.btnConnectClick(Sender: TObject);
var
  aConnected: Boolean;
begin
  if cmbSMSDb.ItemIndex < 0 then Exit;
  SetConnStatus( True );
  try
    aConnected := dtmMain.ConnectToINV( PSMSDb(
      dtmMain.SMSDbList.Items[cmbSMSDb.ItemIndex] ) );
    if aConnected then
    begin
      FDbSID := cmbSMSDb.Properties.Items[cmbSMSDb.ItemIndex].Value;
      GetInvoiceCompany;
      cmbCompany.SetFocus;
    end;
  finally
    SetConnStatus( False );
  end;
end;

{ ---------------------------------------------------------------------------- }

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

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ( Key = VK_RETURN ) then btnOKClick( btnOk );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.cmbCompanyPropertiesChange(Sender: TObject);
begin
  if ( cmbCompany.ItemIndex >= 0 ) then
    txtUserId.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmLogin.BuildDataArear;
  var aIndex : Integer;
    aSoInfo: PSMSDb;
    aItem: TcxImageComboBoxItem;
begin
  dtmMain.GetSMSDb;
  cmbSMSDb.Clear;
  for aIndex := 0 to dtmMain.SMSDbList.Count - 1 do
  begin
    aSoInfo := PSMSDb( dtmMain.SMSDbList[aIndex] );
    aItem := cmbSMSDb.Properties.Items.Add;
    aItem.Value := IntToStr( aSoInfo.ConnSeq );
    aItem.Description := aSoInfo.Description;
    aItem.ImageIndex := 16;
  end;
end;

procedure TfrmLogin.BuildComArear(const aFileName: String);
begin
  dtmMain.GetCOMDb( aFileName );
end;

end.
