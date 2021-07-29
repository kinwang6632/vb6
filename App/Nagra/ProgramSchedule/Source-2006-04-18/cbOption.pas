unit cbOption;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, FileCtrl, OracleCI, Menus,
  cbStyleModule, cbConfigModule, cbLogModule, RzPanel, cxPC, cxControls,
  cxLookAndFeelPainters, StdCtrls, cxButtons, ImgList, cxTextEdit,
  cxMaskEdit, cxButtonEdit, cxContainer, cxEdit, cxGroupBox, cxGraphics,
  cxDropDownEdit, cxCheckBox, cxRadioGroup, cxHyperLinkEdit, cxSpinEdit;

type
  TfmOption = class(TForm)
    RzPanel1: TRzPanel;
    RzPanel2: TRzPanel;
    OptionPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    cxTabSheet3: TcxTabSheet;
    btnConfirm: TcxButton;
    btnCancel: TcxButton;
    DexGroup: TcxGroupBox;
    WrapperGroup: TcxGroupBox;
    AsRunGroup: TcxGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Lable18: TLabel;
    edtDexSrc: TcxButtonEdit;
    edtDexDest: TcxButtonEdit;
    edtDexErr: TcxButtonEdit;
    edtDexBackup: TcxButtonEdit;
    edtWrapperSrc: TcxButtonEdit;
    edtWrapperDest: TcxButtonEdit;
    edtWrapperErr: TcxButtonEdit;
    edtWrapperBackup: TcxButtonEdit;
    edtAsRunSrc: TcxButtonEdit;
    edtAsRunDest: TcxButtonEdit;
    edtAsRunErr: TcxButtonEdit;
    edtAsRunBackup: TcxButtonEdit;
    cxGroupBox2: TcxGroupBox;
    edtDbAccount: TcxTextEdit;
    edtDbPassword: TcxTextEdit;
    cmbDbSid: TcxComboBox;
    PeriodGroup: TcxGroupBox;
    rdoCheckFix1: TcxRadioButton;
    rdoCheckFix2: TcxRadioButton;
    edtMinute: TcxMaskEdit;
    chkClock1: TcxCheckBox;
    chkClock2: TcxCheckBox;
    chkClock3: TcxCheckBox;
    chkClock4: TcxCheckBox;
    chkClock5: TcxCheckBox;
    chkClock6: TcxCheckBox;
    chkClock7: TcxCheckBox;
    chkClock8: TcxCheckBox;
    chkClock9: TcxCheckBox;
    chkClock10: TcxCheckBox;
    chkClock11: TcxCheckBox;
    chkClock12: TcxCheckBox;
    chkClock13: TcxCheckBox;
    chkClock14: TcxCheckBox;
    chkClock15: TcxCheckBox;
    chkClock16: TcxCheckBox;
    chkClock17: TcxCheckBox;
    chkClock18: TcxCheckBox;
    chkClock19: TcxCheckBox;
    chkClock20: TcxCheckBox;
    chkClock21: TcxCheckBox;
    chkClock22: TcxCheckBox;
    chkClock23: TcxCheckBox;
    chkClock24: TcxCheckBox;
    cxTabSheet4: TcxTabSheet;
    cxGroupBox1: TcxGroupBox;
    chkErrorNotify: TcxCheckBox;
    cxGroupBox3: TcxGroupBox;
    chkChangeNotify: TcxCheckBox;
    cmbErrList: TcxComboBox;
    cmbChangeList: TcxComboBox;
    btnAddErr: TcxButton;
    btnAddChange: TcxButton;
    cxGroupBox4: TcxGroupBox;
    Label12: TLabel;
    edtSMTPServer: TcxTextEdit;
    Label13: TLabel;
    edtSMTPEmail: TcxHyperLinkEdit;
    edtSMTPAccount: TcxTextEdit;
    Label14: TLabel;
    Label18: TLabel;
    edtSMTPPassword: TcxTextEdit;
    cxGroupBox5: TcxGroupBox;
    lbDbRetryFrequence: TLabel;
    edtDbRetrySec: TcxSpinEdit;
    cxGroupBox6: TcxGroupBox;
    Label19: TLabel;
    edtActLoadDays: TcxSpinEdit;
    edtMsgLoadDays: TcxSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FilePathButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure PeriodTypeClick(Sender: TObject);
    procedure NotifyPropertiesChange(Sender: TObject);
    procedure EmailModifyClick(Sender: TObject);
    procedure btnConfirmClick(Sender: TObject);
  private
    { Private declarations }
    procedure InitialPageControl;
    procedure ClearEditor;
    procedure OptionToEditor;
    procedure EditorToOption;
    procedure ChangePeriodState(aPeriodType: String);
    procedure ChangeNotifyState(aNotifyType: String; const aEnable: Boolean);
    function ValidateInput: Boolean;
  public
    { Public declarations }
  end;

var
  fmOption: TfmOption;

implementation

uses cbUtilis, cbEmail;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FormCreate(Sender: TObject);
begin
  InitialPageControl;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FormShow(Sender: TObject);
begin
  ClearEditor;
  cmbDbSid.Properties.Items.Assign( OracleAliasList );
  OptionToEditor;
  ChangePeriodState( ConfigModule.Period.PeriodType );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FormDestroy(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.InitialPageControl;
begin
  OptionPage.ActivePageIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.ClearEditor;
var
  aIndex: Integer;
  aControl: TControl;
begin
  edtDexSrc.Text := EmptyStr;
  edtDexDest.Text := EmptyStr;
  edtDexErr.Text := EmptyStr;
  edtDexBackup.Text := EmptyStr;
  edtWrapperSrc.Text := EmptyStr;
  edtWrapperDest.Text := EmptyStr;
  edtWrapperErr.Text := EmptyStr;
  edtWrapperBackup.Text := EmptyStr;
  edtAsRunSrc.Text := EmptyStr;
  edtAsRunDest.Text := EmptyStr;
  edtAsRunErr.Text := EmptyStr;
  edtAsRunBackup.Text := EmptyStr;
  cmbDbSid.Clear;
  edtDbAccount.Text := EmptyStr;
  edtDbPassword.Text := EmptyStr;
  rdoCheckFix1.Checked := False;
  rdoCheckFix2.Checked := False;
  edtMinute.Text := EmptyStr;
  for aIndex := 1 to 24 do
  begin
    aControl := PeriodGroup.FindChildControl( Format( 'chkClock%d', [aIndex] ) );
    if Assigned( aControl ) then TcxCheckBox( aControl ).Checked := False;
  end;
  chkErrorNotify.Checked := False;
  chkChangeNotify.Checked := False;
  cmbErrList.Properties.Items.Clear;
  cmbChangeList.Properties.Items.Clear;
  edtSMTPServer.Text := EmptyStr;
  edtSMTPEmail.Text := EmptyStr;
  edtSMTPAccount.Text := EmptyStr;
  edtSMTPPassword.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.OptionToEditor;
var
  aIndex: Integer;
  aControlName: String;
  aControl: TControl;
begin
  edtDexSrc.Text := ConfigModule.Param.DexSrc;
  edtDexDest.Text := ConfigModule.Param.DexDest;
  edtDexErr.Text := ConfigModule.Param.DexErr;
  edtDexBackup.Text := ConfigModule.Param.DexBackup;
  edtWrapperSrc.Text := ConfigModule.Param.WrapperSrc;
  edtWrapperDest.Text := ConfigModule.Param.WrapperDest;
  edtWrapperErr.Text := ConfigModule.Param.WrapperErr;
  edtWrapperBackup.Text := ConfigModule.Param.WrapperBackup;
  edtAsRunSrc.Text := ConfigModule.Param.AsRunSrc;
  edtAsRunDest.Text := ConfigModule.Param.AsRunDest;
  edtAsRunErr.Text := ConfigModule.Param.AsRunErr;
  edtAsRunBackup.Text := ConfigModule.Param.AsRunBackup;
  cmbDbSid.Text := EmptyStr;
  if ( cmbDbSid.Properties.Items.IndexOf( ConfigModule.Database.DbSid ) >= 0 ) then
    cmbDbSid.Text := ConfigModule.Database.DbSid;
  edtDbAccount.Text := ConfigModule.Database.DbAccount;
  edtDbPassword.Text := ConfigModule.Database.DbPassoword;
  edtDbRetrySec.Value := ConfigModule.Database.DbRetrySec;
  if ( ConfigModule.Period.PeriodType = '1' ) then
  begin
    rdoCheckFix1.Checked := True;
    edtMinute.Text := IntToStr( ConfigModule.Period.PeriodMinute );
  end else
  begin
    rdoCheckFix2.Checked := True;
    for aIndex := 0 to ConfigModule.Period.ClockCount - 1 do
    begin
      if ( aIndex = 0 ) then
        aControlName := Format( 'chkClock%d', [24] )
      else
        aControlName := Format( 'chkClock%d', [aIndex] );
      aControl := PeriodGroup.FindChildControl( aControlName );
      if Assigned( aControl ) then
        TcxCheckBox( aControl ).Checked := ConfigModule.Period.Clocks[aIndex].IsSet;
    end;
  end;
  chkErrorNotify.Checked := ConfigModule.Notify.ErrorNotify;
  cmbErrList.Properties.Items.Assign( ConfigModule.Notify.ErrorEmail );
  if ( cmbErrList.Properties.Items.Count > 0 ) then
    cmbErrList.ItemIndex := 0;
  ChangeNotifyState( '1', ConfigModule.Notify.ErrorNotify );
  chkChangeNotify.Checked := ConfigModule.Notify.ChangeNotiy;
  cmbChangeList.Properties.Items.Assign( ConfigModule.Notify.ChangeEmail );
  if ( cmbChangeList.Properties.Items.Count > 0 ) then
    cmbChangeList.ItemIndex := 0;
  ChangeNotifyState( '2', ConfigModule.Notify.ChangeNotiy );
  edtSMTPServer.Text := ConfigModule.Notify.SMTPSetting.Server;
  edtSMTPEmail.Text := ConfigModule.Notify.SMTPSetting.Email;
  edtSMTPAccount.Text := ConfigModule.Notify.SMTPSetting.Account;
  edtSMTPPassword.Text := ConfigModule.Notify.SMTPSetting.Password;
  {}
  edtActLoadDays.Value := LogModule.Params.ActLoadDays;
  edtMsgLoadDays.Value := LogModule.Params.MsgLoadDays;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.EditorToOption;
var
  aIndex: Integer;
begin
  ConfigModule.Param.DexSrc := edtDexSrc.Text;
  ConfigModule.Param.DexDest := edtDexDest.Text;
  ConfigModule.Param.DexErr := edtDexErr.Text;
  ConfigModule.Param.DexBackup := edtDexBackup.Text;
  ConfigModule.Param.WrapperSrc := edtWrapperSrc.Text;
  ConfigModule.Param.WrapperDest := edtWrapperDest.Text;
  ConfigModule.Param.WrapperErr := edtWrapperErr.Text;
  ConfigModule.Param.WrapperBackup := edtWrapperBackup.Text;
  ConfigModule.Param.AsRunSrc := edtAsRunSrc.Text;
  ConfigModule.Param.AsRunDest := edtAsRunDest.Text;
  ConfigModule.Param.AsRunErr := edtAsRunErr.Text;
  ConfigModule.Param.AsRunBackup := edtAsRunBackup.Text;
  ConfigModule.Database.DbAccount := edtDbAccount.Text;
  ConfigModule.Database.DbPassoword := edtDbPassword.Text;
  ConfigModule.Database.DbSid := cmbDbSid.Text;
  ConfigModule.Database.DbRetrySec := edtDbRetrySec.Value;
  if ( rdoCheckFix1.Checked ) then
  begin
    ConfigModule.Period.PeriodType := '1';
    ConfigModule.Period.PeriodMinute := StrToInt( edtMinute.Text );
    for aIndex := 0 to ConfigModule.Period.ClockCount - 1 do
      ConfigModule.Period.Clocks[aIndex].IsSet := False;
  end else
  begin
    ConfigModule.Period.PeriodType := '2';
    ConfigModule.Period.PeriodMinute := -1;
    ConfigModule.Period.Clocks[1].IsSet := chkClock1.Checked;
    ConfigModule.Period.Clocks[2].IsSet := chkClock2.Checked;
    ConfigModule.Period.Clocks[3].IsSet := chkClock3.Checked;
    ConfigModule.Period.Clocks[5].IsSet := chkClock4.Checked;
    ConfigModule.Period.Clocks[6].IsSet := chkClock6.Checked;
    ConfigModule.Period.Clocks[7].IsSet := chkClock7.Checked;
    ConfigModule.Period.Clocks[9].IsSet := chkClock8.Checked;
    ConfigModule.Period.Clocks[10].IsSet := chkClock10.Checked;
    ConfigModule.Period.Clocks[11].IsSet := chkClock11.Checked;
    ConfigModule.Period.Clocks[12].IsSet := chkClock12.Checked;
    ConfigModule.Period.Clocks[13].IsSet := chkClock13.Checked;
    ConfigModule.Period.Clocks[14].IsSet := chkClock14.Checked;
    ConfigModule.Period.Clocks[15].IsSet := chkClock15.Checked;
    ConfigModule.Period.Clocks[16].IsSet := chkClock16.Checked;
    ConfigModule.Period.Clocks[17].IsSet := chkClock17.Checked;
    ConfigModule.Period.Clocks[18].IsSet := chkClock18.Checked;
    ConfigModule.Period.Clocks[19].IsSet := chkClock19.Checked;
    ConfigModule.Period.Clocks[20].IsSet := chkClock20.Checked;
    ConfigModule.Period.Clocks[21].IsSet := chkClock21.Checked;
    ConfigModule.Period.Clocks[22].IsSet := chkClock22.Checked;
    ConfigModule.Period.Clocks[23].IsSet := chkClock23.Checked;
    ConfigModule.Period.Clocks[0].IsSet := chkClock24.Checked;
  end;
  ConfigModule.Notify.ErrorNotify := chkErrorNotify.Checked;
  ConfigModule.Notify.ErrorEmail.Assign( cmbErrList.Properties.Items );
  ConfigModule.Notify.ChangeNotiy := chkChangeNotify.Checked;
  ConfigModule.Notify.ChangeEmail.Assign( cmbChangeList.Properties.Items );
  ConfigModule.Notify.SMTPSetting.Server := edtSMTPServer.Text;
  ConfigModule.Notify.SMTPSetting.Email := edtSMTPEmail.Text;
  ConfigModule.Notify.SMTPSetting.Account := edtSMTPAccount.Text;
  ConfigModule.Notify.SMTPSetting.Password := edtSMTPPassword.Text;
  {}
  LogModule.Params.ActLoadDays := edtActLoadDays.Value;
  LogModule.Params.MsgLoadDays := edtMsgLoadDays.Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.ChangePeriodState(aPeriodType: String);
var
  aIndex: Integer;
  aControl: TControl;
begin
  edtMinute.Enabled := ( aPeriodType = '1' );
  if ( not edtMinute.Enabled ) then edtMinute.Text := EmptyStr;
  for aIndex := 1 to 24 do
  begin
    aControl := PeriodGroup.FindChildControl( Format( 'chkClock%d', [aIndex] ) );
    aControl.Enabled := ( aPeriodType = '2' );
    if ( not aControl.Enabled ) then TcxCheckBox( aControl ).Checked := False;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.ChangeNotifyState(aNotifyType: String;
  const aEnable: Boolean);
begin
  if ( aNotifyType = '1' ) then
  begin
    cmbErrList.Enabled := aEnable;
    btnAddErr.Enabled := aEnable;
  end else
  if ( aNotifyType = '2' ) then
  begin
    cmbChangeList.Enabled := aEnable;
    btnAddChange.Enabled := aEnable;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmOption.ValidateInput: Boolean;
var
  aIndex: Integer;
  aControl: TControl;
  aBool: Boolean;

  { ---------------------------------------- }

  function GetChangePageIndex(aEdit: TcxCustomTextEdit): Integer;
  begin
    Result := -1;
    if ( aEdit = edtDexSrc ) or ( aEdit = edtDexDest ) or
       ( aEdit = edtDexErr ) or ( aEdit = edtDexBackup ) or
       ( aEdit = edtWrapperSrc ) or ( aEdit = edtWrapperDest ) or
       ( aEdit = edtWrapperErr ) or ( aEdit = edtWrapperBackup ) or
       ( aEdit = edtAsRunSrc ) or ( aEdit = edtAsRunDest ) or
       ( aEdit = edtAsRunErr ) or ( aEdit = edtAsRunBackup ) then
     Result := 0 else
    if ( aEdit = cmbDbSid ) or ( aEdit = edtDbAccount ) or
       ( aEdit = edtDbPassword ) then
     Result := 2;
  end;

  { ---------------------------------------- }

  function CheckIfEmpty(aEdit: TcxCustomTextEdit; const aMsg: String): Boolean;
  var
    aPageIndex: Integer;
  begin
    Result := False;
    if not Assigned( aEdit ) then Exit;
    if ( aEdit.Text = EmptyStr ) then
    begin
      aPageIndex := GetChangePageIndex( aEdit );
      if ( aPageIndex in [0..OptionPage.PageCount-1] ) then
        OptionPage.ActivePageIndex := aPageIndex;
      WarningMsg( aMsg );
      if ( aEdit.CanFocusEx ) then aEdit.SetFocus;
      Exit;
    end;
    Result := True;
  end;

  { ---------------------------------------- }

  function CheckIfSameDir(aSourceEdit, aDestEdit: TcxCustomTextEdit;
    const aMsg: String): Boolean;
  var
    aPageIndex: Integer;
  begin
    Result := False;
    if Assigned( aSourceEdit ) and Assigned( aDestEdit ) then
    begin
      Result := ( aSourceEdit.Text = aDestEdit.Text );
      if Result then
      begin
        aPageIndex := GetChangePageIndex( aSourceEdit );
        if ( aPageIndex in [0..OptionPage.PageCount-1] ) then
          OptionPage.ActivePageIndex := aPageIndex;
        WarningMsg( aMsg );
        if ( aSourceEdit.CanFocusEx ) then aSourceEdit.SetFocus;
      end;
    end;  
  end;

  { ---------------------------------------- }

begin
  Result := False;
  if not CheckIfEmpty( edtDexSrc, '請輸入【Dex檔案來源位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtDexDest, '請輸入【Dex檔案目的位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtDexErr, '請輸入【Dex檔案錯誤位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtDexBackup, '請輸入【Dex檔案備份位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtWrapperSrc, '請輸入【Wrapper檔案來源位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtWrapperDest, '請輸入【Wrapper檔案目的位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtWrapperErr, '請輸入【Wrapper檔案錯誤位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtWrapperBackup, '請輸入【Wrapper檔案備份位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtAsRunSrc, '請輸入【AsRunLog檔案來源位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtAsRunDest, '請輸入【AsRunLog檔案目的位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtAsRunErr, '請輸入【AsRunLog檔案錯誤位置】。' ) then
    Exit;
  if not CheckIfEmpty( edtAsRunBackup, '請輸入【AsRunLog檔案備份位置】。' ) then
    Exit;
  if not CheckIfEmpty( cmbDbSid, '請輸入【資料庫別名】。' ) then
    Exit;
  if not CheckIfEmpty( edtDbAccount, '請輸入【使用者帳號】。' ) then
    Exit;
  if not CheckIfEmpty( edtDbPassword, '請輸入【資料庫密碼】。' ) then
    Exit;
  { Dex 檔案目錄設定不可相同 }
  if CheckIfSameDir( edtDexDest, edtDexBackup, '【Dex檔案目的位置】與【Dex檔案備份位置】不可相同。' ) then
    Exit;
  if CheckIfSameDir( edtDexDest, edtDexErr, '【Dex檔案目的位置】與【Dex檔案錯誤位置】不可相同。' ) then
    Exit;
  if CheckIfSameDir( edtDexBackup, edtDexErr, '【Dex檔案備份位置】與【Dex檔案錯誤位置】不可相同。' ) then
    Exit;
  { Warpper 檔案目錄設定不可相同 }
  if CheckIfSameDir( edtWrapperDest, edtWrapperBackup, '【Wrapper檔案目的位置】與【Wrapper檔案備份位置】不可相同。' ) then
    Exit;
  if CheckIfSameDir( edtWrapperDest, edtWrapperErr, '【Wrapper檔案目的位置】與【Wrapper檔案錯誤位置】不可相同。' ) then
    Exit;
  if CheckIfSameDir( edtWrapperBackup, edtWrapperErr, '【Wrapper檔案備份位置】與【Wrapper檔案錯誤位置】不可相同。' ) then
    Exit;
  { AsRun 檔案目錄設定不可相同 }
  if CheckIfSameDir( edtAsRunDest, edtAsRunBackup, '【AsRun檔案目的位置】與【AsRun檔案備份位置】不可相同。' ) then
    Exit;
  if CheckIfSameDir( edtAsRunDest, edtAsRunErr, '【AsRun檔案目的位置】與【AsRun檔案錯誤位置】不可相同。' ) then
    Exit;
  if CheckIfSameDir( edtAsRunBackup, edtAsRunErr, '【AsRun檔案備份位置】與【AsRun檔案錯誤位置】不可相同。' ) then
    Exit;
  {}  
  if ( rdoCheckFix1.Checked ) then
  begin
    if ( edtMinute.Text = EmptyStr ) then
    begin
      OptionPage.ActivePageIndex := 1;
      WarningMsg( '請輸入【偵測檔案】分鐘值。' );
      if ( edtMinute.CanFocusEx ) then edtMinute.SetFocus;
      Exit;
    end;
  end else
  begin
    aBool := False;
    for aIndex := 1 to 24 do
    begin
      aControl := PeriodGroup.FindChildControl( Format( 'chkClock%d', [aIndex] ) );
      if Assigned( aControl ) then
        aBool := ( aBool or TcxCheckBox( aControl ).Checked );
    end;
    if not aBool then
    begin
      OptionPage.ActivePageIndex := 1;
      WarningMsg( '請至少勾選一項【檔案偵測週期】時段。' );
      if ( chkClock1.CanFocusEx ) then chkClock1.SetFocus;
      Exit;
    end;
  end;
  if ( chkErrorNotify.Checked ) and ( cmbErrList.Properties.Items.Count <= 0 ) then
  begin
    OptionPage.ActivePageIndex := 3;
    WarningMsg( '已啟用【錯誤通知】, 請至少輸入一筆電子郵件位址。' );
    if ( chkErrorNotify.CanFocusEx ) then chkErrorNotify.SetFocus;
    Exit;
  end;
  if ( chkChangeNotify.Checked ) and ( cmbChangeList.Properties.Items.Count <= 0 ) then
  begin
    OptionPage.ActivePageIndex := 3;
    WarningMsg( '已啟用【異動通知】, 請至少輸入一筆電子郵件位址。' );
    if ( chkChangeNotify.CanFocusEx ) then chkChangeNotify.SetFocus;
    Exit;
  end;
  if ( chkErrorNotify.Checked or chkChangeNotify.Checked ) then
  begin
    if ( edtSMTPServer.Text = EmptyStr ) then
    begin
      OptionPage.ActivePageIndex := 3;
      WarningMsg( '請輸入【郵件伺服器】。' );
      if ( edtSMTPServer.CanFocusEx ) then edtSMTPServer.SetFocus;
      Exit;
    end;
    if ( Trim( edtSMTPEmail.Text ) = EmptyStr ) then
    begin
      OptionPage.ActivePageIndex := 3;
      WarningMsg( '請輸入【通知者電子郵件】。' );
      if ( edtSMTPEmail.CanFocusEx ) then edtSMTPEmail.SetFocus;
      Exit;
    end;
    if ( Trim( edtSMTPAccount.Text ) = EmptyStr ) then
    begin
      OptionPage.ActivePageIndex := 3;
      WarningMsg( '請輸入【登入使用者】。' );
      if ( edtSMTPAccount.CanFocusEx ) then edtSMTPAccount.SetFocus;
      Exit;
    end;
    if ( Trim( edtSMTPPassword.Text ) = EmptyStr ) then
    begin
      OptionPage.ActivePageIndex := 3;
      WarningMsg( '請輸入【登入密碼】。' );
      if ( edtSMTPPassword.CanFocusEx ) then edtSMTPPassword.SetFocus;
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FilePathButtonClick(Sender: TObject;
  AButtonIndex: Integer);
var
  aDir: String;
begin
  aDir := TcxButtonEdit( Sender ).Text;
  if SelectDirectory( '選擇目錄', EmptyWideStr, aDir ) then
    TcxButtonEdit( Sender ).Text := aDir;
  TcxButtonEdit( Sender ).SelectAll;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.PeriodTypeClick(Sender: TObject);
begin
  if ( Sender = rdoCheckFix1 ) then
    ChangePeriodState( '1' )
  else
    ChangePeriodState( '2' );
end;

{ ---------------------------------------------------------------------------- }


procedure TfmOption.NotifyPropertiesChange(Sender: TObject);
begin
  if ( Sender = chkErrorNotify ) then
    ChangeNotifyState( '1', chkErrorNotify.Checked )
  else
    ChangeNotifyState( '2', chkChangeNotify.Checked );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.EmailModifyClick(Sender: TObject);
var
  aSource: String;
begin
  aSource := '1';
  if ( Sender = btnAddChange ) then aSource := '2';
  fmEmail := TfmEmail.Create( nil );
  try
    if ( aSource = '1' ) then
      fmEmail.lstEmail.Items.Assign( cmbErrList.Properties.Items )
    else
      fmEmail.lstEmail.Items.Assign( cmbChangeList.Properties.Items );
    if ( fmEmail.ShowModal = mrOk ) then
    begin
      if ( aSource = '1' ) then
      begin
        cmbErrList.Properties.Items.Assign( fmEmail.lstEmail.Items );
        if ( cmbErrList.Properties.Items.Count > 0 ) then
          cmbErrList.ItemIndex := 0;
      end else
      begin
        cmbChangeList.Properties.Items.Assign( fmEmail.lstEmail.Items );
        if ( cmbChangeList.Properties.Items.Count > 0 ) then
          cmbChangeList.ItemIndex := 0;
      end;
    end;
  finally
    fmEmail.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.btnConfirmClick(Sender: TObject);
begin
  if not ValidateInput then Exit;
  EditorToOption;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

end.
