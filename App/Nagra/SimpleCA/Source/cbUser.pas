unit cbUser;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, cbMain, Encryption_TLB,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxCheckListBox, cxCheckBox,
  cxSpinEdit, cxMaskEdit, cxDropDownEdit, cxCalendar, cxStyles,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, DB, cxDBData,
  cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, ADODB;

type
  TDMLMode = ( dmBrowser, dmInsert, dmUpdate, dmDelete, dmSave, dmCancel );

  TfmUser = class(TForm)
    btnInsert: TButton;
    btnUpdate: TButton;
    btnDelete: TButton;
    btnCancel: TButton;
    btnSave: TButton;
    btnExit: TButton;
    Bevel1: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    txtUserId: TcxTextEdit;
    txtPassword: TcxTextEdit;
    txtUserName: TcxTextEdit;
    txtUserLock: TcxCheckBox;
    txtCompStr: TcxCheckListBox;
    Label4: TLabel;
    Label5: TLabel;
    txtExpireDate: TcxDateEdit;
    Label6: TLabel;
    txtChLimitCount: TcxSpinEdit;
    Label7: TLabel;
    txtChLimitDay: TcxSpinEdit;
    txtAdmin: TcxCheckBox;
    UserGrid: TcxGrid;
    UserGridView: TcxGridDBTableView;
    UserGridLevel: TcxGridLevel;
    UserDataSet: TADOQuery;
    UserDataSource: TDataSource;
    UserGridViewUSERID: TcxGridDBColumn;
    UserGridViewUSERNAME: TcxGridDBColumn;
    UserGridViewPASSWORD: TcxGridDBColumn;
    UserGridViewUSERLOCK: TcxGridDBColumn;
    UserGridViewCOMPSTR: TcxGridDBColumn;
    UserGridViewEXPIREDATE: TcxGridDBColumn;
    UserGridViewCHLIMITCOUNT: TcxGridDBColumn;
    UserGridViewCHLIMITDAY: TcxGridDBColumn;
    UserGridViewADMIN: TcxGridDBColumn;
    DataWriter: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure UserDataSetAfterScroll(DataSet: TDataSet);
    procedure btnInsertClick(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
    FMode: TDMLMode;
    FPassObj: _Password;
    FKey: WideString;
    procedure BuildCompList;
    procedure CompStrToCompList(aCompStr: String);
    function CompListToCompStr: String;
    function DataChange(aMode: TDMLMode): Boolean;
    procedure DMLModeChange(aMode: TDMLMode);
    procedure ButtonStateChange(aMode: TDMLMode);
    procedure EdiotrStateChange(aMode: TDMLMode);
    procedure ActivateDefaultEdiotr(aMode: TDMLMode);
    procedure ClearEditor;
    procedure DataToEditor;
    function ApplyDataSet: Boolean;
    function ApplySql: String;
    function RuleValidate: Boolean;
    function RuleConfirm: Boolean;
    function EncryptPassword(aPassword: String): String;
    function DecryptPassword(aPassword: String): String;
    procedure ReloadData;
  public
    { Public declarations }
  end;

var
  fmUser: TfmUser;

implementation

{$R *.dfm}

uses cbUtilis;

procedure TfmUser.FormCreate(Sender: TObject);
begin
  BuildCompList;
  UserDataSet.Close;
  UserDataSet.AfterScroll := nil;
  try
    UserDataSet.Open;
  finally
    UserDataSet.AfterScroll := UserDataSetAfterScroll;
  end;
  FPassObj := CoPassword.Create;
  FKey := 'CS';
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.FormShow(Sender: TObject);
begin
  UserDataSetAfterScroll( UserDataSet );
  DMLModeChange( dmBrowser );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if FMode in [dmInsert, dmUpdate] then
  begin
    CanClose := ConfirmMsg( '確認結束使用者設定?' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.FormDestroy(Sender: TObject);
begin
  UserDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.BuildCompList;
var
  aSo: TSoInfo;
  aIndex: Integer;
  aItem: TcxCheckListBoxItem;
begin
  txtCompStr.Items.Clear;
  for aIndex := 0 to fmMain.SoList.Count - 1 do
  begin
    aSo := TSoInfo( fmMain.SoList[aIndex] );
    aItem := txtCompStr.Items.Add;
    aItem.Text := aSo.CompName;
    aItem.Tag := Integer( aSo );
    aItem.State := cbsUnchecked;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.CompStrToCompList(aCompStr: String);

  { ---------------------------------- }

  function GetItemIndex(aCompCode: String): Integer;
  var
    aSo: TSoInfo;
    aIndex: Integer;
  begin
    Result := -1;
    for aIndex := 0 to txtCompStr.Items.Count - 1 do
    begin
      if ( TSoInfo( txtCompStr.Items[aIndex].Tag ).CompCode = aCompCode ) then
      begin
        Result := aIndex;
        Break;
      end;
    end;
  end;

  { ---------------------------------- }

var
  aCompCode: String;
  aItemIndex, aIndex: Integer;
begin
  if aCompStr = EmptyStr then
  begin
    for aIndex := 0 to txtCompStr.Items.Count - 1 do
      txtCompStr.Items[aIndex].State := cbsUnchecked;
  end else
  begin
    repeat
      aCompCode := ExtractValue( aCompStr );
      if ( aCompCode <> EmptyStr ) then
      begin
        aItemIndex := GetItemIndex( aCompCode );
        if aItemIndex >= 0 then
        begin
          txtCompStr.Items[aItemIndex].State := cbsChecked;
        end;
      end;
    until ( aCompStr = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.CompListToCompStr: String;
var
  aSo: TSoInfo;
  aIndex: Integer;
begin
  Result := EmptyStr;
  for aIndex := 0 to txtCompStr.Items.Count - 1 do
  begin
    if ( txtCompStr.Items[aIndex].State = cbsChecked ) then
    begin
      aSo := TSoInfo( txtCompStr.Items[aIndex].Tag );
      Result := ( Result + Format( '%s,', [aSo.CompCode] ) );
    end;
  end;
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.DMLModeChange(aMode: TDMLMode);
begin
  if aMode in [dmSave, dmCancel, dmDelete] then
  begin
    if not DataChange( aMode ) then Exit;
    aMode := dmBrowser;
  end;
  ButtonStateChange( aMode );
  EdiotrStateChange( aMode );
  ActivateDefaultEdiotr( aMode );
  FMode := aMode;
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.DataChange(aMode: TDMLMode): Boolean;
begin
  Result := False;
  if ( aMode = dmCancel ) then
  begin
    ClearEditor;
    DataToEditor;
    Result := True;
  end else
  if ( aMode = dmSave ) then
  begin
    Result := RuleValidate;
    if Result then
      Result := RuleConfirm;
    if Result then
      Result := ApplyDataSet;
    if Result then
      ReloadData;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.ButtonStateChange(aMode: TDMLMode);
begin
  if ( aMode in [dmBrowser] ) then
  begin
    btnInsert.Enabled := True;
    btnUpdate.Enabled := not UserDataSet.IsEmpty;
    btnDelete.Enabled := not UserDataSet.IsEmpty;
    btnCancel.Enabled := False;
    btnSave.Enabled := False;
  end else
  if ( aMode in [dmInsert, dmUpdate] ) then
  begin
    btnInsert.Enabled := False;
    btnUpdate.Enabled := False;
    btnDelete.Enabled := False;
    btnCancel.Enabled := True;
    btnSave.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TfmUser.EdiotrStateChange(aMode: TDMLMode);
begin
  if ( aMode in [dmBrowser, dmCancel] ) then
  begin
    txtUserId.Properties.ReadOnly := True;
    txtPassword.Properties.ReadOnly := True;
    txtUserName.Properties.ReadOnly := True;
    txtUserLock.Properties.ReadOnly := True;
    txtCompStr.ReadOnly := True;
    txtExpireDate.Properties.ReadOnly := True;
    txtAdmin.Properties.ReadOnly := True;
    txtChLimitCount.Properties.ReadOnly := True;
    txtChLimitDay.Properties.ReadOnly := True;
  end else
  if ( aMode in [dmInsert] ) then
  begin
    ClearEditor;
    txtUserId.Properties.ReadOnly := False;
    txtPassword.Properties.ReadOnly := False;
    txtUserName.Properties.ReadOnly := False;
    txtUserLock.Properties.ReadOnly := False;
    txtCompStr.ReadOnly := False;
    txtExpireDate.Properties.ReadOnly := False;
    txtAdmin.Properties.ReadOnly := False;
    txtChLimitCount.Properties.ReadOnly := False;
    txtChLimitDay.Properties.ReadOnly := False;
  end else
  if ( aMode in [dmUpdate] ) then
  begin
    txtUserId.Properties.ReadOnly := True;
    txtPassword.Properties.ReadOnly := False;
    txtUserName.Properties.ReadOnly := False;
    txtUserLock.Properties.ReadOnly := False;
    txtCompStr.ReadOnly := False;
    txtExpireDate.Properties.ReadOnly := False;
    txtAdmin.Properties.ReadOnly := False;
    txtChLimitCount.Properties.ReadOnly := False;
    txtChLimitDay.Properties.ReadOnly := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.ActivateDefaultEdiotr(aMode: TDMLMode);
begin
  if ( aMode in [dmBrowser, dmInsert] ) then
  begin
    if txtUserId.CanFocusEx then txtUserId.SetFocus;
  end else
  if ( aMode in [dmUpdate] ) then
  begin
    if txtPassword.CanFocusEx then txtPassword.SetFocus;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.ClearEditor;
begin
  txtUserId.Text := EmptyStr;
  txtPassword.Text := EmptyStr;
  txtUserName.Text := EmptyStr;
  txtUserLock.Checked := False;
  CompStrToCompList( EmptyStr );
  txtExpireDate.Text := EmptyStr;
  txtAdmin.Checked := False;
  txtChLimitCount.Value := 0;
  txtChLimitDay.Value := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.DataToEditor;
var
  aPassword: String;
begin
  if not UserDataSet.IsEmpty then
  begin
    aPassword := DecryptPassword( UserDataSet.FieldByName( 'PASSWORD' ).AsString );
    txtUserId.Text := UserDataSet.FieldByName( 'USERID' ).AsString;
    txtPassword.Text := aPassword;
    txtUserName.Text := UserDataSet.FieldByName( 'USERNAME' ).AsString;
    txtUserLock.Checked := ( UserDataSet.FieldByName( 'USERLOCK' ).AsString <> 'N' );
    CompStrToCompList( UserDataSet.FieldByName( 'COMPSTR' ).AsString );
    txtExpireDate.Text := UserDataSet.FieldByName( 'EXPIREDATE' ).AsString;
    txtAdmin.Checked := ( UserDataSet.FieldByName( 'ADMIN' ).AsString = 'Y' );
    txtChLimitCount.Value := UserDataSet.FieldByName( 'CHLIMITCOUNT' ).AsInteger;
    txtChLimitDay.Value := UserDataSet.FieldByName( 'CHLIMITDAY' ).AsInteger;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.UserDataSetAfterScroll(DataSet: TDataSet);
begin
  if ( FMode in [dmBrowser] ) then
  begin
    ClearEditor;
    DataToEditor;
  end else
  begin
    DMLModeChange( dmCancel );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.btnInsertClick(Sender: TObject);
begin
  DMLModeChange( dmInsert );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.btnUpdateClick(Sender: TObject);
begin
  DMLModeChange( dmUpdate );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認刪除此使用者資料?' ) then
    DMLModeChange( dmDelete );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.btnCancelClick(Sender: TObject);
begin
  DMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.ApplyDataSet: Boolean;
begin
  Result := False;
  DataWriter.SQL.Text := ApplySql;
  try
    DataWriter.ExecSQL;
    Result := True;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '使用者設定存檔有誤, 原因:%s。', [E.Message] ) ); 
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.ApplySql: String;
var
  aPassword, aUserLock, aCompStr, aExpireDate, aAdmin: String;
begin
  aUserLock := 'N';
  if txtUserLock.Checked then aUserLock := 'Y';
  aAdmin := 'N';
  if txtAdmin.Checked then aAdmin := 'Y';
  aExpireDate := txtExpireDate.Text;
  aCompStr := CompListToCompStr;
  aPassword := EncryptPassword( txtPassword.Text );
  {}
  if ( FMode in [dmInsert] ) then
  begin
    Result := Format(
      ' INSERT INTO CASIMPLE ( USERID, USERNAME, PASSWORD, ' +
      '    USERLOCK, COMPSTR, EXPIREDATE, CHLIMITCOUNT,    ' +
      '    CHLIMITDAY, ADMIN )                             ' +
      ' VALUES ( ''%s'', ''%s'', ''%s'',                   ' +
      '          ''%s'', ''%s'', ''%s'', ''%s'',           ' +
      '          ''%s'', ''%s'' )                          ',
      [txtUserId.Text, txtUserName.Text, aPassword, aUserLock,
       aCompStr, aExpireDate, txtChLimitCount.Text,
       txtChLimitDay.Text, aAdmin] );
  end else
  if ( FMode in [dmUpdate] ) then
  begin
    Result := Format(
      ' UPDATE CASIMPLE                ' +
      '    SET USERNAME = ''%s'',      ' +
      '        PASSWORD = ''%s'',      ' +
      '        USERLOCK = ''%s'',      ' +
      '        COMPSTR = ''%s'',       ' +
      '        EXPIREDATE = ''%s'',    ' +
      '        CHLIMITCOUNT = ''%s'',  ' +
      '        CHLIMITDAY = ''%s'',    ' +
      '        ADMIN = ''%s''          ' +
      '  WHERE USERID = ''%s''         ',
      [txtUserName.Text, aPassword, aUserLock,
       aCompStr, aExpireDate, txtChLimitCount.Text,
       txtChLimitDay.Text, aAdmin, txtUserId.Text] );
  end;
end;

{ ---------------------------------------------------------------------------- }


function TfmUser.RuleValidate: Boolean;
begin
  Result := False;
  if ( Trim( txtUserId.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入【使用者帳號】' );
    if txtUserId.CanFocusEx then txtUserId.SetFocus;
    Exit;
  end;
  if ( Trim( txtPassword.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入【登入密碼】' );
    if txtPassword.CanFocusEx then txtPassword.SetFocus;
    Exit;
  end;
  if ( Trim( txtUserName.Text ) = EmptyStr ) then
  begin
    WarningMsg( '請輸入【使用者姓名】' );
    if txtUserName.CanFocusEx then txtUserName.SetFocus;
    Exit;
  end;
  if ( CompListToCompStr = EmptyStr ) then
  begin
    WarningMsg( '請選擇【可發送指令系統台】' );
    if txtCompStr.CanFocusEx then txtCompStr.SetFocus;
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.RuleConfirm: Boolean;
begin
  if txtAdmin.Checked then
  begin
    Result := ConfirmMsg( '確認此使用者帳號設定為【管理員】?' );
    if not Result then
      if txtAdmin.CanFocusEx then begin
        txtAdmin.SetFocus;
        Exit;
      end;
  end;
  if ( txtExpireDate.Text = EmptyStr ) then
  begin
    Result := ConfirmMsg( '確認此使用者無【使用到期日】限制?' );
    if not Result then
      if txtExpireDate.CanFocusEx then begin
        txtExpireDate.SetFocus;
        Exit;
      end;
  end;
  if ( Nvl( txtChLimitCount.Text, '0' ) = '0' ) then
  begin
    Result := ConfirmMsg( '確認此使用者無【開頻道次數】限制?' );
    if not Result then
      if txtChLimitCount.CanFocusEx then begin
        txtChLimitCount.SetFocus;
        Exit;
      end;
  end;
  if ( Nvl( txtChLimitDay.Text, '0' ) = '0' ) then
  begin
    Result := ConfirmMsg( '確認此使用者無【開頻道天數】限制?' );
    if not Result then
      if txtChLimitDay.CanFocusEx then begin
        txtChLimitDay.SetFocus;
        Exit;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.btnSaveClick(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.DecryptPassword(aPassword: String): String;
begin
  Result := FPassObj.Decrypt( aPassword, FKey );
end;

{ ---------------------------------------------------------------------------- }

function TfmUser.EncryptPassword(aPassword: String): String;
begin
  Result := FPassObj.Encrypt( aPassword, FKey );
end;

{ ---------------------------------------------------------------------------- }


procedure TfmUser.ReloadData;
var
  aBookMark: TBookmark;
begin
  UserDataSet.AfterScroll := nil;
  try
    aBookMark := UserDataSet.GetBookmark;
    try
      UserDataSet.Close;
      UserDataSet.Open;
      UserDataSet.GotoBookmark( aBookMark );
    finally
      UserDataSet.FreeBookmark( aBookMark );
    end;
  finally
    UserDataSet.AfterScroll := UserDataSetAfterScroll;
  end;
  UserDataSetAfterScroll( UserDataSet );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmUser.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

end.
