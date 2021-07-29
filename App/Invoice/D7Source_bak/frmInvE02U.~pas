unit frmInvE02U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, StdCtrls, Buttons, Grids, DBGrids, Mask, ExtCtrls,
  Encryption_TLB;

type
  TfrmInvE02 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    medUserId: TMaskEdit;
    medUserName: TMaskEdit;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrInv004: TDataSource;
    Label1: TLabel;
    medPassWord: TMaskEdit;
    Label2: TLabel;
    cobGroupID: TComboBox;
    Label3: TLabel;
    medCheckPassWord: TMaskEdit;
    Stt_Show: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv004DataChange(Sender: TObject; Field: TField);
    procedure btnSaveClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    G_GroupIDStrList : TStringList;
    FPassObj: _Password;
    procedure initialComboBox;
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvE02: TfrmInvE02;

implementation

uses frmMainU, dtmMainJU, dtmMainU;

{$R *.dfm}

procedure TfrmInvE02.FormShow(Sender: TObject);
begin
  FPassObj := CoPassword.Create;
  G_GroupIDStrList := TStringList.Create;
  initialComboBox;
  Self.Caption := frmMain.GetFormTitleString('E02','使用者資料維護');
  dtmMainJ.getAllInv004Data;
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE02.FormDestroy(Sender: TObject);
begin
  FPassObj := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE02.btnExitClick(Sender: TObject);
begin
    G_GroupIDStrList.Free;
    dtmMainJ.adoInv004Code.Close;
    Close;
end;

procedure TfrmInvE02.initialComboBox;
var
    sL_SQL : String;
begin
    //群組項目
    sL_SQL := 'select GroupId, Description from inv012 order by GroupId';
    dtmMainJ.createComboBoxItem(cobGroupID,sL_SQL,'GroupId','Description',false,G_GroupIDStrList);
    cobGroupID.ItemIndex := 0;

end;

procedure TfrmInvE02.ChangeBtnStatus;
begin
     with dsrInv004.DataSet do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlGrid.Enabled := false;
          pnlEdit.Enabled := true;
       end
       else
       begin
          if dsrInv004.DataSet.RecordCount>0 then
          begin
            btnAppend.Enabled :=TRUE;
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;
          pnlGrid.Enabled := true;
          pnlEdit.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmInvE02.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('E02',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

     if sL_InsertCompetence = 'Y' then
       btnAppend.Enabled := true
     else
       btnAppend.Enabled := false;


     if sL_DeleteCompetence = 'Y' then
       btnDelete.Enabled := true
     else
       btnDelete.Enabled := false;


     if sL_UpdateCompetence = 'Y' then
       btnEdit.Enabled := true
     else
       btnEdit.Enabled := false;


end;

procedure TfrmInvE02.initialForm;
begin
    with dsrInv004.DataSet do
    begin
      FindField('UserId').DisplayLabel := '使用者代碼';
      FindField('UserName').DisplayLabel := '使用者名稱';
      FindField('UptTime').DisplayLabel := '異動時間';
      FindField('UptEn').DisplayLabel := '異動人員';

      FindField('UserName').DisplayWidth := 30;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('Password').Visible := false;
      FindField('GroupID').Visible := false;
      FindField('DBConnSeq').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;
    end;

end;

procedure TfrmInvE02.btnAppendClick(Sender: TObject);
begin
    dsrInv004.DataSet.Append;

    ChangeBtnStatus;
    medUserId.SetFocus;
    cobGroupID.ItemIndex := 0;
end;

procedure TfrmInvE02.btnEditClick(Sender: TObject);
begin
    dsrInv004.DataSet.Edit;
    ChangeBtnStatus;
    medUserId.Enabled := false;
end;

procedure TfrmInvE02.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg('請確任要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dsrInv004.DataSet.Delete;

      dtmMainJ.getAllInv004Data;
      ChangeBtnStatus;
    end;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvE02.btnCancelClick(Sender: TObject);
begin
    if  dsrInv004.DataSet.State in [dsEdit] then
      medUserId.Enabled := true;

    dsrInv004.DataSet.Cancel;
    dsrInv004.DataSet.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvE02.dsrInv004DataChange(Sender: TObject; Field: TField);
var
  sL_GroupID, sL_TempGroupID, aPass: String;
  aIndex: Integer;
begin

    with dsrInv004.DataSet do
    begin
      if RecordCount > 0 then
      begin
        aPass := EmptyStr;
        if FieldByName('PASSWORD').AsString <> EmptyStr then
          aPass := FPassObj.Decrypt( FieldByName('PASSWORD').AsString, PASSKEY );
        medUserId.Text := FieldByName('USERID').AsString;
        medUserName.Text := FieldByName('USERNAME').AsString;
        medPassWord.Text := aPass;
        medCheckPassWord.Text := aPass;
        sL_GroupID := FieldByName('GROUPID').AsString;

        for aIndex := 0 to G_GroupIDStrList.Count-1 do
        begin
          sL_TempGroupID := (G_GroupIDStrList.Objects[aIndex] as TNormalObj).s_Code;
          if sL_GroupID = sL_TempGroupID then
          begin
            cobGroupID.ItemIndex := aIndex;
            Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);            
            exit;
          end;
        end;
      end
      else
      begin
        medUserId.Text := '';
        medUserName.Text := '';
        medPassWord.Text := '';
        medCheckPassWord.Text := '';
        cobGroupID.ItemIndex := 0;
      end;
    end;

end;

procedure TfrmInvE02.btnSaveClick(Sender: TObject);
var
    sL_UserId,sL_UserName,sL_PassWord,sL_SQL,sL_GroupID,sL_GroupName, aPass: String;
    bL_HavePKValue: Boolean;
begin
    if isDataOK then
    begin
      sL_UserId := Trim(medUserId.Text);
      sL_UserName := Trim(medUserName.Text);
      sL_PassWord := Trim(medPassWord.Text);
      dtmMainJ.getComboBoxCode(cobGroupID,G_GroupIDStrList,sL_GroupID,sL_GroupName);

      //檢核PK值
      sL_SQL := 'select * from inv004 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and UserId=''' + sL_UserId + '''';

      if  dsrInv004.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;


      if bL_HavePKValue then
      begin
        MessageDlg('違反唯一值條件',mtError,[mbOk],0);
      end
      else
      begin
        aPass := FPassObj.Encrypt( sL_Password, PASSKEY );
        with dsrInv004.DataSet do
        begin
          FieldByName('IdentifyId1').AsString := IDENTIFYID1;
          FieldByName('IdentifyId2').AsString := IDENTIFYID2;

          FieldByName('UserId').AsString := sL_UserId;
          FieldByName('UserName').AsString := sL_UserName;
          FieldByName('Password').AsString := aPass;
          FieldByName('GroupID').AsString := sL_GroupID;

          FieldByName('UptTime').AsString := DateTimeToStr(now);
          FieldByName('UptEn').AsString := dtmMain.getLoginUser;

          dsrInv004.DataSet.Post;
        end;
        dtmMainJ.getAllInv012Data;
        medUserId.Enabled := true;
        ChangeBtnStatus;
        setButtonCompetence;
        initialForm;
      end;        
    end;
end;

function TfrmInvE02.isDataOK: Boolean;
var
    sL_UserId,sL_UserName,sL_PassWord,sL_CheckPassWord : String;
begin
    Result := true;
    sL_UserId := Trim(medUserId.Text);
    sL_UserName := Trim(medUserName.Text);
    sL_PassWord := Trim(medPassWord.Text);
    sL_CheckPassWord := Trim(medCheckPassWord.Text);

    if sL_UserId = '' then
    begin
      MessageDlg('請輸入使用者代碼',mtError,[mbOk],0);
      medUserId.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_UserName = '' then
    begin
      MessageDlg('請輸入使用者名稱',mtError,[mbOk],0);
      medUserName.SetFocus;
      Result := false;
      Exit;
    end;


    if sL_PassWord = '' then
    begin
      MessageDlg('請輸入密碼',mtError,[mbOk],0);
      medPassWord.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_CheckPassWord = '' then
    begin
      MessageDlg('請輸入確認密碼',mtError,[mbOk],0);
      medCheckPassWord.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_CheckPassWord <> sL_PassWord then
    begin
      MessageDlg('你輸入的確認密碼不正確!',mtError,[mbOk],0);
      medCheckPassWord.SetFocus;
      Result := false;
      Exit;
    end;
end;

end.
