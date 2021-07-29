unit frmInvD06U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, StdCtrls, Buttons, Grids, DBGrids, Mask, ExtCtrls;

type
  TfrmInvD06 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    medItemId: TMaskEdit;
    medDesc: TMaskEdit;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrInv012: TDataSource;
    Stt_Show: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv012DataChange(Sender: TObject; Field: TField);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
    function CheckCanDelete: Boolean;
    procedure DeleteOtherData;
  public
    { Public declarations }
  end;

var
  frmInvD06: TfrmInvD06;

implementation

uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;

{$R *.dfm}

procedure TfrmInvD06.FormShow(Sender: TObject);
begin
   Self.Caption := frmMain.GetFormTitleString( 'D06', '權限群組別代碼資料維護' );
   dtmMainJ.getAllInv012Data;
   ChangeBtnStatus;
   setButtonCompetence;
   initialForm;
end;

procedure TfrmInvD06.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvD06.ChangeBtnStatus;
begin
     with dsrInv012.DataSet do
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
          if dsrInv012.DataSet.RecordCount>0 then
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

procedure TfrmInvD06.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
    sL_VerifyCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('D06',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,
          sL_UpdateCompetence,sL_ExecuteCompetence,sL_VerifyCompetence);

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

procedure TfrmInvD06.initialForm;
begin
    with dsrInv012.DataSet do
    begin
      FindField('GROUPID').DisplayLabel := '服務種類代碼';
      FindField('DESCRIPTION').DisplayLabel := '服務種類說明';
      FindField('UPTTIME').DisplayLabel := '異動時間';
      FindField('UPTEN').DisplayLabel := '異動人員';

      FindField('DESCRIPTION').DisplayWidth := 30;
      FindField('UPTEN').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;
    end;
end;

procedure TfrmInvD06.btnAppendClick(Sender: TObject);
begin
    dsrInv012.DataSet.Append;

    ChangeBtnStatus;
    medItemId.SetFocus;
end;

procedure TfrmInvD06.btnEditClick(Sender: TObject);
begin
    dsrInv012.DataSet.Edit;
    ChangeBtnStatus;
    medItemId.Enabled := false;
end;

procedure TfrmInvD06.btnDeleteClick(Sender: TObject);
begin
  if not CheckCanDelete then Exit;
  if ConfirmMsg( '確認刪除此筆資料?' ) then
  begin
    DeleteOtherData;
    dsrInv012.DataSet.Delete;
    dtmMainJ.getAllInv012Data;
    ChangeBtnStatus;
  end;
  setButtonCompetence;
  initialForm;
end;

procedure TfrmInvD06.btnCancelClick(Sender: TObject);
begin
    if  dsrInv012.DataSet.State in [dsEdit] then
      medItemId.Enabled := true;

    dsrInv012.DataSet.Cancel;
    dsrInv012.DataSet.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD06.dsrInv012DataChange(Sender: TObject; Field: TField);
begin
    with dsrInv012.DataSet do
    begin
      if RecordCount > 0 then
      begin
        medItemId.Text := FieldByName( 'GROUPID' ).AsString;
        medDesc.Text := FieldByName( 'DESCRIPTION' ).AsString;
        Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);
      end
      else
      begin
        medItemId.Text := '';
        medDesc.Text := '';
      end;
      Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);      
    end;
end;

function TfrmInvD06.isDataOK: Boolean;
var
    sL_GroupId,sL_Description : String;
begin
    Result := true;
    sL_GroupId := Trim(medItemId.Text);
    sL_Description := Trim(medDesc.Text);

    if sL_GroupId = '' then
    begin
      MessageDlg('請輸入群組代碼',mtError,[mbOk],0);
      medItemId.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_Description = '' then
    begin
      MessageDlg('請輸入群組名稱',mtError,[mbOk],0);
      medDesc.SetFocus;
      Result := false;
      Exit;
    end;


end;

procedure TfrmInvD06.btnSaveClick(Sender: TObject);
var
    sL_GroupId,sL_Desc,sL_SQL : String;
    bL_HavePKValue : Boolean;
begin
    if isDataOK then
    begin
      sL_GroupId := Trim(medItemId.Text);
      sL_Desc := Trim(medDesc.Text);


      //檢核PK值
      sL_SQL := 'SELECT * FROM INV012 WHERE IDENTIFYID1=''' + IDENTIFYID1 +
                ''' AND IDENTIFYID2=' + IDENTIFYID2 + ' AND GROUPID=''' + sL_GroupId + '''';


      if  dsrInv012.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;


      if bL_HavePKValue then
      begin
        MessageDlg('違反唯一值條件',mtError,[mbOk],0);
        medDesc.SetFocus;
      end
      else
      begin
        with dsrInv012.DataSet do
        begin
          FieldByName('IDENTIFYID1').AsString := IDENTIFYID1;
          FieldByName('IDENTIFYID2').AsString := IDENTIFYID2;

          FieldByName('GROUPID').AsString := sL_GroupId;
          FieldByName('DESCRIPTION').AsString := sL_Desc;

          FieldByName('UPTTIME').AsString := DateTimeToStr(now);
          FieldByName('UPTEN').AsString := dtmMain.getLoginUser;

          dsrInv012.DataSet.Post;
        end;
        dtmMainJ.getAllInv012Data;
        medItemId.Enabled := true;
        ChangeBtnStatus;
        setButtonCompetence;
        initialForm;
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD06.CheckCanDelete: Boolean;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT COUNT(1) COUNTS FROM INV004 WHERE GROUPID = ''%s'' ',
    [dsrInv012.DataSet.FieldByName( 'GROUPID' ).AsString] );
  dtmMain.adoComm.Open;
  Result := ( dtmMain.adoComm.FieldByName( 'COUNTS' ).AsInteger = 0 );
  dtmMain.adoComm.Close;
  if not Result then
  begin
    WarningMsg( '此群組已被指定至使用者帳號,不可刪除。' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD06.DeleteOtherData;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' DELETE FROM INV013 WHERE GROUPID = ''%s'' ',
    [dsrInv012.DataSet.FieldByName( 'GROUPID' ).AsString] );
  dtmMain.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

end.
