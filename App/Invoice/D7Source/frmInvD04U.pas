unit frmInvD04U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, StdCtrls, Buttons, Grids, DBGrids, Mask, ExtCtrls;

type
  TfrmInvD04 = class(TForm)
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
    dsrInv006: TDataSource;
    Stt_Show: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure dsrInv006DataChange(Sender: TObject; Field: TField);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvD04: TfrmInvD04;

implementation

uses frmMainU, dtmMainJU, dtmMainU, cbUtilis;

{$R *.dfm}

procedure TfrmInvD04.ChangeBtnStatus;
begin
     with dtmMainJ.adoInv006Code do
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
          if dtmMainJ.adoInv006Code.RecordCount>0 then
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

procedure TfrmInvD04.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D04', '發票作廢原因資料維護' );
  dtmMainJ.getAllInv006Data;
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
end;

procedure TfrmInvD04.initialForm;
begin
    with dsrInv006.DataSet do
    begin
      FindField('ItemId').DisplayLabel := '品名代碼';
      FindField('Description').DisplayLabel := '品名';
      FindField('UptTime').DisplayLabel := '異動時間';
      FindField('UptEn').DisplayLabel := '異動人員';

      FindField('Description').DisplayWidth := 30;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;
    end;
end;

procedure TfrmInvD04.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
    sL_VerifyCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('D04',sL_QueryCompetence,sL_InsertCompetence,
          sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence,sL_VerifyCompetence);

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

procedure TfrmInvD04.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmInvD04.dsrInv006DataChange(Sender: TObject; Field: TField);
begin
    with dsrInv006.DataSet do
    begin
      if RecordCount > 0 then
      begin
        medItemId.Text := FieldByName('ItemId').AsString;
        medDesc.Text := FieldByName('Description').AsString;
        Stt_Show.Caption := inttostr(dsrInv006.DataSet.RecNo)+'　/　'+inttostr(dsrInv006.DataSet.recordcount);
      end
      else
      begin
        medItemId.Text := '';
        medDesc.Text := '';
      end;
    end;
end;

procedure TfrmInvD04.btnAppendClick(Sender: TObject);
begin
    dtmMainJ.adoInv006Code.Append;

    ChangeBtnStatus;
    medItemId.SetFocus;
end;

procedure TfrmInvD04.btnEditClick(Sender: TObject);
begin
    dtmMainJ.adoInv006Code.Edit;
    ChangeBtnStatus;
    medItemId.Enabled := false;
    //medDesc.Enabled := false;
end;

procedure TfrmInvD04.btnDeleteClick(Sender: TObject);
begin
  if ConfirmMsg( '確認要刪除此筆資料?' ) then
  begin
    dtmMainJ.adoInv006Code.Delete;
    dtmMainJ.getAllInv006Data;
    ChangeBtnStatus;
  end;
  setButtonCompetence;
  initialForm;
end;

procedure TfrmInvD04.btnCancelClick(Sender: TObject);
begin
    if  dsrInv006.DataSet.State in [dsEdit] then
    begin
      medItemId.Enabled := true;
      //medDesc.Enabled := true;
    end;

    dtmMainJ.adoInv006Code.Cancel;
    dtmMainJ.adoInv006Code.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

function TfrmInvD04.isDataOK: Boolean;
var
    sL_ItemId,sL_Description : String;
begin
    Result := true;
    sL_ItemId := Trim(medItemId.Text);
    sL_Description := Trim(medDesc.Text);

    if sL_ItemId = '' then
    begin
      MessageDlg('請輸入品名代碼',mtError,[mbOk],0);
      medItemId.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_Description = '' then
    begin
      MessageDlg('請輸入品名',mtError,[mbOk],0);
      medDesc.SetFocus;
      Result := false;
      Exit;
    end;
end;

procedure TfrmInvD04.btnSaveClick(Sender: TObject);
var
    sL_ItemId,sL_Desc,sL_SQL : String;
    bL_HavePKValue : Boolean;
begin
    if isDataOK then
    begin
      sL_ItemId := Trim(medItemId.Text);
      sL_Desc := Trim(medDesc.Text);


      if  dsrInv006.DataSet.State in [dsInsert] then
      begin
        bL_HavePKValue := false;
        //檢核PK值
        sL_SQL := 'select * from inv006 where IdentifyId1=''' + IDENTIFYID1 +
                  ''' and IdentifyId2=' + IDENTIFYID2 + ' and ItemId=''' + sL_ItemId + '''';
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL);
        if bL_HavePKValue then
        begin
          WarningMsg( '違反唯一值條件。' );
          medItemId.SetFocus;
          Exit;
        end;
        sL_SQL := 'select * from inv006 where IdentifyId1=''' + IDENTIFYID1 +
                  ''' and IdentifyId2=' + IDENTIFYID2 + ' and description=''' + sL_Desc + '''';
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL);
        if bL_HavePKValue then
        begin
          WarningMsg( '作廢品名不可重覆。' );
          medDesc.SetFocus;
          Exit;
        end;
      end;

      with dsrInv006.DataSet do
      begin
        FieldByName('IdentifyId1').AsString := IDENTIFYID1;
        FieldByName('IdentifyId2').AsString := IDENTIFYID2;

        FieldByName('ItemId').AsString := sL_ItemId;
        FieldByName('Description').AsString := sL_Desc;

        FieldByName('UptTime').AsString := DateTimeToStr(now);
        FieldByName('UptEn').AsString := dtmMain.getLoginUser;

        dsrInv006.DataSet.Post;
      end;
      dtmMainJ.getAllInv006Data;
      medItemId.Enabled := true;
      //medDesc.Enabled := true;
      ChangeBtnStatus;
      setButtonCompetence;
      initialForm;
    end;
end;

end.
