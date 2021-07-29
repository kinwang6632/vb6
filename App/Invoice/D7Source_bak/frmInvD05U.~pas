unit frmInvD05U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Buttons, Grids, DBGrids, Mask, ExtCtrls;

type
  TfrmInvD05 = class(TForm)
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
    dsrInv022: TDataSource;
    Label23: TLabel;
    cobIsSelfCreated: TComboBox;
    Stt_Show: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure dsrInv022DataChange(Sender: TObject; Field: TField);
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
  frmInvD05: TfrmInvD05;

implementation

uses frmMainU, dtmMainJU, dtmMainU;

{$R *.dfm}

procedure TfrmInvD05.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'D05', '業者服務種類資料維護' );
  dtmMainJ.getAllInv022Data;
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
end;

procedure TfrmInvD05.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvD05.ChangeBtnStatus;
begin
     with dtmMainJ.adoInv022Code do
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
          if dtmMainJ.adoInv022Code.RecordCount>0 then
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

procedure TfrmInvD05.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('D05',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

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

procedure TfrmInvD05.initialForm;
begin
    with dsrInv022.DataSet do
    begin
      FindField('ItemId').DisplayLabel := '服務種類代碼';
      FindField('Description').DisplayLabel := '服務種類說明';
      FindField('UptTime').DisplayLabel := '異動時間';
      FindField('UptEn').DisplayLabel := '異動人員';

      FindField('Description').DisplayWidth := 30;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('IsSelfCreated').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;
    end;
end;

procedure TfrmInvD05.dsrInv022DataChange(Sender: TObject; Field: TField);
var
    sL_IsSelfCreated : String;
begin
    with dsrInv022.DataSet do
    begin
      if RecordCount > 0 then
      begin
        medItemId.Text := FieldByName('ItemId').AsString;
        medDesc.Text := FieldByName('Description').AsString;

        sL_IsSelfCreated := FieldByName('IsSelfCreated').AsString;
        if sL_IsSelfCreated = 'Y' then
          cobIsSelfCreated.ItemIndex := 0
        else
          cobIsSelfCreated.ItemIndex := 1;
      end
      else
      begin
        medItemId.Text := '';
        medDesc.Text := '';
        cobIsSelfCreated.ItemIndex := 0;
      end;
      Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);      
    end;
end;

procedure TfrmInvD05.btnAppendClick(Sender: TObject);
begin
    dtmMainJ.adoInv022Code.Append;

    ChangeBtnStatus;
    medItemId.SetFocus;
end;

procedure TfrmInvD05.btnEditClick(Sender: TObject);
begin
    dtmMainJ.adoInv022Code.Edit;
    ChangeBtnStatus;
    medItemId.Enabled := false;
    //medDesc.Enabled := false;
end;

procedure TfrmInvD05.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg('請確任要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMainJ.adoInv022Code.Delete;

      dtmMainJ.getAllInv022Data;
      ChangeBtnStatus;
    end;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD05.btnCancelClick(Sender: TObject);
begin
    if  dsrInv022.DataSet.State in [dsEdit] then
    begin
      medItemId.Enabled := true;
      //medDesc.Enabled := true;
    end;

    dtmMainJ.adoInv022Code.Cancel;
    dtmMainJ.adoInv022Code.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD05.btnSaveClick(Sender: TObject);
var
    sL_ItemId,sL_Desc,sL_SQL,sL_IsSelfCreated : String;
    bL_HavePKValue : Boolean;
begin
    if isDataOK then
    begin
      sL_ItemId := Trim(medItemId.Text);
      sL_Desc := Trim(medDesc.Text);

      if cobIsSelfCreated.ItemIndex = 0 then
        sL_IsSelfCreated := 'Y'
      else
        sL_IsSelfCreated := 'N';


      //檢核PK值
      sL_SQL := 'select * from inv022 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and ItemId=''' + sL_ItemId +
                ''' and Description=''' + sL_Desc + '''';

      if  dsrInv022.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;

      //bL_HavePKValue := dtmMainJ.checkCdsPK(dsrInv022.DataSet,'ItemId','Description',sL_ItemId,sL_Desc);

      if bL_HavePKValue then
      begin
        MessageDlg('違反唯一值條件',mtError,[mbOk],0);
        medDesc.SetFocus;
      end
      else
      begin
        with dsrInv022.DataSet do
        begin
          FieldByName('IdentifyId1').AsString := IDENTIFYID1;
          FieldByName('IdentifyId2').AsString := IDENTIFYID2;

          FieldByName('ItemId').AsString := sL_ItemId;
          FieldByName('Description').AsString := sL_Desc;

          FieldByName('IsSelfCreated').AsString := sL_IsSelfCreated;

          FieldByName('UptTime').AsString := DateTimeToStr(now);
          FieldByName('UptEn').AsString := dtmMain.getLoginUser;

          dsrInv022.DataSet.Post;
        end;
        dtmMainJ.getAllInv022Data;
        medItemId.Enabled := true;
        //medDesc.Enabled := true;
        ChangeBtnStatus;
        setButtonCompetence;
        initialForm;
      end;
    end;
end;

function TfrmInvD05.isDataOK: Boolean;
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

end.

