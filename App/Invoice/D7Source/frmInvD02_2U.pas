unit frmInvD02_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, StdCtrls, Buttons, Grids, DBGrids, Mask, ExtCtrls;

type
  TfrmInvD02_2 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    medTitleId: TMaskEdit;
    medTitleSName: TMaskEdit;
    medTitleName: TMaskEdit;
    medBusinessId: TMaskEdit;
    medMailAddr: TMaskEdit;
    medMZipCode: TMaskEdit;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrInv019: TDataSource;
    medInvAddr: TMaskEdit;
    Label1: TLabel;
    Memo1: TMemo;
    Stt_Show: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv019DataChange(Sender: TObject; Field: TField);
    procedure medBusinessIdExit(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure medTitleNameExit(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
    procedure setNoRecordCountForm;
  public
    { Public declarations }
    sG_CompID,sG_CustID: String;
    CustName, CustSName: String;
  end;

var
  frmInvD02_2: TfrmInvD02_2;

implementation

uses cbUtilis, frmMainU, dtmMainJU, dtmMainU, dtmSOU;

{$R *.dfm}

procedure TfrmInvD02_2.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D02_2', '客戶明細資料維護' );
  dtmMainJ.getInv019Data(sG_CompID,sG_CustID);
  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
  setNoRecordCountForm;
end;

procedure TfrmInvD02_2.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvD02_2.ChangeBtnStatus;
begin
     with dtmMainJ.adoInv019Code do
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
          if dtmMainJ.adoInv019Code.RecordCount>0 then
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

procedure TfrmInvD02_2.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
    sL_VerifyCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('D02',sL_QueryCompetence,sL_InsertCompetence,
        sL_DeleteCompetence,sL_UpdateCompetence,
        sL_ExecuteCompetence,sL_VerifyCompetence);

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

procedure TfrmInvD02_2.initialForm;
begin
    with dsrInv019.DataSet do
    begin
      FindField('TitleId').DisplayLabel := '抬頭代碼';
      FindField('TitleSName').DisplayLabel := '抬頭簡稱';
      FindField('BusinessId').DisplayLabel := '統一編號';
      FindField('UptTime').DisplayLabel := '異動時間';
      FindField('UptEn').DisplayLabel := '異動人員';


      FindField('TitleSName').DisplayWidth := 30;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('CompId').Visible := false;
      FindField('CustId').Visible := false;
      FindField('TitleName').Visible := false;
      FindField('MZipCode').Visible := false;
      FindField('MailAddr').Visible := false;
      FindField('InvAddr').Visible := false;
      FindField('Memo').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;      
    end;

end;

procedure TfrmInvD02_2.btnAppendClick(Sender: TObject);
var
    sL_MaxTitleID : String;
begin
    dtmMainJ.adoInv019Code.Append;

    //取得抬頭代碼
    sL_MaxTitleID := IntToStr(dtmMainJ.adoInv019Code.RecordCount + 1);
    medTitleId.Text := sL_MaxTitleID;
    medTitleName.Text := CustName;
    medTitleSName.Text := CustSName;


    ChangeBtnStatus;
    medTitleSName.SetFocus;

end;

procedure TfrmInvD02_2.btnEditClick(Sender: TObject);
begin
    dtmMainJ.adoInv019Code.Edit;
    if Trim( medTitleName.Text ) = EmptyStr then
      medTitleName.Text := CustName;
    if Trim( medTitleSName.Text ) = EmptyStr then
      medTitleSName.Text := CustSName;
    ChangeBtnStatus;
    medTitleSName.SetFocus;
end;

procedure TfrmInvD02_2.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg('請確認要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMainJ.adoInv019Code.Delete;

      ChangeBtnStatus;
    end;
    setButtonCompetence;
    initialForm;
    setNoRecordCountForm;
end;

procedure TfrmInvD02_2.btnCancelClick(Sender: TObject);
begin
    dtmMainJ.adoInv019Code.Cancel;
    dtmMainJ.adoInv019Code.First;
    ChangeBtnStatus;
    setButtonCompetence;
    setNoRecordCountForm;
    initialForm;
end;

procedure TfrmInvD02_2.dsrInv019DataChange(Sender: TObject; Field: TField);
begin
    with dsrInv019.DataSet do
    begin
      if RecordCount > 0 then
      begin
        medTitleId.Text := FieldByName('TitleId').AsString;
        medTitleSName.Text := FieldByName('TitleSName').AsString;
        medTitleName.Text := FieldByName('TitleName').AsString;
        medBusinessId.Text := FieldByName('BusinessId').AsString;
        medMZipCode.Text := FieldByName('MZipCode').AsString;
        medMailAddr.Text := FieldByName('MailAddr').AsString;
        medInvAddr.Text := FieldByName('InvAddr').AsString;
        Memo1.Text := FieldByName('Memo').AsString;
      end
      else
      begin
        medTitleId.Text := '';
        medTitleSName.Text := '';
        medTitleName.Text := '';
        medBusinessId.Text := '';
        medMZipCode.Text := '';
        medMailAddr.Text := '';
        medInvAddr.Text := '';
        Memo1.Text := '';
      end;
      Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);
    end;
end;

procedure TfrmInvD02_2.medBusinessIdExit(Sender: TObject);
var
    sL_BusinessId : String;
    nL_BusinessIdLen : Integer;
begin
    sL_BusinessId := Trim(medBusinessId.Text);

    if sL_BusinessId <> '' then
    begin
      nL_BusinessIdLen := Length(sL_BusinessId);
      if nL_BusinessIdLen = 8 then
      begin
        if not dtmMain.checkBusinessID(sL_BusinessId) then
        begin
          MessageDlg('統一編號錯誤',mtError,[mbOk],0);
          medBusinessId.SetFocus;
        end;
      end
      else
      begin
        MessageDlg('統一編號錯誤',mtError,[mbOk],0);
        medBusinessId.SetFocus;
      end;
    end;
end;

function TfrmInvD02_2.isDataOK: Boolean;
var
    sL_BusinessId,sL_TitleId : String;
    nL_BusinessIdLen : Integer;
begin
    Result := true;

    sL_TitleId := Trim(medTitleId.Text);
    if sL_TitleId = '' then
    begin
      MessageDlg('請輸入抬頭代碼',mtError,[mbOk],0);
      Result := false;
      Exit;
    end;


    sL_BusinessId := Trim(medBusinessId.Text);

    if sL_BusinessId <> '' then
    begin
      nL_BusinessIdLen := Length(sL_BusinessId);
      if nL_BusinessIdLen = 8 then
      begin
        if not dtmMain.checkBusinessID(sL_BusinessId) then
        begin
          MessageDlg('統一編號錯誤',mtError,[mbOk],0);
          medBusinessId.SetFocus;
          Result := false;
          Exit;
        end;
      end
      else
      begin
        MessageDlg('統一編號錯誤',mtError,[mbOk],0);
        medBusinessId.SetFocus;
        Result := false;
        Exit;
      end;
    end;

end;

procedure TfrmInvD02_2.btnSaveClick(Sender: TObject);
var
    sL_TitleId,sL_TitleSName,sL_TitleName,sL_BusinessId,sL_MZipCode : String;
    sL_MailAddr,sL_InvAddr,sL_Memo,sL_SQL : String;
    bL_HavePKValue : Boolean;
begin
    if isDataOK then
    begin
      sL_TitleId := Trim(medTitleId.Text);
      sL_TitleSName := Trim(medTitleSName.Text);
      sL_TitleName := Trim(medTitleName.Text);

      sL_BusinessId := Trim(medBusinessId.Text);
      sL_MZipCode := Trim(medMZipCode.Text);
      sL_MailAddr := Trim(medMailAddr.Text);
      sL_InvAddr := Trim(medInvAddr.Text);
      sL_Memo := Trim(Memo1.Text);


      //檢核PK值
      sL_SQL := 'select * from inv019 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and CompID=''' + sG_CompID +
                ''' and CustId=''' + sG_CustID + ''' and TitleId=''' + sL_TitleId + '''';


      if  dsrInv019.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;


      if bL_HavePKValue then
      begin
        MessageDlg('違反唯一值條件',mtError,[mbOk],0);
      end
      else
      begin
        with dsrInv019.DataSet do
        begin
          FieldByName('IdentifyId1').AsString := IDENTIFYID1;
          FieldByName('IdentifyId2').AsString := IDENTIFYID2;
          FieldByName('CompID').AsString := sG_CompID;
          FieldByName('CustId').AsString := sG_CustId;
          FieldByName('TitleId').AsString := sL_TitleId;
          FieldByName('TitleSName').AsString := sL_TitleSName;
          FieldByName('TitleName').AsString := sL_TitleName;
          FieldByName('BusinessId').AsString := sL_BusinessId;
          FieldByName('MZipCode').AsString := sL_MZipCode;
          FieldByName('MailAddr').AsString := sL_MailAddr;
          FieldByName('InvAddr').AsString := sL_InvAddr;
          FieldByName('Memo').AsString := sL_Memo;
          FieldByName('UptTime').AsString := DateTimeToStr(now);
          FieldByName('UptEn').AsString := dtmMain.getLoginUser;
          dsrInv019.DataSet.Post;
        end;
        ChangeBtnStatus;
        setButtonCompetence;
        dtmMainJ.getInv019Data(sG_CompID,sG_CustID);
        initialForm;
      end;
    end;
end;

procedure TfrmInvD02_2.setNoRecordCountForm;
begin
    if dsrInv019.DataSet.RecordCount = 0 then
    begin
      medTitleId.Text := '';
      medTitleSName.Text := '';
      medTitleName.Text := '';
      medBusinessId.Text := '';
      medMZipCode.Text := '';
      medMailAddr.Text := '';
      medInvAddr.Text := '';
      Memo1.Text := '';

      btnDelete.Enabled := false;
      btnEdit.Enabled := false;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD02_2.medTitleNameExit(Sender: TObject);
begin
  medTitleSName.Text := medTitleName.Text;
end;

{ ---------------------------------------------------------------------------- }

end.
