unit frmInvA10_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, StdCtrls, Buttons, Grids, DBGrids, Mask, ExtCtrls;

type
  TfrmInvA10_1 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    medIsLockedYear: TMaskEdit;
    cobIsLockedMonth: TComboBox;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrInv018: TDataSource;
    Label1: TLabel;
    rgpIsLocked: TRadioGroup;
    Stt_Show: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure dsrInv018DataChange(Sender: TObject; Field: TField);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
  public
    { Public declarations }
    sG_CompID,sG_QueryYear : String;
  end;

var
  frmInvA10_1: TfrmInvA10_1;

implementation

uses frmMainU, dtmMainJU, dtmMainU;

{$R *.dfm}

procedure TfrmInvA10_1.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.GetFormTitleString( 'A0A_1', '發票鎖帳作業' );

    dtmMainJ.getInv018Data(sG_CompID,sG_QueryYear);

    ChangeBtnStatus;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvA10_1.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvA10_1.ChangeBtnStatus;
begin
     with dsrInv018.DataSet do
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
          if dsrInv018.DataSet.RecordCount>0 then
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

procedure TfrmInvA10_1.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('A0A',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

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

procedure TfrmInvA10_1.initialForm;
begin
    with dsrInv018.DataSet do
    begin
      FindField('IsLocked').DisplayLabel := '是否鎖帳';
      FindField('YearMonth').DisplayLabel := '年月';
      FindField('UptTime').DisplayLabel := '異動時間';
      FindField('UptEn').DisplayLabel := '異動人員';


      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('CompID').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;
    end;
end;

procedure TfrmInvA10_1.btnAppendClick(Sender: TObject);
begin
    dtmMainJ.adoQryInv018.Append;

    cobIsLockedMonth.ItemIndex := 0;
    rgpIsLocked.ItemIndex := 0;

    ChangeBtnStatus;
    medIsLockedYear.SetFocus;
end;

procedure TfrmInvA10_1.btnEditClick(Sender: TObject);
begin
    dtmMainJ.adoQryInv018.Edit;
    ChangeBtnStatus;
    medIsLockedYear.Enabled := false;
    cobIsLockedMonth.Enabled := false;
end;

procedure TfrmInvA10_1.btnDeleteClick(Sender: TObject);
var
    sL_IsLockedYear,sL_IsLockedMonth : String;
begin
    if MessageDlg('請確任要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sL_IsLockedYear := Trim(medIsLockedYear.Text);
      sL_IsLockedMonth := Trim(cobIsLockedMonth.Text);
      sL_IsLockedMonth := Format('%.2d',[StrToInt(sL_IsLockedMonth)]);

      dtmMainJ.adoQryInv018.Delete;

      //Inv018刪除異動時,還會影響到Inv007及Inv099(當作不鎖帳)
      dtmMainJ.handleIsLockedData(sG_CompID,sL_IsLockedYear,sL_IsLockedMonth,'N');

      dtmMainJ.getInv018Data(sG_CompID,sG_QueryYear);
      ChangeBtnStatus;
    end;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvA10_1.btnCancelClick(Sender: TObject);
begin
    if  dsrInv018.DataSet.State in [dsEdit] then
    begin
      medIsLockedYear.Enabled := true;
      cobIsLockedMonth.Enabled := true;
    end;

    dtmMainJ.adoQryInv018.Cancel;
    dtmMainJ.adoQryInv018.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

function TfrmInvA10_1.isDataOK: Boolean;
var
    sL_IsLockedYear,sL_IsLockedMonth : String;
begin
    Result := true;
    sL_IsLockedYear := Trim(medIsLockedYear.Text);
    sL_IsLockedMonth := Trim(cobIsLockedMonth.Text);

    if sL_IsLockedYear = '' then
    begin
      MessageDlg('請輸入鎖帳年度',mtError,[mbOk],0);
      medIsLockedYear.SetFocus;
      Result := false;
      Exit;
    end;



    if sL_IsLockedMonth = '' then
    begin
      MessageDlg('請輸入鎖帳月份',mtError,[mbOk],0);
      cobIsLockedMonth.SetFocus;
      Result := false;
      Exit;
    end;
end;

procedure TfrmInvA10_1.btnSaveClick(Sender: TObject);
var
    sL_IsLockedYear,sL_IsLockedMonth,sL_IsLocked,sL_SQL,sL_YM : String;
    bL_HavePKValue : Boolean;
begin
    if isDataOK then
    begin
      //執行等待狀態
      dtmMain.setWaitingCursor;

      sL_IsLockedYear := Trim(medIsLockedYear.Text);
      sL_IsLockedMonth := Trim(cobIsLockedMonth.Text);

      if rgpIsLocked.ItemIndex = 0 then
        sL_IsLocked := 'Y'
      else
        sL_IsLocked := 'N';

     sL_IsLockedMonth := Format('%.2d',[StrToInt(sL_IsLockedMonth)]);
     sL_YM := sL_IsLockedYear + sL_IsLockedMonth;


      //檢核PK值
      sL_SQL := 'select * from inv018 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and CompId=''' + sG_CompID +
                ''' and YearMonth=''' + sL_YM + '''';

      if  dsrInv018.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;

      if bL_HavePKValue then
      begin
        MessageDlg('違反唯一值條件',mtError,[mbOk],0);
        cobIsLockedMonth.SetFocus;
      end
      else
      begin
        with dsrInv018.DataSet do
        begin
          FieldByName('IdentifyId1').AsString := IDENTIFYID1;
          FieldByName('IdentifyId2').AsString := IDENTIFYID2;

          FieldByName('CompID').AsString := sG_CompID;
          FieldByName('YearMonth').AsString := sL_YM;
          FieldByName('IsLocked').AsString := sL_IsLocked;

          FieldByName('UptTime').AsString := DateTimeToStr(now);
          FieldByName('UptEn').AsString := dtmMain.getLoginUserName;

          dsrInv018.DataSet.Post;
        end;

        //Inv018新增或異動時,還會影響到Inv007及Inv099
        dtmMainJ.handleIsLockedData(sG_CompID,sL_IsLockedYear,sL_IsLockedMonth,sL_IsLocked);
        
        dtmMainJ.getInv018Data(sG_CompID,sG_QueryYear);
        medIsLockedYear.Enabled := true;
        cobIsLockedMonth.Enabled := true;
        ChangeBtnStatus;
        setButtonCompetence;
        initialForm;

        //回復原狀態
        dtmMain.setDefaultCursor;
      end;
    end;

end;

procedure TfrmInvA10_1.dsrInv018DataChange(Sender: TObject; Field: TField);
var
    sL_YM,sL_Year,sL_Month,sL_IsLocked : String;
begin
    with dsrInv018.DataSet do
    begin
      if RecordCount > 0 then
      begin
        sL_YM := FieldByName('YearMonth').AsString;
        sL_Year := Copy(sL_YM,0,4);
        sL_Month := Copy(sL_YM,5,2);

        medIsLockedYear.Text := sL_Year;
        if sL_Month = '01' then
          cobIsLockedMonth.ItemIndex :=0
        else if sL_Month = '03' then
          cobIsLockedMonth.ItemIndex :=1
        else if sL_Month = '05' then
          cobIsLockedMonth.ItemIndex :=2
        else if sL_Month = '07' then
          cobIsLockedMonth.ItemIndex :=3
        else if sL_Month = '09' then
          cobIsLockedMonth.ItemIndex :=4
        else if sL_Month = '11' then
          cobIsLockedMonth.ItemIndex :=5;

        sL_IsLocked := FieldByName('IsLocked').AsString;
        if sL_IsLocked = 'Y' then
          rgpIsLocked.ItemIndex := 0
        else
          rgpIsLocked.ItemIndex := 1;
      end
      else
      begin
        medIsLockedYear.Text := '';
        cobIsLockedMonth.ItemIndex := 0;
        rgpIsLocked.ItemIndex := 0;
      end;
    end;

end;

end.
