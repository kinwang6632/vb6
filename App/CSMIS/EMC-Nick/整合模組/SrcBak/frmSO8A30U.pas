unit frmSO8A30U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, ExtCtrls, StdCtrls, Buttons, DBCtrls, Mask,DB,
  DBActns, ActnList ,DBClient;

type
  TfrmSO8A30 = class(TForm)
    Groupbox1: TGroupBox;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCloseCancel: TBitBtn;
    btnSave: TBitBtn;
    btnNext: TBitBtn;
    gbxSo112: TGroupBox;
    dlcProductId: TDBLookupComboBox;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    dedType: TDBEdit;
    dedPrice: TDBEdit;
    Label7: TLabel;
    btnPrior: TBitBtn;
    btnLast: TBitBtn;
    BtnFirst: TBitBtn;
    dedRef_no: TDBEdit;
    Label4: TLabel;
    pnlSo112: TPanel;
    dgrSo112M: TDBGrid;
    pnlSo111: TPanel;
    dgrSo111D: TDBGrid;
    dsrSo113Cd039: TDataSource;
    dsrSo112M: TDataSource;
    dsrSo111D: TDataSource;
    DBTEXT1: TDbText;
    DBTEXT2: TDbText;
    DBTEXT3: TDbText;
    DBTEXT4: TDbText;
    DBTEXT5: TDbText;
    DBTEXT6: TDbText;
    DBTEXT7: TDbText;
    DBTEXT8: TDbText;
    DBTEXT9: TDbText;
    DBTEXT10: TDbText;
    dsrSo113Look: TDataSource;
    dsrSo110: TDataSource;
    dsrCd024CT: TDataSource;
    Label8: TLabel;
    Label10: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Panel1: TPanel;
    BitBtn2: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    btnDAppend: TBitBtn;
    btnDEdit: TBitBtn;
    btnDDelete: TBitBtn;
    DataSource1: TDataSource;
    dsrCd024Look: TDataSource;
    dsrSo110Look: TDataSource;
    Label1: TLabel;
    Label12: TLabel;
    dlcPRoviderId: TDBLookupComboBox;
    Label9: TLabel;
    Label11: TLabel;
    Label16: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBCheckBox1: TDBCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCloseCancelClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure dgrSo112MCellClick(Column: TColumn);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dlcProductIdCloseUp(Sender: TObject);
    procedure btnPriorClick(Sender: TObject);
    procedure BtnFirstClick(Sender: TObject);
    procedure btnLastClick(Sender: TObject);
    procedure CursorMoveState;
    procedure btn_Control(btn_First,btn_Prior,btn_Next,btn_Last,btn_Save,btn_Append,btn_Edit,btn_Delete,btn_CloseCancel:Boolean);
    procedure dedRef_noExit(Sender: TObject);
    procedure dedPriceExit(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure dgrSo111DEditButtonClick(Sender: TObject);
    procedure btnDAppendClick(Sender: TObject);
    procedure btnDEditClick(Sender: TObject);
    procedure btnDDeleteClick(Sender: TObject);
    procedure dgrSo111DDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dgrSo112MDblClick(Sender: TObject);
    procedure dlcFormulaIdCloseUp(Sender: TObject);
    procedure dlcPRoviderIdCloseUp(Sender: TObject);
    procedure dgrSo112MTitleClick(Column: TColumn);

  private
    { Private declarations }

  public
    { Public declarations }
  end;

var
  frmSO8A30: TfrmSO8A30;

implementation

uses frmSO8A301U, frmSO8A3011U, dtmMain2U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A30.FormCreate(Sender: TObject);
var sL_SQL:String;
begin
    //sG_Name:=Paramstr(1);
    //sG_Pwd:=Paramstr(2);
    with dtmMain2 do
    begin
      cdsSo112M.Open;
      cdsSo111D.Open;
      cdsCd024CT.Open;
      cdsSo113CT.Open;
      cdsSo113Look.Open;
      cdsSo110Look.Open;
      cdsCd024Look.Open;
    end;
    with dtmMain2.cdsSo111D do
    begin
        Close;
        sL_SQL:='select * from SO111 where PRODUCT_ID='''+dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString+'''';
        CommandText:=sL_SQL;
        //CommandText:='select * from SO111 where PRODUCT_ID=:PRODUCT_ID';
        //PARAMS.ParamByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
        OPen;
    end;

    if dtmMain2.cdsSo112M.RecordCount=0 then
    begin
      btnEdit.Enabled := false;
      btnDelete.Enabled := false;

      BtnFirst.Enabled := false;
      btnPrior.Enabled := false;
      btnNext.Enabled := false;
      btnLast.Enabled := false;
    end;
end;

procedure TfrmSO8A30.btnEditClick(Sender: TObject);
begin
  dtmMain2.cdsSo112M.Edit;
  gbxSo112.Enabled:=True;
  pnlSo112.Enabled:=False;
  dtmMain2.cdsSo112M.FieldByName('UseBaseFormula').AsString:='Y';
  DBCheckBox1.Checked:=True;
  dlcProductId.Enabled:=False;
  btn_Control(False,False,False,False,True,False,False,False,True);
  btnCloseCancel.Caption:='取消';
  pnlSo111.Enabled:=False;
  dlcPRoviderId.SetFocus;
end;

procedure TfrmSO8A30.btnAppendClick(Sender: TObject);
begin
    dtmMain2.cdsSo112M.Append;
    dtmMain2.cdsSo112M.FieldByName('UseBaseFormula').AsString:='Y';
    DBCheckBox1.Checked:=True;
    gbxSo112.Enabled:=True;
    pnlSo112.Enabled:=False;
    dlcProductId.Enabled:=True;
    btn_Control(False,False,False,False,True,False,False,False,True);
    btnCloseCancel.Caption:='取消';
    pnlSo111.Enabled:=False;
    dlcProductId.SetFocus;
end;

procedure TfrmSO8A30.btnSaveClick(Sender: TObject);
var sL_SQL:String;
begin
    if dtmMain2.cdsSo112M.State=dsBrowse then
       Exit;

    if Length(dtmMain2.cdsSo112M.FieldByName('REF_NO').AsString)>1 then
    begin
    MessageDlg('參考號最多1碼',mtError, [mbOK],0);
    dedRef_no.SetFocus;
    dtmMain2.cdsSo112M.FieldByName('REF_NO').AsString:='';
    Exit;
    end;
{JACKAL911129
    if dtmMain2.cdsSo112M.FieldByName('REF_NO').AsString='1' then
    begin
    dtmMain2.cdsSo112M.FieldByName('FORMULA_ID').AsString:='1';
    end;
}
    if dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString='' then
     begin
        MessageDlg('請輸入產品代碼',mtError, [mbOK],0);
        dlcProductId.SetFocus;
        Exit;
     end;



     if Length(dtmMain2.cdsSo112M.FieldByName('PRICE').AsString)>6 then
     begin
       MessageDlg('原定價最多6碼',mtError, [mbOK],0);
       dedPrice.SetFocus;
       Exit;
     end;

     if dtmMain2.cdsSo112M.FieldByName('UseBaseFormula').AsString='Y' then
     begin
       if (dtmMain2.cdsSo112M.FieldByName('Provider_Percent').AsInteger<=0) or
          (dtmMain2.cdsSo112M.FieldByName('Provider_Percent').AsInteger>100) then
       begin
         MessageDlg('頻道商拆帳百分比必須大於0且小於等於100',mtError, [mbOK],0);
         DBEdit1.SetFocus;
         Exit;
       end;
       if (dtmMain2.cdsSo112M.FieldByName('So_Percent').AsInteger<=0) or
          (dtmMain2.cdsSo112M.FieldByName('So_Percent').AsInteger>100) then
       begin
         MessageDlg('So拆帳百分比拆帳百分比必須大於0且小於等於100',mtError, [mbOK],0);
         DBEdit2.SetFocus;
         Exit;
       end;
     end;
     if dtmMain2.cdsSo112M.FieldByName('UseBaseFormula').AsString='N' then
     begin
       if (dtmMain2.cdsSo112M.FieldByName('Provider_Percent').AsInteger>100) then
         begin
           MessageDlg('頻道商拆帳百分比必須小於等於100',mtError, [mbOK],0);
           Exit;
         end;
       if (dtmMain2.cdsSo112M.FieldByName('So_Percent').AsInteger>100) then
         begin
           MessageDlg('So拆帳百分比拆帳百分比必須小於等於100',mtError, [mbOK],0);
           Exit;
         end;
     end;
     {
     if (dtmMain2.cdsSo112M.FieldByName('UseBaseFormula').AsString='N') and
         ((DbEdit1.Text<>'') or (DbEdit2.Text<>'')) then
     begin
       ShowMessage('請勾選啟用基本公式');
       Exit;
     end;
     }



    //INSERT PK CHECK
    if dtmMain2.cdsSo112M.State=dsInsert then
    begin
        with dtmMain2.cdssqlstatement do
        begin
        Close;
        sL_SQL:='Select count(*) as CNT from SO112 where PRODUCT_ID='''+COPY((dlcProductId.Text),1,2)+'''';
        CommandText:=sL_SQL;
        OPEN;
        if FieldByName('CNT').Value>0 then
            begin
            MessageDlg('產品代碼重覆',mtError, [mbOK],0);
            dlcProductId.SetFocus;
            btn_Control(False,False,False,False,True,False,False,False,True);
            EXIT;
            end
        else
            try
                with dtmMain2 do
                begin
                cdsSo112M.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
                cdsSo112M.FieldByName('UPDTIME').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
                cdsSo112M.ApplyUpdates(-1);
                end;
            except
                begin
                MessageDlg('存檔失敗',mtError, [mbOK],0);
                Exit;
                end;
            end;
        end;
    end;
    if dtmMain2.cdsSo112M.State=dsEdit then
    begin
    try
        dtmMain2.cdsSo112M.FieldByName('OPERATOR').AsString:= frmMainMenu.sG_User;
        dtmMain2.cdsSo112M.FieldByName('UPDTIME').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
        dtmMain2.saveToDB(dtmMain2.cdsSo112M);
        dtmMain2.saveToDB(dtmMain2.cdsSo111D);
    except
        begin
        MessageDlg('存檔失敗',mtError, [mbOK],0);
        Exit;
        end;
    end;
    end;
    CursorMoveState;
    btn_Control(True,True,True,True,False,True,True,True,True);
    btnCloseCancel.Caption:='(&X)結束';
    pnlSo111.Enabled:=True;
    pnlSo112.Enabled:=True;
    gbxSo112.Enabled:=False;
end;

procedure TfrmSO8A30.btnCloseCancelClick(Sender: TObject);
begin
    gbxSo112.Enabled:=False;
    if (btnCloseCancel.Caption='(&X)結束') then
    begin
      dtmMain2.cdsSo112M.Close;
      dtmMain2.cdsSo111D.Close;
      dtmMain2.cdsSo110Look.Close;
      dtmMain2.cdsCd024Look.Close;
      dtmMain2.cdsSo113Look.Close;
      frmSO8A30.ModalResult:=mrOK;
    end;
    if (btnCloseCancel.Caption='取消') then
    begin
      dtmMain2.cdsSo112M.Cancel;
      btn_Control(True,True,True,True,False,True,True,True,True);
      btnCloseCancel.Caption:='(&X)結束';
      pnlSo111.Enabled:=True;
      pnlSo112.Enabled:=True;
    end;

end;

procedure TfrmSO8A30.btnDeleteClick(Sender: TObject);
var sL_SQL:String;
begin
    if dtmMain2.cdsSo112M.IsEmpty then
    begin
    Exit;
    end;
    dlcProductId.Enabled:=False;
    with dtmMain2 do
    begin
    cdssqlStatement.Close;
    sL_SQL:='Select count(*) as CNT from SO111 where PRODUCT_ID='''+cdsSo112M.FieldByName('PRODUCT_ID').AsString+'''';
    cdssqlStatement.CommandText:=sL_SQL;
    //cdssqlStatement.CommandText:='Select count(*) as CNT from SO111 where PRODUCT_ID=:PRODUCT_ID';
    //cdssqlStatement.PARAMS.ParamByName('PRODUCT_ID').AsString:=cdsSo112M.FieldByName('PRODUCT_ID').AsString;
    cdssqlStatement.Open;
    end;

    if dtmMain2.cdssqlStatement.FieldByName('CNT').Value>0 then
    begin
     if MessageDlg('明細資料檔內尚有相關資料,是否一起刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
     begin
        with dtmMain2.cdsSqlDelete do
        begin
        Close;
        sL_SQL:='DELETE from SO111 where PRODUCT_ID='''+dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString+'''';
        CommandText:=sL_SQL;
        //CommandText:='DELETE from SO111 where PRODUCT_ID=:PRODUCT_ID';
        //PARAMS.ParamByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
        Execute;
        end;

       //dtmMain2.getCdsSqlDelete_So111Result;
        try
            dtmMain2.cdsSo111D.ApplyUpdates(-1);
            dtmMain2.cdsSo111D.Refresh;
            dtmMain2.cdsSo112M.Delete;
            dtmMain2.cdsSo112M.ApplyUpdates(-1);
        except
            begin
            MessageDlg('刪除失敗',mtError, [mbOK],0);
            Exit;
            end;
        end;
     end;
    end;

    if dtmMain2.cdssqlStatement.FieldByName('CNT').Value=0 then
     begin
        if MessageDlg('確定將此筆資料刪除嗎?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
            dtmMain2.cdsSo112M.Delete;
            dtmMain2.cdsSo112M.ApplyUpdates(-1);
        end;
     end;


    dtmMain2.cdsSo112M.Open;
    if dtmMain2.cdsSo112M.RecordCount=0 then
    begin
      btnEdit.Enabled := false;
      btnDelete.Enabled := false;

      BtnFirst.Enabled := false;
      btnPrior.Enabled := false;
      btnNext.Enabled := false;
      btnLast.Enabled := false;
    end;
CursorMoveState;
end;

procedure TfrmSO8A30.btnNextClick(Sender: TObject);
begin
dtmMain2.cdsSo112M.Next;
CursorMoveState;

end;


procedure TfrmSO8A30.dgrSo112MCellClick(Column: TColumn);
begin
CursorMoveState;
end;

procedure TfrmSO8A30.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    with dtmMain2 do
    begin
    cdsSo112M.Close;
    cdsSo111D.Close;
    cdsCd024CT.Close;
    cdsSo113CT.Close;
    END;
end;

procedure TfrmSO8A30.dlcProductIdCloseUp(Sender: TObject);
begin
    dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsCd024Look.FieLdByName('CODENO').AsString;
    dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsCd024Look.FieLdByName('DESCRIPTION').AsString;
end;

procedure TfrmSO8A30.btnPriorClick(Sender: TObject);
begin
dtmMain2.cdsSo112M.Prior;
CursorMoveState;
end;

procedure TfrmSO8A30.BtnFirstClick(Sender: TObject);
begin
dtmMain2.cdsSo112M.First;
CursorMoveState;
end;

procedure TfrmSO8A30.btnLastClick(Sender: TObject);
begin
dtmMain2.cdsSo112M.Last;
CursorMoveState;
end;

procedure TfrmSO8A30.CursorMoveState;
var sL_SQL:String;
begin
    with dtmMain2 do
    begin
    cdsSo111D.Close;
    sL_SQL:='Select * from SO111 where PRODUCT_ID='''+cdsSo112M.FieldbyName('PRODUCT_ID').AsString+'''';
    cdsSo111D.CommandText:=sL_SQL;
    //cdsSo111D.CommandText:='Select * from SO111 where PRODUCT_ID=:PRODUCT_ID';
    //cdsSo111D.Params.ParamByName('PRODUCT_ID').AsString:=cdsSo112M.FieldbyName('PRODUCT_ID').AsString;
    cdsSo111D.Open;
    end;
end;

procedure TfrmSO8A30.btn_Control(btn_First,btn_Prior,btn_Next,btn_Last,btn_Save,btn_Append, btn_Edit,btn_Delete,
  btn_CloseCancel: Boolean);
begin
btnFirst.Enabled:=btn_First;
btnPrior.Enabled:=btn_Prior;
btnNext.Enabled:=btn_Next;
btnLast.Enabled:=btn_Last;
btnSave.Enabled:=btn_Save;
btnAppend.Enabled:=btn_Append;
btnEdit.Enabled:=btn_Edit;
btnDelete.Enabled:=btn_Delete;
btnCloseCancel.Enabled:=btn_CloseCancel;

end;

procedure TfrmSO8A30.dedRef_noExit(Sender: TObject);
begin
    {if TDBEdit(Sender).Text='' then
    begin
        ShowMessage('所有欄位都必須有值');
        TDBEdit(Sender).SetFocus;
        EXIT;
    end;}


end;

procedure TfrmSO8A30.dedPriceExit(Sender: TObject);
begin
 if Length(dedPrice.Text)>6 then
 begin
 MessageDlg('原定價最多6碼',mtError, [mbOK],0);
 dedPrice.SetFocus;
 dtmMain2.cdsSo112M.FieldByName('PRICE').AsString:='';
 Exit;
 end;
 {if TDBEdit(Sender).Text='' then
 begin
 ShowMessage('所有欄位都必須有值');
 TDBEdit(Sender).SetFocus;
 EXIT;
 end;}
end;

procedure TfrmSO8A30.BitBtn2Click(Sender: TObject);
begin
    dtmMain2.cdsSo111D.First;
end;

procedure TfrmSO8A30.BitBtn9Click(Sender: TObject);
begin
    dtmMain2.cdsSo111D.Prior;
end;

procedure TfrmSO8A30.BitBtn10Click(Sender: TObject);
begin
    dtmMain2.cdsSo111D.Next;
end;

procedure TfrmSO8A30.BitBtn11Click(Sender: TObject);
begin
    dtmMain2.cdsSo111D.Last;
end;

procedure TfrmSO8A30.dgrSo111DEditButtonClick(Sender: TObject);
{var
sL_SelectedFieldName : String;}
begin
    //inherited;
    {if (dsrSo111D.DataSet.State=dsBrowse) and (dsrSo111D.DataSet.RecordCount =0) then
      dsrSo111D.DataSet.Append;

    sL_SelectedFieldName := UpperCase(dgrSo111D.SelectedField.FieldName);

    if sL_SelectedFieldName='SO_PROVIDER_DESC' then //選取部門
    begin
    if SelectRecord('請點選公司別/供應商', dsrSo113Cd039.DataSet, 'CLASSIFY_ID;CODE;DESC','') = mrOk then
    begin
    if not (dsrSo111D.DataSet.State in [dsEdit, dsInsert]) then
    dsrSo111D.DataSet.Edit;
    dsrSo111D.DataSet.FieldByName('SO_PROVIDER_DESC').AsString :=dsrSo113Cd039.DataSet.FieldByName('DESC').AsString;
    dsrSo111D.DataSet.FieldByName('COMP_TYPE').AsString :=dsrSo113Cd039.DataSet.FieldByName('CLASSIFY_ID').AsString;
    dsrSo111D.DataSet.FieldByName('So_Provider_ID').AsString :=dsrSo113Cd039.DataSet.FieldByName('CODE').AsString;
    end;
    end;}
end;



procedure TfrmSO8A30.btnDAppendClick(Sender: TObject);
begin
    if dtmMain2.cdsSo112M.IsEmpty then
    begin
      MessageDlg('主檔內無資料,無法新增?',mtError, [mbOK],0);
      Exit;
    end;
    bPanelControlG:=True;
    Application.CreateForm(TfrmSO8A301,frmSO8A301);
    frmSO8A301.ShowModal;

end;

procedure TfrmSO8A30.btnDEditClick(Sender: TObject);
begin
    if dtmMain2.cdsSo112M.IsEmpty then
    begin
      MessageDlg('主檔內無資料,無法修改',mtError, [mbOK],0);
      Exit;
    end;
    if dtmMain2.cdsSo111D.IsEmpty then
    begin
      if MessageDlg('目前無資料,是否新增',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        btnDAppendClick(Sender);
      end;
    end
    else
    begin
    bPanelControlG:=False;
    Application.CreateForm(TfrmSO8A301,frmSO8A301);
    frmSO8A301.ShowModal;

    end;
end;

procedure TfrmSO8A30.btnDDeleteClick(Sender: TObject);
begin
    if dtmMain2.cdsSo111D.IsEmpty then
    begin
    Exit;
    end;
    dtmMain2.cdsSo111D.Delete;
    if MessageDlg('您確定要刪除此筆資料?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      dtmMain2.cdsSo111D.ApplyUpdates(-1);
    frmSO8A30.CursorMoveState;
end;


procedure TfrmSO8A30.dgrSo111DDblClick(Sender: TObject);
begin
    if dtmMain2.cdsSo111D.IsEmpty then
    begin
      if MessageDlg('目前無資料,是否新增?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        btnDAppendClick(Sender);
      end;
    end
    else btnDEditClick(Sender);
end;

procedure TfrmSO8A30.FormShow(Sender: TObject);
begin
  self.Caption := frmMainMenu.setCaption('SO8A30','[拆帳公式設定]');
end;

procedure TfrmSO8A30.dgrSo112MDblClick(Sender: TObject);
begin
    if dtmMain2.cdsSo112M.IsEmpty then
    begin
      if MessageDlg('目前無資料,是否新增?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        btnAppendClick(Sender);
      end;
    end
    else btnEditClick(Sender);
end;

procedure TfrmSO8A30.dlcFormulaIdCloseUp(Sender: TObject);
begin
{JACKAL911129
    dtmMain2.cdsSo112M.FieldByName('FORMULA_ID').AsString:=dtmMain2.cdsSo110Look.FieLdByName('FORMULA_ID').AsString;
    dtmMain2.cdsSo112M.FieldByName('FORMULA_NAME').AsString:=dtmMain2.cdsSo110Look.FieLdByName('FORMULA_NAME').AsString;
}    
end;

procedure TfrmSO8A30.dlcPRoviderIdCloseUp(Sender: TObject);
begin
    dtmMain2.cdsSo112M.FieldByName('PROVIDER_ID').AsString:=dtmMain2.cdssO113Look.FieLdByName('PROVIDER_ID').AsString;
    dtmMain2.cdsSo112M.FieldByName('PROVIDER_NAME').AsString:=dtmMain2.cdssO113Look.FieLdByName('PROVIDER_NAME').AsString;
end;



procedure TfrmSO8A30.dgrSo112MTitleClick(Column: TColumn);
var
   ii : Integer;
   bFlag : Boolean;
begin

  bFlag := False;
  //for ii:=0 to FDataSet.FieldCount-1 do
  for ii:=0 to dtmMain2.cdsSo112M.FieldCount-1 do
  begin
   if (dtmMain2.cdsSo112M.Fields[ii].FieldName=Column.FieldName) and
      (dtmMain2.cdsSo112M.Fields[ii].FieldKind=fkData) then
   begin
     bFlag := True;
     break;
   end;
  end;

  if not bFlag then
   Exit;//該FieldName不存在於FDataSet中
  if dtmMain2.cdsSo112M IS TClientDataSet then
  begin
   if TClientDataSet(dtmMain2.cdsSo112M).IndexName = 'TmpIndex' then
     TClientDataSet(dtmMain2.cdsSo112M).DeleteIndex('TmpIndex');
   TClientDataSet(dtmMain2.cdsSo112M).AddIndex('TmpIndex', Column.FieldName, [ixCaseInsensitive],'','',0);
   TClientDataSet(dtmMain2.cdsSo112M).IndexName := 'TmpIndex';
  end;

end;


end.
