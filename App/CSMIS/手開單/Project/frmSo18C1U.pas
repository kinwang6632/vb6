unit frmSo18C1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids, DB, fraChineseYMDU,
  Mask, DBCtrls, DBClient;

type
  TfrmSo18C1 = class(TForm)
    pnlCommand: TPanel;
    pnlMultiData: TPanel;
    Splitter1: TSplitter;
    pnlSingleData: TPanel;
    dsrSo126: TDataSource;
    DBGrid1: TDBGrid;
    pnlQuery: TPanel;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    btnQuery: TBitBtn;
    fraChineseYMD1: TfraChineseYMD;
    edtPaperNum: TEdit;
    StaticText4: TStaticText;
    fraChineseYMD2: TfraChineseYMD;
    lbxEmp: TListBox;
    StaticText5: TStaticText;
    fraChineseYMD3: TfraChineseYMD;
    StaticText6: TStaticText;
    StaticText7: TStaticText;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    StaticText8: TStaticText;
    StaticText9: TStaticText;
    DBEdit4: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancelExit: TBitBtn;
    btnSave: TBitBtn;
    DBEdit3: TDBEdit;
    btnGetPaperUser: TButton;
    btnEmp: TBitBtn;
    StaticText10: TStaticText;
    DBEdit5: TDBEdit;
    StaticText11: TStaticText;
    edtPrefix: TEdit;
    StaticText12: TStaticText;
    StaticText13: TStaticText;
    btnReUse: TBitBtn;
    StaticText14: TStaticText;
    fraChineseYMD4: TfraChineseYMD;
    StaticText15: TStaticText;
    fraChineseYMD5: TfraChineseYMD;
    StaticText16: TStaticText;
    DBMemo1: TDBMemo;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelExitClick(Sender: TObject);
    procedure dsrSo126DataChange(Sender: TObject; Field: TField);
    procedure btnGetPaperUserClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure fraChineseYMD3mseYMDExit(Sender: TObject);
    procedure btnEmpClick(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure btnReUseClick(Sender: TObject);
    procedure fraChineseYMD4mseYMDExit(Sender: TObject);
    procedure fraChineseYMD5mseYMDExit(Sender: TObject);

  private
    { Private declarations }
    sG_EmpSQL : String;
    function IsDataOK:boolean;
    procedure ChangeBtnState;    
  public
    { Public declarations }
  end;

var
  frmSo18C1: TfrmSo18C1;

implementation

uses dtmMainU, frmSo18C0U, frmDbMultiSelectu, Ustru, frmDbSelectu,
  UdateTimeu, frmSo18C6U;

{$R *.dfm}

procedure TfrmSo18C1.FormShow(Sender: TObject);
begin
    sG_EmpSQL := '';
    self.Caption := frmSo18C0.getCaption('SO8C10','手開單據管理-領單作業');
    dtmMain.activeDataSet(dtmMain.cdsSo126);    
    ChangeBtnState;
    fraChineseYMD1.mseYMD.SetFocus;
end;

procedure TfrmSo18C1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

procedure TfrmSo18C1.ChangeBtnState;
begin
     with dsrSo126.DataSet do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancelExit.Caption := '取消';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnQuery.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          btnReUse.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
          pnlQuery.Enabled := FALSE;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
            btnReUse.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
            btnReUse.Enabled := FALSE;            
          end;
          btnCancelExit.Caption := '結束';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := FALSE;
          btnQuery.Enabled := TRUE;
          btnAppend.Enabled := TRUE;
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
          pnlQuery.Enabled := TRUE;
       end;
     end;
end;

procedure TfrmSo18C1.btnAppendClick(Sender: TObject);
begin
    if dsrSo126.DataSet.State = dsInactive then
      dsrSo126.DataSet.Active := True;
    dsrSo126.DataSet.Append;
    ChangeBtnState;    
    fraChineseYMD3.setYMD('');
    fraChineseYMD4.setYMD('');
    fraChineseYMD5.setYMD('');
    fraChineseYMD3.mseYMD.SetFocus;

end;

procedure TfrmSo18C1.btnEditClick(Sender: TObject);
var
    nL_Seq : Integer;
begin
    if dsrSo126.DataSet.State = dsInactive then Exit;
    if dsrSo126.DataSet.RecordCount =0 then Exit;
    nL_Seq := dsrSo126.DataSet.fieldByName('SEQ').AsInteger;
    if dtmMain.canModifyGetPaperData(nL_Seq) then
    begin
      dsrSo126.DataSet.Edit;
      ChangeBtnState;
      fraChineseYMD3.mseYMD.SetFocus;
    end
    else
      MessageDlg('此組單據號碼,已經有對應之回單資料,所以無法修改!', mtWarning, [mbOK],0); 
end;

procedure TfrmSo18C1.btnDeleteClick(Sender: TObject);
var
    nL_Seq : Integer;
begin
    if dsrSo126.DataSet.State = dsInactive then Exit;
    if dsrSo126.DataSet.RecordCount =0 then Exit;
    nL_Seq := dsrSo126.DataSet.fieldByName('SEQ').AsInteger;
    if dtmMain.canModifyGetPaperData(nL_Seq) then
    begin
      if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;
      dsrSo126.DataSet.Delete;
      dtmMain.saveToDB(dsrSo126.DataSet as TClientDataSet);      
      dtmMain.deletePaperData(nL_Seq);
      ChangeBtnState;
    end
    else
      MessageDlg('此組單據號碼,已經有對應之回單資料,所以無法刪除!', mtWarning, [mbOK],0);
end;

procedure TfrmSo18C1.btnSaveClick(Sender: TObject);
begin
    if dsrSo126.DataSet.State = dsInactive then Exit;
    if IsDataOK then
    begin
      with dsrSo126.DataSet do
      begin
        FieldByName('GETPAPERDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD3.getYMD);
        if Trim(TUstr.replaceStr(fraChineseYMD4.getYMD,'/',''))<>'' then
        begin
          FieldByName('RETURNDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD4.getYMD);
        end else
        begin
          FieldByName('RETURNDATE').Value := Null;
        end;
        if Trim(TUstr.replaceStr(fraChineseYMD5.getYMD,'/',''))<>'' then
        begin
          FieldByName('CLEARDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD5.getYMD);
        end else
        begin
          FieldByName('CLEARDATE').Value := Null;
        end;
        Post;
        dtmMain.saveToDB(dsrSo126.DataSet as TClientDataSet);
      end;
      ChangeBtnState;
    end;



end;

procedure TfrmSo18C1.btnCancelExitClick(Sender: TObject);
begin
    if dsrSo126.DataSet.State in [dsEdit, dsInsert ] then
    begin
      dsrSo126.DataSet.Cancel;
      ChangeBtnState;
    end
    else if dsrSo126.DataSet.State in [dsInactive, dsBrowse] then
      Close;


end;

procedure TfrmSo18C1.dsrSo126DataChange(Sender: TObject; Field: TField);
var
    dL_TmpDate : TDate; 
begin
    if dsrSo126.DataSet.State = dsBrowse then
    begin
      if dsrSo126.DataSet.FieldByName('GETPAPERDATE').Value <> null then
      begin
        fraChineseYMD3.setYMD(TUdateTime.CDateStr(dsrSo126.DataSet.FieldByName('GETPAPERDATE').AsDateTime,9));
      end;
      if dsrSo126.DataSet.FieldByName('RETURNDATE').Value <> null then
      begin
        fraChineseYMD4.setYMD(TUdateTime.CDateStr(dsrSo126.DataSet.FieldByName('RETURNDATE').AsDateTime,9));
      end;
      if dsrSo126.DataSet.FieldByName('CLEARDATE').Value <> null then
      begin
        fraChineseYMD5.setYMD(TUdateTime.CDateStr(dsrSo126.DataSet.FieldByName('CLEARDATE').AsDateTime,9));
      end;
    end;
end;

function TfrmSo18C1.IsDataOK: boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    if Trim(TUstr.replaceStr(fraChineseYMD3.getYMD,'/',''))='' then
    begin
      MessageDlg('請輸入領單日期!', mtWarning, [mbOK],0);
      fraChineseYMD3.mseYMD.SetFocus;
      bL_Result := false;
    end
    else if dsrSo126.DataSet.FieldByName('EMPNO').AsString='' then
    begin
      MessageDlg('請點選領單人員!', mtWarning, [mbOK],0);
      btnGetPaperUser.SetFocus;
      bL_Result := false;
    end
    else if dsrSo126.DataSet.FieldByName('BEGINNUM').AsString='' then
    begin
      MessageDlg('請輸入起始單號!', mtWarning, [mbOK],0);
      dsrSo126.DataSet.FieldByName('BEGINNUM').FocusControl;
      bL_Result := false;
    end
    else if dsrSo126.DataSet.FieldByName('TOTALPAPERCOUNT').AsString='' then
    begin
      MessageDlg('請輸入張數!', mtWarning, [mbOK],0);
      dsrSo126.DataSet.FieldByName('TOTALPAPERCOUNT').FocusControl;
      bL_Result := false;
    end;
    result := bL_Result;
end;

procedure TfrmSo18C1.btnGetPaperUserClick(Sender: TObject);
begin
    dtmMain.activeDataSet(dtmMain.qryCm003);
    dtmMain.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
    dtmMain.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

    if SelectRecord('請點選領單人員', dtmMain.qryCm003, 'EMPNO;EMPNAME') = mrOk then
    begin
      if not (dsrSo126.DataSet.State in [dsEdit, dsInsert]) then
        dsrSo126.DataSet.Edit;
      dsrSo126.DataSet.FieldByName('EMPNO').AsString := dtmMain.qryCm003.FieldByName('EMPNO').AsString;
      dsrSo126.DataSet.FieldByName('EMPNAME').AsString := dtmMain.qryCm003.FieldByName('EMPNAME').AsString;
    end;
end;

procedure TfrmSo18C1.btnQueryClick(Sender: TObject);
var
    sL_Prefix, sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_PaperNum:String;
begin
    sL_BeginDate := trim(TUstr.replaceStr(fraChineseYMD1.getYMD,'/',''));
    sL_EndDate := trim(TUstr.replaceStr(fraChineseYMD2.getYMD,'/',''));
    if ((sL_BeginDate='') and (sL_EndDate<>'')) or
       ((sL_BeginDate<>'') and (sL_EndDate=''))then
    begin
      MessageDlg('請輸入領單之起始與截止日期!', mtWarning, [mbOK],0);
      Exit;
    end;


    sL_EmpSQL := sG_EmpSQL;
    sL_PaperNum := Trim(edtPaperNum.Text);
    sL_Prefix := Trim(edtPrefix.text);
    if sL_PaperNum <>'' then
      sL_PaperNum := TUstr.AddString(sL_PaperNum,'0',false,PAPER_NUMBER_LENGTH);
    if TClientDataSet(dsrSo126.DataSet).IndexName='TmpIndex' then
      TClientDataSet(dsrSo126.DataSet).DeleteIndex('TmpIndex');

    if not dtmMain.activeGetPaperData(sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_PaperNum, sL_Prefix) then
    begin
      fraChineseYMD3.setYMD('');
      fraChineseYMD4.setYMD('');
      fraChineseYMD5.setYMD('');
      
      ChangeBtnState;
      fraChineseYMD1.mseYMD.SetFocus;
      MessageDlg('查無資料!', mtInformation, [mbOK],0);
      Exit;
    end
    else
      ChangeBtnState;
end;

procedure TfrmSo18C1.fraChineseYMD3mseYMDExit(Sender: TObject);
begin
    fraChineseYMD3.mseYMDExit(Sender);

    if trim(TUstr.replaceStr(fraChineseYMD3.getYMD,'/',''))<>'' then
      dsrSo126.DataSet.FieldByName('GETPAPERDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD3.getYMD);
end;

procedure TfrmSo18C1.btnEmpClick(Sender: TObject);
var
    ii : Integer;
    L_TargetFieldNamesStrList : TStringList;
    L_TargetFieldValueStrList : TStringList;

    bL_SelectedAllData : boolean;
    sL_TargetSQL, sL_Tmp : String;
    sL_Code, sL_Desc, sG_EmpsCode, sG_EmpsName : String;
    L_TempStrList : TStringList;
begin

  L_TargetFieldNamesStrList:=TStringList.Create;
  L_TargetFieldNamesStrList.Add('EMPNO');
  L_TargetFieldNamesStrList.Add('EMPNAME');

  L_TargetFieldValueStrList:=TStringList.Create;
  bL_SelectedAllData := false;


  dtmMain.activeDataSet(dtmMain.qryCm003);
  dtmMain.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
  dtmMain.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

  if SelectMultiRecords('請點選員工', dtmMain.qryCm003, 'EMPNO;EMPNAME', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
  begin
    lbxEmp.Items.Clear;
    sG_EmpsCode := '';
    for ii:=0 to L_TargetFieldValueStrList.Count -1 do
    begin
      sL_Tmp := Trim(L_TargetFieldValueStrList.Strings[ii]);
      L_TempStrList := TUstr.ParseStrings(sL_Tmp,',');
      sL_Code := L_TempStrList.Strings[0];
      sL_Desc := L_TempStrList.Strings[1];
      if sG_EmpsCode='' then
      begin
        sG_EmpsCode := sL_Code;
        sG_EmpsName := sL_Desc;
      end
      else
      begin
        sG_EmpsCode := sG_EmpsCode + ',' + sL_Code;
        sG_EmpsName := sG_EmpsName + ',' + sL_Desc;
      end;
      lbxEmp.Items.Add(sL_Desc);
    end;
    sG_EmpSQL := sL_TargetSQL;

  end;

end;

procedure TfrmSo18C1.DBGrid1TitleClick(Column: TColumn);
var
   ii : Integer;
   bFlag : Boolean;
begin

     bFlag := False;
     for ii:=0 to dsrSo126.DataSet.FieldCount-1 do
     begin
       if (dsrSo126.DataSet.Fields[ii].FieldName=Column.FieldName) and
          (dsrSo126.DataSet.Fields[ii].FieldKind=fkData) then
       begin
         bFlag := True;
         break;
       end;
     end;

     if not bFlag then
       Exit;//該FieldName不存在於FDataSet中
     if dsrSo126.DataSet IS TClientDataSet then
     begin
       if TClientDataSet(dsrSo126.DataSet).IndexName = 'TmpIndex' then
         TClientDataSet(dsrSo126.DataSet).DeleteIndex('TmpIndex');
       TClientDataSet(dsrSo126.DataSet).AddIndex('TmpIndex', Column.FieldName, [ixCaseInsensitive],'','',0);
       TClientDataSet(dsrSo126.DataSet).IndexName := 'TmpIndex';
     end;
end;

procedure TfrmSo18C1.btnReUseClick(Sender: TObject);
var
    L_Frm : TfrmSo18C6;
begin
    if dsrSo126.DataSet.State in [dsEdit, dsInsert] then
    begin
      MessageDlg('目前資料並非處於瀏覽模式,無法執行此功能!', mtWarning, [mbOk],0);
      Exit;
    end
    else if dsrSo126.DataSet.State in [dsBrowse] then
    begin
      L_Frm := TfrmSo18C6.Create(Application);
      L_Frm.sG_BeginNum := dsrSo126.DataSet.FieldByName('BeginNum').AsString;
      L_Frm.sG_EndNum := dsrSo126.DataSet.FieldByName('EndNum').AsString;
      L_Frm.sG_Counts := dsrSo126.DataSet.FieldByName('TotalPaperCount').AsString;
      L_Frm.sG_Prefix := dsrSo126.DataSet.FieldByName('PREFIX').AsString;
      L_Frm.sG_OriginalSeq := dsrSo126.DataSet.FieldByName('SEQ').AsString;
      L_Frm.sG_CompCode := dsrSo126.DataSet.FieldByName('COMPCODE').AsString;
      L_Frm.dG_OriginalGetPaperDate := dsrSo126.DataSet.FieldByName('GETPAPERDATE').AsDateTime;
      L_Frm.sG_OriginalEmpNo := dsrSo126.DataSet.FieldByName('EMPNO').AsString;
      L_Frm.sG_OriginalEmpName := dsrSo126.DataSet.FieldByName('EMPNAME').AsString;
      L_Frm.sG_OriginalCount := dsrSo126.DataSet.FieldByName('TOTALPAPERCOUNT').AsString;
      
      L_Frm.ShowModal;
      L_Frm.Free;
      btnQueryClick(nil);

    end
    else if dsrSo126.DataSet.State in [dsInactive] then
    begin
      MessageDlg('請先查詢出相關資料,再執行此功能!', mtWarning, [mbOk],0);
      fraChineseYMD1.mseYMD.SetFocus;
      Exit;
    end;
end;

procedure TfrmSo18C1.fraChineseYMD4mseYMDExit(Sender: TObject);
begin
  fraChineseYMD4.mseYMDExit(Sender);
  if trim(TUstr.replaceStr(fraChineseYMD4.getYMD,'/',''))<>'' then
    dsrSo126.DataSet.FieldByName('RETURNDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD4.getYMD);
end;

procedure TfrmSo18C1.fraChineseYMD5mseYMDExit(Sender: TObject);
begin
  fraChineseYMD5.mseYMDExit(Sender);
  if trim(TUstr.replaceStr(fraChineseYMD5.getYMD,'/',''))<>'' then
    dsrSo126.DataSet.FieldByName('CLEARDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD5.getYMD);
end;

end.

