unit frmSO8C10U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids, DB, fraChineseYMDU,
  Mask, DBCtrls, DBClient;

type
  TfrmSO8C10 = class(TForm)
    pnlMultiData: TPanel;
    pnlSingleData: TPanel;
    dsrSo126: TDataSource;
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
    DBEdit3: TDBEdit;
    btnGetPaperUser: TButton;
    btnEmp: TBitBtn;
    StaticText10: TStaticText;
    DBEdit5: TDBEdit;
    StaticText11: TStaticText;
    edtPrefix: TEdit;
    StaticText12: TStaticText;
    StaticText13: TStaticText;
    StaticText14: TStaticText;
    StaticText15: TStaticText;
    fraChineseYMD5: TfraChineseYMD;
    DBMemo1: TDBMemo;
    StaticText16: TStaticText;
    pnlCommand: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancelExit: TBitBtn;
    btnSave: TBitBtn;
    btnReUse: TBitBtn;
    btnPrint: TBitBtn;
    DBGrid1: TDBGrid;
    Splitter1: TSplitter;
    fraChineseYMD4: TfraChineseYMD;
    procedure FormShow(Sender: TObject);
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
    procedure btnPrintClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  private
    { Private declarations }
    sG_EmpSQL : String;
    sG_EmpsCode: String;
    sG_EmpsName: String;
    function IsDataOK:boolean;
    procedure ChangeBtnState;
    procedure ChangeEditorState;
  public
    { Public declarations }
  end;

var
  frmSO8C10: TfrmSO8C10;

implementation

uses dtmMain3U, frmDbMultiSelectu, Ustru, frmDbSelectu,
  UdateTimeu, frmMainMenuU, frmSo8C60U, rptSO8C10U, QuickRpt;

{$R *.dfm}

procedure TfrmSO8C10.FormShow(Sender: TObject);
begin
    sG_EmpSQL := '';
    self.Caption := frmMainMenu.setCaption('SO8C10','手開單據管理-領單作業');
    dtmMain3.activeDataSet(dtmMain3.cdsSo126);
    ChangeBtnState;
    fraChineseYMD1.mseYMD.SetFocus;
end;

procedure TfrmSO8C10.ChangeBtnState;
begin
     with dsrSo126.DataSet do
     begin
       if state in [dsInactive] then
       begin
         btnPrint.Enabled := False;
         Exit;
       end
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancelExit.Caption := '取消';
          btnCancelExit.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnQuery.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
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
            btnPrint.Enabled := True;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
            btnPrint.Enabled := False;
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

procedure TfrmSO8C10.btnAppendClick(Sender: TObject);
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

procedure TfrmSO8C10.btnEditClick(Sender: TObject);
var
    nL_Seq : Integer;
begin
    if dsrSo126.DataSet.State = dsInactive then Exit;
    if dsrSo126.DataSet.RecordCount =0 then Exit;
//    nL_Seq := dsrSo126.DataSet.fieldByName('SEQ').AsInteger;
//    if dtmMain3.canModifyGetPaperData(nL_Seq) then
//    begin
//      dsrSo126.DataSet.Edit;
//      ChangeBtnState;
//      fraChineseYMD3.mseYMD.SetFocus;
//    end
//    else
//      MessageDlg('此組單據號碼,已經有對應之回單資料,所以無法修改!', mtWarning, [mbOK],0);
    dsrSo126.DataSet.Edit;
    ChangeBtnState;
    ChangeEditorState;
    if ( fraChineseYMD4.mseYMD.CanFocus ) then
      fraChineseYMD4.mseYMD.SetFocus;
end;

procedure TfrmSO8C10.btnDeleteClick(Sender: TObject);
var
    nL_Seq : Integer;
begin
    if dsrSo126.DataSet.State = dsInactive then Exit;
    if dsrSo126.DataSet.RecordCount =0 then Exit;
    nL_Seq := dsrSo126.DataSet.fieldByName('SEQ').AsInteger;
    if dtmMain3.canModifyGetPaperData(nL_Seq) then
    begin
      if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;
      dsrSo126.DataSet.Delete;
      dtmMain3.saveToDB(dsrSo126.DataSet as TClientDataSet);      
      dtmMain3.deletePaperData(nL_Seq);
      ChangeBtnState;
    end
    else
      MessageDlg('此組單據號碼,已經有對應之回單資料,所以無法刪除!', mtWarning, [mbOK],0);
end;

procedure TfrmSO8C10.btnSaveClick(Sender: TObject);
begin
    if dsrSo126.DataSet.State = dsInactive then Exit;
    if IsDataOK then
    begin
      with dsrSo126.DataSet do
      begin
        FieldByName('GETPAPERDATE').AsDateTime := StrToDate( fraChineseYMD3.getYMD );
        if Trim(TUstr.replaceStr(fraChineseYMD4.getYMD,'/',''))<>'' then
        begin
          FieldByName('RETURNDATE').AsDateTime := StrToDate( fraChineseYMD4.getYMD );
        end else
        begin
          FieldByName('RETURNDATE').Value := Null;
        end;
        if Trim(TUstr.replaceStr(fraChineseYMD5.getYMD,'/',''))<>'' then
        begin
          FieldByName('CLEARDATE').AsDateTime := StrToDate( fraChineseYMD5.getYMD );
        end else
        begin
          FieldByName('CLEARDATE').Value := Null;
        end;
        Post;
        dtmMain3.saveToDB(dsrSo126.DataSet as TClientDataSet);
      end;
      ChangeBtnState;
      ChangeEditorState;
    end;
end;

procedure TfrmSO8C10.btnCancelExitClick(Sender: TObject);
begin
    if dsrSo126.DataSet.State in [dsEdit, dsInsert ] then
    begin
      dsrSo126.DataSet.Cancel;
      ChangeBtnState;
      ChangeEditorState;
    end
    else if dsrSo126.DataSet.State in [dsInactive, dsBrowse] then
      Close;
end;

procedure TfrmSO8C10.dsrSo126DataChange(Sender: TObject; Field: TField);
var
    dL_TmpDate : TDate; 
begin
    if dsrSo126.DataSet.State = dsBrowse then
    begin
      if not VarIsNull( dsrSo126.DataSet.FieldByName('GETPAPERDATE').Value ) then
      begin
        fraChineseYMD3.setYMD( FormatDateTime( 'yyyy/mm/dd',
          dsrSo126.DataSet.FieldByName('GETPAPERDATE').AsDateTime ) );
      end;
      if not VarIsNull( dsrSo126.DataSet.FieldByName('RETURNDATE').Value ) then
      begin
        fraChineseYMD4.setYMD( FormatDateTime( 'yyyy/mm/dd',
          dsrSo126.DataSet.FieldByName('RETURNDATE').AsDateTime ) );
      end;
      if not VarIsNull( dsrSo126.DataSet.FieldByName('CLEARDATE').Value ) then
      begin
        fraChineseYMD5.setYMD( FormatDateTime( 'yyyy/mm/dd',
          dsrSo126.DataSet.FieldByName('CLEARDATE').AsDateTime ) );
      end;
    end;
end;

function TfrmSO8C10.IsDataOK: boolean;
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

procedure TfrmSO8C10.btnGetPaperUserClick(Sender: TObject);
begin
    dtmMain3.activeDataSet(dtmMain3.qryCm003);
    dtmMain3.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
    dtmMain3.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

    if SelectRecord('請點選領單人員', dtmMain3.qryCm003, 'EMPNO;EMPNAME') = mrOk then
    begin
      if not (dsrSo126.DataSet.State in [dsEdit, dsInsert]) then
        dsrSo126.DataSet.Edit;
      dsrSo126.DataSet.FieldByName('EMPNO').AsString := dtmMain3.qryCm003.FieldByName('EMPNO').AsString;
      dsrSo126.DataSet.FieldByName('EMPNAME').AsString := dtmMain3.qryCm003.FieldByName('EMPNAME').AsString;
    end;
end;

procedure TfrmSO8C10.btnQueryClick(Sender: TObject);
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

    if ( sL_BeginDate <> EmptyStr ) then sL_BeginDate := TUstr.ToROCYMD( sL_BeginDate );
    if (sL_EndDate <> EmptyStr ) then sL_EndDate := TUstr.ToROCYMD( sL_EndDate );

    sL_EmpSQL := sG_EmpSQL;
    sL_PaperNum := Trim(edtPaperNum.Text);
    sL_Prefix := Trim(edtPrefix.text);
    if sL_PaperNum <>'' then
      sL_PaperNum := TUstr.AddString(sL_PaperNum,'0',false,PAPER_NUMBER_LENGTH);
    if TClientDataSet(dsrSo126.DataSet).IndexName='TmpIndex' then
      TClientDataSet(dsrSo126.DataSet).DeleteIndex('TmpIndex');

    if not dtmMain3.activeGetPaperData(sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_PaperNum, sL_Prefix) then
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

procedure TfrmSO8C10.fraChineseYMD3mseYMDExit(Sender: TObject);
begin
    fraChineseYMD3.mseYMDExit(Sender);

    if trim(TUstr.replaceStr(fraChineseYMD3.getYMD,'/',''))<>'' then
      dsrSo126.DataSet.FieldByName('GETPAPERDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD3.getYMD);
end;

procedure TfrmSO8C10.btnEmpClick(Sender: TObject);
var
    ii : Integer;
    L_TargetFieldNamesStrList : TStringList;
    L_TargetFieldValueStrList : TStringList;

    bL_SelectedAllData : boolean;
    sL_TargetSQL, sL_Tmp : String;
    sL_Code, sL_Desc: String;
    L_TempStrList : TStringList;
begin

  L_TargetFieldNamesStrList:=TStringList.Create;
  L_TargetFieldNamesStrList.Add('EMPNO');
  L_TargetFieldNamesStrList.Add('EMPNAME');

  L_TargetFieldValueStrList:=TStringList.Create;
  bL_SelectedAllData := false;


  dtmMain3.activeDataSet(dtmMain3.qryCm003);
  dtmMain3.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
  dtmMain3.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

  if SelectMultiRecords('請點選員工', dtmMain3.qryCm003, 'EMPNO;EMPNAME', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
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

procedure TfrmSO8C10.DBGrid1TitleClick(Column: TColumn);
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

procedure TfrmSO8C10.btnReUseClick(Sender: TObject);
var
    L_Frm : TfrmSo8C60;
begin
    if dsrSo126.DataSet.State in [dsEdit, dsInsert] then
    begin
      MessageDlg('目前資料並非處於瀏覽模式,無法執行此功能!', mtWarning, [mbOk],0);
      Exit;
    end
    else if dsrSo126.DataSet.State in [dsBrowse] then
    begin
      L_Frm := TfrmSo8C60.Create(Application);
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

{ ---------------------------------------------------------------------------- }

procedure TfrmSO8C10.fraChineseYMD4mseYMDExit(Sender: TObject);
begin
  fraChineseYMD4.mseYMDExit(Sender);
  if trim(TUstr.replaceStr(fraChineseYMD4.getYMD,'/',''))<>'' then
    dsrSo126.DataSet.FieldByName('RETURNDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD4.getYMD);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSO8C10.fraChineseYMD5mseYMDExit(Sender: TObject);
begin
  fraChineseYMD5.mseYMDExit(Sender);
  if trim(TUstr.replaceStr(fraChineseYMD5.getYMD,'/',''))<>'' then
    dsrSo126.DataSet.FieldByName('CLEARDATE').AsDateTime := TUdateTime.CDate2Date(fraChineseYMD5.getYMD);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSO8C10.ChangeEditorState;
begin
  if ( dsrSo126.DataSet.State = dsEdit ) then
  begin
    fraChineseYMD3.mseYMD.Enabled := False;
    btnGetPaperUser.Enabled := False;
    DBEdit5.ReadOnly := True;
    DBEdit5.Color := clBtnFace;
    DBEdit2.ReadOnly := True;
    DBEdit2.Color := clBtnFace;
    DBEdit3.ReadOnly := True;
    DBEdit4.ReadOnly := True;
    DBEdit4.Color := clBtnFace;
  end else
  begin
    fraChineseYMD3.mseYMD.Enabled := True;
    btnGetPaperUser.Enabled := True;
    DBEdit5.ReadOnly := False;
    DBEdit5.Color := clWindow;
    DBEdit2.ReadOnly := False;
    DBEdit2.Color := clWindow;
    DBEdit3.ReadOnly := False;
    DBEdit4.ReadOnly := False;
    DBEdit4.Color := clWindow;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSO8C10.btnPrintClick(Sender: TObject);
var
  aBeginDate, aEndDate: String;
  aPrefix, aPaperNum: String;
begin
  aBeginDate := trim(TUstr.replaceStr(fraChineseYMD1.getYMD,'/',''));
  aEndDate := trim(TUstr.replaceStr(fraChineseYMD2.getYMD,'/',''));
  {}
  aBeginDate := TUstr.ToROCYMD( aBeginDate );
  aEndDate := TUstr.ToROCYMD( aEndDate );
  {}
  aPrefix := EmptyStr;
  aPaperNum := EmptyStr;
  if ( edtPrefix.Text <> EmptyStr ) then aPrefix := edtPrefix.Text;
  if ( edtPaperNum.Text <> EmptyStr ) then aPaperNum := edtPaperNum.Text; 
  {}
  dtmMain3.rtpSO8C10.Data := dtmMain3.cdsSo126.Data;
  dtmMain3.rtpSO8C10.First;
  {}
  rptSO8C10 := TrptSO8C10.Create( nil );
  try
    rptSO8C10.aCompName := frmMainMenu.sG_CompCode;
    rptSO8C10.aPaperDate := Format( '%s~%s', [aBeginDate,aEndDate] );
    rptSO8C10.aOperator := sG_EmpsName;
    rptSO8C10.aPaperNo := aPrefix + '/' + aPaperNum;
    if ( rptSO8C10.aPaperNo = '/' ) then
      rptSO8C10.aPaperNo := EmptyStr;
    rptSO8C10.Preview;  
  finally
    rptSO8C10.Free;
    dtmMain3.rtpSO8C10.Data := Null;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSO8C10.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  dsrSo126.DataSet.Cancel;
end;

end.

