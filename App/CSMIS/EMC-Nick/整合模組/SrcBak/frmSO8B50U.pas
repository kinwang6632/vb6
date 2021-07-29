unit frmSO8B50U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, DBGrids, ExtCtrls, ComCtrls, Mask,
  DBCtrls, DB, ADODB;

const
    CITEM_SEP = '~';
    STR_SEP = '''';
type
  TfrmSO8B50 = class(TForm)
    StatusBar1: TStatusBar;
    pnlQuery: TPanel;
    Panel2: TPanel;
    pnlMultiData: TPanel;
    DBGrid1: TDBGrid;
    rgpQryCond: TRadioGroup;
    edtCustID: TEdit;
    edtCItem: TEdit;
    edtCm003: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    pnlSingleData: TPanel;
    StaticText1: TStaticText;
    dbtCustID: TDBEdit;
    StaticText2: TStaticText;
    dbtCm0032: TDBEdit;
    memCItem: TMemo;
    StaticText3: TStaticText;
    pnlCmd: TPanel;
    BitBtn9: TBitBtn;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    dsrSO132: TDataSource;
    BitBtn4: TBitBtn;
    dsrCm003: TDataSource;
    dsrCD019: TDataSource;
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure dsrSO132DataChange(Sender: TObject; Field: TField);
    procedure BitBtn3Click(Sender: TObject);
    procedure kBeforePost(DataSet: TDataSet);
  private
    { Private declarations }
    procedure ChangeState;
  public
    { Public declarations }
    sG_CurEmpName, sG_CurEmpNo : String;
    sG_CompCode, sG_User : String;
  end;

var
  frmSO8B50: TfrmSO8B50;

implementation

uses frmDbMultiSelectu, Ustru, frmDbSelectu, dtmMain1U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8B50.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B50.btnAppendClick(Sender: TObject);
begin
    if dtmMain1.qrySO132.State = dsInactive then
      dtmMain1.qrySO132.Active := True;
    dtmMain1.qrySO132.Append;
    ChangeState;
    dbtCustID.ReadOnly := false;
    dbtCustID.SetFocus;

end;

procedure TfrmSO8B50.btnEditClick(Sender: TObject);
begin
    if dtmMain1.qrySO132.State = dsInactive then Exit;
    dtmMain1.qrySO132.Edit;
    ChangeState;
    dbtCustID.ReadOnly := true;
    BitBtn4.SetFocus;

end;

procedure TfrmSO8B50.btnDeleteClick(Sender: TObject);
begin
    if dtmMain1.qrySO132.State = dsInactive then Exit;
     if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
       Exit;
     dtmMain1.qrySO132.Delete;
     ChangeState;

end;

procedure TfrmSO8B50.btnCancelClick(Sender: TObject);
begin
    if dtmMain1.qrySO132.State = dsInactive then Exit;
    dtmMain1.qrySO132.Cancel;
    ChangeState;

end;

procedure TfrmSO8B50.btnSaveClick(Sender: TObject);
var
    L_CItem, L_StrList : TStringList;
    ii : Integer;
    sL_DataStr : String;
    nL_CustID : Integer;
begin


    if not (dtmMain1.qrySO132.State in [dsEdit, dsInsert]) then Exit;

    if dtmMain1.qrySO132.FieldByName('CUSTID').AsString ='' then
    begin
      MessageDlg('請輸入客編!', mtWarning, [mbOK],0);
      dbtCustID.SetFocus;
      Exit;
    end;

    if dtmMain1.qrySO132.State in [dsEdit] then
    begin
      if (dtmMain1.qrySO132.FieldByName('EMPNO').AsString='') or
         (dtmMain1.qrySO132.FieldByName('EMPNAME').AsString='') then
      begin
        MessageDlg('請點選佣金歸屬者!', mtWarning, [mbOK],0);
        BitBtn4.SetFocus;
        Exit;
      end;
    end;

    if memCItem.Lines.Count=0 then
    begin
      MessageDlg('請點選收費項目!', mtWarning, [mbOK],0);
      BitBtn9.SetFocus;
      Exit;
    end;

    L_Citem := TSTringList.Create;
    for ii:=0 to memCItem.Lines.Count -1 do
    begin
      L_Citem.Add(memCItem.Lines.Strings[ii]);
    end;
    
    nL_CustID := dtmMain1.qrySO132.FieldByName('CUSTID').AsInteger;

    if (dtmMain1.qrySO132.State in [dsInsert]) then
    begin
      dtmMain1.qrySO132.Cancel;

      for ii:=0 to L_Citem.Count-1 do
      begin

        sL_DataStr := L_Citem.Strings[ii];
        L_StrList := TUstr.ParseStrings(sL_DataStr,CITEM_SEP);

        dtmMain1.qrySO132.Append;
        dtmMain1.qrySO132.FieldByName('COMPCODE').asInteger := StrToInt(sG_CompCode);
        dtmMain1.qrySO132.FieldByName('CUSTID').asInteger := nL_CustID;

        dtmMain1.qrySO132.FieldByName('CITEMCODE').asInteger := StrToInt(L_StrList.Strings[0]);
        dtmMain1.qrySO132.FieldByName('CITEMNAME').AsString := L_StrList.Strings[1];

        dtmMain1.qrySO132.FieldByName('EMPNO').AsString := sG_CurEmpNo;
        dtmMain1.qrySO132.FieldByName('EMPNAME').AsString := sG_CurEmpName;

            
        dtmMain1.qrySO132.FieldByName('DATACREATOR').AsString := frmMainMenu.sG_User;
        dtmMain1.qrySO132.FieldByName('UPDEN').AsString := frmMainMenu.sG_User;
        dtmMain1.qrySO132.FieldByName('UPDTIME').AsDateTime := now;

        dtmMain1.qrySO132.Post;
      end;
      ChangeState;
    end
    else if (dtmMain1.qrySO132.State in [dsEdit]) then
    begin

      sL_DataStr := L_Citem.Strings[0];
      L_StrList := TUstr.ParseStrings(sL_DataStr,CITEM_SEP);


      dtmMain1.qrySO132.FieldByName('CITEMCODE').asInteger := StrToInt(L_StrList.Strings[0]);
      dtmMain1.qrySO132.FieldByName('CITEMNAME').AsString := L_StrList.Strings[1];

      dtmMain1.qrySO132.FieldByName('UPDEN').AsString := frmMainMenu.sG_User;
      dtmMain1.qrySO132.FieldByName('UPDTIME').AsDateTime := now;

      dtmMain1.qrySO132.Post;

      ChangeState;

    end;
    if Assigned(L_Citem) then
      L_Citem.Free;    
end;

procedure TfrmSO8B50.ChangeState;
begin
     with dtmMain1.qrySO132 do
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
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
       end
       else
       begin
          if dtmMain1.qrySO132.RecordCount>0 then
          begin
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
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;

end;

procedure TfrmSO8B50.BitBtn9Click(Sender: TObject);
var
    ii : Integer;
    sL_Code, sL_Desc : String;
    sL_TargetSQL, sL_Tmp, sL_SQL : String;
    L_TempStrList, L_TargetFieldNamesStrList, L_TargetFieldValueStrList : TStringList;
begin
    with dtmMain1.qryCodeCD019_2 do
    begin
      sL_SQL := 'select CODENO,DESCRIPTION from CD019 where StopFlag=0 and RefNo=2';
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      FieldByName('DESCRIPTION').DisplayLabel := '收費項目代碼';
      FieldByName('DESCRIPTION').DisplayLabel := '收費項目';
    end;
    if dtmMain1.qryCodeCD019_2.RecordCount = 0 then
    begin
      dtmMain1.qryCodeCD019_2.Close;
      MessageDlg('沒有收費項目可供選取!!', mtWarning, [mbOK],0);
      Exit;
    end
    else
    begin
      L_TargetFieldNamesStrList := TStringList.Create;
      L_TargetFieldValueStrList := TStringList.Create;

      L_TargetFieldNamesStrList.Add('CODENO');
      L_TargetFieldNamesStrList.Add('DESCRIPTION');

      if SelectMultiRecords('請點選公司別', dtmMain1.qryCodeCD019_2, 'CODENO;DESCRIPTION', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
      begin
        memCItem.Clear;

        if dtmMain1.qrySO132.State = dsInsert then
        begin
          for ii:=0 to L_TargetFieldValueStrList.Count -1 do
          begin
            sL_Tmp := Trim(L_TargetFieldValueStrList.Strings[ii]);
            L_TempStrList := TUstr.ParseStrings(sL_Tmp,',');

            sL_Code := L_TempStrList.Strings[0];
            sL_Desc := L_TempStrList.Strings[1];
            memCItem.Lines.Add(sL_Code + CITEM_SEP +sL_Desc);


          end;
        end
        else if dtmMain1.qrySO132.State = dsEdit then
        begin
          if L_TargetFieldValueStrList.Count<>1 then
          begin
            MessageDlg('修改時,只能選取一筆', mtWarning, [mbOK],0);
            BitBtn9.SetFocus;
            Exit;
          end
          else
          begin
            sL_Tmp := Trim(L_TargetFieldValueStrList.Strings[0]);
            L_TempStrList := TUstr.ParseStrings(sL_Tmp,',');

            sL_Code := L_TempStrList.Strings[0];
            sL_Desc := L_TempStrList.Strings[1];
            memCItem.Lines.Add(sL_Code + CITEM_SEP +sL_Desc);

          end;

        end;
      end;
      L_TargetFieldNamesStrList.Free;
      L_TargetFieldValueStrList.Free;

    end;

end;

procedure TfrmSO8B50.BitBtn1Click(Sender: TObject);
var
    sL_SQL : String;
begin
    if dtmMain1.qryCodeCD019_2.State = dsInactive then
    begin
      sL_SQL := 'select CODENO,DESCRIPTION from CD019 where StopFlag=0 and RefNo=2';
      dtmMain1.qryCodeCD019_2.Close;
      dtmMain1.qryCodeCD019_2.SQL.Clear;
      dtmMain1.qryCodeCD019_2.SQL.Add(sL_SQL);
      dtmMain1.qryCodeCD019_2.Open;
    end;
      
    if SelectRecord('請點選收費項目', dtmMain1.qryCodeCD019_2, 'CODENO;DESCRIPTION','') = mrOk then
    begin
      edtCItem.Text := dtmMain1.qryCodeCD019_2.FieldByName('CodeNo').AsString + CITEM_SEP + dtmMain1.qryCodeCD019_2.FieldByName('DESCRIPTION').AsString;
    end;

end;

procedure TfrmSO8B50.FormShow(Sender: TObject);
begin
    sG_CompCode := frmMainMenu.sG_CompCode;
    edtCustID.SetFocus;
    ChangeState;
    self.Caption:=frmMainMenu.setCaption('SO8B50','[信用卡付費續收佣金歸屬人員設定]');


    //將畫面清空
    dtmMain1.qrySO132.Close;
    memCItem.Clear;
end;

procedure TfrmSO8B50.BitBtn2Click(Sender: TObject);
begin
    if dtmMain1.qryCodeCM003.State = dsInactive then
      dtmMain1.qryCodeCM003.Open;
    if SelectRecord('請點選佣金歸屬者', dtmMain1.qryCodeCM003, 'EMPNO;EMPNAME','') = mrOk then
    begin
      edtCm003.Text := dtmMain1.qryCodeCM003.FieldByName('EMPNO').AsString + CITEM_SEP + dtmMain1.qryCodeCM003.FieldByName('EMPNAME').AsString;
    end;

end;


procedure TfrmSO8B50.BitBtn4Click(Sender: TObject);
begin
    if dtmMain1.qryCodeCM003.State = dsInactive then
    begin
      dtmMain1.qryCodeCM003.Open;
    end;
      
    if SelectRecord('請點選佣金歸屬者', dtmMain1.qryCodeCM003, 'EMPNO;EMPNAME','') = mrOk then
    begin
      if dtmMain1.qrySO132.State = dsInsert then
      begin
        sG_CurEmpNo := dtmMain1.qryCodeCM003.FieldByName('EMPNO').AsString;
        sG_CurEmpName := dtmMain1.qryCodeCM003.FieldByName('EMPNAME').AsString;
        dtmMain1.qrySO132.FieldByName('EMPNAME').AsString := dtmMain1.qryCodeCM003.FieldByName('EMPNAME').AsString;
      end
      else if dtmMain1.qrySO132.State = dsEdit then
      begin
        dtmMain1.qrySO132.FieldByName('EMPNO').AsString := dtmMain1.qryCodeCM003.FieldByName('EMPNO').AsString;
        dtmMain1.qrySO132.FieldByName('EMPNAME').AsString := dtmMain1.qryCodeCM003.FieldByName('EMPNAME').AsString;
      end;
    end;

end;

procedure TfrmSO8B50.dsrSO132DataChange(Sender: TObject; Field: TField);
var
    sL_Data : String;
begin
    if (dsrSO132.State = dsBrowse) and (dsrSO132.DataSet.RecordCount>0) then
    begin

      sL_Data := dsrSO132.DataSet.fieldByName('CITEMCODE').AsString + CITEM_SEP + dsrSO132.DataSet.fieldByName('CITEMNAME').AsString;

      memCItem.Clear;
      memCItem.Lines.Add(sL_Data);
    end;
end;

procedure TfrmSO8B50.BitBtn3Click(Sender: TObject);
var
    L_StrList : TStringList;
    sL_SQL, sL_Where : String;
begin
    case rgpQryCond.ItemIndex of
      0://by 客編
       begin
         if edtCustID.Text='' then
         begin
           MessageDlg('請輸入客編!', mtWarning, [mbOK],0);
           edtCustID.SetFocus;
           Exit;
         end
         else
         begin
           sL_Where := 'CUSTID=' + edtCustID.Text;
         end;
       end;
      1://by 收費項目
       begin
         if edtCItem.Text='' then
         begin
           MessageDlg('請點選收費項目!', mtWarning, [mbOK],0);
           BitBtn1.SetFocus;
           Exit;
         end
         else
         begin
           L_StrList := TUstr.ParseStrings(edtCItem.Text,CITEM_SEP);
           sL_Where := 'CITEMCODE=' + L_StrList.Strings[0];
         end;
       end;
      2://by 佣金歸屬者
       begin
         if edtCm003.Text='' then
         begin
           MessageDlg('請點選佣金歸屬者!', mtWarning, [mbOK],0);
           BitBtn2.SetFocus;
           Exit;
         end
         else
         begin
           L_StrList := TUstr.ParseStrings(edtCm003.Text,CITEM_SEP);

           sL_Where := 'EMPNO=' + STR_SEP + L_StrList.Strings[0] + STR_SEP ;
         end;
       end;   
    end;

    sL_SQL := 'select * from SO132 where compcode=' + sG_CompCode + ' and ' + sL_Where;

    memCItem.Clear;
    with dtmMain1.qrySO132 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;
    ChangeState;    
end;

procedure TfrmSO8B50.kBeforePost(DataSet: TDataSet);
var
    nL_CompCode, nL_SeqNo : Integer;
    sL_SQL : String;
    bL_BeModified : boolean;
    nL_OldCItemCode, nL_NewCItemCode : Integer;
    sL_OldCItemName, sL_NewCItemName : String;
    sL_OldEmpNo, sL_OldEmpName, sL_NewEmpNo, sL_NewEmpName : String;
begin
    bL_BeModified := false;

    if DataSet.State = dsEdit then
    begin
//      nL_OldCItemCode := 'NULL';
      sL_OldCItemName := 'NULL';
//      sL_NewCItemCode := 'NULL';
      sL_NewCItemName := 'NULL';
      sL_OldEmpNo := 'NULL';
      sL_OldEmpName := 'NULL';
      sL_NewEmpNo := 'NULL';
      sL_NewEmpName := 'NULL';

      //down, 檢測 收費項目 是否被異動..
      if VarToStr(DataSet.FieldByName('CITEMCODE').OldValue) <> VarToStr(DataSet.FieldByName('CITEMCODE').NewValue) then
      begin
        bL_BeModified := true;
        nL_OldCItemCode := StrToInt(VarToStr(DataSet.FieldByName('CITEMCODE').OldValue));
        sL_OldCItemName := VarToStr(DataSet.FieldByName('CITEMNAME').OldValue);
        nL_NewCItemCode := StrToInt(VarToStr(DataSet.FieldByName('CITEMCODE').NewValue));
        sL_NewCItemName := VarToStr(DataSet.FieldByName('CITEMNAME').NewValue);
        {
        Memo1.Lines.Add('sL_OldCItemCode=' + sL_OldCItemCode);
        Memo1.Lines.Add('sL_OldCItemName=' + sL_OldCItemName);
        Memo1.Lines.Add('sL_NewCItemCode=' + sL_NewCItemCode);
        Memo1.Lines.Add('sL_NewCItemName=' + sL_NewCItemName);
        }
      end;
      //up, 檢測 收費項目 是否被異動..

      //down, 檢測 收費項目 是否被異動..
      if VarToStr(DataSet.FieldByName('EMPNO').OldValue) <> VarToStr(DataSet.FieldByName('EMPNO').NewValue) then
      begin
        bL_BeModified := true;
        sL_OldEmpNo := VarToStr(DataSet.FieldByName('EMPNO').OldValue);
        sL_OldEmpName := VarToStr(DataSet.FieldByName('EMPNAME').OldValue);
        sL_NewEmpNo := VarToStr(DataSet.FieldByName('EMPNO').NewValue);
        sL_NewEmpName := VarToStr(DataSet.FieldByName('EMPNAME').NewValue);
        {
        Memo1.Lines.Add('sL_OldEmpNo=' + sL_OldEmpNo);
        Memo1.Lines.Add('sL_OldEmpName=' + sL_OldEmpName);
        Memo1.Lines.Add('sL_NewEmpNo=' + sL_NewEmpNo);
        Memo1.Lines.Add('sL_NewEmpName=' + sL_NewEmpName);
        }
      end;
      //up, 檢測 收費項目 是否被異動..

      if bL_BeModified then
      begin
        with dtmMain1.qryCommon do
        begin
          SQL.Clear;
          SQL.Add('select nvl(max(seq),1) seq from so133');
          Open;
          nL_SeqNo := FieldByName('seq').AsInteger + 1;
          Close;
        end;

        nL_CompCode := DataSet.FieldByName('COMPCODE').AsInteger;
        with dtmMain1.qrySo133 do
        begin
          if state=dsInactive then
            Open;
          append;
          FieldByName('SEQ').AsInteger := nL_SeqNo;
          FieldByName('COMPCODE').AsInteger := nL_CompCode;
          FieldByName('OLDCITEMCODE').AsInteger := nL_OldCItemCode;
          FieldByName('OLDCITEMNAME').Value := sL_OldCItemName;
          FieldByName('NEWCITEMCODE').AsInteger := nL_NewCItemCode;
          FieldByName('NEWCITEMNAME').Value := sL_NewCItemName;

          FieldByName('OLDEMPNO').Value := sL_OldEmpNo;
          FieldByName('OLDEMPNAME').Value := sL_OldEmpName;
          FieldByName('NEWEMPNO').Value := sL_NewEmpNo;
          FieldByName('NEWEMPNAME').Value := sL_NewEmpName;
          FieldByName('OPERATIONMODE').AsString := 'U';
          FieldByName('UPDEN').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsDateTime := now;
          post;                              
        end;

{
        with ADOCommand1 do
        begin
          sL_SQL := 'insert into SO133(SEQ,COMPCODE,OLDCITEMCODE,OLDCITEMNAME,NEWCITEMCODE,NEWCITEMNAME,OLDEMPNO,OLDEMPNAME,NEWEMPNO,NEWEMPNAME,OPERATIONMODE,UPDEN,UPDTIME)';
          sL_SQL := sL_SQL + ' values((select nvl(max(seq),1) from so133)+1,' + IntToStr(nL_CompCode) + ',' +  sL_OldCItemCode + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_OldCItemName + STR_SEP + ',' + sL_NewCItemCode + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_NewCItemName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_OldEmpNo + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_OldEmpName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_NewEmpNo + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_NewEmpName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + 'U' + STR_SEP + ',' + STR_SEP + sG_User + STR_SEP + ',';
          sL_SQL := sL_SQL + TUstr.getOracleSQLDateTimeStr(now) + ')';
          Memo1.Lines.Add(sL_SQL);
          ADOCommand1.CommandText := sL_SQL;
          Memo1.Lines.Add(ADOCommand1.CommandText);
          ADOCommand1.Execute;
          
        end;

        Showmessage('modified..');
}        
      end;
    end;
end;

end.
